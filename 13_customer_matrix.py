import pandas as pd
import os

base_dir = os.path.dirname(__file__)
data_dir = os.path.join(base_dir, 'data')

path_sales = os.path.join(data_dir, 'sales.xlsx')
path_sales_details = os.path.join(data_dir, 'sales_details.xlsx')

sales = pd.read_excel(path_sales)
sales_details = pd.read_excel(path_sales_details)


# Count number of orders made by each customer
customer_order_counts = sales.groupby("CustomerID")["SalesOrderID"].count().reset_index()
customer_order_counts.columns = ["CustomerID", "OrderCount"]

# Merge sales with sales_details to get monetary values per order
sales_with_details = pd.merge(sales_details, sales[["SalesOrderID", "CustomerID"]], on="SalesOrderID")

# Calculate total spent per customer
customer_spending = sales_with_details.groupby("CustomerID")["LineTotal"].sum().reset_index()
customer_spending.columns = ["CustomerID", "TotalSpent"]

# Merge order count and total spent
customer_summary = pd.merge(customer_order_counts, customer_spending, on="CustomerID", how="outer").fillna(0)


# Classify customer based on number of orders
def classify_frequency(count):
    if count == 1:
        return "New"
    elif count <= 3:
        return "Repeated"
    else:
        return "Fan"


customer_summary["frequency"] = customer_summary["OrderCount"].apply(classify_frequency)


#  Classify customer based on amount spent
def classify_spending(total):
    if total < 100:
        return "Frugal Spender"
    elif total < 10000:
        return "Medium Spender"
    else:
        return "High Spender"


customer_summary["monetary_value"] = customer_summary["TotalSpent"].apply(classify_spending)

# matrix
row_order = ["New", "Repeated", "Fan"]
col_order = ["Frugal Spender", "Medium Spender", "High Spender"]

matrix = pd.crosstab(
    customer_summary["frequency"],
    customer_summary["monetary_value"]
).reindex(index=row_order, columns=col_order)

print(matrix)

# monetary_value  Frugal Spender  Medium Spender  High Spender
# frequency
# New                       6832            4817             0
# Repeated                   848            5825             4
# Fan                          3             317           473

# Business logic
# Best customers: "Fan" + "High Spender" — loyal and valuable.
# Worst customers: "New" + "Frugal Spender" — low engagement and low value.
