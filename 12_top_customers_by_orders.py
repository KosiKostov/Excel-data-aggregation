import pandas as pd
import os

base_dir = os.path.dirname(__file__)
data_dir = os.path.join(base_dir, 'data')

path_sales = os.path.join(data_dir, 'sales.xlsx')

sales = pd.read_excel(path_sales)

# Group by customer and count orders
customer_order_counts = sales.groupby("CustomerID")["SalesOrderID"].count()

# Find max number of orders
max_orders = customer_order_counts.max()

# Get all customers with that max number
top_customers = customer_order_counts[customer_order_counts == max_orders]

print("Customers ID with the most orders and the number of orders:")
for customer_id, order_count in top_customers.items():
    print(f"CustomerID: {customer_id} - Orders: {order_count}")

# Customers ID with the most orders and the number of orders:
# CustomerID: 11091 - Orders: 28
# CustomerID: 11176 - Orders: 28
