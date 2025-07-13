import pandas as pd
import os

base_dir = os.path.dirname(__file__)
data_dir = os.path.join(base_dir, 'data')

path_product = os.path.join(data_dir, 'product.xlsx')
path_sales = os.path.join(data_dir, 'sales.xlsx')
path_sales_details = os.path.join(data_dir, 'sales_details.xlsx')
path_special_offer = os.path.join(data_dir, 'special_offer.xlsx')

product = pd.read_excel(path_product)
sales = pd.read_excel(path_sales)
sales_details = pd.read_excel(path_sales_details)
special_offer = pd.read_excel(path_special_offer)


# Merge sales_details with OrderDate from sales to check if discount is valid
merged = sales_details.merge(sales[['SalesOrderID', 'OrderDate']], on='SalesOrderID', how='left')
# Merge with product to get product names
merged = merged.merge(product[['ProductID', 'Name']], on='ProductID', how='left')

# Drop metadata columns from special_offer to prevent column name conflicts during merge
special_offer_drop = special_offer.drop(columns=['rowguid', 'ModifiedDate'], errors='ignore')
merged = merged.merge(special_offer_drop[['SpecialOfferID', 'DiscountPct', 'StartDate', 'EndDate']],
                       on='SpecialOfferID', how='left')

# Filter for valid discounts
on_discount = merged[
    (merged['OrderDate'] >= merged['StartDate']) &
    (merged['OrderDate'] <= merged['EndDate']) &
    (merged['DiscountPct'] > 0)
    ]

# Group by product name and sum quantity
quantity_on_discount = on_discount.groupby('Name')['OrderQty'].sum().reset_index()
quantity_on_discount = quantity_on_discount.sort_values(by='OrderQty', ascending=False)

print(quantity_on_discount)

# Finds most sold product on discount
if not quantity_on_discount.empty:
    most = quantity_on_discount.iloc[0]
    print(f" Most sold product on discount: {most['Name']} with {most['OrderQty']} units.")
else:
    print("No products were sold with a valid discount.")

# Most sold product on discount: Classic Vest, S with 2218 units.
