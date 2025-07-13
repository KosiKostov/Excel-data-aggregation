import pandas as pd
import os

base_dir = os.path.dirname(__file__)
data_dir = os.path.join(base_dir, 'data')

path_product = os.path.join(data_dir, 'product.xlsx')
path_sales_details = os.path.join(data_dir, 'sales_details.xlsx')

product = pd.read_excel(path_product)
sales_details = pd.read_excel(path_sales_details)

# Joins sales details with product data on ProductID
merged = pd.merge(sales_details, product, on='ProductID')

# Group by product name and calculate total sold quantity
quantity_per_product = merged.groupby('Name')['OrderQty'].sum().reset_index()

# Sort by quantity descending
quantity_per_product = quantity_per_product.sort_values(by='OrderQty', ascending=False)

print(quantity_per_product)

# Get highest and lowest sold products
most_sold = quantity_per_product.iloc[0]
least_sold = quantity_per_product.iloc[-1]

print(f"Most sold product: {most_sold['Name']} with {most_sold['OrderQty']} units sold.")
print(f"Least sold product: {least_sold['Name']} with {least_sold['OrderQty']} units sold.")

# Most sold product: AWC Logo Cap with 8311 units sold.
# Least sold product: LL Touring Frame - Blue, 58 with 4 units sold.
