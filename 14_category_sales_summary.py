import pandas as pd
import os

base_dir = os.path.dirname(__file__)
data_dir = os.path.join(base_dir, 'data')

path_product = os.path.join(data_dir, 'product.xlsx')
path_product_category = os.path.join(data_dir, 'product_category.xlsx')
path_product_subcategory = os.path.join(data_dir, 'product_subcategory.xlsx')
path_sales_details = os.path.join(data_dir, 'sales_details.xlsx')
path_special_offer = os.path.join(data_dir, 'special_offer.xlsx')

product = pd.read_excel(path_product)
product_category = pd.read_excel(path_product_category)
product_subcategory = pd.read_excel(path_product_subcategory)
sales_details = pd.read_excel(path_sales_details)
special_offer = pd.read_excel(path_special_offer)

# Merge sales details with discounts
merged = sales_details.merge(special_offer[['SpecialOfferID', 'DiscountPct']], on='SpecialOfferID', how='left')
merged['DiscountPct'] = merged['DiscountPct'].fillna(0)

# Discount amount per one unit of the product
merged['unit_discount'] = merged['UnitPrice'] * merged['DiscountPct']
# Total discount applied to the entire order quantity
merged['total_discount'] = merged['unit_discount'] * merged['OrderQty']
# Total price without any discounts (before applying any offers)
merged['original_price'] = merged['UnitPrice'] * merged['OrderQty']
# Total price after the discount has been applied
merged['discounted_price'] = merged['original_price'] - merged['total_discount']

# Merge each product to its category name
merged = merged.merge(product[['ProductID', 'ProductSubcategoryID']], on="ProductID", how="left") \
               .merge(product_subcategory[['ProductSubcategoryID', 'ProductCategoryID']], on="ProductSubcategoryID", how="left") \
               .merge(product_category[['ProductCategoryID', 'Name']], on="ProductCategoryID", how="left")

# Group by category and calculate totals
category_summary = merged.groupby("Name").agg(
    TotalSalesAmount=pd.NamedAgg(column="original_price", aggfunc="sum"),
    TotalDiscountAmount=pd.NamedAgg(column="total_discount", aggfunc="sum"),
    TotalOrders=pd.NamedAgg(column="SalesOrderID", aggfunc="nunique")
).reset_index()

# Format columns for display
category_summary["TotalSalesAmount"] = category_summary["TotalSalesAmount"].apply(lambda x: f"${x:,.2f}")
category_summary["TotalDiscountAmount"] = category_summary["TotalDiscountAmount"].apply(lambda x: f"${x:,.2f}")
category_summary["TotalOrders"] = category_summary["TotalOrders"].apply(lambda x: f"{x:,}")

print(category_summary)

# Find top categories
raw_summary = merged.groupby("Name").agg(
    raw_sales=("original_price", "sum"),
    raw_orders=("SalesOrderID", "nunique")
).reset_index()

best_sales = raw_summary.loc[raw_summary["raw_sales"].idxmax()]
best_orders = raw_summary.loc[raw_summary["raw_orders"].idxmax()]

print(f"\n Highest sales category: {best_sales['Name']} — ${best_sales['raw_sales']:,.2f}")
print(f" Most orders category: {best_orders['Name']} — {best_orders['raw_orders']:,} orders")

#           Name TotalSalesAmount TotalDiscountAmount TotalOrders
# 0  Accessories    $1,278,760.91           $6,726.64      19,524
# 1        Bikes   $95,145,813.35         $543,005.58      18,368
# 2     Clothing    $2,141,507.02          $21,091.20       9,877
# 3   Components   $11,807,808.02           $5,214.74       2,650
#
#  Highest sales category: Bikes — $95,145,813.35
#  Most orders category: Accessories — 19,524 orders
