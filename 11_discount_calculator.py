import pandas as pd
import os

base_dir = os.path.dirname(__file__)
data_dir = os.path.join(base_dir, 'data')

path_sales_details = os.path.join(data_dir, 'sales_details.xlsx')
path_special_offer = os.path.join(data_dir, 'special_offer.xlsx')

sales_details = pd.read_excel(path_sales_details)
special_offer = pd.read_excel(path_special_offer)

merged = sales_details.merge(special_offer[['SpecialOfferID', 'DiscountPct']], on='SpecialOfferID', how='left')

# Calculate the discount for one unit
merged['unit_discount'] = merged['UnitPrice'] * merged['DiscountPct']
# Calculate the total discount for the order
merged['total_discount'] = merged['unit_discount'] * merged['OrderQty']
# Calculate the original total price before discount
merged['original_price'] = merged['UnitPrice'] * merged['OrderQty']
# Calculate the final discounted price
merged['discounted_price'] = merged['original_price'] - merged['total_discount']

# Show only rows where DiscountPct > 0
discounted = merged[merged['DiscountPct'] > 0]
pd.set_option('display.max_columns', None)

print(discounted[['ProductID', 'UnitPrice', 'OrderQty', 'DiscountPct',
           'unit_discount', 'total_discount', 'original_price', 'discounted_price']]
      .round(2)  # Round to 2 decimals for readability
      .to_string(index=False))

# Example
# ProductID  UnitPrice  OrderQty  DiscountPct  unit_discount  total_discount  original_price  discounted_price
#       773    1971.99        12         0.02          39.44          473.28        23663.93          23190.65
#       775    1957.49        13         0.02          39.15          508.95        25447.42          24938.48
#       712       5.01        13         0.02           0.10            1.30           65.18             63.87
#       709       5.22        21         0.05           0.26            5.49          109.72            104.24
