# Excel data aggregation

This project performs analysis on retail sales data using Python and pandas. It consists of six scripts, each handling a different analytical task including product sales ranking, discount impact, customer segmentation, and category-level performance.

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Scripts Summary](#scripts-summary)


## Overview

The scripts analyze Excel files containing information on products, sales transactions, customers, and special offers. The goal is to extract actionable business insights regarding sales volume, discount effectiveness, customer value, and product category performance.

## Features

- Analyze best and worst selling products
- Identify the most sold product on discount 
- Calculate revenue before and after discounts
- Find top customers by order frequency
- Segment customers by purchase frequency and spending
- Summarize total sales, discounts, and order volume by product category

## Installation

### Requirements

- Python 3.x
- pandas

Use pip to install the required libraries:
```pip install pandas```


### File Setup

Place the following Excel files in the project directory or update their paths in the code:

- product.xlsx
- product_category.xlsx
- product_subcategory.xlsx
- sales.xlsx
- sales_details.xlsx
- special_offer.xlsx

## Usage

Each script is self-contained and can be run independently. 

Example:

python 13_customer_matrix.py

## Scripts Summary

### 9. Most and least sold products 

- Merges product and sales data.
- Groups by product name and sums total quantity sold.
- Identifies the most and least sold products.

### 10. Most sold product on discount

- Merges sales data with special offers.
- Filters orders that occurred during valid discount periods.
- Aggregates total quantity sold under discount conditions.

### 11. Discount calculator

- Calculates discount amount per unit and total discount per order.
- Computes original and discounted prices.
- Displays all transactions where a discount was applied.

### 12. Top Customers by Orders

- Groups sales data by customer ID.
- Identifies customers with the highest number of orders.

### 13. Customer matrix

- Classifies customers by order count: New, Repeated, Fan.
- Classifies customers by spending: Frugal Spender, Medium Spender, High Spender.
- Outputs a matrix for frequency and monetary categories.

### 14. Category sales summary

- Merges sales data with product category.
- Aggregates total sales, total discounts, and order count by category.
- Identifies the category with highest sales and the one with most orders.


