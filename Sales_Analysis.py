# -*- coding: utf-8 -*-
#SUPERSTORE SALES ANALYSIS PROJECT  
# ---------------------------------------------------------

# STEP 1 — Import Required Libraries
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

save_path = r"C:\Users\User\Downloads\Sales Analysis using Python for Data Science\plots"
plt.style.use("seaborn")
# STEP 2 — Load Superstore "Orders" Sheet
df = pd.read_excel("Superstore.xls", sheet_name="Orders")

df.head()

# STEP 3 — Basic Dataset Information
print(df.shape)
df.info()
df.isnull().sum()

# STEP 4 — Data Cleaning

# Remove duplicates
df = df.drop_duplicates()

# Ensure numeric columns are correct
df["Sales"] = pd.to_numeric(df["Sales"], errors="coerce")
df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
df["Discount"] = pd.to_numeric(df["Discount"], errors="coerce")
df["Profit"] = pd.to_numeric(df["Profit"], errors="coerce")

df.head()

# STEP 5 — Feature Engineering

# Convert Order Date to datetime
df["Order Date"] = pd.to_datetime(df["Order Date"], errors="coerce")

# Extract Date Components
df["Year"] = df["Order Date"].dt.year
df["Month"] = df["Order Date"].dt.month
df["Quarter"] = df["Order Date"].dt.quarter

# Profit Margin %
df["Profit Margin %"] = (df["Profit"] / df["Sales"]) * 100

df.head()

# STEP 6 — Exploratory Data Analysis (EDA)

# Total Sales & Profit
print("Total Sales:", df["Sales"].sum())
print("Total Profit:", df["Profit"].sum())

# Region-wise summary
region_summary = df.groupby("Region")[["Sales", "Profit"]].sum()
region_summary

# STEP 7 — Category and Subcategory Analysis

# Category Performance
category_perf = df.groupby("Category")[["Sales", "Profit"]].sum()
category_perf

# Sub-category performance (sorted)
subcategory_perf = df.groupby("Sub-Category")["Sales"].sum().sort_values(ascending=False)
subcategory_perf.head(10)

# Top 10 Products by Sales
top_products = df.groupby("Product Name")["Sales"].sum().sort_values(ascending=False).head(10)
top_products

# STEP 8 — Visualizations

# 1. Sales by Category
plt.figure(figsize=(8,5))
category_perf["Sales"].plot(kind="bar", color="skyblue")
plt.title("Sales by Category")
plt.ylabel("Sales")

plt.savefig(save_path + "sales_by_category.png", dpi=300, bbox_inches='tight')
plt.show()

# 2. Profit by Region
plt.figure(figsize=(8,5))
region_summary["Profit"].plot(kind="bar", color="orange")
plt.title("Profit by Region")
plt.ylabel("Profit")
plt.savefig(save_path + "profit_by_region.png", dpi=300, bbox_inches='tight')
plt.show()

# 3. Monthly Sales Trend
monthly_sales = df.groupby("Month")["Sales"].sum()

plt.figure(figsize=(10,5))
plt.plot(monthly_sales, marker='o')
plt.title("Monthly Sales Trend")
plt.xlabel("Month")
plt.ylabel("Sales")
plt.grid(True)
plt.savefig(save_path + "monthly_sales_trend.png", dpi=300, bbox_inches='tight')
plt.show()

# 4. Discount vs Profit (Scatter Plot)
plt.figure(figsize=(8,5))
plt.scatter(df["Discount"], df["Profit"], alpha=0.5)
plt.title("Discount vs Profit")
plt.xlabel("Discount")
plt.ylabel("Profit")
plt.savefig(save_path + "discount_vs_profit.png", dpi=300, bbox_inches='tight')
plt.show()

# 5. Correlation Heatmap
plt.figure(figsize=(8,6))
sns.heatmap(df.corr(), annot=True, cmap="coolwarm")
plt.title("Correlation Heatmap")
plt.savefig(save_path + "correlation_heatmap.png", dpi=300, bbox_inches='tight')
plt.show()

# STEP 9 — Export Cleaned Data & Summary Report
df.to_excel("Superstore_Sales_Report.xlsx", index=False)

print("Excel report created successfully: Superstore_Sales_Report.xlsx")

# STEP 10 — Key Insights (Auto Summary)

print("--------- KEY INSIGHTS ---------")
print("Best Sales Region:", df.groupby("Region")["Sales"].sum().idxmax())
print("Best Profit Region:", df.groupby("Region")["Profit"].sum().idxmax())
print("Best Category:", df.groupby("Category")["Sales"].sum().idxmax())
print("Best Sub-Category:", df.groupby("Sub-Category")["Sales"].sum().idxmax())
print("Highest Sales Month:", df.groupby("Month")["Sales"].sum().idxmax())
print("--------------------------------")