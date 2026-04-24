```python
import pandas as pd
import matplotlib.pyplot as plt
import os

# Create output folder
if not os.path.exists("output"):
    os.makedirs("output")

# Load Excel file
df = pd.read_excel("data/sales_data.xlsx")

print("\n📌 Data Preview:")
print(df.head())

# -------------------------------
# 📊 DATA PROCESSING
# -------------------------------

# Convert Date column to datetime
df["Date"] = pd.to_datetime(df["Date"])

# Create Month column
df["Month"] = df["Date"].dt.to_period("M")

# Calculate Revenue
df["Revenue"] = df["Quantity"] * df["Price"]

# -------------------------------
# 📈 ANALYSIS
# -------------------------------

# Total Revenue
total_revenue = df["Revenue"].sum()
print("\n💰 Total Revenue:", total_revenue)

# Product-wise Revenue
product_sales = df.groupby("Product")["Revenue"].sum()

# Category-wise Revenue
category_sales = df.groupby("Category")["Revenue"].sum()

# Monthly Revenue
monthly_sales = df.groupby("Month")["Revenue"].sum()

print("\n📦 Product-wise Revenue:\n", product_sales)
print("\n📂 Category-wise Revenue:\n", category_sales)
print("\n📅 Monthly Revenue:\n", monthly_sales)

# -------------------------------
# 📊 VISUALIZATION
# -------------------------------

# Bar Chart - Product Sales
plt.figure()
product_sales.plot(kind="bar")
plt.title("Revenue by Product")
plt.xlabel("Product")
plt.ylabel("Revenue")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig("output/product_sales.png")

# Pie Chart - Category Sales
plt.figure()
category_sales.plot(kind="pie", autopct='%1.1f%%')
plt.title("Category-wise Revenue")
plt.ylabel("")
plt.savefig("output/category_sales.png")

# Line Chart - Monthly Sales
plt.figure()
monthly_sales.plot(kind="line", marker='o')
plt.title("Monthly Revenue Trend")
plt.xlabel("Month")
plt.ylabel("Revenue")
plt.grid()
plt.savefig("output/monthly_sales.png")

# -------------------------------
# 💾 SAVE SUMMARY
# -------------------------------

summary = pd.DataFrame({
    "Product": product_sales.index,
    "Revenue": product_sales.values
})

summary.to_excel("output/summary.xlsx", index=False)

print("\n✅ Charts saved in 'output/' folder")
print("📄 Summary saved as summary.xlsx")

# -------------------------------
# 🧠 INSIGHTS
# -------------------------------

top_product = product_sales.idxmax()
top_category = category_sales.idxmax()

print("\n🔍 Insights:")
print(f"📌 Top Product: {top_product}")
print(f"📌 Top Category: {top_category}")
```
