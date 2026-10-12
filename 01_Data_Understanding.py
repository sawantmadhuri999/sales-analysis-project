import pandas as pd
df = pd.read_csv("Sales_Analysis_Project_Raw.csv")
print("Dataset Shape:")
print(df.shape)
print("\nColumn Names:")
print(df.columns.tolist())

print("\nData Types:")
print(df.dtypes)

print("\nMissing Values:")
print(df.isnull().sum())

print("\nDuplicate Rows:")
print(df.duplicated().sum())

from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt

# ------------------------------------------
# 1. SETUP
# ------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
INPUT_FILE = BASE_DIR / "Sales_Analysis_Project_Raw.csv"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

if not INPUT_FILE.exists():
    raise FileNotFoundError(
        f"Dataset not found: {INPUT_FILE}\n"
        "Place the CSV file beside this Python script."
    )

# ------------------------------------------
# 2. LOAD DATA
# ------------------------------------------
df = pd.read_csv(INPUT_FILE)

print("=" * 55)
print("SALES DATA ANALYSIS PROJECT")
print("=" * 55)

print("\n--- ORIGINAL DATA ---")
print("Rows and columns:", df.shape)
print("\nColumn names:")
print(df.columns.tolist())
print("\nData types:")
print(df.dtypes)
print("\nMissing values:")
print(df.isnull().sum())
print("\nExact duplicate rows:", df.duplicated().sum())

# Save original data quality report
missing_before = df.isnull().sum()
duplicates_before = int(df.duplicated().sum())
duplicate_order_ids_before = int(
    df["Order_ID"].duplicated().sum()
)

# ------------------------------------------
# 3. DATA CLEANING
# ------------------------------------------
df["Order_Date"] = pd.to_datetime(
    df["Order_Date"], errors="coerce"
)

# Remove extra spaces from text columns
text_columns = df.select_dtypes(
    include=["object", "string"]
).columns

for col in text_columns:
    df[col] = df[col].str.strip()

# Standardize text formatting
df["Customer_Name"] = df["Customer_Name"].str.title()
df["City"] = df["City"].str.title()
df["State"] = df["State"].str.title()
df["Product_Category"] = df["Product_Category"].str.title()
df["Product_Name"] = df["Product_Name"].str.title()
df["Payment_Mode"] = df["Payment_Mode"].str.title()
df["Sales_Channel"] = df["Sales_Channel"].str.title()

# Convert numeric columns
numeric_columns = ["Quantity", "Unit_Price", "Discount"]

for col in numeric_columns:
    df[col] = pd.to_numeric(df[col], errors="coerce")

# Remove completely duplicated records
df = df.drop_duplicates()

# Remove records without essential identifiers/date
df = df.dropna(subset=["Order_ID", "Order_Date"])

# For repeated order IDs, retain the first record
df = df.drop_duplicates(subset=["Order_ID"], keep="first")

# Fill missing numeric values
df["Quantity"] = df["Quantity"].fillna(
    df["Quantity"].median()
)
df["Discount"] = df["Discount"].fillna(0)

# Exclude invalid quantities and prices
df = df[
    (df["Quantity"] > 0) &
    (df["Unit_Price"] >= 0) &
    (df["Discount"] >= 0)
].copy()

# Convert Quantity to integer if all values are whole numbers
if (df["Quantity"] % 1 == 0).all():
    df["Quantity"] = df["Quantity"].astype(int)

# ------------------------------------------
# 4. CALCULATE SALES
# ------------------------------------------
df["Gross_Sales"] = df["Quantity"] * df["Unit_Price"]
df["Total_Sales"] = df["Gross_Sales"] - df["Discount"]

df["Year"] = df["Order_Date"].dt.year
df["Month"] = df["Order_Date"].dt.month
df["Month_Name"] = df["Order_Date"].dt.strftime("%b")
df["Year_Month"] = df["Order_Date"].dt.to_period("M").astype(str)

# Categorize orders using sales quartiles
df["Order_Value_Category"] = pd.qcut(
    df["Total_Sales"].rank(method="first"),
    q=3,
    labels=["Low", "Medium", "High"]
)

# ------------------------------------------
# 5. BUSINESS ANALYSIS
# ------------------------------------------
total_orders = df["Order_ID"].nunique()
total_sales = df["Total_Sales"].sum()
average_order_value = (
    total_sales / total_orders if total_orders else 0
)
total_units = df["Quantity"].sum()
unique_customers = df["Customer_Name"].nunique()

kpi_report = pd.DataFrame({
    "Metric": [
        "Total Orders",
        "Total Sales",
        "Average Order Value",
        "Total Units Sold",
        "Unique Customers"
    ],
    "Value": [
        total_orders,
        round(total_sales, 2),
        round(average_order_value, 2),
        total_units,
        unique_customers
    ]
})

sales_by_city = (
    df.groupby("City", as_index=False)
      .agg(
          Total_Sales=("Total_Sales", "sum"),
          Total_Orders=("Order_ID", "nunique"),
          Units_Sold=("Quantity", "sum")
      )
      .sort_values("Total_Sales", ascending=False)
)

sales_by_category = (
    df.groupby("Product_Category", as_index=False)
      .agg(
          Total_Sales=("Total_Sales", "sum"),
          Total_Orders=("Order_ID", "nunique"),
          Units_Sold=("Quantity", "sum")
      )
      .sort_values("Total_Sales", ascending=False)
)

sales_by_product = (
    df.groupby("Product_Name", as_index=False)
      .agg(
          Total_Sales=("Total_Sales", "sum"),
          Units_Sold=("Quantity", "sum")
      )
      .sort_values("Total_Sales", ascending=False)
)

sales_by_month = (
    df.groupby("Year_Month", as_index=False)
      .agg(
          Total_Sales=("Total_Sales", "sum"),
          Total_Orders=("Order_ID", "nunique")
      )
      .sort_values("Year_Month")
)

sales_by_payment = (
    df.groupby("Payment_Mode", as_index=False)
      .agg(Total_Sales=("Total_Sales", "sum"))
      .sort_values("Total_Sales", ascending=False)
)

sales_by_channel = (
    df.groupby("Sales_Channel", as_index=False)
      .agg(
          Total_Sales=("Total_Sales", "sum"),
          Total_Orders=("Order_ID", "nunique")
      )
      .sort_values("Total_Sales", ascending=False)
)

top_customers = (
    df.groupby("Customer_Name", as_index=False)
      .agg(
          Total_Sales=("Total_Sales", "sum"),
          Total_Orders=("Order_ID", "nunique")
      )
      .sort_values("Total_Sales", ascending=False)
      .head(10)
)

# ------------------------------------------
# 6. EXPORT CLEANED DATA AND REPORTS
# ------------------------------------------
cleaned_csv = OUTPUT_DIR / "Cleaned_Sales_Data.csv"
df.to_csv(cleaned_csv, index=False)

report_xlsx = OUTPUT_DIR / "Sales_Analysis_Reports.xlsx"

with pd.ExcelWriter(report_xlsx, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Cleaned_Data", index=False)
    kpi_report.to_excel(writer, sheet_name="KPIs", index=False)
    sales_by_city.to_excel(writer, sheet_name="Sales_by_City", index=False)
    sales_by_category.to_excel(
        writer, sheet_name="Sales_by_Category", index=False
    )
    sales_by_product.to_excel(
        writer, sheet_name="Sales_by_Product", index=False
    )
    sales_by_month.to_excel(
        writer, sheet_name="Monthly_Sales", index=False
    )
    sales_by_payment.to_excel(
        writer, sheet_name="Payment_Analysis", index=False
    )
    sales_by_channel.to_excel(
        writer, sheet_name="Channel_Analysis", index=False
    )
    top_customers.to_excel(
        writer, sheet_name="Top_Customers", index=False
    )

# ------------------------------------------
# 7. CREATE CHARTS
# ------------------------------------------
def save_bar_chart(data, x_col, y_col, title, filename):
    if data.empty:
        return

    plt.figure(figsize=(10, 6))
    plt.bar(data[x_col].astype(str), data[y_col])
    plt.title(title)
    plt.xlabel(x_col.replace("_", " "))
    plt.ylabel(y_col.replace("_", " "))
    plt.xticks(rotation=35, ha="right")
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / filename, dpi=150)
    plt.close()

save_bar_chart(
    sales_by_city, "City", "Total_Sales",
    "Sales by City", "Sales_by_City.png"
)

save_bar_chart(
    sales_by_category, "Product_Category", "Total_Sales",
    "Sales by Product Category", "Sales_by_Category.png"
)

save_bar_chart(
    sales_by_month, "Year_Month", "Total_Sales",
    "Monthly Sales Trend", "Monthly_Sales_Trend.png"
)

save_bar_chart(
    sales_by_product.head(10), "Product_Name", "Total_Sales",
    "Top 10 Products by Sales", "Top_10_Products.png"
)

# ------------------------------------------
# 8. FINAL QUALITY REPORT
# ------------------------------------------
quality_report = pd.DataFrame({
    "Check": [
        "Original rows",
        "Original columns",
        "Missing values before cleaning",
        "Exact duplicate rows before cleaning",
        "Duplicate Order_IDs before cleaning",
        "Final rows",
        "Final columns",
        "Missing values after cleaning",
        "Duplicate Order_IDs after cleaning"
    ],
    "Result": [
        502,
        12,
        int(missing_before.sum()),
        duplicates_before,
        duplicate_order_ids_before,
        len(df),
        len(df.columns),
        int(df.isnull().sum().sum()),
        int(df["Order_ID"].duplicated().sum())
    ]
})

quality_report.to_csv(
    OUTPUT_DIR / "Data_Quality_Report.csv", index=False
)

# ------------------------------------------
# 9. PRINT SUMMARY
# ------------------------------------------
print("\n--- FINAL PROJECT SUMMARY ---")
print("Cleaned rows:", len(df))
print("Cleaned columns:", len(df.columns))
print("Total orders:", total_orders)
print("Total sales:", round(total_sales, 2))
print("Average order value:", round(average_order_value, 2))
print("Total units sold:", total_units)
print("Unique customers:", unique_customers)

print("\nTop 5 cities by sales:")
print(sales_by_city.head().to_string(index=False))

print("\nTop 5 product categories:")
print(sales_by_category.head().to_string(index=False))

print("\nTop 5 products:")
print(sales_by_product.head().to_string(index=False))

print("\nFiles saved in:", OUTPUT_DIR)
print("Project completed successfully!")

