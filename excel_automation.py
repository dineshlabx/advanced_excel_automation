

import pandas as pd
from openpyxl import load_workbook

# Step 1 Read Excel file
df = pd.read_excel("sales_data.xlsx")

# Step 2 Clean column names
df.columns = df.columns.str.strip().str.lower()

# Step 3 Handling missing values
df["quantity"] = df["quantity"].fillna(0)
df["quantity"] = df["price"].fillna(0)

# Step 4 Create total sales column
df["total_sales"] = df["quantity"] * df["price"]

# Step 5 Save processed file
df.to_excel("processed_sales.xlsx", index=False)

print("Excel automation completed successfully")


