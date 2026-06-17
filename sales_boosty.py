import json, gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from tabulate import tabulate
from gspread_formatting import (
    get_conditional_format_rules,
     ConditionalFormatRule,
    BooleanRule,
    BooleanCondition,
    CellFormat,
    set_frozen,
    set_column_width,
    format_cell_range,
    CellFormat,
    TextFormat,
    Borders,
    Border,
    Color,
    GradientRule,
    GridRange
)

# -----------------------------
# CONNECT TO GOOGLE SHEETS
# -----------------------------
with open("google-sheet.json", "r", encoding="utf-8") as f:
    sheet_info = json.load(f)

sheet_url = sheet_info["services-sheet-url"]

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

creds = Credentials.from_service_account_file(
    "google-api-credentials.json",
    scopes=SCOPES
)

client = gspread.authorize(creds)
spreadsheet = client.open_by_url(sheet_url)

print("Connected to Google Sheet successfully.")

# -----------------------------
# LOAD TRACKER DATA
# -----------------------------
tracker_sheet = spreadsheet.worksheet("Boosty")
tracker_df = pd.DataFrame(tracker_sheet.get_all_records())

# Clean columns
tracker_df.columns = tracker_df.columns.str.strip()

# Normalize fields
tracker_df["PRODUCT"] = tracker_df["PRODUCT"].astype(str).str.strip()
tracker_df["DATE"] = tracker_df["DATE"].astype(str).str.strip()
# -----------------------------
# PREPARE DATA TYPES
# -----------------------------
tracker_df["$"] = pd.to_numeric(tracker_df["$"], errors="coerce")
tracker_df["DATE"] = pd.to_datetime(tracker_df["DATE"], errors="coerce")

# -----------------------------
# GROUP BY PRODUCT (STATISTICS)
# -----------------------------
product_stats = tracker_df.groupby("PRODUCT").agg(
    total_revenue=("$", "sum"),
    sales_count=("$", "count"),
    avg_check=("$", "mean"),
    first_sale=("DATE", "min"),
    last_sale=("DATE", "max")
).reset_index()

# Optional: sort by revenue
product_stats = product_stats.sort_values(by="total_revenue", ascending=False)

# -----------------------------
# OUTPUT
# -----------------------------
print("\nProduct Statistics:\n")
print(tabulate(product_stats, headers="keys", tablefmt="grid", showindex=False))