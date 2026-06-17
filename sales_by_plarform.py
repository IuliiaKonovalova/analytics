import json, gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from collections import Counter
import re
from tabulate import tabulate
from normalize_input import normalize_phrase
from gspread_formatting import (
    set_frozen, set_column_width, format_cell_range, CellFormat, Color, TextFormat,
    Borders, Border, ConditionalFormatRule, BooleanRule, GridRange, BooleanCondition,
    get_conditional_format_rules
)


with open("google-sheet.json", "r", encoding="utf-8") as f:
    sheet_info = json.load(f)

sheet_url = sheet_info["services-sheet-url"]


SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

creds = Credentials.from_service_account_file("google-api-credentials.json", scopes=SCOPES)
client = gspread.authorize(creds)

spreadsheet = client.open_by_url(sheet_url)
print("Connected to Google Sheet successfully.")

worksheet = spreadsheet.sheet1
data = pd.DataFrame(worksheet.get_all_records())
tracker_sheet = spreadsheet.worksheet("Tracker") 




# --- Assume spreadsheet and final_table already defined ---
sales_by_platform_tab = spreadsheet.worksheet("Sales by Platform")

# -----------------------------
# LOAD TRACKER
# -----------------------------
tracker_df = pd.DataFrame(tracker_sheet.get_all_records())

# Clean columns
tracker_df.columns = tracker_df.columns.str.strip()

print("Tracker rows:", len(tracker_df))

# -----------------------------
# CLEAN DATA
# -----------------------------
tracker_df["FROM PLATFORM"] = (
    tracker_df["FROM PLATFORM"]
    .astype(str)
    .str.strip()
    .str.lower()
)

tracker_df["SERVICE"] = (
    tracker_df["SERVICE"]
    .astype(str)
    .str.strip()
)

tracker_df["MONTH"] = (
    tracker_df["MONTH"]
    .astype(str)
    .str.strip()
)

# ❗ Remove empty rows (VERY IMPORTANT)
tracker_df = tracker_df[
    (tracker_df["SERVICE"] != "") &
    (tracker_df["FROM PLATFORM"] != "") &
    (tracker_df["MONTH"] != "")
]

print("After cleaning:", len(tracker_df))

# -----------------------------
# TOP 10 SERVICES (GLOBAL)
# -----------------------------
top_services = (
    tracker_df["SERVICE"]
    .value_counts()
    .head(10)
    .index
)

# -----------------------------
# GROUP DATA
# -----------------------------
sales_counts = (
    tracker_df
    .groupby(["MONTH", "FROM PLATFORM", "SERVICE"])
    .size()
    .reset_index(name="Sales Count")
)


print("sales_counts rows:", len(sales_counts))
# Keep only top services
sales_counts = sales_counts[
    sales_counts["SERVICE"].isin(top_services)
]
print("After filtering top services, rows:", len(sales_counts))
print("Grouped rows:", len(sales_counts))

# -----------------------------
# PIVOT TABLE
# -----------------------------
pivot_table = sales_counts.pivot_table(
    index=["MONTH", "FROM PLATFORM"],
    columns="SERVICE",
    values="Sales Count",
    fill_value=0
)

# Add TOTAL column (important for visibility)
pivot_table["TOTAL"] = pivot_table.sum(axis=1)
print("Pivot table shape:", pivot_table.shape)

# Sort by best performance
pivot_table = pivot_table.sort_values("TOTAL", ascending=False)

final_table = pivot_table.reset_index()

print("Final table shape:", final_table.shape)
print(final_table.head())
# print final table in a nice format
print("\n📊 SALES BY PLATFORM\n")
print(tabulate(
    final_table,
    headers="keys",
    tablefmt="fancy_grid"
))


# -----------------------------
# WRITE TO GOOGLE SHEET
# -----------------------------
data_to_write = [final_table.columns.tolist()] + final_table.astype(str).values.tolist()

print("Writing to sheet:", sales_by_platform_tab.title)
print([ws.title for ws in spreadsheet.worksheets()])
try:
    sales_by_platform_tab = spreadsheet.worksheet("Sales by Platform")
    print("Found existing 'Sales by Platform' tab, clearing it.")
    sales_by_platform_tab.clear()
    # sales_by_platform_tab.update('A1', data_to_write)
    sales_by_platform_tab.update('A1', data_to_write)

except:
    sales_by_platform_tab = spreadsheet.add_worksheet(
        title="Sales by Platform",
        rows="100",
        cols="20"
    )
print("✅ Data written to 'Sales by Platform'")
