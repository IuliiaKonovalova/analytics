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
tracker_sheet = spreadsheet.worksheet("Tracker")
tracker_df = pd.DataFrame(tracker_sheet.get_all_records())

# Clean columns
tracker_df.columns = tracker_df.columns.str.strip()

# Normalize fields
tracker_df["SERVICE"] = tracker_df["SERVICE"].astype(str).str.strip()
tracker_df["MONTH"] = tracker_df["MONTH"].astype(str).str.strip()

# Remove empty rows
tracker_df = tracker_df[
    (tracker_df["SERVICE"] != "") &
    (tracker_df["MONTH"] != "")
]

print("Rows after cleaning:", len(tracker_df))

# -----------------------------
# PARSE MONTH (ROBUST)
# -----------------------------
def parse_month(x):
    try:
        return pd.to_datetime(x, format="%B %Y")  # January 2026
    except:
        try:
            return pd.to_datetime(x, format="%b %Y")  # Jan 2026
        except:
            return pd.NaT

tracker_df["MONTH"] = tracker_df["MONTH"].apply(parse_month)

# Remove invalid dates
tracker_df = tracker_df[tracker_df["MONTH"].notna()]

# Convert to clean format
tracker_df["MONTH_STR"] = tracker_df["MONTH"].dt.strftime("%b %Y")

# -----------------------------
# PIVOT TABLE (SERVICE x MONTH)
# -----------------------------
sales_by_month = tracker_df.pivot_table(
    index="SERVICE",
    columns="MONTH_STR",
    aggfunc="size",
    fill_value=0
)

# Add TOTAL column
sales_by_month["Total"] = sales_by_month.sum(axis=1)

# Sort rows by total
sales_by_month = sales_by_month.sort_values("Total", ascending=False)

# Sort columns chronologically
month_cols = sales_by_month.columns[:-1]  # exclude Total

sorted_months = sorted(
    month_cols,
    key=lambda x: pd.to_datetime(x, format="%b %Y")
)

sales_by_month = sales_by_month[sorted_months + ["Total"]]

# Reset index
final_month_table = sales_by_month.reset_index()


# -----------------------------
# WRITE TO GOOGLE SHEETS
# -----------------------------
data_to_write = [
    final_month_table.columns.tolist()
] + final_month_table.values.tolist()

try:
    sheet = spreadsheet.worksheet("Sales by Month")
    sheet.clear()
except:
    sheet = spreadsheet.add_worksheet(
        title="Sales by Month",
        rows="100",
        cols="20"
    )

sheet.update("A1", data_to_write)


# -----------------------------
# STYLING
# -----------------------------

# Get worksheet ID (required!)
sheet_id = sheet._properties['sheetId']

# Convert A1 → GridRange
grid_range = GridRange.from_a1_range("B2:Z100", sheet)

rules = get_conditional_format_rules(sheet)
rules.clear()


# 🔴 Zero = red
rules.append(ConditionalFormatRule(
    ranges=[grid_range],
    booleanRule=BooleanRule(
        condition=BooleanCondition('NUMBER_EQ', ['0']),
        format=CellFormat(backgroundColor=Color(1, 0.8, 0.8))
    )
))

# 🔵 Greater than 0 = blue
rules.append(ConditionalFormatRule(
    ranges=[grid_range],
    booleanRule=BooleanRule(
        condition=BooleanCondition('NUMBER_GREATER', ['0']),
        format=CellFormat(backgroundColor=Color(0.8, 0.9, 1))
    )
))

# ✅ FINALLY: apply
rules.save()
# 1️⃣ Freeze header row
set_frozen(sheet, rows=1)

# 2️⃣ Set column width
set_column_width(sheet, "A:A", 250)   # SERVICE column wider
set_column_width(sheet, "B:Z", 120)   # months

# 3️⃣ Bold header + center
header_format = CellFormat(
    textFormat=TextFormat(bold=True),
    horizontalAlignment="CENTER"
)

format_cell_range(sheet, "A1:Z1", header_format)

# 4️⃣ Left align service names
format_cell_range(sheet, "A2:A100", CellFormat(
    horizontalAlignment="LEFT"
))

# 5️⃣ Center numbers
format_cell_range(sheet, "B2:Z100", CellFormat(
    horizontalAlignment="CENTER"
))

# 6️⃣ Add borders
border_style = Border("SOLID")

format_cell_range(
    sheet,
    "A1:Z100",
    CellFormat(
        borders=Borders(
            top=border_style,
            bottom=border_style,
            left=border_style,
            right=border_style
        )
    )
)
print("✅ Data written to 'Sales by Month'")