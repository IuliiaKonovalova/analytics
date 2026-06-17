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
# print(spreadsheet.sheet1.get_all_records()[:5])  # Print first 5 rows for preview

worksheet = spreadsheet.sheet1
data = pd.DataFrame(worksheet.get_all_values())
tracker_sheet = spreadsheet.worksheet("Tracker") 
# get DATE SUM column
date_sum_col = tracker_sheet.col_values(tracker_sheet.find("DATE SUM").col)[1:]  # Skip header
print("date_sum_col length:", len(date_sum_col))  # Debugging line to check length of date_sum_col
# get column TOTAL
total_col = tracker_sheet.col_values(tracker_sheet.find("TOTAL").col)[1:] 
print("total_col length:", len(total_col))  # Debugging line to check length of total_col
# create a table with columns "Date Sum", "Total" and count of each date sum
date_sum_counts = Counter(date_sum_col)
table_data = [
    {"Date Sum": date_sum_col[i], "Total": total_col[i]}
    for i in range(len(date_sum_col))
]



df = pd.DataFrame({
    "Date": date_sum_col,
    "Total": total_col
})


df["Date"] = pd.to_datetime(df["Date"])
# convert totals to numbers
df["Total"] = pd.to_numeric(df["Total"], errors="coerce").fillna(0)
df["Weekday"] = df["Date"].dt.day_name()


weekdays = [
    "Monday", "Tuesday", "Wednesday",
    "Thursday", "Friday", "Saturday", "Sunday"
]

weekday_data = {}

for day in weekdays:
    temp = df[df["Weekday"] == day][["Date", "Total"]]
    weekday_data[day] = temp.reset_index(drop=True)

# print("debug weekday_data:", weekday_data)

max_len = max(len(weekday_data[day]) for day in weekdays)

# print("max_len:", max_len)  # Debugging line to check max_len value


final_table = []

# Header row
header = []
for day in weekdays:
    header.extend([f"All {day}s", f"{day}s Total", ""])

final_table.append(header)

# Data rows
for i in range(max_len):
    row = []
    for day in weekdays:
        if i < len(weekday_data[day]):
            row.append(str(weekday_data[day].iloc[i]["Date"].date()))
            row.append(int(weekday_data[day].iloc[i]["Total"]))
        else:
            row.extend(["", ""])
        row.append("")  # spacer column
    final_table.append(row)
# --- Add SUM row ---
sum_row = []

for day in weekdays:
    day_total = int(weekday_data[day]["Total"].sum())

    sum_row.append("SUM")
    sum_row.append(day_total)
    sum_row.append("")  # spacer

final_table.append(sum_row)

week_days_tab = spreadsheet.worksheet("Sales/week day")

week_days_tab.clear()
week_days_tab.update("A1", final_table)



# Format header row (row 1)
header_format = CellFormat(
    backgroundColor=Color(0.9, 0.9, 0.9),  # light gray
    textFormat=TextFormat(bold=True),
    horizontalAlignment='CENTER'
)

format_cell_range(week_days_tab, 'A1:Z1', header_format)


# --- Assume spreadsheet and final_table already defined ---
week_days_tab = spreadsheet.worksheet("Sales/week day")

# --- Clear the worksheet ---
week_days_tab.clear()

# --- Update sheet with final_table ---
week_days_tab.update("A1", final_table)

# --- Define helper for column letters ---
def column_letter(n):
    """Convert 1-based column number to Excel-style letter"""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result

num_rows = len(final_table)
num_cols = len(final_table[0])

# --- Freeze header row ---
set_frozen(week_days_tab, rows=1)

# --- Set column widths ---
for col in range(1, num_cols + 1):
    col_letter = column_letter(col)
    set_column_width(week_days_tab, f"{col_letter}:{col_letter}", 120)

# --- Header formatting ---
header_format = CellFormat(
    backgroundColor=Color(0.2, 0.4, 0.6),  # dark blue
    textFormat=TextFormat(bold=True, foregroundColor=Color(1,1,1)),
    horizontalAlignment='CENTER'
)
format_cell_range(week_days_tab, f"A1:{column_letter(num_cols)}1", header_format)

# --- Table borders ---
border = Border('SOLID')
table_format = CellFormat(
    borders=Borders(top=border, bottom=border, left=border, right=border)
)
format_cell_range(week_days_tab, f"A1:{column_letter(num_cols)}{num_rows}", table_format)

# --- Conditional formatting for totals > 50k ---
# First, get current rules
rules = get_conditional_format_rules(week_days_tab)

# Apply conditional formatting to every "Total" column
# Total columns are every 3rd column starting from B (1-based index 2)
for col in range(2, num_cols + 1, 3):
    col_letter = column_letter(col)
    rule = ConditionalFormatRule(
        ranges=[GridRange.from_a1_range(f"{col_letter}2:{col_letter}{num_rows}", week_days_tab)],
        booleanRule=BooleanRule(
            condition=BooleanCondition('NUMBER_GREATER', ['50000']),
            # dark green background for totals > 50k
            format=CellFormat(backgroundColor=Color(0.2, 0.4, 0.6))  # dark blue
        )
    )
    rules.append(rule)

rules.save()

