import json
import gspread

from google.oauth2.service_account import Credentials
from gspread.exceptions import WorksheetNotFound

# -----------------------------
# LOAD CONFIG
# -----------------------------
with open("google-sheet.json", "r", encoding="utf-8") as f:
    sheet_info = json.load(f)

clients_db_url = sheet_info["sheet-url"]
tracker_url = sheet_info["services-sheet-url"]

# -----------------------------
# AUTHENTICATION
# -----------------------------
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets"
]

creds = Credentials.from_service_account_file(
    "google-api-credentials.json",
    scopes=SCOPES
)

client = gspread.authorize(creds)

# -----------------------------
# OPEN SPREADSHEETS
# -----------------------------
clients_spreadsheet = client.open_by_url(clients_db_url)
tracker_spreadsheet = client.open_by_url(tracker_url)

print("Connected successfully.")

# -----------------------------
# COPY FUNCTION
# -----------------------------
def copy_sheet(
    source_spreadsheet,
    target_spreadsheet,
    source_tab_name,
    target_tab_name=None
):
    if target_tab_name is None:
        target_tab_name = source_tab_name

    print(f"\nCopying '{source_tab_name}'...")

    # Load source worksheet
    source_sheet = source_spreadsheet.worksheet(source_tab_name)
    data = source_sheet.get_all_values()

    print(f"Loaded {len(data)} rows")

    # Remove existing tab if present
    try:
        old_tab = target_spreadsheet.worksheet(target_tab_name)
        target_spreadsheet.del_worksheet(old_tab)
        print(f"Deleted existing '{target_tab_name}'")
    except WorksheetNotFound:
        pass

    # Determine worksheet size
    rows = max(1000, len(data) + 100)

    cols = max(
        26,
        max(len(row) for row in data) if data else 26
    )

    # Create worksheet
    new_tab = target_spreadsheet.add_worksheet(
        title=target_tab_name,
        rows=rows,
        cols=cols
    )

    # Copy values
    if data:
        new_tab.update(data)

    print(f"Created '{target_tab_name}' successfully")


# -----------------------------
# COPY CLIENTS DB (FIRST SHEET)
# -----------------------------
copy_sheet(
    source_spreadsheet=clients_spreadsheet,
    target_spreadsheet=tracker_spreadsheet,
    source_tab_name=clients_spreadsheet.sheet1.title,
    target_tab_name="Clients DB"
)

# -----------------------------
# COPY ЛИЧНЫЙ КОД
# -----------------------------
copy_sheet(
    source_spreadsheet=clients_spreadsheet,
    target_spreadsheet=tracker_spreadsheet,
    source_tab_name="Личный код",
    target_tab_name="Личный код"
)

print("\nAll sheets copied successfully.")