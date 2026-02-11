import json, gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from collections import Counter
import re


with open("google-sheet.json", "r", encoding="utf-8") as f:
    sheet_info = json.load(f)

sheet_url = sheet_info["sheet-url"]


SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

creds = Credentials.from_service_account_file("google-api-credentials.json", scopes=SCOPES)
client = gspread.authorize(creds)

spreadsheet = client.open_by_url(sheet_url)

worksheet = spreadsheet.sheet1
data = pd.DataFrame(worksheet.get_all_records())

column_name = 'Боль клиента'
texts = data[column_name].dropna().astype(str)

counter = Counter()

for text in texts:
    words = re.split(r'[,\n]', text)
    words = [w.strip().lower() for w in words if w.strip()] 
    counter.update(words)

for word, count in counter.most_common():
    print(f"{word} {count}")
