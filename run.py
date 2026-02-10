import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from collections import Counter
import re

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

creds = Credentials.from_service_account_file("google-api-credentials.json", scopes=SCOPES)
client = gspread.authorize(creds)

spreadsheet = client.open_by_url(
    "https://docs.google.com/spreadsheets/d/1MP_F7sicbKw1uqijPU1NFNL9LYGV_g1souTzGI7v2BA/edit"
)

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
