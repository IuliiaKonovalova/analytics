import json, gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from collections import Counter
import re
from tabulate import tabulate
from normalize_input import normalize_phrase


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

# for text in texts:
#     words = re.split(r'[,\n]', text)
#     words = [w.strip().lower() for w in words if w.strip()] 
#     counter.update(words)
for text in texts:
    words = re.split(r'[,\n]', text)
    print("words:", words)
    words = [w.strip() for w in words if w.strip()]
    print("stripped words:", words)

    for phrase in words:
        normalized = normalize_phrase(phrase)
        counter.update([normalized])

# for word, count in counter.most_common():
#     print(f"{word} {count}")


# Convert counter to sorted list
most_common_words = counter.most_common()

# Prepare data for tabulate
table_data = [
    {"Word": word, "Count": count}
    for word, count in most_common_words
]

print("\n WORD FREQUENCY ANALYSIS\n")


for word, count in most_common_words:
    print(f"{word} {count}")
print(tabulate(
    table_data,
    headers="keys",
    tablefmt="fancy_grid",
    showindex=True
))


data["normalized_pain"] = data["Боль клиента"].astype(str).apply(normalize_phrase)
print("\nNormalized Pain Analysis:")
print(data[["Боль клиента", "normalized_pain"]].head(10))


data["platform"] = data["platform"].astype(str).str.strip().str.lower()
print("\nPlatform Analysis:")
platform_counts = data["platform"].value_counts()
print(tabulate(
    platform_counts.reset_index().rename(columns={"platform": "Platform", "count": "Count"}),
    headers="keys",
    tablefmt="fancy_grid",
    showindex=False
))



# Top 5 pain categories
top_categories = data["normalized_pain"].value_counts().head(10).index

# Crosstab with pain as rows and platform as columns
cross_table = pd.crosstab(
    data["normalized_pain"],
    data["platform"]
)

# Filter only top categories
filtered_cross = cross_table.loc[top_categories]

print("\n📊 CLIENT PAIN vs PLATFORM\n")

print(tabulate(
    filtered_cross,
    headers="keys",
    tablefmt="fancy_grid"
))

data["date"] = pd.to_datetime(
    data["date"],
    format="%d.%m.%Y",
    errors="coerce"
)

print(data["date"].isna().sum())

data["week_of_month"] = (
    (data["date"].dt.day - 1) // 7 + 1
)

cross_week = pd.crosstab(
    data["week_of_month"],
    data["normalized_pain"]
)

top_categories = data["normalized_pain"].value_counts().head(5).index

filtered_week = cross_week[top_categories]

print("\n📊 WEEK OF MONTH vs CLIENT PAIN\n")

print(tabulate(
    filtered_week,
    headers="keys",
    tablefmt="fancy_grid"
))
