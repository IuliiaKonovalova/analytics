import os
from datetime import datetime
from telethon import TelegramClient
from telethon.tl.functions.messages import GetHistoryRequest
from dotenv import load_dotenv
import pandas as pd

load_dotenv()

api_id = int(os.getenv("API_ID"))
api_hash = os.getenv("API_HASH")
channel_username = os.getenv("CHANNEL_USERNAME")

client = TelegramClient("session", api_id, api_hash)


async def fetch_messages():
    await client.start()

    channel = await client.get_entity(channel_username)

    all_messages = []
    offset_id = 0
    limit = 100

    while True:
        history = await client(GetHistoryRequest(
            peer=channel,
            offset_id=offset_id,
            offset_date=None,
            add_offset=0,
            limit=limit,
            max_id=0,
            min_id=0,
            hash=0
        ))

        messages = history.messages
        if not messages:
            break

        for msg in messages:
            if not msg.message:
                continue

            reactions = 0
            if msg.reactions:
                reactions = sum(r.count for r in msg.reactions.results)

            all_messages.append({
                "id": msg.id,
                "date": msg.date,
                "text": msg.message,
                "views": msg.views or 0,
                "forwards": msg.forwards or 0,
                "reactions": reactions
            })

        offset_id = messages[-1].id

    return all_messages


async def main():
    data = await fetch_messages()

    df = pd.DataFrame(data)

    # Convert date
    df["date"] = pd.to_datetime(df["date"])

    # Filter months (Jan, Feb, Mar)
    df = df[
        (df["date"] >= "2026-01-01") &
        (df["date"] <= "2026-03-31")
    ]

    # Add useful columns
    df["month"] = df["date"].dt.to_period("M")
    df["week"] = df["date"].dt.isocalendar().week

    # Save raw data
    df.to_csv("data/telegram_posts_raw.csv", index=False)

    print("✅ Data saved!")

with client:
    client.loop.run_until_complete(main())