import requests
import smtplib
import os
from datetime import datetime, timedelta, timezone
from email.mime.text import MIMEText

API_KEY = os.environ["CONGRESS_API_KEY"]
KEYWORDS = ["artificial intelligence", "semiconductor", "semiconductors"]
SENDER_EMAIL = os.environ["SENDER_EMAIL"]
SENDER_PASSWORD = os.environ["SENDER_PASSWORD"]
RECIPIENT_EMAIL = os.environ["RECIPIENT_EMAIL"]

def fetch_recent_bills():
    yesterday = datetime.now(timezone.utc) - timedelta(days=1)
    from_date = yesterday.strftime("%Y-%m-%dT%H:%M:%SZ")
    url = "https://api.congress.gov/v3/bill"
    params = {
        "api_key": API_KEY,
        "fromDateTime": from_date,
        "sort": "updateDate+desc",
        "limit": 20,
    }
    response = requests.get(url, params=params)
    data = response.json()
    return data.get("bills", [])

def matches_keywords(bill):
    title = bill.get("title", "").lower()
    return [kw for kw in KEYWORDS if kw.lower() in title]

def send_email(subject, body):
    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = SENDER_EMAIL
    msg["To"] = RECIPIENT_EMAIL
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, RECIPIENT_EMAIL, msg.as_string())
    print("Email sent!")

def main():
    print("Fetching bills updated in the last 24 hours...\n")
    bills = fetch_recent_bills()

    matches = []
    for bill in bills:
        hits = matches_keywords(bill)
        if hits:
            matches.append((bill, hits))

    if not matches:
        body = "No matching bills found today."
        print(body)
    else:
        lines = [f"Found {len(matches)} matching bill(s):\n"]
        for bill, hits in matches:
            lines.append(f"Title:    {bill.get('title')}")
            lines.append(f"Keywords: {', '.join(hits)}")
            lines.append(f"Updated:  {bill.get('updateDate')}")
            lines.append(f"URL:      https://congress.gov/bill/{bill.get('congress')}th-congress/{bill.get('type', '').lower()}-bill/{bill.get('number')}")
            lines.append("")
        body = "\n".join(lines)
        print(body)

    send_email("Congress Digest", body)

main()