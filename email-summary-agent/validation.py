import httpx
from azure.identity import ClientSecretCredential
from dotenv import load_dotenv
import os
import json

load_dotenv()

TENANT_ID     = os.getenv("AZURE_TENANT_ID")
CLIENT_ID     = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
USER_EMAIL    = os.getenv("TARGET_USER_EMAIL")

def get_token():
    credential = ClientSecretCredential(
        tenant_id=TENANT_ID,
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET
    )
    token = credential.get_token("https://graph.microsoft.com/.default")
    return token.token

def get_emails(token: str, count: int = 5):
    headers = {"Authorization": f"Bearer {token}"}
    params = {
        "$top": count,
        "$select": "subject,from,receivedDateTime,bodyPreview",
        "$orderby": "receivedDateTime desc"
    }
    url = f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/messages"
    
    with httpx.Client() as client:
        response = client.get(url, headers=headers, params=params)
        response.raise_for_status()
        return response.json()

def main():
    print("🔐 Getting token...")
    token = get_token()
    print(f"✅ Token acquired: {token[:40]}...")

    print(f"\n📬 Fetching emails for {USER_EMAIL}...")
    result = get_emails(token)
    
    emails = result.get("value", [])
    print(f"✅ Got {len(emails)} emails\n")
    
    for i, email in enumerate(emails, 1):
        print(f"--- Email {i} ---")
        print(f"  Subject : {email.get('subject', '(no subject)')}")
        print(f"  From    : {email['from']['emailAddress']['address']}")
        print(f"  Date    : {email.get('receivedDateTime', '')}")
        print(f"  Preview : {email.get('bodyPreview', '')[:100]}...")
        print()

if __name__ == "__main__":
    main()