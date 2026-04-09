# mcp_server.py
import os
import re
import json
import httpx
import asyncio
import logging
from datetime import datetime, timezone, timedelta
from contextlib import asynccontextmanager
from azure.identity import ClientSecretCredential
from dotenv import load_dotenv
from fastmcp import FastMCP
from fastapi import FastAPI, Request, Response, BackgroundTasks
import uvicorn

load_dotenv()

# ── Config ─────────────────────────────────────────────────────────────────────

TENANT_ID     = os.getenv("AZURE_TENANT_ID")
CLIENT_ID     = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
USER_EMAIL    = os.getenv("TARGET_USER_EMAIL")
WEBHOOK_URL   = os.getenv("WEBHOOK_URL")
CLIENT_STATE  = os.getenv("WEBHOOK_CLIENT_STATE", "email-agent-secret-2026")
GRAPH_BASE    = "https://graph.microsoft.com/v1.0"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
log = logging.getLogger(__name__)

# ── Auth ───────────────────────────────────────────────────────────────────────

def get_graph_token() -> str:
    credential = ClientSecretCredential(
        tenant_id=TENANT_ID,
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET
    )
    return credential.get_token("https://graph.microsoft.com/.default").token

# ── Graph Subscription Manager ─────────────────────────────────────────────────

subscription_id: str | None = None

def register_webhook_subscription() -> str | None:
    """Register a Graph change notification subscription for new emails."""
    global subscription_id

    token = get_graph_token()
    expiry = (datetime.now(timezone.utc) + timedelta(minutes=4230)).strftime(
        "%Y-%m-%dT%H:%M:%S.000Z"
    )

    payload = {
        "changeType": "created",
        "notificationUrl": WEBHOOK_URL,
        "resource": f"users/{USER_EMAIL}/messages",
        "expirationDateTime": expiry,
        "clientState": CLIENT_STATE
    }

    log.info(f"📡 Registering Graph webhook subscription...")
    log.info(f"   notificationUrl : {WEBHOOK_URL}")
    log.info(f"   resource        : users/{USER_EMAIL}/messages")
    log.info(f"   expires         : {expiry}")

    with httpx.Client(timeout=30.0) as client:
            response = client.post(
                f"{GRAPH_BASE}/subscriptions",
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/json"
                },
                json=payload
            )

    if response.status_code == 201:
        data = response.json()
        subscription_id = data["id"]
        log.info(f"✅ Subscription registered: {subscription_id}")
        return subscription_id
    else:
        log.error(f"❌ Subscription failed: {response.status_code} {response.text}")
        return None


def renew_webhook_subscription() -> bool:
    """Renew existing subscription before it expires (call every ~2 days)."""
    global subscription_id
    if not subscription_id:
        log.warning("⚠️  No subscription to renew — registering fresh")
        register_webhook_subscription()
        return True

    token = get_graph_token()
    expiry = (datetime.now(timezone.utc) + timedelta(minutes=4230)).strftime(
        "%Y-%m-%dT%H:%M:%S.000Z"
    )

    with httpx.Client() as client:
        response = client.patch(
            f"{GRAPH_BASE}/subscriptions/{subscription_id}",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json"
            },
            json={"expirationDateTime": expiry}
        )

    if response.status_code == 200:
        log.info(f"✅ Subscription renewed until {expiry}")
        return True
    else:
        log.error(f"❌ Renewal failed: {response.status_code} {response.text}")
        register_webhook_subscription()
        return False


def delete_webhook_subscription():
    """Clean up subscription on shutdown."""
    global subscription_id
    if not subscription_id:
        return

    token = get_graph_token()
    with httpx.Client() as client:
        response = client.delete(
            f"{GRAPH_BASE}/subscriptions/{subscription_id}",
            headers={"Authorization": f"Bearer {token}"}
        )
    if response.status_code == 204:
        log.info(f"🗑️  Subscription deleted: {subscription_id}")
    else:
        log.warning(f"⚠️  Subscription delete returned: {response.status_code}")
    subscription_id = None


# ── Renewal Background Task ────────────────────────────────────────────────────

async def subscription_renewal_loop():
    """Renew subscription every 2 days (subscription max is ~3 days)."""
    while True:
        await asyncio.sleep(60 * 60 * 48)  # 48 hours
        log.info("🔄 Renewing webhook subscription...")
        renew_webhook_subscription()

# ── Email Helpers ──────────────────────────────────────────────────────────────

def fetch_email_by_id(email_id: str) -> dict:
    """Fetch full email content by ID from Graph."""
    token = get_graph_token()
    with httpx.Client() as client:
        response = client.get(
            f"{GRAPH_BASE}/users/{USER_EMAIL}/messages/{email_id}",
            headers={
                "Authorization": f"Bearer {token}",
                "Prefer": 'outlook.body-content-type="text"'
            },
            params={"$select": "id,subject,from,receivedDateTime,body,toRecipients,bodyPreview"}
        )
        response.raise_for_status()

    email = response.json()
    body_content = email.get("body", {}).get("content", "")
    body_clean = re.sub(r'<[^>]+>', ' ', body_content)
    body_clean = re.sub(r'\s+', ' ', body_clean).strip()

    return {
        "id": email["id"],
        "subject": email.get("subject", "(no subject)"),
        "from_name": email["from"]["emailAddress"].get("name", ""),
        "from_email": email["from"]["emailAddress"]["address"],
        "received": email.get("receivedDateTime", ""),
        "preview": email.get("bodyPreview", "")[:300],
        "body": body_clean[:3000]
    }


def summarize_email(email: dict) -> str:
    """Produce a formatted summary of a single email."""
    return (
        f"\n{'='*60}\n"
        f"🔵 NEW EMAIL RECEIVED\n"
        f"{'='*60}\n"
        f"  Subject : {email['subject']}\n"
        f"  From    : {email['from_name']} <{email['from_email']}>\n"
        f"  Date    : {email['received'][:19].replace('T', ' ')} UTC\n"
        f"  Preview : {email['preview']}\n"
        f"{'='*60}\n"
        f"  Summary :\n"
        f"  {email['body'][:500]}...\n"
        f"{'='*60}\n"
    )


async def handle_new_email(email_id: str):
    """Background task: fetch and summarize a newly arrived email."""
    log.info(f"📨 New email detected, fetching ID: {email_id[:40]}...")
    try:
        email = fetch_email_by_id(email_id)
        summary = summarize_email(email)
        print(summary)
        log.info(f"✅ Summarized: '{email['subject']}'")
    except Exception as e:
        log.error(f"❌ Failed to process email {email_id}: {e}")


async def fetch_latest_unread_email():
    """Fallback: fetch latest unread email when notification has no resource ID."""
    token = get_graph_token()
    with httpx.Client() as client:
        response = client.get(
            f"{GRAPH_BASE}/users/{USER_EMAIL}/messages",
            headers={"Authorization": f"Bearer {token}"},
            params={
                "$top": 1,
                "$filter": "isRead eq false",
                "$orderby": "receivedDateTime desc",
                "$select": "id"
            }
        )
    emails = response.json().get("value", [])
    if emails:
        await handle_new_email(emails[0]["id"])


# ── FastAPI lifespan ───────────────────────────────────────────────────────────

@asynccontextmanager
async def lifespan(app: FastAPI):
    log.info("🚀 Starting Email Summarizer Agent...")
    if not WEBHOOK_URL:
        log.warning("⚠️  WEBHOOK_URL not set — webhook disabled")
    else:
        log.info("ℹ️  Server ready. Call POST /admin/register to subscribe.")
    yield
    log.info("🛑 Shutting down...")
    delete_webhook_subscription()


# ── FastAPI app ────────────────────────────────────────────────────────────────

app = FastAPI(title="Email Summarizer Agent", lifespan=lifespan)


# ── Webhook: GET — Graph validation handshake ──────────────────────────────────
# Graph sends GET /webhook?validationToken=xxx to verify your endpoint.
# Must echo back the raw token as plain text with 200 OK within 10 seconds.

@app.get("/webhook")
async def webhook_get(request: Request):
    """Handle GET validation requests from Microsoft Graph."""
    params = dict(request.query_params)
    if "validationToken" in params:
        token = params["validationToken"]
        log.info(f"🤝 Graph GET validation handshake received")
        log.info(f"   Echoing token: {token[:40]}...")
        return Response(
            content=token,
            media_type="text/plain",
            status_code=200
        )
    # Plain health check
    return Response(content="webhook OK", media_type="text/plain", status_code=200)


# ── Webhook: POST — Graph change notifications ─────────────────────────────────

@app.post("/webhook")
async def webhook_post(request: Request, background_tasks: BackgroundTasks):
    """
    Handle POST change notifications from Microsoft Graph.
    Must return 202 Accepted quickly — Graph retries if response is slow.
    """
    # Some Graph versions send validation via POST too
    params = dict(request.query_params)
    if "validationToken" in params:
        token = params["validationToken"]
        log.info(f"🤝 Graph POST validation handshake received")
        return Response(
            content=token,
            media_type="text/plain",
            status_code=200
        )

    try:
        body = await request.json()
    except Exception:
        log.warning("⚠️  Webhook received non-JSON body")
        return Response(status_code=400)

    notifications = body.get("value", [])
    for notification in notifications:
        received_state = notification.get("clientState", "")
        if received_state != CLIENT_STATE:
            log.warning(f"⚠️  Invalid clientState: {received_state}")
            continue

        resource_data = notification.get("resourceData", {})
        email_id = resource_data.get("id")

        if email_id:
            log.info(f"📬 New email notification received (ID present)")
            background_tasks.add_task(handle_new_email, email_id)
        else:
            log.info("📬 Notification received (no ID) — fetching latest unread")
            background_tasks.add_task(fetch_latest_unread_email)

    return Response(status_code=202)

@app.post("/admin/register")
async def admin_register(background_tasks: BackgroundTasks):
    """Manually trigger webhook subscription registration in background."""
    background_tasks.add_task(register_webhook_subscription)
    return {"status": "registering", "message": "Check server logs for result"}

# ── MCP Tools ──────────────────────────────────────────────────────────────────

mcp = FastMCP("Email Summarizer Agent")

@mcp.tool()
def get_recent_emails(count: int = 10) -> str:
    """Fetch the most recent emails. Returns subject, sender, date, preview."""
    count = min(count, 25)
    token = get_graph_token()
    with httpx.Client() as client:
        response = client.get(
            f"{GRAPH_BASE}/users/{USER_EMAIL}/messages",
            headers={"Authorization": f"Bearer {token}"},
            params={
                "$top": count,
                "$select": "id,subject,from,receivedDateTime,bodyPreview,isRead",
                "$orderby": "receivedDateTime desc"
            }
        )
        response.raise_for_status()

    emails = response.json().get("value", [])
    result = []
    for i, email in enumerate(emails, 1):
        result.append({
            "index": i,
            "id": email["id"],
            "subject": email.get("subject", "(no subject)"),
            "from_name": email["from"]["emailAddress"].get("name", ""),
            "from_email": email["from"]["emailAddress"]["address"],
            "received": email.get("receivedDateTime", ""),
            "preview": email.get("bodyPreview", "")[:300],
            "is_read": email.get("isRead", True)
        })
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
def get_email_body(email_id: str) -> str:
    """Fetch the full body of a specific email by its ID."""
    email = fetch_email_by_id(email_id)
    return json.dumps(email, ensure_ascii=False, indent=2)


@mcp.tool()
def summarize_emails(count: int = 5) -> str:
    """Fetch and summarize the most recent emails as a digest."""
    count = min(count, 10)
    token = get_graph_token()
    with httpx.Client() as client:
        response = client.get(
            f"{GRAPH_BASE}/users/{USER_EMAIL}/messages",
            headers={"Authorization": f"Bearer {token}"},
            params={
                "$top": count,
                "$select": "subject,from,receivedDateTime,bodyPreview,isRead",
                "$orderby": "receivedDateTime desc"
            }
        )
        response.raise_for_status()

    emails = response.json().get("value", [])
    lines = [f"📬 Email Digest — Last {len(emails)} emails\n"]
    for i, email in enumerate(emails, 1):
        sender_name = email["from"]["emailAddress"].get("name", "")
        sender_addr = email["from"]["emailAddress"]["address"]
        subject     = email.get("subject", "(no subject)")
        received    = email.get("receivedDateTime", "")[:10]
        preview     = email.get("bodyPreview", "")[:150]
        is_read     = "📖" if email.get("isRead") else "🔵"
        lines.append(
            f"{is_read} {i}. {subject}\n"
            f"   From: {sender_name} <{sender_addr}>\n"
            f"   Date: {received}\n"
            f"   {preview}...\n"
        )
    return "\n".join(lines)


# ── Entry point ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import threading

    log.info(f"📧 Target mailbox : {USER_EMAIL}")
    log.info(f"🔗 Webhook URL    : {WEBHOOK_URL or 'NOT SET'}")
    log.info(f"🔌 FastAPI/Webhook : port 8000")
    log.info(f"🔌 MCP Server      : port 8001")

    def run_mcp():
        mcp.run(transport="sse", host="0.0.0.0", port=8001)

    mcp_thread = threading.Thread(target=run_mcp, daemon=True)
    mcp_thread.start()

    uvicorn.run(app, host="0.0.0.0", port=8000)