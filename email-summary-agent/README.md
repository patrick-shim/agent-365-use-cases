# 📧 Outlook Email Summarizer Agent

A real-time email monitoring and summarization agent built with:
- **Microsoft Graph API** — email access via webhook change notifications
- **FastMCP 3.x** — MCP tool server exposing email tools
- **FastAPI** — webhook receiver for Graph notifications
- **Azure OpenAI (GPT-5)** — natural language agent with function calling
- **VS Code Dev Tunnels** — secure public HTTPS endpoint for Graph webhooks

---

## Architecture

```
New email arrives in Outlook
        ↓
Microsoft Graph webhook notification
        ↓
VS Code Dev Tunnel (*.devtunnels.ms — trusted by Graph)
        ↓
FastAPI /webhook (port 8000)
        ↓
Fetches full email via Graph API
        ↓
Prints real-time summary to console

Console Chat Agent:
You: "summarize my emails"
        ↓
Azure OpenAI GPT-5 (function calling)
        ↓
Calls tool functions directly (get_recent_emails, etc.)
        ↓
Graph API fetches emails
        ↓
GPT-5 summarizes and responds
```

---

## Project Structure

```
mail-reader/
├── outlook_mcp_server.py   # FastAPI webhook + FastMCP tool server
├── email_agent.py          # Console chat agent (Azure OpenAI + function calling)
├── validation.py           # Standalone Graph auth test script
├── .env                    # Credentials (never commit — gitignored)
├── .env.example            # Template for .env
├── .gitignore
├── requirements.txt
└── README.md
```

---

## Prerequisites

| Requirement | Details |
|---|---|
| Python 3.12+ | |
| VS Code | With Microsoft account signed in (required for Dev Tunnels) |
| Azure CLI (`az`) | Logged in with appropriate account |
| Azure AI Services resource | With GPT-5 (or GPT-4o) deployment |
| Microsoft 365 tenant | Admin access — use M365 Developer Program sandbox |
| Entra App Registration | With Graph Mail.Read application permission |

---

## Part 1 — Entra App Registration

### 1.1 Create the App

1. Go to https://entra.microsoft.com
2. Sign in as tenant admin
3. **App registrations → New registration**
   - Name: `email-summarizer-agent`
   - Supported account types: Single tenant
   - Click **Register**
4. Note down:
   - **Application (client) ID**
   - **Directory (tenant) ID**

### 1.2 Add Graph API Permissions

1. Your app → **API Permissions → Add a permission**
2. **Microsoft Graph → Application permissions**
3. Search and add: `Mail.Read`
4. Click **Grant admin consent for [your tenant]**
5. Confirm green checkmarks appear

### 1.3 Create Client Secret

1. Your app → **Certificates & secrets → New client secret**
2. Description: `email-agent-dev`
3. Expiry: 180 days
4. **Copy the Value immediately** — shown only once
5. Store securely — put in `.env` as `AZURE_CLIENT_SECRET`

---

## Part 2 — Project Setup

### 2.1 Clone and Create Virtual Environment

```bash
git clone https://github.com/patrick-shim/email-agent-with-outlook-mcp-server.git
cd email-agent-with-outlook-mcp-server

python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS/Linux
source .venv/bin/activate
```

### 2.2 Install Dependencies

```bash
pip install -r requirements.txt
```

### 2.3 Create `.env` File

Copy `.env.example` to `.env` and fill in your values:

```bash
# Windows
Copy-Item .env.example .env

# macOS/Linux
cp .env.example .env
```

Edit `.env`:

```bash
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
TARGET_USER_EMAIL=admin@yourtenant.onmicrosoft.com
GRAPH_ENDPOINT=https://graph.microsoft.com/v1.0
WEBHOOK_CLIENT_STATE=email-agent-secret-2026

# Fill in after Dev Tunnel setup (Part 3)
WEBHOOK_URL=https://YOUR-TUNNEL-URL.devtunnels.ms/webhook

# Azure OpenAI
AZURE_OPENAI_ENDPOINT=https://your-ai-resource.cognitiveservices.azure.com/
AZURE_OPENAI_DEPLOYMENT=gpt-5-default
```

### 2.4 Validate Graph Auth

Test that your credentials work before proceeding:

```bash
python validation.py
```

Expected output:
```
🔐 Getting token...
✅ Token acquired
📬 Fetching emails...
✅ Got 5 emails
--- Email 1 ---
  Subject : ...
```

If this works, your Entra app and Graph permissions are correct.

---

## Part 3 — VS Code Dev Tunnel Setup

### Why Dev Tunnels?

Microsoft Graph webhooks POST change notifications to a public HTTPS URL.
Your local server runs on `localhost` which is not publicly accessible.
VS Code Dev Tunnels create a secure public URL that forwards to your localhost.
The `*.devtunnels.ms` domain is fully trusted by Microsoft Graph.

> ⚠️ Other tunnel tools (ngrok free tier, Cloudflare free tunnels) are
> blocked by Microsoft Graph or fail DNS resolution from Graph's servers.
> VS Code Dev Tunnels are the recommended solution for M365 development.

### 3.1 Sign into VS Code with Microsoft Account

```
VS Code → Accounts icon (bottom left) → Sign in with Microsoft Account
```

Must use a Microsoft account (not GitHub).

### 3.2 Forward Port 8000

```
VS Code → Bottom panel → PORTS tab
→ Click "+ Forward a Port" → Type: 8000 → Enter
→ Right-click 8000 → Port Visibility → Public
→ Right-click 8000 → Make Persistent  ← keeps same URL across restarts
```

### 3.3 Copy the Tunnel URL

In the PORTS panel, copy the **Forwarded Address**:
```
https://xxxxxxxx-8000.jpe1.devtunnels.ms
```

### 3.4 Update `.env`

```bash
WEBHOOK_URL=https://xxxxxxxx-8000.jpe1.devtunnels.ms/webhook
```

> ⚠️ If the tunnel URL changes (restart without Make Persistent),
> update `.env` and restart `outlook_mcp_server.py`.

---

## Part 4 — Running the Servers

You need **two terminals** running simultaneously.

### Terminal 1 — Start MCP + Webhook Server

```bash
.venv\Scripts\activate
python outlook_mcp_server.py
```

Expected output:
```
📧 Target mailbox : admin@yourtenant.onmicrosoft.com
🔗 Webhook URL    : https://xxxxxxxx-8000.jpe1.devtunnels.ms/webhook
🔌 FastAPI/Webhook : port 8000
🔌 MCP Server      : port 8001
🚀 Starting Email Summarizer Agent...
ℹ️  Server ready. Call POST /admin/register to subscribe.
Uvicorn running on http://0.0.0.0:8000
```

### Terminal 2 — Register Graph Webhook Subscription

After the server is fully running, register the subscription:

```bash
# Windows PowerShell
Invoke-RestMethod -Method POST -Uri http://localhost:8000/admin/register

# macOS/Linux
curl -X POST http://localhost:8000/admin/register
```

Expected response:
```json
{"status": "registering", "message": "Check server logs for result"}
```

Watch Terminal 1 for:
```
🤝 Graph POST validation handshake received
✅ Subscription registered: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
```

> ⚠️ Why manual registration?
> Graph validates your webhook during subscription registration.
> Auto-registering at startup causes a race condition — the server is
> not ready to respond to Graph's validation request in time (10 second limit).
> Manual registration after the server is fully up avoids this entirely.

---

## Part 5 — Console Chat Agent

In Terminal 2:

```bash
python email_agent.py
```

Example session:
```
==================================================
  📧 Email Assistant Agent
  Model    : gpt-5-default
==================================================

You: summarize my emails
  🔧 Calling: [summarize_emails]
Agent: Here are your 5 most recent emails:

1. 🔵 mail agent test — Patrick Shim — Apr 9
   Brief test message with signature block.

2. 📖 Microsoft Entra ID Protection Weekly Digest — MSSecurity — Apr 7
   Weekly security digest for your tenant.
...

You: get the full content of the first email
  🔧 Calling: [get_recent_emails]
  🔧 Calling: [get_email_body]
Agent: ...

You: quit
👋 Goodbye!
```

> Note: `email_agent.py` imports tool functions directly from
> `outlook_mcp_server.py`. The MCP server does NOT need to be running
> separately to use the chat agent — but it DOES need to be running
> for real-time webhook notifications.

---

## Part 6 — Real-time Email Monitoring

With `outlook_mcp_server.py` running and webhook registered,
send an email to `TARGET_USER_EMAIL`. Within 2-3 seconds:

```
📬 New email notification received
📨 New email detected, fetching ID: AAMkAD...

============================================================
🔵 NEW EMAIL RECEIVED
============================================================
  Subject : Hello from Patrick
  From    : sender@example.com
  Date    : 2026-04-09 04:36:18 UTC
  Preview : Hi, just testing the webhook...
============================================================
  Summary :
  A brief test email...
============================================================

✅ Summarized: 'Hello from Patrick'
```

---

## Available MCP Tools

| Tool | Description | Parameters |
|---|---|---|
| `get_recent_emails` | Fetch recent emails with preview | `count` (default 10, max 25) |
| `get_email_body` | Get full body of a specific email | `email_id` (from get_recent_emails) |
| `summarize_emails` | Fetch and format email digest | `count` (default 5, max 10) |

---

## API Endpoints

| Method | Path | Description |
|---|---|---|
| `GET` | `/webhook` | Health check + Graph validation handshake |
| `POST` | `/webhook` | Graph change notification receiver |
| `POST` | `/admin/register` | Trigger webhook subscription registration |
| `GET` | `/docs` | FastAPI Swagger UI (auto-generated) |

---

## Azure OpenAI RBAC Setup

This project uses `DefaultAzureCredential` (via `az login`) — API keys are not used.
Your account needs the **Cognitive Services OpenAI User** role:

```bash
az role assignment create \
  --role "Cognitive Services OpenAI User" \
  --assignee YOUR_EMAIL \
  --scope /subscriptions/YOUR_SUB_ID/resourceGroups/YOUR_RG/providers/Microsoft.CognitiveServices/accounts/YOUR_AI_RESOURCE
```

Wait 1-2 minutes for role propagation, then run `email_agent.py`.

---

## Troubleshooting

### Subscription validation timed out
Graph gives 10 seconds to respond to validation. Causes:
- Server not fully started when `/admin/register` was called — wait a few seconds and retry
- Dev Tunnel set to Private — set to **Public** in VS Code PORTS tab
- Wrong URL in `.env` — verify it includes `/webhook` at the end

### HTTP 530 from Cloudflare tunnel
Cloudflare free tunnels (`trycloudflare.com`) are blocked or fail DNS from Graph's servers.
Use VS Code Dev Tunnels instead (`*.devtunnels.ms`).

### Unauthorized on Dev Tunnel
```
VS Code → PORTS tab → right-click 8000 → Port Visibility → Public
```

### PermissionDenied on Azure OpenAI (401)
Add Cognitive Services OpenAI User role — see RBAC Setup section above.

### Webhook subscription expires
Graph subscriptions expire after ~3 days. The server auto-renews every 48 hours
while running. After restarting the server, call `/admin/register` again.

### Dev Tunnel URL changed after restart
Right-click port 8000 in PORTS tab → **Make Persistent**.
If it changed, update `.env` → restart `outlook_mcp_server.py` → re-register.

### `ManagedIdentityCredential` warnings in logs
These are normal — `DefaultAzureCredential` tries multiple auth methods.
It will eventually succeed via `EnvironmentCredential` or `AzureCliCredential`.

---

## Security Notes

- Never commit `.env` — it contains secrets
- Rotate `AZURE_CLIENT_SECRET` regularly (set a calendar reminder before expiry)
- `WEBHOOK_CLIENT_STATE` is used to verify notifications come from your subscription
- This uses **app-only** (non-OBO) auth — the agent acts as itself, not as the signed-in user
- For production: deploy to Azure Container Apps — no Dev Tunnel needed, permanent HTTPS URL

---

## Dependencies

```
fastmcp>=3.2.2
fastapi
uvicorn
httpx
azure-identity
python-dotenv
openai
```

Install all:
```bash
pip install -r requirements.txt
```
