# 📧 Outlook Email Summary Agent
### Microsoft Agent Framework 1.0

A console-based email summarization agent built with Microsoft Agent Framework 1.0. Monitors an Outlook mailbox in real time via Microsoft Graph webhooks and answers natural language questions about emails.

> **Note:** This is the baseline version with no A365 governance layer.
> For the enterprise version with observability, Purview DLP, and A365 registration
> see [`../email-summary-agent-with-a365`](../email-summary-agent-with-a365/).

---

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│  REAL-TIME MONITORING  (outlook_mcp_server.py)              │
│                                                             │
│  New email in Outlook                                       │
│       ↓                                                     │
│  Microsoft Graph Webhook                                    │
│       ↓                                                     │
│  VS Code Dev Tunnel (*.devtunnels.ms)                       │
│       ↓                                                     │
│  FastAPI /webhook (port 8000)                               │
│       ↓                                                     │
│  Graph API fetches email → console summary                  │
│                                                             │
│  FastMCP Tool Server (port 8001)                            │
│    ├── fetch_recent_emails()                                │
│    ├── fetch_email_body()                                   │
│    └── fetch_email_digest()                                 │
└─────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────┐
│  CONSOLE CHAT AGENT  (email_agent.py)                       │
│                                                             │
│  You: "summarize my latest email"                           │
│       ↓                                                     │
│  Agent Framework Agent.run()                                │
│       ↓                                                     │
│  OpenAIChatCompletionClient → Azure OpenAI GPT-5            │
│       ↓                                                     │
│  fetch_email_digest() tool → Graph API → email data         │
│       ↓                                                     │
│  Console output                                             │
└─────────────────────────────────────────────────────────────┘
```

---

## Project Files

```
email-summary-agent/
├── outlook_mcp_server.py   # FastAPI webhook + FastMCP tool server
├── email_agent.py          # Agent Framework 1.0 console agent
├── validation.py           # Graph auth test script
├── .env                    # Credentials (gitignored)
├── .env.example            # Template
├── requirements.txt
└── README.md
```

---

## SDK Stack

### Microsoft Agent Framework 1.0
```
pip install agent-framework
```

| Component | Purpose |
|---|---|
| `Agent` | Orchestrates LLM calls and tool execution |
| `OpenAIChatCompletionClient` | Azure OpenAI via Chat Completions API |
| Tool functions | Plain Python functions with `Annotated` type hints |

```python
from agent_framework import Agent
from agent_framework.openai import OpenAIChatCompletionClient

chat_client = OpenAIChatCompletionClient(
    model="gpt-5-default",
    azure_endpoint="https://your-resource.cognitiveservices.azure.com/",
    credential=AzureCliCredential(),
    api_version="2025-03-01-preview",
)

agent = Agent(
    client=chat_client,
    instructions="You are a helpful email assistant...",
    tools=[fetch_recent_emails, fetch_email_body, fetch_email_digest],
)

response = await agent.run("Summarize my latest email")
```

### FastMCP 3.x
```
pip install fastmcp
```

Exposes email tools over MCP protocol (port 8001) and as plain Python functions importable directly by the agent.

### Microsoft Graph API
Email access via app-only authentication (`ClientSecretCredential`).
Webhook change notifications for real-time monitoring.

---

## Prerequisites

| Requirement | Details |
|---|---|
| Python 3.12+ | |
| VS Code | Microsoft account signed in (Dev Tunnels) |
| Azure CLI | `az login` |
| Azure AI Services | GPT-5 deployment |
| M365 tenant (admin) | [M365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program) |
| Entra App Registration | `Mail.Read` application permission |

---

## Part 1 — Entra App Registration

### 1.1 Create the App

1. Go to https://entra.microsoft.com → sign in as tenant admin
2. **App registrations → New registration**
   - Name: `email-summarizer-agent`
   - Supported account types: Single tenant
3. Note down **Application (client) ID** and **Directory (tenant) ID**

### 1.2 Add Graph Permissions

```
API Permissions → Add a permission → Microsoft Graph
→ Application permissions:
  ✅ Mail.Read
→ Grant admin consent
```

### 1.3 Create Client Secret

```
Certificates & secrets → New client secret
Description: email-agent-dev  |  Expiry: 180 days
Copy the Value immediately — shown only once
```

---

## Part 2 — Azure OpenAI RBAC

Uses `AzureCliCredential` — no API keys needed.

```bash
az role assignment create \
  --role "Cognitive Services OpenAI User" \
  --assignee YOUR_EMAIL \
  --scope /subscriptions/YOUR_SUB/resourceGroups/YOUR_RG/providers/Microsoft.CognitiveServices/accounts/YOUR_RESOURCE
```

---

## Part 3 — Project Setup

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS/Linux
source .venv/bin/activate

pip install -r requirements.txt
```

### Create `.env`

```bash
Copy-Item .env.example .env   # Windows
cp .env.example .env          # macOS/Linux
```

Fill in:
```bash
# Entra / Graph
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
TARGET_USER_EMAIL=admin@yourtenant.onmicrosoft.com
GRAPH_ENDPOINT=https://graph.microsoft.com/v1.0
WEBHOOK_CLIENT_STATE=email-agent-secret-2026

# Webhook — fill in after Dev Tunnel setup
WEBHOOK_URL=https://YOUR-TUNNEL-ID.devtunnels.ms/webhook

# Azure OpenAI
AZURE_OPENAI_ENDPOINT=https://your-resource.cognitiveservices.azure.com/
AZURE_OPENAI_DEPLOYMENT=gpt-5-default
AZURE_OPENAI_API_VERSION=2025-03-01-preview
```

### Validate Graph Auth

```bash
python validation.py
```

Expected:
```
🔐 Getting token... ✅
📬 Fetching emails... ✅ Got 5 emails
--- Email 1 --- Subject: ...
```

---

## Part 4 — VS Code Dev Tunnel

Graph webhooks require a public HTTPS URL. VS Code Dev Tunnels expose your
localhost using `*.devtunnels.ms` — a domain Microsoft Graph fully trusts.

> ⚠️ Do not use Cloudflare free tunnels or ngrok free tier —
> these are blocked by Microsoft Graph.

```
VS Code → Accounts → Sign in with Microsoft Account

VS Code → PORTS tab
→ Forward a Port → 8000
→ Right-click → Port Visibility → Public
→ Right-click → Make Persistent
→ Copy Forwarded Address: https://xxxxxxxx-8000.jpe1.devtunnels.ms
```

Update `.env`:
```bash
WEBHOOK_URL=https://xxxxxxxx-8000.jpe1.devtunnels.ms/webhook
```

---

## Part 5 — Running

### Terminal 1 — MCP + Webhook Server

```bash
python outlook_mcp_server.py
```

```
📧 Target mailbox : admin@yourtenant.onmicrosoft.com
🔗 Webhook URL    : https://xxxxxxxx-8000.jpe1.devtunnels.ms/webhook
🔌 FastAPI/Webhook : port 8000
🔌 MCP Server      : port 8001
ℹ️  Server ready. Call POST /admin/register to subscribe.
```

### Terminal 2 — Register Webhook

```bash
# Windows PowerShell
Invoke-RestMethod -Method POST -Uri http://localhost:8000/admin/register

# macOS/Linux
curl -X POST http://localhost:8000/admin/register
```

Watch Terminal 1 for:
```
🤝 Graph POST validation handshake received
✅ Subscription registered: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
```

### Terminal 2 — Run Agent

```bash
python email_agent.py
```

```
=======================================================
  📧 Outlook Email Agent
  Microsoft Agent Framework 1.0
  Model    : gpt-5-default
=======================================================
  SDK Stack:
    ✅ Agent Framework 1.0
       Agent + OpenAIChatCompletionClient + tool functions
  For full A365 SDK version see: email_agent_a365.py
=======================================================

You: summarize my latest email
Agent: Your latest email is from Patrick Shim at Microsoft
       with the subject "mail agent test"...

You: get the full content of the first email
Agent: ...

You: quit
👋 Goodbye!
```

---

## Real-time Monitoring

With `outlook_mcp_server.py` running and webhook registered,
send an email to `TARGET_USER_EMAIL`:

```
📬 New email notification received
============================================================
🔵 NEW EMAIL RECEIVED
============================================================
  Subject : Hello
  From    : sender@example.com
  Date    : 2026-04-09 04:36:18 UTC
============================================================
✅ Summarized: 'Hello'
```

---

## MCP Tools

| Tool | Description | Parameters |
|---|---|---|
| `fetch_recent_emails` | Recent emails with preview | `count` (default 10, max 25) |
| `fetch_email_body` | Full body of specific email | `email_id` |
| `fetch_email_digest` | Formatted email digest | `count` (default 5, max 10) |

## API Endpoints

| Method | Path | Description |
|---|---|---|
| `GET` | `/webhook` | Health check + Graph validation |
| `POST` | `/webhook` | Graph change notification receiver |
| `POST` | `/admin/register` | Register webhook subscription |
| `GET` | `/docs` | FastAPI Swagger UI |

---

## Troubleshooting

**Subscription validation timed out**
- Wait for server to fully start before calling `/admin/register`
- Dev Tunnel must be set to Public (not Private)
- `.env` `WEBHOOK_URL` must include `/webhook` at the end

**API version not supported (400)**
```bash
AZURE_OPENAI_API_VERSION=2025-03-01-preview
```

**PermissionDenied on Azure OpenAI (401)**
Add `Cognitive Services OpenAI User` role — see Part 2.

**Webhook subscription expires**
Graph subscriptions expire after ~3 days. Server auto-renews every 48 hours.
After restart call `/admin/register` again.

---

## Security Notes

- Never commit `.env`
- Rotate `AZURE_CLIENT_SECRET` before expiry
- `WEBHOOK_CLIENT_STATE` verifies notifications are from your subscription
- Uses app-only (non-OBO) auth — agent acts as itself, not as a user
- For production: deploy to Azure Container Apps (permanent HTTPS, no Dev Tunnel)
