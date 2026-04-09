# 📧 Outlook Email Summarizer Agent
### Microsoft Agent Framework 1.0 + Agent 365 SDK + MCP

A real-time email monitoring and summarization agent that demonstrates the full Microsoft enterprise agent stack:

- **Microsoft Agent Framework 1.0** — agent orchestration, tool calling, middleware pipeline
- **Agent 365 Observability SDK** — OTel-based tracing (BaggageBuilder, InvokeAgentScope, InferenceScope)
- **Agent 365 Notifications SDK** — Teams-hosted event routing (architecture documented)
- **Purview DLP Middleware** — content policy enforcement on all agent inputs/outputs
- **FastMCP 3.x** — MCP tool server exposing Graph email tools
- **Microsoft Graph API** — real-time email access via webhook change notifications
- **VS Code Dev Tunnels** — secure public HTTPS endpoint trusted by Microsoft Graph

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────────┐
│  REAL-TIME EMAIL MONITORING (outlook_mcp_server.py)                 │
│                                                                     │
│  New email in Outlook                                               │
│       ↓                                                             │
│  Microsoft Graph Webhook → VS Code Dev Tunnel (devtunnels.ms)       │
│       ↓                                                             │
│  FastAPI /webhook (port 8000)                                       │
│       ↓                                                             │
│  Graph API fetches email → Prints real-time summary to console      │
│                                                                     │
│  FastMCP Tool Server (port 8001)                                    │
│    ├── get_recent_emails()    → Graph /users/{id}/messages          │
│    ├── get_email_body()       → Graph /users/{id}/messages/{id}     │
│    └── summarize_emails()     → Graph + formatting                  │
└─────────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────────┐
│  CONSOLE CHAT AGENT (email_agent_a365.py)                           │
│                                                                     │
│  You: "summarize my latest email"                                   │
│       ↓                                                             │
│  BaggageBuilder (OTel context: tenant_id, agent_id, correlation_id) │
│       ↓                                                             │
│  InvokeAgentScope (outer OTel span: full agent invocation)          │
│       ↓                                                             │
│  Agent Framework 1.0 Agent.run()                                    │
│    → Purview DLP pre-check (input scanned)                          │
│    → InferenceScope (inner OTel span: LLM call)                     │
│    → OpenAIChatCompletionClient → Azure OpenAI GPT-5                │
│        → fetch_email_digest() tool                                  │
│        → Graph API → email data                                     │
│    → Purview DLP post-check (output scanned)                        │
│       ↓                                                             │
│  Agent365Exporter → A365 Observability Backend                      │
│       ↓                                                             │
│  Console output                                                     │
└─────────────────────────────────────────────────────────────────────┘
```

---

## Project Structure

```
mail-reader-a365/
├── outlook_mcp_server.py    # FastAPI webhook receiver + FastMCP tool server
├── email_agent_a365.py      # Agent Framework 1.0 + A365 SDK console agent
├── email_agent.py           # Simple console agent (Azure OpenAI, no A365 SDK)
├── validation.py            # Standalone Graph auth test script
├── a365.config.json         # A365 CLI generated config (agent registration)
├── a365.state.json          # A365 CLI state (blueprint ID, consents)
├── .env                     # Credentials (never commit — gitignored)
├── .env.example             # Template for .env
├── .gitignore
├── requirements.txt
└── README.md
```

---

## SDK Stack Details

### Microsoft Agent Framework 1.0
```
pip install agent-framework
```

The core agent orchestration layer. Key components used:

| Component | Purpose |
|---|---|
| `Agent` | Main agent class — wraps LLM client, tools, middleware |
| `OpenAIChatCompletionClient` | Azure OpenAI client (Chat Completions API) |
| Tool functions | Plain Python functions with `Annotated` type hints |
| `PurviewPolicyMiddleware` | DLP middleware wired into agent pipeline |

```python
from agent_framework import Agent
from agent_framework.openai import OpenAIChatCompletionClient
from agent_framework.microsoft import PurviewPolicyMiddleware, PurviewSettings

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
    middleware=[PurviewPolicyMiddleware(credential=..., settings=PurviewSettings(...))]
)

response = await agent.run("Summarize my latest email")
```

### Agent 365 Observability SDK
```
pip install microsoft-agents-a365-observability-core
```

OpenTelemetry-based observability for A365. Every agent invocation is traced:

| Primitive | Purpose |
|---|---|
| `BaggageBuilder` | Sets OTel context baggage (tenant_id, agent_id, blueprint_id, correlation_id) |
| `InvokeAgentScope` | Outer span — tracks full agent invocation from user input to response |
| `InferenceScope` | Inner span — tracks each LLM call (tokens, model, finish reason) |
| `ExecuteToolScope` | Per-tool span — tracks each tool call execution |
| `Agent365Exporter` | Ships OTel spans to A365 observability backend |

```python
from microsoft_agents_a365.observability.core.middleware.baggage_builder import BaggageBuilder
from microsoft_agents_a365.observability.core.invoke_agent_scope import InvokeAgentScope
from microsoft_agents_a365.observability.core.inference_scope import InferenceScope

with BaggageBuilder() \
        .tenant_id(TENANT_ID) \
        .agent_id(AGENT_BLUEPRINT_ID) \
        .agent_upn(AGENT_UPN) \
        .correlation_id(str(uuid.uuid4())) \
        .build():
    
    invoke_scope = InvokeAgentScope.start(invoke_details, tenant_details)
    invoke_scope.record_input_messages([user_input])
    
    inference_scope = InferenceScope.start(inference_details, agent_details, tenant_details)
    
    response = await agent.run(user_input)
    
    inference_scope.record_finish_reasons(["stop"])
    invoke_scope.record_response(str(response))
```

### Agent 365 Notifications SDK
```
pip install microsoft-agents-a365-notifications
```

Receives real-time Microsoft 365 push events and routes them to agent handlers.

> ⚠️ **Requires M365 Agents SDK hosting** (TurnContext/TurnState/AgentApplication).
> For this console agent, real-time notifications are handled by the Graph webhook
> in `outlook_mcp_server.py`. The Notifications SDK is documented for Teams deployment.

**Teams deployment pattern:**
```python
from microsoft_agents_a365.notifications import AgentNotification, AgentNotificationActivity
from microsoft_agents.hosting.core import AgentApplication, TurnContext, TurnState

bot_app = AgentApplication(auth_handler=..., storage=...)
agent_notification = AgentNotification(app=bot_app)

@agent_notification.on_email()
async def handle_new_email(
    context: TurnContext,
    state: TurnState,
    notification_activity: AgentNotificationActivity
) -> None:
    email = notification_activity.email
    result = await agent.run(f"Summarize: {email.subject}\n{email.html_body}")
    await context.send_activity(str(result))

@agent_notification.on_user_created()
async def handle_user_onboarded(context, state, notification_activity):
    await context.send_activity("Welcome! I'm your email assistant.")
```

### Purview DLP Middleware
```
pip install agent-framework  # includes agent_framework.microsoft
```

Intercepts ALL agent inputs and outputs for DLP policy enforcement:

```
User input → Purview pre-check → Agent LLM → Purview post-check → Response
```

- **Pre-check**: scans user's prompt for sensitive data
- **Post-check**: scans agent's response before it's returned to the user
- **Blocking**: if content violates DLP policy, response is blocked and exception raised
- **Audit**: all interactions logged to Microsoft Purview audit trail

**Required Graph permissions** (delegated) on your Entra app:
- `ProtectionScopes.Compute.All`
- `ContentActivity.Write`
- `Content.Process.All`

---

## Prerequisites

| Requirement | Details |
|---|---|
| Python 3.12+ | |
| VS Code | With Microsoft account signed in (required for Dev Tunnels) |
| Azure CLI (`az`) | Logged in: `az login` |
| Azure AI Services resource | With GPT-5 deployment |
| Microsoft 365 Developer tenant | Admin access — [M365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program) |
| Entra App Registration | Mail.Read + Purview permissions |
| A365 CLI | `npm install -g @microsoft/agent365-cli` |

---

## Part 1 — Entra App Registration

### 1.1 Create the App

1. Go to https://entra.microsoft.com
2. Sign in as tenant admin
3. **App registrations → New registration**
   - Name: `email-summarizer-agent`
   - Supported account types: Single tenant
   - Click **Register**
4. Note down **Application (client) ID** and **Directory (tenant) ID**

### 1.2 Add Graph API Permissions

**Application permissions** (for Graph webhook + mail access):
```
Microsoft Graph → Application permissions:
  ✅ Mail.Read
  ✅ Mail.ReadBasic
  ✅ User.Read.All
→ Grant admin consent
```

**Delegated permissions** (for Purview DLP middleware):
```
Microsoft Graph → Delegated permissions:
  ✅ ProtectionScopes.Compute.All
  ✅ ContentActivity.Write
  ✅ Content.Process.All
→ Grant admin consent
```

### 1.3 Create Client Secret

1. Your app → **Certificates & secrets → New client secret**
2. Description: `email-agent-dev` | Expiry: 180 days
3. **Copy the Value immediately** — shown only once

---

## Part 2 — A365 Agent Registration

```bash
# Install A365 CLI
npm install -g @microsoft/agent365-cli

# Login
a365 auth login

# Initialize agent configuration (interactive wizard)
a365 init

# Setup Azure resources and register agent blueprint
a365 setup all
```

This creates:
- **Entra Agent Identity** with its own UPN (`agentname@yourtenant.onmicrosoft.com`)
- **Agent Blueprint** in A365 control plane
- **OAuth2 grants** for observability, notifications, messaging APIs
- `a365.config.json` and `a365.state.json` in your project

---

## Part 3 — Project Setup

### 3.1 Clone and Create Virtual Environment

```bash
git clone https://github.com/patrick-shim/email-agent-with-outlook-mcp-server.git
cd email-agent-with-outlook-mcp-server

python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS/Linux
source .venv/bin/activate
```

### 3.2 Install Dependencies

```bash
pip install -r requirements.txt
```

Full dependency list:
```
fastmcp>=3.2.2
fastapi
uvicorn
httpx
azure-identity
python-dotenv
openai
agent-framework
microsoft-agents-a365-observability-core
microsoft-agents-a365-notifications
wrapt
pydantic
```

### 3.3 Create `.env` File

```bash
# Windows
Copy-Item .env.example .env

# macOS/Linux
cp .env.example .env
```

Edit `.env` with your values:

```bash
# ── Entra / Graph ──────────────────────────────────────────────
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
TARGET_USER_EMAIL=admin@yourtenant.onmicrosoft.com
GRAPH_ENDPOINT=https://graph.microsoft.com/v1.0

# ── Webhook ────────────────────────────────────────────────────
# Get from VS Code PORTS tab after forwarding port 8000
WEBHOOK_URL=https://YOUR-TUNNEL-ID-8000.devtunnels.ms/webhook
WEBHOOK_CLIENT_STATE=email-agent-secret-2026

# ── Azure OpenAI ───────────────────────────────────────────────
AZURE_OPENAI_ENDPOINT=https://your-ai-resource.cognitiveservices.azure.com/
AZURE_OPENAI_DEPLOYMENT=gpt-5-default
AZURE_OPENAI_API_VERSION=2025-03-01-preview

# ── A365 Agent Identity (from a365.state.json after a365 setup all) ──
A365_AGENT_BLUEPRINT_ID=your-blueprint-id
A365_AGENT_UPN=youragent@yourtenant.onmicrosoft.com
A365_MANAGER_EMAIL=admin@yourtenant.onmicrosoft.com

# ── Purview (optional — set to enable DLP middleware) ──────────
# PURVIEW_CLIENT_APP_ID=your-client-id
# Requires delegated permissions: ProtectionScopes.Compute.All,
# ContentActivity.Write, Content.Process.All
```

### 3.4 Validate Graph Auth

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
  Subject : Microsoft Entra ID Protection Weekly Digest
  From    : MSSecurity-noreply@microsoft.com
```

---

## Part 4 — Azure OpenAI RBAC

This project uses `DefaultAzureCredential` / `AzureCliCredential` — API keys are not used.

```bash
az role assignment create \
  --role "Cognitive Services OpenAI User" \
  --assignee YOUR_EMAIL \
  --scope /subscriptions/YOUR_SUB_ID/resourceGroups/YOUR_RG/providers/Microsoft.CognitiveServices/accounts/YOUR_AI_RESOURCE
```

Wait 1-2 minutes for propagation, then verify:
```bash
az cognitiveservices account list --output table
```

---

## Part 5 — VS Code Dev Tunnel Setup

Microsoft Graph webhooks require a public HTTPS URL. VS Code Dev Tunnels
expose your localhost securely using `*.devtunnels.ms` — a Microsoft domain
that Graph fully trusts.

> ⚠️ Do NOT use Cloudflare free tunnels (`trycloudflare.com`) or ngrok free tier.
> These are blocked by Microsoft Graph DNS or fail validation.

### 5.1 Sign into VS Code with Microsoft Account

```
VS Code → Accounts icon (bottom left) → Sign in with Microsoft Account
```

### 5.2 Forward Port 8000

```
VS Code → Bottom panel → PORTS tab
→ "+ Forward a Port" → 8000 → Enter
→ Right-click 8000 → Port Visibility → Public
→ Right-click 8000 → Make Persistent   ← same URL across restarts
```

### 5.3 Copy Tunnel URL and Update `.env`

```bash
# In PORTS tab, copy the Forwarded Address, e.g.:
WEBHOOK_URL=https://xxxxxxxx-8000.jpe1.devtunnels.ms/webhook
```

---

## Part 6 — Running the MCP Server + Webhook

### Terminal 1 — Start MCP + Webhook Server

```bash
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

> ⚠️ **Why manual registration?**
> Graph validates your webhook endpoint during subscription registration
> (10 second limit). Auto-registering at startup causes a race condition.
> Manual registration after the server is fully up avoids this entirely.

---

## Part 7 — Running the A365 Agent

```bash
python email_agent_a365.py
```

Expected startup:
```
✅ Agent365Exporter configured
✅ ConsoleSpanExporter configured (dev mode)
✅ Purview middleware wired into agent pipeline

==============================================================
  📧 Outlook Email Agent
  Microsoft Agent Framework 1.0 + A365 SDK
  Model        : gpt-5-default
  Agent ID     : c3d385f8-83d4-4650-8f3e-e1e3e80b163e
  Agent UPN    : outlookemailsummary@yourtenant.onmicrosoft.com
==============================================================
  SDK Stack:
    ✅ Agent Framework 1.0  (Agent + OpenAIChatCompletionClient + tools)
    ✅ Purview Middleware   (Active — all inputs/outputs DLP-checked)
    ✅ A365 Observability  (BaggageBuilder + InvokeAgentScope + InferenceScope)
    ✅ Agent365Exporter    (OTel → A365 observability backend)
    ℹ️  Notifications SDK  (requires M365 Agents hosting)
==============================================================

You: summarize my latest email
Agent: Your latest email is from Patrick Shim at Microsoft
       with the subject "mail agent test."...
```

Each response generates OTel spans visible in console:
```json
{
  "name": "purview.get_protection_scopes",
  "status": {"status_code": "UNSET"},
  "attributes": {"correlation_id": "abc123@AF"}
}
```

---

## Part 8 — Real-time Email Monitoring

With `outlook_mcp_server.py` running and webhook registered,
send an email to `TARGET_USER_EMAIL`. Within 2-3 seconds:

```
📬 New email notification received
📨 New email detected, fetching ID: AAMkAD...

============================================================
🔵 NEW EMAIL RECEIVED
============================================================
  Subject : Hello from Patrick
  From    : patrick.shim@microsoft.com
  Date    : 2026-04-09 04:36:18 UTC
  Preview : Testing the real-time webhook...
============================================================
  Summary :
  Brief test email with signature...
============================================================

✅ Summarized: 'Hello from Patrick'
```

---

## Part 9 — Purview DLP Demo

To demonstrate Purview blocking sensitive content:

### 9.1 Create DLP Policy

```
https://compliance.microsoft.com
→ Data loss prevention → Policies → Create policy
→ Custom policy
→ Name: "Agent Sensitive Data Test"
→ Locations: AI apps and services
→ Rule: detect Credit Card Number or U.S. SSN
→ Action: Block
→ Enable policy
```

### 9.2 Send Test Email With Sensitive Data

Send an email containing:
```
SSN: 123-45-6789
Credit card: 4111-1111-1111-1111
```

### 9.3 Ask the Agent

```
You: summarize my latest email
```

Expected: Purview intercepts the response, raises `PurviewBlockedResponseError`,
and you see a DLP violation in the Purview audit log at:
```
https://compliance.microsoft.com → Audit → Search
```

### 9.4 Verify in Audit Logs

All agent interactions are logged regardless of blocking:
```
Audit → Activities: ContentActivity.Write
Filter by App ID: your-client-id
```

---

## API Endpoints (MCP Server)

| Method | Path | Description |
|---|---|---|
| `GET` | `/webhook` | Health check + Graph validation handshake |
| `POST` | `/webhook` | Graph change notification receiver |
| `POST` | `/admin/register` | Trigger webhook subscription registration |
| `GET` | `/docs` | FastAPI Swagger UI |

## MCP Tools

| Tool | Description | Parameters |
|---|---|---|
| `get_recent_emails` | Fetch recent emails with preview | `count` (default 10, max 25) |
| `get_email_body` | Get full body of a specific email | `email_id` |
| `summarize_emails` | Fetch and format email digest | `count` (default 5, max 10) |

---

## Troubleshooting

### Subscription validation timed out
- Server not fully started — wait and retry `/admin/register`
- Dev Tunnel set to Private — set to **Public** in VS Code PORTS tab
- Wrong URL in `.env` — must include `/webhook` at the end

### API version not supported (400)
```bash
# Use this version for Chat Completions API with GPT-5
AZURE_OPENAI_API_VERSION=2025-03-01-preview
```
Note: `OpenAIChatClient` uses Responses API (requires `/openai/v1/` base URL).
Use `OpenAIChatCompletionClient` for standard Chat Completions endpoint.

### Purview ReadTimeout
Purview DLP endpoint latency can be high from certain regions.
Retry or temporarily disable by commenting out `PURVIEW_CLIENT_APP_ID`.

### PermissionDenied on Azure OpenAI (401)
```bash
az role assignment create \
  --role "Cognitive Services OpenAI User" \
  --assignee YOUR_EMAIL \
  --scope /subscriptions/.../accounts/YOUR_AI_RESOURCE
```

### Webhook subscription expires
Graph subscriptions expire after ~3 days. Server auto-renews every 48 hours.
After restart, call `/admin/register` again.

### `ManagedIdentityCredential` warnings
Normal — `DefaultAzureCredential` tries multiple auth methods in order.
Will succeed via `AzureCliCredential`.

### ImportError: No module named 'wrapt'
```bash
pip install wrapt
```

---

## Security Notes

- Never commit `.env` — contains secrets
- Rotate `AZURE_CLIENT_SECRET` before expiry (set calendar reminder)
- `WEBHOOK_CLIENT_STATE` verifies notifications are from your subscription
- Uses **app-only** (non-OBO) auth for Graph — agent acts as itself
- Purview `InteractiveBrowserCredential` uses **delegated** auth — acts as signed-in user
- For production: deploy to Azure Container Apps for permanent HTTPS URL

---

## Terminal Layout

```
Terminal 1:  python outlook_mcp_server.py    ← MCP tools + webhook (port 8000/8001)
Terminal 2:  python email_agent_a365.py      ← A365 agent chat interface
Terminal 3:  (optional) Register webhook:
             Invoke-RestMethod -Method POST -Uri http://localhost:8000/admin/register
```

---

## References

- [Microsoft Agent Framework 1.0](https://learn.microsoft.com/en-us/agent-framework/)
- [Agent 365 SDK Documentation](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/)
- [Agent 365 Observability](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/observability)
- [Purview + Agent Framework](https://learn.microsoft.com/en-us/agent-framework/integrations/purview)
- [Microsoft Graph Change Notifications](https://learn.microsoft.com/en-us/graph/webhooks)
- [VS Code Dev Tunnels](https://code.visualstudio.com/docs/remote/tunnels)
- [M365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program)
