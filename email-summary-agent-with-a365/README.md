# 📧 Outlook Email Summary Agent
### Microsoft Agent Framework 1.0 + Agent 365 SDK

A production-grade email summarization agent demonstrating the full Microsoft enterprise agent governance stack. Built on Agent Framework 1.0 and extended with Agent 365 observability, Purview DLP middleware, and A365 agent registration.

> **Note:** This is the full enterprise version with A365 SDK.
> For the simpler baseline version without A365 see
> [`../email-summary-agent`](../email-summary-agent/).

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│  REAL-TIME MONITORING  (outlook_mcp_server.py)                      │
│                                                                     │
│  New email in Outlook                                               │
│       ↓                                                             │
│  Microsoft Graph Webhook → VS Code Dev Tunnel (devtunnels.ms)       │
│       ↓                                                             │
│  FastAPI /webhook (port 8000)                                       │
│       ↓                                                             │
│  Graph API fetches email → console summary                          │
│                                                                     │
│  FastMCP Tool Server (port 8001)                                    │
│    ├── fetch_recent_emails()                                        │
│    ├── fetch_email_body()                                           │
│    └── fetch_email_digest()                                         │
└─────────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────────┐
│  CONSOLE CHAT AGENT  (email_agent_a365.py)                          │
│                                                                     │
│  You: "summarize my latest email"                                   │
│       ↓                                                             │
│  BaggageBuilder                                                     │
│    sets OTel context: tenant_id, agent_id, blueprint_id,            │
│    agent_upn, correlation_id, conversation_id                       │
│       ↓                                                             │
│  InvokeAgentScope (outer OTel span)                                 │
│       ↓                                                             │
│  Agent Framework Agent.run()                                        │
│    → Purview DLP pre-check (input scanned against DLP policies)     │
│    → InferenceScope (inner OTel span: LLM call)                     │
│    → OpenAIChatCompletionClient → Azure OpenAI GPT-5                │
│        → fetch_email_digest() tool → Graph API → email data         │
│    → Purview DLP post-check (output scanned)                        │
│       ↓                                                             │
│  Agent365Exporter → A365 Observability Backend                      │
│       ↓                                                             │
│  Console output                                                     │
└─────────────────────────────────────────────────────────────────────┘
```

---

## Project Files

```
email-summary-agent-with-a365/
├── outlook_mcp_server.py    # FastAPI webhook + FastMCP tool server
├── email_agent_a365.py      # Agent Framework 1.0 + A365 SDK agent
├── email_agent.py           # Agent Framework 1.0 agent (no A365 SDK)
├── validation.py            # Graph auth test script
├── a365.config.json         # A365 CLI agent configuration
├── a365.state.json          # A365 CLI state (blueprint ID, consents)
├── .env                     # Credentials (gitignored)
├── .env.example             # Template
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
| `Agent` | Orchestrates LLM calls, tool execution, middleware pipeline |
| `OpenAIChatCompletionClient` | Azure OpenAI via Chat Completions API |
| Tool functions | Plain Python functions with `Annotated` type hints |
| `PurviewPolicyMiddleware` | DLP middleware wired into agent pipeline |

### Agent 365 Observability SDK
```
pip install microsoft-agents-a365-observability-core
```

| Primitive | Purpose |
|---|---|
| `BaggageBuilder` | Sets OTel context baggage per request (tenant, agent, correlation IDs) |
| `InvokeAgentScope` | Outer OTel span — tracks full agent invocation |
| `InferenceScope` | Inner OTel span — tracks each LLM call (tokens, model, finish reason) |
| `Agent365Exporter` | Ships OTel spans to A365 observability backend |

### Agent 365 Notifications SDK
```
pip install microsoft-agents-a365-notifications
```

Receives real-time Microsoft 365 push events and routes them to agent handlers.
Requires M365 Agents SDK hosting (TurnContext/TurnState) for full deployment.
See code comments in `email_agent_a365.py` for the Teams deployment pattern.

### Purview DLP Middleware
Built into `agent-framework` via `agent_framework.microsoft`.

Intercepts all agent inputs and outputs:
```
User input → Purview pre-check → LLM → tool calls → Purview post-check → Response
```

- Pre-check: scans user prompt against DLP policies
- Post-check: scans agent response before returning to user
- Blocking: raises exception if content violates policy
- Audit: all interactions logged to Microsoft Purview audit trail

---

## Prerequisites

| Requirement | Details |
|---|---|
| Python 3.12+ | |
| VS Code | Microsoft account signed in (Dev Tunnels) |
| Azure CLI | `az login` |
| Azure AI Services | GPT-5 deployment |
| M365 tenant (admin) | [M365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program) |
| Entra App Registration | `Mail.Read` + Purview permissions |
| A365 CLI | `npm install -g @microsoft/agent365-cli` |
| Microsoft Purview | E5 license (included in M365 Developer Program) |

---

## Part 1 — Entra App Registration

### 1.1 Create the App

1. Go to https://entra.microsoft.com → sign in as tenant admin
2. **App registrations → New registration**
   - Name: `email-summarizer-agent`
   - Supported account types: Single tenant
3. Note down **Application (client) ID** and **Directory (tenant) ID**

### 1.2 Add Graph Permissions

**Application permissions** (Graph webhook + mail access):
```
API Permissions → Add → Microsoft Graph → Application permissions:
  ✅ Mail.Read
  ✅ Mail.ReadBasic
  ✅ User.Read.All
→ Grant admin consent
```

**Delegated permissions** (Purview DLP middleware):
```
API Permissions → Add → Microsoft Graph → Delegated permissions:
  ✅ ProtectionScopes.Compute.All
  ✅ ContentActivity.Write
  ✅ Content.Process.All
→ Grant admin consent
```

### 1.3 Create Client Secret

```
Certificates & secrets → New client secret
Description: email-agent-dev  |  Expiry: 180 days
Copy the Value immediately — shown only once
```

---

## Part 2 — A365 Agent Registration

```bash
# Install A365 CLI
npm install -g @microsoft/agent365-cli

# Login
a365 auth login

# Interactive setup wizard
a365 init
# Provide: agent name, resource group, messaging endpoint, manager email

# Create Azure resources + register agent blueprint
a365 setup all
```

This creates:
- **Entra Agent Identity** with UPN (`agentname@yourtenant.onmicrosoft.com`)
- **Agent Blueprint** in A365 control plane
- **OAuth2 grants** for observability, notifications, messaging APIs
- `a365.config.json` and `a365.state.json` in your project

The Blueprint ID from `a365.state.json` goes into your `.env` as `A365_AGENT_BLUEPRINT_ID`.

---

## Part 3 — Azure OpenAI RBAC

Uses `AzureCliCredential` — no API keys needed.

```bash
az role assignment create \
  --role "Cognitive Services OpenAI User" \
  --assignee YOUR_EMAIL \
  --scope /subscriptions/YOUR_SUB/resourceGroups/YOUR_RG/providers/Microsoft.CognitiveServices/accounts/YOUR_RESOURCE
```

---

## Part 4 — Project Setup

```bash
python -m venv .venv
.venv\Scripts\activate        # Windows
source .venv/bin/activate     # macOS/Linux

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

# A365 Agent Identity (from a365.state.json after a365 setup all)
A365_AGENT_BLUEPRINT_ID=your-blueprint-id
A365_AGENT_UPN=youragent@yourtenant.onmicrosoft.com
A365_MANAGER_EMAIL=admin@yourtenant.onmicrosoft.com

# Purview DLP (optional — comment out to disable)
# Requires delegated Graph permissions above
PURVIEW_CLIENT_APP_ID=your-client-id
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

## Part 5 — VS Code Dev Tunnel

Graph webhooks require a public HTTPS URL. VS Code Dev Tunnels expose your
localhost using `*.devtunnels.ms` — a domain Microsoft Graph fully trusts.

> ⚠️ Do not use Cloudflare free tunnels or ngrok free tier —
> these are blocked by Microsoft Graph DNS.

```
VS Code → Accounts → Sign in with Microsoft Account

VS Code → PORTS tab
→ Forward a Port → 8000
→ Right-click → Port Visibility → Public
→ Right-click → Make Persistent
→ Copy Forwarded Address
```

Update `.env`:
```bash
WEBHOOK_URL=https://xxxxxxxx-8000.jpe1.devtunnels.ms/webhook
```

---

## Part 6 — Running

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

### Terminal 2 — Run A365 Agent

```bash
python email_agent_a365.py
```

```
✅ Agent365Exporter configured
✅ ConsoleSpanExporter configured (dev mode)
✅ Purview middleware wired into agent pipeline

==============================================================
  📧 Outlook Email Agent
  Microsoft Agent Framework 1.0 + A365 SDK
  Model        : gpt-5-default
  Agent ID     : c3d385f8-...
  Agent UPN    : youragent@yourtenant.onmicrosoft.com
==============================================================
  SDK Stack:
    ✅ Agent Framework 1.0  (Agent + OpenAIChatCompletionClient + tools)
    ✅ Purview Middleware   (Active — all inputs/outputs DLP-checked)
    ✅ A365 Observability  (BaggageBuilder + InvokeAgentScope + InferenceScope)
    ✅ Agent365Exporter    (OTel → A365 observability backend)
    ℹ️  Notifications SDK  (requires M365 Agents hosting)
==============================================================

You: summarize my latest email
Agent: Your latest email is from Patrick Shim at Microsoft...
```

---

## Purview DLP Demo

To demonstrate Purview blocking sensitive content in emails:

### Create a DLP Policy

```
https://compliance.microsoft.com
→ Data loss prevention → Policies → Create policy
→ Custom policy
→ Name: "Agent Sensitive Data Test"
→ Locations: AI apps and services
→ Rule: detect Credit Card Number or U.S. SSN
→ Action: Block
→ Enable immediately
```

### Send a Test Email With Sensitive Data

Send an email to `TARGET_USER_EMAIL` containing:
```
SSN: 123-45-6789
Credit card: 4111-1111-1111-1111
```

### Ask the Agent

```
You: summarize my latest email
```

Expected: Purview post-check detects the sensitive data in the agent's response and raises a `PurviewBlockedResponseError`.

### Verify in Audit Logs

```
https://compliance.microsoft.com → Audit → Search
Filter: App ID = your-client-id, last 1 hour
```

All agent interactions are logged here regardless of blocking.

---

## Observability in Console

Each agent run produces OTel spans visible in the console:

```json
{
  "name": "purview.get_protection_scopes",
  "status": {"status_code": "UNSET"},
  "attributes": {"correlation_id": "abc123@AF"}
}
```

`status_code: UNSET` = success (content allowed).
`status_code: ERROR` = DLP violation or connectivity issue.

Spans are also shipped to the A365 observability backend via `Agent365Exporter`
where they appear in the Microsoft Admin Center agent monitoring view.

---

## A365 Notifications SDK — Teams Deployment Pattern

The full Notifications SDK requires M365 Agents SDK hosting. For Teams deployment:

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
    result = await agent.run(
        f"New email from {email.sender}: {email.subject}\n{email.html_body}"
    )
    await context.send_activity(str(result))

@agent_notification.on_user_created()
async def handle_onboarded(context, state, notification_activity):
    await context.send_activity("Welcome! I'm your email assistant.")
```

For this console agent, real-time notifications are handled by the Graph webhook
in `outlook_mcp_server.py` which serves the same purpose.

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
- `.env` `WEBHOOK_URL` must include `/webhook`

**API version not supported (400)**
```bash
# Use Chat Completions API version (not Responses API)
AZURE_OPENAI_API_VERSION=2025-03-01-preview
```

Note: Use `OpenAIChatCompletionClient` (Chat Completions) not `OpenAIChatClient`
(Responses API). The Responses API requires a different base URL format on Azure.

**Purview ReadTimeout**
Purview endpoint latency can be high from certain regions. Retry or
temporarily disable by commenting out `PURVIEW_CLIENT_APP_ID` in `.env`.

**Purview InsufficientGraphPermissions (403)**
Add the three delegated Graph permissions — see Part 1.2.

**PermissionDenied on Azure OpenAI (401)**
Add `Cognitive Services OpenAI User` role — see Part 3.

**Webhook subscription expires**
Graph subscriptions expire after ~3 days. Server auto-renews every 48 hours.
After restart call `/admin/register` again.

**ImportError: No module named 'wrapt'**
```bash
pip install wrapt
```

---

## Security Notes

- Never commit `.env` — contains secrets
- Rotate `AZURE_CLIENT_SECRET` before expiry
- `WEBHOOK_CLIENT_STATE` verifies webhook notifications are from your subscription
- Graph access uses app-only (non-OBO) auth — agent acts as itself
- Purview DLP uses delegated auth (`InteractiveBrowserCredential`) — acts as signed-in user
- For production: deploy to Azure Container Apps for permanent HTTPS URL

---

## References

- [Microsoft Agent Framework 1.0](https://learn.microsoft.com/en-us/agent-framework/)
- [Agent 365 SDK](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/)
- [Agent 365 Observability](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/observability)
- [Purview + Agent Framework](https://learn.microsoft.com/en-us/agent-framework/integrations/purview)
- [Microsoft Graph Change Notifications](https://learn.microsoft.com/en-us/graph/webhooks)
- [VS Code Dev Tunnels](https://code.visualstudio.com/docs/remote/tunnels)
- [A365 CLI Reference](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/cli)
