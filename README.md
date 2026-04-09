# 🤖 Agent 365 Use Cases

A collection of Outlook email agent implementations demonstrating the Microsoft enterprise agent stack — from a baseline Agent Framework agent to a fully governed Agent 365 deployment.

---

## Repository Structure

```
agent-365-use-cases/
├── README.md                             ← you are here
├── email-summary-agent/                  ← Agent Framework 1.0, no A365 SDK
│   ├── outlook_mcp_server.py
│   ├── email_agent.py
│   ├── validation.py
│   ├── .env.example
│   ├── requirements.txt
│   └── README.md
└── email-summary-agent-with-a365/        ← Agent Framework 1.0 + full A365 SDK
    ├── outlook_mcp_server.py
    ├── email_agent_a365.py
    ├── email_agent.py
    ├── validation.py
    ├── a365.config.json
    ├── .env.example
    ├── requirements.txt
    └── README.md
```

---

## The Two Agents

### 📧 [`email-summary-agent`](./email-summary-agent/)
**Agent Framework 1.0 — baseline**

A clean, minimal agent that summarizes Outlook emails using Microsoft Agent Framework 1.0 with no A365 governance layer. Good starting point for understanding core Agent Framework patterns.

| | |
|---|---|
| **Agent Framework** | Agent + OpenAIChatCompletionClient + tool functions |
| **LLM** | Azure OpenAI GPT-5 (Chat Completions API) |
| **Email access** | Microsoft Graph API (app-only auth) |
| **Real-time** | Graph webhooks via VS Code Dev Tunnels |
| **MCP tools** | FastMCP 3.x |
| **Observability** | ❌ None |
| **DLP / Compliance** | ❌ None |
| **A365 registration** | ❌ None |

→ [Full setup and usage guide](./email-summary-agent/README.md)

---

### 🏢 [`email-summary-agent-with-a365`](./email-summary-agent-with-a365/)
**Agent Framework 1.0 + Agent 365 SDK — enterprise**

The same email agent extended with the full Agent 365 enterprise governance stack: observability, Purview DLP middleware, agent identity registration, and Notifications SDK.

| | |
|---|---|
| **Agent Framework** | Agent + OpenAIChatCompletionClient + tool functions |
| **LLM** | Azure OpenAI GPT-5 (Chat Completions API) |
| **Email access** | Microsoft Graph API (app-only auth) |
| **Real-time** | Graph webhooks via VS Code Dev Tunnels |
| **MCP tools** | FastMCP 3.x |
| **Observability** | ✅ BaggageBuilder + InvokeAgentScope + InferenceScope + Agent365Exporter |
| **DLP / Compliance** | ✅ Purview DLP Middleware (pre + post content checks) |
| **A365 registration** | ✅ Blueprint + Entra Agent Identity via A365 CLI |
| **Notifications SDK** | ✅ Architecture documented (Teams deployment pattern) |

→ [Full setup and usage guide](./email-summary-agent-with-a365/README.md)

---

## Side-by-Side Comparison

```
                                    email-summary-agent   email-summary-agent-with-a365
                                    ───────────────────   ─────────────────────────────
Agent Framework 1.0                       ✅                          ✅
OpenAIChatCompletionClient                ✅                          ✅
Tool functions (Annotated pattern)        ✅                          ✅
Graph webhook + MCP server                ✅                          ✅
──────────────────────────────────────────────────────────────────────────────────────
A365 Agent Registration                   ❌                          ✅
Entra Agent Identity (UPN + Blueprint)    ❌                          ✅
BaggageBuilder (OTel context)             ❌                          ✅
InvokeAgentScope (outer span)             ❌                          ✅
InferenceScope (LLM span)                 ❌                          ✅
Agent365Exporter (telemetry backend)      ❌                          ✅
Purview DLP Middleware                    ❌                          ✅
Notifications SDK (Teams pattern)         ❌                          ✅
```

---

## Adoption Path

These two agents represent a natural enterprise AI adoption journey:

```
Phase 1 — Build                      Phase 2 — Govern
email-summary-agent                  email-summary-agent-with-a365
───────────────────                  ──────────────────────────────
Learn Agent Framework patterns       Add enterprise observability
Understand tool calling              Enforce DLP policies via Purview
Build core email logic               Register with A365 control plane
Minimal Azure dependencies           Full lifecycle management
Quick setup for prototyping          Production-ready governance
```

---

## Prerequisites

**Both agents:**

| Requirement | Notes |
|---|---|
| Python 3.12+ | |
| VS Code | Microsoft account signed in (Dev Tunnels) |
| Azure CLI | `az login` |
| Azure AI Services | GPT-5 deployment |
| M365 tenant (admin) | [M365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program) recommended |
| Entra App Registration | `Mail.Read` application permission |

**A365 agent additionally:**

| Requirement | Notes |
|---|---|
| A365 CLI | `npm install -g @microsoft/agent365-cli` |
| Microsoft Purview | E5 license (included in M365 Developer Program) |
| Additional Graph permissions | `ProtectionScopes.Compute.All`, `ContentActivity.Write`, `Content.Process.All` |

---

## Quick Start

**Simple agent:**
```bash
cd email-summary-agent
python -m venv .venv && .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env        # fill in your values
python validation.py         # verify Graph auth works
python outlook_mcp_server.py # terminal 1: start MCP + webhook server
# terminal 2: register webhook, then run agent
Invoke-RestMethod -Method POST -Uri http://localhost:8000/admin/register
python email_agent.py
```

**A365 agent:**
```bash
cd email-summary-agent-with-a365
python -m venv .venv && .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env         # fill in your values
a365 init                    # configure agent
a365 setup all               # register with A365 control plane
python validation.py          # verify Graph auth works
python outlook_mcp_server.py  # terminal 1: start MCP + webhook server
# terminal 2: register webhook, then run agent
Invoke-RestMethod -Method POST -Uri http://localhost:8000/admin/register
python email_agent_a365.py
```

---

## References

- [Microsoft Agent Framework 1.0](https://learn.microsoft.com/en-us/agent-framework/)
- [Agent 365 SDK](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/)
- [Agent 365 Observability](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/observability)
- [Purview + Agent Framework](https://learn.microsoft.com/en-us/agent-framework/integrations/purview)
- [Microsoft Graph Change Notifications](https://learn.microsoft.com/en-us/graph/webhooks)
- [VS Code Dev Tunnels](https://code.visualstudio.com/docs/remote/tunnels)
