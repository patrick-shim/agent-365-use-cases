# email_agent_a365.py
#
# Outlook Email Summarizer Agent
# ─────────────────────────────────────────────────────────────────────────────
# SDK Stack:
#   Microsoft Agent Framework 1.0    pip install agent-framework
#     - Agent + tool functions + OpenAIChatCompletionClient (Azure Chat Completions)
#     - PurviewPolicyMiddleware (DLP on all inputs/outputs)
#
#   Agent 365 Observability SDK 0.1   pip install microsoft-agents-a365-observability-core
#     - BaggageBuilder   → OTel context propagation per request
#     - InvokeAgentScope → outer span: full agent invocation
#     - InferenceScope   → inner span: each LLM call
#     - Agent365Exporter → ships spans to A365 observability endpoint
#
#   Agent 365 Notifications SDK 0.1   pip install microsoft-agents-a365-notifications
#     - AgentNotification: requires M365 Agents SDK hosting (TurnContext/TurnState)
#     - For Teams deployment: routes email/Word/Excel events to agent handlers
#     - For this console agent: architecture stub with full code comments
#
# Architecture:
#   User prompt
#       → BaggageBuilder (OTel context: tenant, agent, correlation IDs)
#           → InvokeAgentScope (outer span)
#               → Agent Framework Agent.run()
#                   → PurviewPolicyMiddleware (DLP check: input)
#                   → OpenAIChatCompletionClient → Azure OpenAI GPT-5
#                       → tool functions (Graph API email fetch)
#                   → PurviewPolicyMiddleware (DLP check: output)
#               → InferenceScope (inner span: LLM call telemetry)
#           → Agent365Exporter (OTel spans → A365 observability)
#   Console output
# ─────────────────────────────────────────────────────────────────────────────

import asyncio
import os
import uuid
import logging
import sys
from typing import Annotated
from pydantic import Field
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.WARNING,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
log = logging.getLogger("email_agent_a365")
log.setLevel(logging.INFO)

# ── Config ─────────────────────────────────────────────────────────────────────
AZURE_OPENAI_ENDPOINT   = os.getenv("AZURE_OPENAI_ENDPOINT", "https://cdx-ai-foundary.cognitiveservices.azure.com/")
AZURE_OPENAI_DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-5-default")
AZURE_OPENAI_API_VER    = os.getenv("AZURE_OPENAI_API_VERSION", "2025-01-01-preview")
TENANT_ID               = os.getenv("AZURE_TENANT_ID")
AGENT_BLUEPRINT_ID      = os.getenv("A365_AGENT_BLUEPRINT_ID", "c3d385f8-83d4-4650-8f3e-e1e3e80b163e")
AGENT_UPN               = os.getenv("A365_AGENT_UPN", "outlookemailsummary@SCIPCP05985001.onmicrosoft.com")
AGENT_NAME              = "outlookemailsummary"
MANAGER_EMAIL           = os.getenv("A365_MANAGER_EMAIL", "admin@SCIPCP05985001.onmicrosoft.com")
PURVIEW_CLIENT_APP_ID   = os.getenv("PURVIEW_CLIENT_APP_ID")  # optional

# ── Azure Credentials ──────────────────────────────────────────────────────────
from azure.identity import AzureCliCredential, DefaultAzureCredential

cli_credential     = AzureCliCredential()
default_credential = DefaultAzureCredential()

# ── A365 Observability SDK ─────────────────────────────────────────────────────
from opentelemetry import trace
from opentelemetry.sdk.trace import TracerProvider
from opentelemetry.sdk.trace.export import BatchSpanProcessor, ConsoleSpanExporter

from microsoft_agents_a365.observability.core.agent_details import AgentDetails
from microsoft_agents_a365.observability.core.tenant_details import TenantDetails
from microsoft_agents_a365.observability.core.inference_call_details import InferenceCallDetails
from microsoft_agents_a365.observability.core.inference_scope import InferenceScope
from microsoft_agents_a365.observability.core.invoke_agent_scope import InvokeAgentScope
from microsoft_agents_a365.observability.core.invoke_agent_details import InvokeAgentDetails
from microsoft_agents_a365.observability.core.inference_operation_type import InferenceOperationType
from microsoft_agents_a365.observability.core.middleware.baggage_builder import BaggageBuilder
from microsoft_agents_a365.observability.core.exporters.agent365_exporter import Agent365Exporter

# ── A365 shared identity objects ───────────────────────────────────────────────
AGENT_DETAILS = AgentDetails(
    agent_id=AGENT_BLUEPRINT_ID,
    agent_name=AGENT_NAME,
    agent_description="Outlook email summarizer — Agent Framework 1.0 + A365 SDK",
    agent_upn=AGENT_UPN,
    agent_blueprint_id=AGENT_BLUEPRINT_ID,
    tenant_id=TENANT_ID,
)
TENANT_DETAILS = TenantDetails(tenant_id=TENANT_ID)


def get_observability_token(agent_id: str, tenant_id: str) -> str | None:
    """Token resolver for Agent365Exporter."""
    try:
        token = default_credential.get_token(
            "https://observability.agent365.microsoft.com/.default"
        )
        return token.token
    except Exception as e:
        log.warning(f"⚠️  Observability token failed (console only): {e}")
        return None


def setup_observability() -> TracerProvider:
    """Initialize OTel with Agent365Exporter + ConsoleSpanExporter."""
    provider = TracerProvider()

    try:
        a365_exporter = Agent365Exporter(
            token_resolver=get_observability_token,
            cluster_category="prod"
        )
        provider.add_span_processor(BatchSpanProcessor(a365_exporter))
        log.info("✅ Agent365Exporter configured")
    except Exception as e:
        log.warning(f"⚠️  Agent365Exporter setup failed: {e}")

    provider.add_span_processor(BatchSpanProcessor(ConsoleSpanExporter()))
    log.info("✅ ConsoleSpanExporter configured (dev mode)")

    trace.set_tracer_provider(provider)
    return provider


# ── Email tool functions ───────────────────────────────────────────────────────
# Imported from the MCP server module.
# Agent Framework 1.0 accepts plain functions as tools.
# Function docstring = tool description sent to the LLM.

sys.path.insert(0, os.path.dirname(__file__))
from outlook_mcp_server import get_recent_emails, get_email_body, summarize_emails


def fetch_recent_emails(
    count: Annotated[int, Field(description="Number of emails to fetch (default 10, max 25)")] = 10,
) -> str:
    """Fetch the most recent emails from the user's Outlook mailbox.
    Returns subject, sender, date, and a short preview for each email."""
    return get_recent_emails(count=min(count, 25))


def fetch_email_body(
    email_id: Annotated[str, Field(description="The email ID returned from fetch_recent_emails")],
) -> str:
    """Fetch the full body of a specific email by its ID.
    Use this when you need the complete content of an email for detailed analysis."""
    return get_email_body(email_id=email_id)


def fetch_email_digest(
    count: Annotated[int, Field(description="Number of emails to summarize (default 5, max 10)")] = 5,
) -> str:
    """Fetch and format a digest of the most recent emails.
    Best for overview requests like 'summarize my emails'."""
    return summarize_emails(count=min(count, 10))


# ── Agent Framework 1.0 — build agent ─────────────────────────────────────────

from agent_framework import Agent
from agent_framework.openai import OpenAIChatCompletionClient  # Chat Completions API


def build_agent() -> tuple[Agent, bool]:
    """
    Build the Agent Framework 1.0 agent.

    Uses OpenAIChatCompletionClient with Azure routing:
      - Uses the Chat Completions API (not Responses API)
      - azure_endpoint + credential → forces Azure routing
      - api_version → standard Azure OpenAI Chat Completions version

    Note on client choice:
      OpenAIChatClient       → uses Responses API (/openai/v1/ base URL required)
      OpenAIChatCompletionClient → uses Chat Completions API (standard versioned endpoint)
      Your gpt-5-default deployment supports Chat Completions with 2025-01-01-preview.

    Middleware pipeline:
      PurviewPolicyMiddleware intercepts ALL agent inputs/outputs automatically.
      DLP policy checks happen inside agent.run() — no extra code needed.

    Returns: (agent, purview_enabled)
    """

    chat_client = OpenAIChatCompletionClient(
        model=AZURE_OPENAI_DEPLOYMENT,
        azure_endpoint=AZURE_OPENAI_ENDPOINT,
        credential=cli_credential,
        api_version=AZURE_OPENAI_API_VER,
    )

    middleware = []
    purview_enabled = False

    if PURVIEW_CLIENT_APP_ID:
        try:
            from agent_framework.microsoft import PurviewPolicyMiddleware, PurviewSettings
            from azure.identity import InteractiveBrowserCredential

            purview_credential = InteractiveBrowserCredential(
                client_id=PURVIEW_CLIENT_APP_ID
            )
            purview_mw = PurviewPolicyMiddleware(
                credential=purview_credential,
                settings=PurviewSettings(app_name=AGENT_NAME)
            )
            middleware.append(purview_mw)
            purview_enabled = True
            log.info("✅ Purview middleware wired into agent pipeline")
        except ImportError:
            log.warning("⚠️  agent_framework.microsoft not available")
        except Exception as e:
            log.warning(f"⚠️  Purview middleware init failed: {e}")
    else:
        log.info("ℹ️  Purview middleware skipped (PURVIEW_CLIENT_APP_ID not set)")

    agent = Agent(
        client=chat_client,
        name=AGENT_NAME,
        instructions="""You are a helpful email assistant with access to the user's Outlook mailbox.

Available tools:
- fetch_email_digest: Use for overview requests ('summarize my emails', 'what's new')
- fetch_recent_emails: Use to list emails with previews
- fetch_email_body: Use when you need the full content of a specific email

Be concise. For each email highlight: sender, subject, and the key point.""",
        tools=[fetch_recent_emails, fetch_email_body, fetch_email_digest],
        middleware=middleware,
    )

    return agent, purview_enabled


# ── A365 Observability — instrumented agent run ────────────────────────────────

async def run_with_observability(
    agent: Agent,
    user_input: str,
    session_id: str,
) -> str:
    """
    Run one agent turn wrapped in A365 observability scopes.

    BaggageBuilder   → propagates tenant/agent/correlation IDs through all OTel spans
    InvokeAgentScope → outer span: full agent invocation
    InferenceScope   → inner span: LLM inference call
    """
    from urllib.parse import urlparse

    correlation_id = str(uuid.uuid4())

    with BaggageBuilder() \
            .tenant_id(TENANT_ID) \
            .agent_id(AGENT_BLUEPRINT_ID) \
            .agent_upn(AGENT_UPN) \
            .agent_blueprint_id(AGENT_BLUEPRINT_ID) \
            .agent_name(AGENT_NAME) \
            .correlation_id(correlation_id) \
            .conversation_id(session_id) \
            .caller_upn(MANAGER_EMAIL) \
            .build():

        # InvokeAgentScope — outer span
        invoke_scope = None
        try:
            invoke_details = InvokeAgentDetails(
                details=AGENT_DETAILS,
                endpoint=urlparse(AZURE_OPENAI_ENDPOINT),
                session_id=session_id,
            )
            invoke_scope = InvokeAgentScope.start(
                invoke_agent_details=invoke_details,
                tenant_details=TENANT_DETAILS,
            )
            invoke_scope.record_input_messages([user_input])
            log.info(f"📡 InvokeAgentScope started (correlation: {correlation_id[:8]}...)")
        except Exception as e:
            log.warning(f"⚠️  InvokeAgentScope failed: {e}")

        # InferenceScope — inner span
        inference_scope = None
        try:
            inference_details = InferenceCallDetails(
                operationName=InferenceOperationType.CHAT,
                model=AZURE_OPENAI_DEPLOYMENT,
                providerName="azure_openai",
            )
            inference_scope = InferenceScope.start(
                details=inference_details,
                agent_details=AGENT_DETAILS,
                tenant_details=TENANT_DETAILS,
            )
        except Exception as e:
            log.warning(f"⚠️  InferenceScope failed: {e}")

        try:
            # Agent Framework 1.0 agent.run()
            # Handles the full agentic loop internally:
            #   1. Sends message to LLM
            #   2. Detects tool calls
            #   3. Executes tools
            #   4. Sends tool results back to LLM
            #   5. Repeats until final answer
            # Purview middleware intercepts at step 1 (input) and step 5 (output)
            response = await agent.run(user_input)
            answer = str(response) if response else ""

            # Record inference telemetry
            if inference_scope:
                try:
                    inference_scope.record_finish_reasons(["stop"])
                    inference_scope.record_output_messages([answer])
                    inference_scope.__exit__(None, None, None)
                except Exception:
                    pass

            # Record invoke telemetry
            if invoke_scope:
                try:
                    invoke_scope.record_response(answer)
                    invoke_scope.__exit__(None, None, None)
                except Exception:
                    pass

            return answer

        except Exception as e:
            if inference_scope:
                try:
                    inference_scope.__exit__(type(e), e, None)
                except Exception:
                    pass
            if invoke_scope:
                try:
                    invoke_scope.__exit__(type(e), e, None)
                except Exception:
                    pass
            raise


# ── Agent 365 Notifications SDK — architecture note ───────────────────────────
#
# microsoft_agents_a365.notifications.AgentNotification
#
# PURPOSE:
#   Receives real-time push events from Microsoft 365 and routes them to
#   agent handler functions. Events: new email, Word comments, Excel changes,
#   Teams messages, PowerPoint, agent lifecycle events.
#
# HOW IT WORKS:
#   Wraps a Bot Framework AgentApplication (M365 Agents SDK).
#   Registers route selectors matching Activity objects by channel_id/sub_channel.
#   Dispatches to your async handler functions.
#
# WHY NOT WIRED HERE:
#   Requires M365 Agents SDK hosting stack:
#     - microsoft_agents.hosting.core.AgentApplication
#     - TurnContext / TurnState (Bot Framework activity protocol)
#     - HTTP server handling /api/messages Bot Framework endpoint
#   This console agent uses Graph webhook (outlook_mcp_server.py) instead.
#
# FOR TEAMS DEPLOYMENT — wire it like this:
#
#   from microsoft_agents_a365.notifications import (
#       AgentNotification, AgentNotificationActivity, AgentLifecycleEvent
#   )
#   from microsoft_agents.hosting.core import AgentApplication, TurnContext, TurnState
#
#   bot_app = AgentApplication(auth_handler=..., storage=...)
#   agent_notification = AgentNotification(app=bot_app)
#
#   @agent_notification.on_email()
#   async def handle_new_email(
#       context: TurnContext,
#       state: TurnState,
#       notification_activity: AgentNotificationActivity
#   ) -> None:
#       email = notification_activity.email
#       if not email:
#           return
#       result = await agent.run(
#           f"New email from {email.sender}: {email.subject}\n\n{email.html_body}"
#       )
#       await context.send_activity(str(result))
#
#   @agent_notification.on_user_created()
#   async def handle_user_onboarded(context, state, notification_activity):
#       await context.send_activity("Welcome! I'm your Outlook email assistant.")


# ── Main ───────────────────────────────────────────────────────────────────────

async def main():
    tracer_provider = setup_observability()
    agent, purview_enabled = build_agent()

    session_id = str(uuid.uuid4())

    print("\n" + "=" * 62)
    print("  📧 Outlook Email Agent")
    print("  Microsoft Agent Framework 1.0 + A365 SDK")
    print(f"  Model        : {AZURE_OPENAI_DEPLOYMENT}")
    print(f"  Endpoint     : {AZURE_OPENAI_ENDPOINT}")
    print(f"  API Version  : {AZURE_OPENAI_API_VER}")
    print(f"  Agent ID     : {AGENT_BLUEPRINT_ID}")
    print(f"  Agent UPN    : {AGENT_UPN}")
    print(f"  Session      : {session_id[:8]}...")
    print("=" * 62)
    print("  SDK Stack:")
    print("    ✅ Agent Framework 1.0")
    print("       Agent + OpenAIChatCompletionClient (Azure) + tool functions")
    print(f"    {'✅' if purview_enabled else '⚠️ '} Purview Middleware (DLP)")
    print(f"       {'Active — all inputs/outputs DLP-checked' if purview_enabled else 'Set PURVIEW_CLIENT_APP_ID to enable'}")
    print("    ✅ A365 Observability SDK")
    print("       BaggageBuilder + InvokeAgentScope + InferenceScope")
    print("       Agent365Exporter → A365 observability backend")
    print("    ℹ️  A365 Notifications SDK")
    print("       Requires M365 Agents hosting — see code comments")
    print("       (Graph webhook in outlook_mcp_server.py serves same purpose)")
    print("=" * 62)
    print("  Commands: 'quit' → exit  |  'clear' → reset session")
    print("=" * 62 + "\n")

    while True:
        try:
            user_input = input("You: ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\n👋 Goodbye!")
            break

        if not user_input:
            continue

        if user_input.lower() in ("quit", "exit", "q"):
            print("👋 Goodbye!")
            break

        if user_input.lower() == "clear":
            session_id = str(uuid.uuid4())
            print(f"🗑️  Session reset: {session_id[:8]}...\n")
            continue

        try:
            print("Agent: ", end="", flush=True)
            answer = await run_with_observability(agent, user_input, session_id)
            print(answer)
            print()
        except Exception as e:
            print(f"\n❌ Error: {e}\n")
            import traceback
            traceback.print_exc()

    log.info("🔄 Flushing telemetry...")
    tracer_provider.force_flush()
    tracer_provider.shutdown()
    log.info("✅ Shutdown complete")


if __name__ == "__main__":
    asyncio.run(main())
