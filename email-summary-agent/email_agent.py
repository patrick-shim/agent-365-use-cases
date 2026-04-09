# email_agent.py
#
# Outlook Email Summarizer Agent — Simple Version
# ─────────────────────────────────────────────────────────────────────────────
# Built with Microsoft Agent Framework 1.0 only.
# No A365 SDK (observability, notifications, Purview) — see email_agent_a365.py
# for the full enterprise version.
#
# SDK Stack:
#   Microsoft Agent Framework 1.0   pip install agent-framework
#     - Agent + tool functions
#     - OpenAIChatCompletionClient (Azure Chat Completions)
#
# Architecture:
#   User prompt
#       → Agent Framework Agent.run()
#           → OpenAIChatCompletionClient → Azure OpenAI GPT-5
#               → tool functions (Graph API email fetch)
#   Console output
# ─────────────────────────────────────────────────────────────────────────────

import asyncio
import os
import sys
import logging
from typing import Annotated
from pydantic import Field
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.WARNING,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
log = logging.getLogger("email_agent")
log.setLevel(logging.INFO)

# ── Config ─────────────────────────────────────────────────────────────────────
AZURE_OPENAI_ENDPOINT   = os.getenv("AZURE_OPENAI_ENDPOINT", "https://cdx-ai-foundary.cognitiveservices.azure.com/")
AZURE_OPENAI_DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-5-default")
AZURE_OPENAI_API_VER    = os.getenv("AZURE_OPENAI_API_VERSION", "2025-03-01-preview")

# ── Azure Credentials ──────────────────────────────────────────────────────────
from azure.identity import AzureCliCredential

cli_credential = AzureCliCredential()

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
    Use this when you need the complete content of an email."""
    return get_email_body(email_id=email_id)


def fetch_email_digest(
    count: Annotated[int, Field(description="Number of emails to summarize (default 5, max 10)")] = 5,
) -> str:
    """Fetch and format a digest of the most recent emails.
    Best for overview requests like 'summarize my emails'."""
    return summarize_emails(count=min(count, 10))


# ── Agent Framework 1.0 — build agent ─────────────────────────────────────────

from agent_framework import Agent
from agent_framework.openai import OpenAIChatCompletionClient


def build_agent() -> Agent:
    """
    Build the Agent Framework 1.0 agent.

    Uses OpenAIChatCompletionClient with Azure routing:
      - azure_endpoint + credential → forces Azure routing
      - api_version → Azure OpenAI Chat Completions version
    """
    chat_client = OpenAIChatCompletionClient(
        model=AZURE_OPENAI_DEPLOYMENT,
        azure_endpoint=AZURE_OPENAI_ENDPOINT,
        credential=cli_credential,
        api_version=AZURE_OPENAI_API_VER,
    )

    agent = Agent(
        client=chat_client,
        name="email-summarizer",
        instructions="""You are a helpful email assistant with access to the user's Outlook mailbox.

Available tools:
- fetch_email_digest: Use for overview requests ('summarize my emails', 'what's new')
- fetch_recent_emails: Use to list emails with previews
- fetch_email_body: Use when you need the full content of a specific email

Be concise. For each email highlight: sender, subject, and the key point.""",
        tools=[fetch_recent_emails, fetch_email_body, fetch_email_digest],
    )

    return agent


# ── Main ───────────────────────────────────────────────────────────────────────

async def main():
    agent = build_agent()

    print("\n" + "=" * 55)
    print("  📧 Outlook Email Agent")
    print("  Microsoft Agent Framework 1.0")
    print(f"  Model    : {AZURE_OPENAI_DEPLOYMENT}")
    print(f"  Endpoint : {AZURE_OPENAI_ENDPOINT}")
    print("=" * 55)
    print("  SDK Stack:")
    print("    ✅ Agent Framework 1.0")
    print("       Agent + OpenAIChatCompletionClient + tool functions")
    print("  For full A365 SDK version see: email_agent_a365.py")
    print("=" * 55)
    print("  Commands: 'quit' → exit  |  'clear' → reset session")
    print("=" * 55 + "\n")

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
            # Agent Framework manages session state internally per agent.run() call
            # Each call is stateless by default — clear is a no-op here
            print("🗑️  Ready for new conversation.\n")
            continue

        try:
            print("Agent: ", end="", flush=True)
            response = await agent.run(user_input)
            print(str(response))
            print()
        except Exception as e:
            print(f"\n❌ Error: {e}\n")
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    asyncio.run(main())
