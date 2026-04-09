# agent.py - Console email agent using Azure OpenAI + direct tool calls
import os
import sys
import json
from openai import AzureOpenAI
from azure.identity import DefaultAzureCredential, get_bearer_token_provider
from dotenv import load_dotenv

# Load .env from same directory as this script
load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

# Import tool functions directly from mcp_server
# (no MCP protocol needed — just call the Python functions)
sys.path.insert(0, os.path.dirname(__file__))
from outlook_mcp_server import get_recent_emails, get_email_body, summarize_emails

ENDPOINT   = os.getenv("AZURE_OPENAI_ENDPOINT", "https://cdx-ai-foundary.cognitiveservices.azure.com/")
DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-5-default")

# ── Azure OpenAI client ────────────────────────────────────────────────────────

token_provider = get_bearer_token_provider(
    DefaultAzureCredential(),
    "https://cognitiveservices.azure.com/.default"
)

client = AzureOpenAI(
    azure_endpoint=ENDPOINT,
    azure_ad_token_provider=token_provider,
    api_version="2025-01-01-preview"
)

# ── Tool registry ──────────────────────────────────────────────────────────────

TOOL_REGISTRY = {
    "get_recent_emails": get_recent_emails,
    "get_email_body":    get_email_body,
    "summarize_emails":  summarize_emails,
}

def call_tool(tool_name: str, arguments: dict) -> str:
    """Call a tool function directly."""
    fn = TOOL_REGISTRY.get(tool_name)
    if not fn:
        return f"Unknown tool: {tool_name}"
    try:
        return fn(**arguments)
    except Exception as e:
        return f"Tool error ({tool_name}): {e}"

# ── Tool definitions for OpenAI function calling ───────────────────────────────

TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "get_recent_emails",
            "description": "Fetch the most recent emails from the user's Outlook mailbox. Returns subject, sender, date, and preview.",
            "parameters": {
                "type": "object",
                "properties": {
                    "count": {
                        "type": "integer",
                        "description": "Number of emails to fetch (default 10, max 25)",
                        "default": 10
                    }
                },
                "required": []
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_email_body",
            "description": "Fetch the full body of a specific email by its ID. Use when you need complete email content.",
            "parameters": {
                "type": "object",
                "properties": {
                    "email_id": {
                        "type": "string",
                        "description": "The email ID returned from get_recent_emails"
                    }
                },
                "required": ["email_id"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "summarize_emails",
            "description": "Fetch and summarize the most recent emails as a digest. Best for overview requests.",
            "parameters": {
                "type": "object",
                "properties": {
                    "count": {
                        "type": "integer",
                        "description": "Number of emails to summarize (default 5, max 10)",
                        "default": 5
                    }
                },
                "required": []
            }
        }
    }
]

SYSTEM_PROMPT = """You are a helpful email assistant with access to the user's Outlook mailbox.
You can fetch recent emails, get full email content, and provide summaries.
Be concise and clear. When summarizing emails, highlight the key information:
sender, subject, and main point of each email.
If asked about a specific email, use get_email_body to get the full content."""

# ── Agent loop ─────────────────────────────────────────────────────────────────

def run_agent(user_input: str, conversation_history: list) -> str:
    """Run one turn of the agent loop with tool calling."""

    conversation_history.append({
        "role": "user",
        "content": user_input
    })

    while True:
        response = client.chat.completions.create(
            model=DEPLOYMENT,
            messages=[{"role": "system", "content": SYSTEM_PROMPT}] + conversation_history,
            tools=TOOLS,
            tool_choice="auto"
        )

        message = response.choices[0].message
        finish_reason = response.choices[0].finish_reason

        conversation_history.append(message.model_dump(exclude_none=True))

        # Final answer
        if finish_reason == "stop" or not message.tool_calls:
            return message.content

        # Process tool calls
        print(f"\n  🔧 Calling:", end="")
        tool_results = []

        for tool_call in message.tool_calls:
            tool_name = tool_call.function.name
            try:
                arguments = json.loads(tool_call.function.arguments)
            except json.JSONDecodeError:
                arguments = {}

            print(f" [{tool_name}]", end="", flush=True)
            result = call_tool(tool_name, arguments)

            tool_results.append({
                "role": "tool",
                "tool_call_id": tool_call.id,
                "content": result
            })

        print()
        conversation_history.extend(tool_results)

# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print("\n" + "=" * 50)
    print("  📧 Email Assistant Agent")
    print(f"  Model    : {DEPLOYMENT}")
    print(f"  Endpoint : {ENDPOINT}")
    print("=" * 50)
    print("  'clear' → reset conversation")
    print("  'quit'  → exit")
    print("=" * 50 + "\n")

    conversation_history = []

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
            conversation_history = []
            print("🗑️  Conversation cleared.\n")
            continue

        try:
            print("Agent: ", end="", flush=True)
            answer = run_agent(user_input, conversation_history)
            print(answer)
            print()
        except Exception as e:
            print(f"\n❌ Error: {e}\n")
            import traceback
            traceback.print_exc()

if __name__ == "__main__":
    main()