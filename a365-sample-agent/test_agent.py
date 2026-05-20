"""Simple test script to verify the agent works via direct HTTP POST"""
import requests
import json

# Bot Framework Activity format
activity = {
    "type": "message",
    "id": "test-123",
    "timestamp": "2026-05-19T06:30:00.000Z",
    "channelId": "emulator",
    "from": {
        "id": "user-test-123",
        "name": "Test User"
    },
    "conversation": {
        "id": "test-conversation-123"
    },
    "recipient": {
        "id": "bot-123",
        "name": "AgentFrameworkAgent",
        "tenant_id": "test-tenant",
        "agentic_app_id": "test-agent"
    },
    "text": "Hello, can you help me?",
    "serviceUrl": "http://localhost:56150"
}

print("Sending test message to agent...")
print(f"Message: {activity['text']}")

try:
    response = requests.post(
        "http://localhost:3979/api/messages",
        json=activity,
        headers={"Content-Type": "application/json"},
        timeout=30
    )
    
    print(f"\nStatus Code: {response.status_code}")
    print(f"Response: {response.text}")
    
    if response.status_code == 200:
        print("\n✅ Agent responded successfully!")
    else:
        print(f"\n❌ Agent returned error: {response.status_code}")
        
except Exception as e:
    print(f"\n❌ Error: {e}")
