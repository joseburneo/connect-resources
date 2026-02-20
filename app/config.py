import os
from dotenv import load_dotenv

load_dotenv()

def get_all_api_keys():
    """Retrieves all Instantly API keys from environment variables."""
    keys = {}
    for key, val in os.environ.items():
        if key.startswith("INSTANTLY_API_KEY_"):
            # Extract client name from var name, e.g. INSTANTLY_API_KEY_LUXVANCE -> LUXVANCE
            client_name = key.replace("INSTANTLY_API_KEY_", "").replace("_", " ").title()
            keys[client_name] = val
            
    if not keys:
        raise ValueError("No environment variables starting with INSTANTLY_API_KEY_ found.")
    return keys

def get_notion_api_key():
    """Retrieves the Notion API key from environment variables."""
    key = os.getenv("NOTION_API_KEY")
    if not key:
        raise ValueError("Environment variable NOTION_API_KEY is not set.")
    return key

def get_notion_database_id():
    """Retrieves the Notion Database ID from environment variables."""
    val = os.getenv("NOTION_DATABASE_ID")
    if not val:
        raise ValueError("Environment variable NOTION_DATABASE_ID is not set.")
    return val
