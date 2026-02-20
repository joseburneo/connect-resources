import requests
import os

class InstantlyClient:
    BASE_URL = "https://api.instantly.ai/api/v2"

    def __init__(self, api_key):
        self.api_key = api_key
        self.headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        }

    def request(self, method, path, params=None, json=None):
        """
        Makes a request to the Instantly API.
        
        Args:
            method (str): HTTP method (GET, POST, etc.)
            path (str): API endpoint path (e.g., '/campaigns')
            params (dict, optional): Query parameters.
            json (dict, optional): JSON body.

        Returns:
            dict or str: Parsed JSON response if available, else raw text.

        Raises:
            requests.exceptions.HTTPError: If the response status code is not 2xx.
        """
        url = f"{self.BASE_URL}{path}"
        
        try:
            response = requests.request(
                method=method,
                url=url,
                headers=self.headers,
                params=params,
                json=json,
                timeout=30 
            )
            response.raise_for_status()
            
            try:
                return response.json()
            except ValueError:
                return response.text
                
        except requests.exceptions.HTTPError as e:
            # Enhance error message with response text if available
            error_msg = f"HTTP Error: {e}"
            if e.response is not None:
                error_msg += f" | Body: {e.response.text}"
            raise requests.exceptions.HTTPError(error_msg, response=e.response) from e
        except requests.exceptions.RequestException as e:
            raise RuntimeError(f"Network error occurred: {e}") from e

    def get_campaigns(self):
        """Helper to fetch campaigns."""
        return self.request("GET", "/campaigns")
        
def load_clients():
    """
    Dynamically loads all Instantly clients from environment variables.
    Looks for keys starting with INSTANTLY_API_KEY_
    """
    clients = {}
    
    # Load all keys starting with INSTANTLY_API_KEY_
    for key, val in os.environ.items():
        if key.startswith("INSTANTLY_API_KEY_"):
            # Extract client name from key (e.g. INSTANTLY_API_KEY_LUXVANCE -> Luxvance)
            client_name_raw = key.replace("INSTANTLY_API_KEY_", "")
            
            # Formatting: LUXVANCE -> Luxvance, GLOBAL_FOOD_VENTURES -> Global Food Ventures
            client_name = client_name_raw.replace("_", " ").title()
            
            # Special manual fixups
            if client_name.upper() == "CAMB AI": client_name = "CAMB.ai"
            if client_name.upper() == "KCAL": client_name = "Kcal" 
            if client_name.upper() == "CAPQUEST": client_name = "CapQuest"
            if client_name.upper() == "INSURANCE MARKET": client_name = "Insurance Market"
            
            clients[client_name] = InstantlyClient(val)
            
    return clients
