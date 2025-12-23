import os
import sys
import ssl
import json
import httpx
import msal
from dotenv import load_dotenv

# Windows OpenSSL Applink Fix
try:
    if sys.platform == "win32":
        ssl.create_default_context()
except Exception:
    pass

# Load environment variables
load_dotenv()

# Helper to get dynamic scopes based on environment variables
def get_scopes():
    scopes = ['User.Read']
    if os.getenv("ENABLE_CALENDAR", "true").lower() in ("true", "1", "yes"):
        scopes.append('Calendars.ReadWrite')
    if os.getenv("ENABLE_TASKS", "true").lower() in ("true", "1", "yes"):
        scopes.append('Tasks.ReadWrite')
    if os.getenv("ENABLE_EMAIL", "true").lower() in ("true", "1", "yes"):
        scopes.append('Mail.ReadWrite')
        scopes.append('Mail.Send')
    return scopes

class GraphClient:
    def __init__(self, client_id, client_secret=None, redirect_uri=None, token_path=None):
        self.client_id = client_id
        self.client_secret = client_secret
        self.redirect_uri = redirect_uri or 'https://login.microsoftonline.com/common/oauth2/nativeclient'
        self.token_path = token_path or 'graph_token.json'
        self.authority = "https://login.microsoftonline.com/common"
        self.base_url = "https://graph.microsoft.com/v1.0"
        
        # Use SerializableTokenCache for persistence
        self._token_cache = msal.SerializableTokenCache()
        self._load_cache()

        if client_secret:
            self.app = msal.ConfidentialClientApplication(
                client_id, 
                client_credential=client_secret, 
                authority=self.authority,
                token_cache=self._token_cache
            )
        else:
            self.app = msal.PublicClientApplication(
                client_id, 
                authority=self.authority,
                token_cache=self._token_cache
            )

    def _load_cache(self):
        if os.path.exists(self.token_path):
            with open(self.token_path, 'r') as f:
                try:
                    self._token_cache.deserialize(f.read())
                except:
                    pass

    def _save_cache(self):
        if self._token_cache.has_state_changed:
            with open(self.token_path, 'w') as f:
                f.write(self._token_cache.serialize())

    def get_token(self):
        accounts = self.app.get_accounts()
        result = None
        scopes = get_scopes()
        if accounts:
            result = self.app.acquire_token_silent(scopes, account=accounts[0])
        
        if not result:
            return None
        
        self._save_cache()
        return result.get("access_token")

    def request(self, method, endpoint, **kwargs):
        token = self.get_token()
        if not token:
            raise RuntimeError("Not authenticated. Please run m365-auth first.")
        
        url = f"{self.base_url}{endpoint}" if endpoint.startswith('/') else endpoint
        headers = kwargs.pop('headers', {})
        headers['Authorization'] = f"Bearer {token}"
        # Set default timezone to China Standard Time (UTC+8)
        headers['Prefer'] = 'outlook.timezone="China Standard Time"'
        
        with httpx.Client() as client:
            try:
                response = client.request(method, url, headers=headers, **kwargs)
                response.raise_for_status()
                return response
            except httpx.HTTPStatusError as e:
                # Attempt to parse Graph API error message
                try:
                    error_data = e.response.json()
                    error_msg = error_data.get('error', {}).get('message', str(e))
                    error_code = error_data.get('error', {}).get('code', 'UnknownError')
                    raise RuntimeError(f"Microsoft Graph API Error ({error_code}): {error_msg}")
                except Exception:
                    raise RuntimeError(f"HTTP Error {e.response.status_code}: {str(e)}")

    @property
    def is_authenticated(self):
        return self.get_token() is not None

def get_client():
    client_id = os.getenv('MS_GRAPH_CLIENT_ID')
    client_secret = os.getenv('MS_GRAPH_CLIENT_SECRET')
    redirect_uri = os.getenv('MS_GRAPH_REDIRECT_URI')
    token_path = os.getenv('MS_GRAPH_TOKEN_PATH')

    if not client_id:
        print("Error: MS_GRAPH_CLIENT_ID must be set in .env file or environment variables.")
        sys.exit(1)
    
    return GraphClient(
        client_id=client_id,
        client_secret=client_secret if client_secret else None,
        redirect_uri=redirect_uri,
        token_path=token_path if token_path else 'graph_token.json'
    )

def authenticate_interactive():
    client = get_client()
    scopes = get_scopes()
    
    # MSAL Interactive flow
    flow = client.app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        # Fallback to auth code flow if device flow is not supported/configured
        auth_url = client.app.get_authorization_url(scopes, redirect_uri=client.redirect_uri)
        print(f"Please visit this URL to authorize: {auth_url}")
        code = input("Enter the code from the redirect URL: ")
        result = client.app.acquire_token_by_authorization_code(code, scopes=scopes, redirect_uri=client.redirect_uri)
    else:
        print(flow["message"])
        result = client.app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        client._save_cache()
        print("Authentication successful!")
    else:
        print(f"Authentication failed: {result.get('error_description')}")

if __name__ == "__main__":
    authenticate_interactive()
