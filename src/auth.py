import os
import sys
import ssl
import json
import httpx
import msal
from dotenv import load_dotenv

# Windows OpenSSL Applink 修复
try:
    if sys.platform == "win32":
        ssl.create_default_context()
except Exception:
    pass

# 加载环境变量
load_dotenv()

# 根据环境变量获取动态权限范围的辅助函数
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
        
        # 使用 SerializableTokenCache 进行持久化存储
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
            raise RuntimeError("账号未认证。请先运行 m365-auth。")
        
        url = f"{self.base_url}{endpoint}" if endpoint.startswith('/') else endpoint
        headers = kwargs.pop('headers', {})
        headers['Authorization'] = f"Bearer {token}"
        # 设置默认时区为中国标准时间 (UTC+8)
        headers['Prefer'] = 'outlook.timezone="China Standard Time"'
        
        with httpx.Client() as client:
            try:
                response = client.request(method, url, headers=headers, **kwargs)
                response.raise_for_status()
                return response
            except httpx.HTTPStatusError as e:
                # 尝试解析 Graph API 错误信息
                try:
                    error_data = e.response.json()
                    error_msg = error_data.get('error', {}).get('message', str(e))
                    error_code = error_data.get('error', {}).get('code', 'UnknownError')
                    raise RuntimeError(f"Microsoft Graph API 错误 ({error_code}): {error_msg}")
                except Exception:
                    raise RuntimeError(f"HTTP 错误 {e.response.status_code}: {str(e)}")

    @property
    def is_authenticated(self):
        return self.get_token() is not None

def get_client():
    client_id = os.getenv('MS_GRAPH_CLIENT_ID')
    client_secret = os.getenv('MS_GRAPH_CLIENT_SECRET')
    redirect_uri = os.getenv('MS_GRAPH_REDIRECT_URI')
    token_path = os.getenv('MS_GRAPH_TOKEN_PATH')

    if not client_id:
        print("错误：必须在 .env 文件或环境变量中设置 MS_GRAPH_CLIENT_ID。")
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
    
    # MSAL 交互式流程
    flow = client.app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        # 如果不支持/未配置设备流，则回退到授权码流程
        auth_url = client.app.get_authorization_url(scopes, redirect_uri=client.redirect_uri)
        print(f"请访问此 URL 进行授权：{auth_url}")
        code = input("请输入重定向 URL 中的代码：")
        result = client.app.acquire_token_by_authorization_code(code, scopes=scopes, redirect_uri=client.redirect_uri)
    else:
        print(flow["message"])
        result = client.app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        client._save_cache()
        print("认证成功！")
    else:
        print(f"认证失败：{result.get('error_description')}")

if __name__ == "__main__":
    authenticate_interactive()
