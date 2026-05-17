import os
import sys
import json
import httpx
import msal
from datetime import datetime, timedelta
from dotenv import load_dotenv

# Load .env file from the same directory as the script
script_dir = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(script_dir, ".env"))

class M365Client:
    def __init__(self):
        self.client_id = os.getenv('MS_GRAPH_CLIENT_ID')
        # Use absolute path for token to avoid WinError 3
        default_token_path = os.path.join(os.path.dirname(script_dir), "graph_token.json")
        self.token_path = os.getenv('MS_GRAPH_TOKEN_PATH', default_token_path)
        if not os.path.isabs(self.token_path):
             self.token_path = os.path.join(script_dir, self.token_path)
             
        self.authority = "https://login.microsoftonline.com/common"
        self.base_url = "https://graph.microsoft.com/v1.0"
        
        self._token_cache = msal.SerializableTokenCache()
        if os.path.exists(self.token_path):
            with open(self.token_path, 'r') as f:
                try:
                    self._token_cache.deserialize(f.read())
                except:
                    pass
        
        self.app = msal.PublicClientApplication(
            self.client_id, 
            authority=self.authority, 
            token_cache=self._token_cache
        )

    def get_token(self):
        if not self.client_id:
            raise RuntimeError("MS_GRAPH_CLIENT_ID not set in environment.")
        
        scopes = ['User.Read', 'Calendars.ReadWrite', 'Tasks.ReadWrite', 'Mail.ReadWrite', 'Mail.Send']
        accounts = self.app.get_accounts()
        result = None
        if accounts:
            result = self.app.acquire_token_silent(scopes, account=accounts[0])
        
        if not result:
            return None
            
        if self._token_cache.has_state_changed:
            dir_name = os.path.dirname(self.token_path)
            if dir_name:
                os.makedirs(dir_name, exist_ok=True)
            with open(self.token_path, 'w') as f:
                f.write(self._token_cache.serialize())
        return result.get("access_token")

    def request(self, method, endpoint, **kwargs):
        token = self.get_token()
        if not token:
            raise RuntimeError("Authentication required. Run: python m365_client.py auth")
        
        url = f"{self.base_url}{endpoint}" if endpoint.startswith('/') else endpoint
        headers = kwargs.pop('headers', {})
        headers['Authorization'] = f"Bearer {token}"
        headers['Prefer'] = 'outlook.timezone="China Standard Time"'
        
        with httpx.Client() as client:
            resp = client.request(method, url, headers=headers, **kwargs)
            resp.raise_for_status()
            return resp.json() if resp.content else None

def main():
    if len(sys.argv) < 2:
        print("Usage: python m365_client.py <command> [args]")
        sys.exit(1)

    client = M365Client()
    cmd = sys.argv[1]

    try:
        # --- AUTH ---
        if cmd == "auth":
            scopes = ['User.Read', 'Calendars.ReadWrite', 'Tasks.ReadWrite', 'Mail.ReadWrite', 'Mail.Send']
            flow = client.app.initiate_device_flow(scopes=scopes)
            if "user_code" not in flow:
                print("Could not initiate device flow.")
                sys.exit(1)
            print(flow["message"])
            result = client.app.acquire_token_by_device_flow(flow)
            if "access_token" in result:
                dir_name = os.path.dirname(client.token_path)
                if dir_name:
                    os.makedirs(dir_name, exist_ok=True)
                with open(client.token_path, 'w') as f:
                    f.write(client._token_cache.serialize())
                print("Successfully authenticated.")
            else:
                print(f"Authentication failed: {result.get('error_description')}")

        # --- CALENDAR ---
        elif cmd == "calendar-list":
            start = sys.argv[2] if len(sys.argv) > 2 else datetime.now().isoformat()
            end = sys.argv[3] if len(sys.argv) > 3 else (datetime.fromisoformat(start[:19]) + timedelta(days=7)).isoformat()
            res = client.request("GET", f"/me/calendar/calendarView?startDateTime={start}&endDateTime={end}")
            events = []
            for item in res.get("value", []):
                events.append({
                    "id": item.get("id"),
                    "subject": item.get("subject"),
                    "start": item.get("start", {}).get("dateTime"),
                    "end": item.get("end", {}).get("dateTime"),
                    "location": item.get("location", {}).get("displayName")
                })
            print(json.dumps(events, indent=2, ensure_ascii=False))

        elif cmd == "calendar-create":
            subject, start, end = sys.argv[2:5]
            location = sys.argv[5] if len(sys.argv) > 5 else None
            payload = {
                "subject": subject,
                "start": {"dateTime": start, "timeZone": "China Standard Time"},
                "end": {"dateTime": end, "timeZone": "China Standard Time"},
                "location": {"displayName": location} if location else None
            }
            res = client.request("POST", "/me/events", json=payload)
            print(json.dumps({"status": "success", "id": res.get("id")}))

        elif cmd == "calendar-update":
            event_id = sys.argv[2]
            updates = json.loads(sys.argv[3]) # Pass as JSON string
            payload = {}
            if 'subject' in updates: payload['subject'] = updates['subject']
            if 'start' in updates: payload['start'] = {"dateTime": updates['start'], "timeZone": "China Standard Time"}
            if 'end' in updates: payload['end'] = {"dateTime": updates['end'], "timeZone": "China Standard Time"}
            if 'location' in updates: payload['location'] = {"displayName": updates['location']}
            
            client.request("PATCH", f"/me/events/{event_id}", json=payload)
            print(json.dumps({"status": "success"}))

        elif cmd == "calendar-delete":
            event_id = sys.argv[2]
            client.request("DELETE", f"/me/events/{event_id}")
            print(json.dumps({"status": "success"}))

        elif cmd == "calendar-freebusy":
            emails = sys.argv[2].split(',')
            start, end = sys.argv[3:5]
            payload = {
                "schedules": emails,
                "startTime": {"dateTime": start, "timeZone": "China Standard Time"},
                "endTime": {"dateTime": end, "timeZone": "China Standard Time"},
                "availabilityViewInterval": 30
            }
            res = client.request("POST", "/me/calendar/getSchedule", json=payload)
            print(json.dumps(res, indent=2, ensure_ascii=False))

        # --- TASKS ---
        elif cmd == "tasks-list":
            lists = client.request("GET", "/me/todo/lists")
            list_id = next((l["id"] for l in lists.get("value", []) if l.get("wellKnownName") == "defaultList"), None) or lists["value"][0]["id"]
            res = client.request("GET", f"/me/todo/lists/{list_id}/tasks")
            tasks = [{"id": t.get("id"), "title": t.get("title"), "status": t.get("status")} for t in res.get("value", [])]
            print(json.dumps(tasks, indent=2, ensure_ascii=False))

        elif cmd == "tasks-create":
            title = sys.argv[2]
            due_date = sys.argv[3] if len(sys.argv) > 3 else None
            lists = client.request("GET", "/me/todo/lists")
            list_id = next((l["id"] for l in lists.get("value", []) if l.get("wellKnownName") == "defaultList"), None) or lists["value"][0]["id"]
            payload = {"title": title}
            if due_date: payload["dueDateTime"] = {"dateTime": due_date, "timeZone": "China Standard Time"}
            res = client.request("POST", f"/me/todo/lists/{list_id}/tasks", json=payload)
            print(json.dumps({"status": "success", "id": res.get("id")}))

        elif cmd == "tasks-update":
            task_id = sys.argv[2]
            updates = json.loads(sys.argv[3])
            lists = client.request("GET", "/me/todo/lists")
            list_id = next((l["id"] for l in lists.get("value", []) if l.get("wellKnownName") == "defaultList"), None) or lists["value"][0]["id"]
            payload = {}
            if 'title' in updates: payload['title'] = updates['title']
            if 'status' in updates: payload['status'] = updates['status']
            if 'due_date' in updates: payload['dueDateTime'] = {"dateTime": updates['due_date'], "timeZone": "China Standard Time"}
            client.request("PATCH", f"/me/todo/lists/{list_id}/tasks/{task_id}", json=payload)
            print(json.dumps({"status": "success"}))

        elif cmd == "tasks-delete":
            task_id = sys.argv[2]
            lists = client.request("GET", "/me/todo/lists")
            list_id = next((l["id"] for l in lists.get("value", []) if l.get("wellKnownName") == "defaultList"), None) or lists["value"][0]["id"]
            client.request("DELETE", f"/me/todo/lists/{list_id}/tasks/{task_id}")
            print(json.dumps({"status": "success"}))

        # --- MAIL ---
        elif cmd == "mail-list":
            limit = sys.argv[2] if len(sys.argv) > 2 else 10
            res = client.request("GET", f"/me/messages?$top={limit}")
            mails = [{"id": m.get("id"), "subject": m.get("subject"), "from": m.get("from", {}).get("emailAddress", {}).get("address")} for m in res.get("value", [])]
            print(json.dumps(mails, indent=2, ensure_ascii=False))

        elif cmd == "mail-send":
            to_recipients = sys.argv[2].split(',')
            subject = sys.argv[3]
            body = sys.argv[4]
            payload = {
                "message": {
                    "subject": subject,
                    "body": {"contentType": "Text", "content": body},
                    "toRecipients": [{"emailAddress": {"address": addr}} for addr in to_recipients]
                }
            }
            client.request("POST", "/me/sendMail", json=payload)
            print(json.dumps({"status": "success"}))

        elif cmd == "mail-delete":
            msg_id = sys.argv[2]
            client.request("DELETE", f"/me/messages/{msg_id}")
            print(json.dumps({"status": "success"}))

        elif cmd == "mail-move":
            msg_id, folder_id = sys.argv[2:4]
            client.request("POST", f"/me/messages/{msg_id}/move", json={"destinationId": folder_id})
            print(json.dumps({"status": "success"}))

        else:
            print(f"Unknown command: {cmd}")
            sys.exit(1)

    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
