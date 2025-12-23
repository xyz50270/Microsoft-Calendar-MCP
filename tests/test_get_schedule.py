import sys
import os
from datetime import datetime, timedelta

# Add src to sys.path to import modules
sys.path.append(os.path.join(os.getcwd(), 'src'))

from auth import get_client
from capabilities import calendar_tools

def test_get_schedule():
    print("--- Starting GetSchedule Test ---")
    client = get_client()
    
    # Get current user email
    me = client.request("GET", "/me").json()
    my_email = me.get('mail') or me.get('userPrincipalName')
    
    start = datetime.utcnow().isoformat()
    end = (datetime.utcnow() + timedelta(days=1)).isoformat()
    
    print(f"Checking availability for: {my_email}")
    result = calendar_tools.get_user_schedules(client, [my_email], start, end)
    
    # print(result)
    if 'value' in result:
        print("Success! Received schedule information.")
        for item in result['value']:
            print(f"Schedule ID: {item.get('scheduleId')}")
            print(f"Availability View: {item.get('availabilityView')[:50]}...")
    else:
        print(f"Error or unexpected response: {result}")

if __name__ == "__main__":
    test_get_schedule()
