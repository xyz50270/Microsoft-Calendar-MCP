import sys
import os
from datetime import datetime, timedelta

# Add src to sys.path to import modules
sys.path.append(os.path.join(os.getcwd(), 'src'))

from auth import get_client
from capabilities import calendar_tools

def run_calendar_tests():
    print("--- Starting Calendar Manual Tests ---")
    
    # 1. Authenticate
    try:
        client = get_client()
        if not client.is_authenticated:
            print("Error: Not authenticated. Run src/auth.py first.")
            return
    except Exception as e:
        print(f"Error loading client: {e}")
        return

    # 2. Create Event
    print("\n[Test] Creating Event with multiple properties...")
    start_time = (datetime.now() + timedelta(hours=1)).replace(microsecond=0).isoformat()
    end_time = (datetime.now() + timedelta(hours=2)).replace(microsecond=0).isoformat()
    subject = "MCP Advanced Test Event"
    body = "<b>Important</b> meeting regarding P402B."
    
    result = calendar_tools.create_event(
        client, subject, start_time, end_time, 
        location="Meeting Room A", 
        body=body, 
        body_type="HTML",
        importance="high"
    )
    print(f"Create Result: {result}")
    print(f"Create Result: {result}")
    input("Press Enter to continue to Listing...")
    
    if result.get("status") != "success":
        print("Failed to create event. Aborting tests.")
        return
        
    event_id = result["id"]
    print(f"Event ID: {event_id}")

    # 3. List Events
    print("\n[Test] Listing Events...")
    # List events for today/tomorrow to ensure we catch the new one
    list_start = (datetime.now() - timedelta(days=1)).isoformat()
    list_end = (datetime.now() + timedelta(days=2)).isoformat()
    
    print(f"Debug: Listing from {list_start} to {list_end}")
    
    events = calendar_tools.list_events(client, list_start, list_end)
    print(f"Debug: Found {len(events)} events.")
    found = False
    for event in events:
        print(f"Debug: Checking event {event['id']} - {event['subject']} ({event['start']})")
        if event["id"] == event_id:
            print(f"Found created event: {event['subject']} at {event['start']}")
            found = True
            break
    
    input("Press Enter to continue to Update...")
    
    if not found:
        print("Error: Created event not found in list.")
    else:
        print("Verification successful.")

    # 4. Update Event
    print("\n[Test] Updating Event...")
    new_subject = "MCP Test Event (Updated)"
    update_result = calendar_tools.update_event(client, event_id, subject=new_subject)
    print(f"Update Result: {update_result}")
    
    # Verify update
    events_after_update = calendar_tools.list_events(client, list_start, list_end)
    updated_found = False
    for event in events_after_update:
        if event["id"] == event_id and event["subject"] == new_subject:
            print(f"Verified update: Subject is now '{event['subject']}'")
            updated_found = True
            break
            
    input("Press Enter to continue to Deletion...")
            
    if not updated_found:
        print("Error: Event update verification failed.")

    # 5. Delete Event
    print("\n[Test] Deleting Event...")
    delete_result = calendar_tools.delete_event(client, event_id)
    print(f"Delete Result: {delete_result}")
    
    # Verify deletion
    events_after_delete = calendar_tools.list_events(client, list_start, list_end)
    deleted_found = False
    for event in events_after_delete:
        if event["id"] == event_id:
            deleted_found = True
            break
            
    if not deleted_found:
        print("Verification successful: Event no longer exists.")
    else:
        print("Error: Event still exists after deletion.")

    input("Press Enter to finish Calendar tests...")

    print("\n--- Calendar Tests Completed ---")

if __name__ == "__main__":
    run_calendar_tests()
