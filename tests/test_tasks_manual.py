import sys
import os
from datetime import datetime, timedelta

# Add src to sys.path to import modules
sys.path.append(os.path.join(os.getcwd(), 'src'))

from auth import get_client
from capabilities import tasks_tools

def run_tasks_tests():
    print("--- Starting Tasks Manual Tests ---")
    
    # 1. Authenticate
    try:
        client = get_client()
        if not client.is_authenticated:
            print("Error: Not authenticated. Run src/auth.py first.")
            return
    except Exception as e:
        print(f"Error loading client: {e}")
        return

    # 2. Create Task
    print("\n[Test] Creating Task with multiple properties...")
    title = "MCP Comprehensive Test Task"
    body = "This is a detailed test task with importance and reminder."
    due_date = (datetime.now() + timedelta(days=1)).replace(hour=17, minute=0, second=0).isoformat()
    reminder_date = (datetime.now() + timedelta(days=1)).replace(hour=9, minute=0, second=0).isoformat()
    
    result = tasks_tools.create_task(
        client, 
        title, 
        due_date=due_date, 
        body=body, 
        importance='high', 
        reminder_date=reminder_date
    )
    print(f"Create Result: {result}")
    input("Press Enter to continue to Listing...")
    
    if result.get("status") != "success":
        print("Failed to create task. Aborting tests.")
        return
        
    task_id = result["id"]
    print(f"Task ID: {task_id}")

    # 3. List Tasks
    print("\n[Test] Listing Tasks...")
    tasks = tasks_tools.list_tasks(client)
    found = False
    for task in tasks:
        if task["id"] == task_id:
            print(f"Found created task: {task['title']} (Due: {task['due']})")
            found = True
            break
    
    input("Press Enter to continue to Completion...")

    if not found:
        print("Error: Created task not found in list.")
    else:
        print("Verification successful.")

    # 4. Update Task
    print("\n[Test] Updating Task (Title, Body, Importance)...")
    new_title = "MCP Comprehensive Test Task (UPDATED)"
    new_body = "This body has been updated via PATCH request."
    update_result = tasks_tools.update_task(
        client, 
        task_id, 
        title=new_title, 
        body=new_body, 
        importance='low'
    )
    print(f"Update Result: {update_result}")
    
    # Verify update
    tasks_after_update = tasks_tools.list_tasks(client)
    updated_found = False
    for task in tasks_after_update:
        if task["id"] == task_id:
            print(f"Verified update: Title is now '{task['title']}', Importance is '{task['importance']}'")
            updated_found = True
            break
            
    input("Press Enter to continue to Deletion...")

    # 5. Delete Task
    print("\n[Test] Deleting Task...")
    delete_result = tasks_tools.delete_task(client, task_id)
    print(f"Delete Result: {delete_result}")
    
    # Verify deletion
    try:
        # Try to update it - should fail
        check_result = tasks_tools.update_task(client, task_id, title="Should Fail")
        if check_result.get("status") == "error":
             print("Verification successful: Task no longer accessible.")
        else:
             print("Error: Task still exists.")
    except Exception as e:
        print(f"Verification successful (Expected failure): {e}")

    input("Press Enter to finish Tasks tests...")

    print("\n--- Tasks Tests Completed ---")

if __name__ == "__main__":
    run_tasks_tests()
