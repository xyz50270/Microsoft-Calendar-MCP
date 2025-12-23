import sys
import os
from datetime import datetime

# Add src to sys.path to import modules
sys.path.append(os.path.join(os.getcwd(), 'src'))

from auth import get_client
from capabilities import email_tools

def run_email_tests():
    print("--- Starting Email Manual Tests ---")
    
    # 1. Authenticate
    try:
        client = get_client()
        if not client.is_authenticated:
            print("Error: Not authenticated. Run src/auth.py first.")
            return
    except Exception as e:
        print(f"Error loading client: {e}")
        return

    # 2. List Emails
    print("\n[Test] Listing Emails...")
    emails = email_tools.list_emails(client, limit=5)
    print(f"Found {len(emails)} emails.")
    for msg in emails:
        print(f" - {msg['subject']} from {msg['sender']}")
    
    input("Press Enter to continue to Sending...")

    # 3. Send Email (to self)
    print("\n[Test] Sending Email to self...")
    # Raw API response format might be different for user data
    user_resp = client.request("GET", "/me")
    current_user = user_resp.json()
    
    if current_user:
        my_email = current_user.get('mail') or current_user.get('userPrincipalName')
        print(f"Sending to: {my_email}")
        
        subject = f"MCP Test Email {datetime.now().isoformat()}"
        body = "This is a test email sent from the MCP service."
        
        result = email_tools.send_email(client, my_email, subject, body)
        print(f"Send Result: {result}")
        
        if result.get("status") == "success":
            print("Email sent successfully. Waiting a few seconds for it to arrive...")
            import time
            time.sleep(5) # Wait for email to arrive
            
            input("Press Enter to continue to Verifying/Deleting...")
            
            # Try to find and delete it
            print("\n[Test] Verifying and Deleting the sent email...")
            recent_emails = email_tools.list_emails(client, limit=5)
            target_id = None
            for msg in recent_emails:
                if msg['subject'] == subject:
                    target_id = msg['id']
                    break
            
            if target_id:
                print(f"Found email to delete. ID: {target_id}")
                del_result = email_tools.delete_email(client, target_id)
                print(f"Delete Result: {del_result}")
            else:
                print("Could not find the sent email to delete (might be delayed).")
        else:
            print(f"Failed to send email: {result.get('message')}")
    else:
        print("Could determine current user email. Skipping send test.")

    input("Press Enter to finish Email tests...")

    print("\n--- Email Tests Completed ---")

if __name__ == "__main__":
    run_email_tests()
