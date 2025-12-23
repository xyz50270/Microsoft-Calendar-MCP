def list_emails(client, limit=10):
    """列出最近的邮件。"""
    response = client.request("GET", f"/me/messages?$top={limit}")
    data = response.json()
    
    return [
        {
            "id": msg.get("id"),
            "subject": msg.get("subject"),
            "sender": msg.get("from", {}).get("emailAddress", {}).get("address"),
            "received": msg.get("receivedDateTime"),
            "body_preview": msg.get("bodyPreview")
        }
        for msg in data.get("value", [])
    ]

def send_email(client, to_recipients, subject, body):
    """发送电子邮件。"""
    if isinstance(to_recipients, str):
        to_recipients = [to_recipients]
        
    payload = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "Text",
                "content": body
            },
            "toRecipients": [
                {"emailAddress": {"address": addr}} for addr in to_recipients
            ]
        }
    }
    
    client.request("POST", "/me/sendMail", json=payload)
    return {"status": "success"}

def delete_email(client, message_id):
    """删除邮件。"""
    client.request("DELETE", f"/me/messages/{message_id}")
    return {"status": "success"}

def move_email(client, message_id, folder_id):
    """将邮件移动到指定文件夹。"""
    payload = {"destinationId": folder_id}
    client.request("POST", f"/me/messages/{message_id}/move", json=payload)
    return {"status": "success"}