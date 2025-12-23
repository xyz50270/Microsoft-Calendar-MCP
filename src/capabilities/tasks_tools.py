def _get_default_todo_list_id(client):
    response = client.request("GET", "/me/todo/lists")
    lists = response.json().get("value", [])
    for lst in lists:
        if lst.get("wellKnownName") == "defaultList":
            return lst.get("id")
    return lists[0].get("id") if lists else None

def list_tasks(client):
    """List tasks from the default To Do list."""
    list_id = _get_default_todo_list_id(client)
    if not list_id:
        return []
    
    response = client.request("GET", f"/me/todo/lists/{list_id}/tasks")
    data = response.json()
    
    return [
        {
            "id": task.get("id"),
            "title": task.get("title"),
            "status": task.get("status"),
            "due": task.get("dueDateTime", {}).get("dateTime"),
            "importance": task.get("importance"),
            "is_completed": task.get("status") == "completed"
        }
        for task in data.get("value", [])
    ]

def create_task(client, title, body=None, body_type="text", categories=None, due_date=None, start_date=None, reminder_date=None, importance=None, status=None, completed_date=None):
    """
    Create a new task with all supported Microsoft Graph API properties.
    """
    list_id = _get_default_todo_list_id(client)
    if not list_id:
        return {"status": "error", "message": "No default todo list found"}

    # Timezone for DateTimeTimeZone objects
    tz = "UTC" 
    
    payload = {"title": title}
    
    if body:
        payload["body"] = {"content": body, "contentType": body_type}
    
    if categories:
        payload["categories"] = categories
        
    if due_date:
        payload["dueDateTime"] = {"dateTime": due_date, "timeZone": tz}
        
    if start_date:
        payload["startDateTime"] = {"dateTime": start_date, "timeZone": tz}
        
    if reminder_date:
        payload["reminderDateTime"] = {"dateTime": reminder_date, "timeZone": tz}
        payload["isReminderOn"] = True
        
    if importance:
        # low, normal, high
        payload["importance"] = importance.lower()
        
    if status:
        # notStarted, inProgress, completed, waitingOnOthers, deferred
        payload["status"] = status
        
    if completed_date:
        payload["completedDateTime"] = {"dateTime": completed_date, "timeZone": tz}
        
    response = client.request("POST", f"/me/todo/lists/{list_id}/tasks", json=payload)
    data = response.json()
    return {"status": "success", "id": data.get("id")}

def update_task(client, task_id, **kwargs):
    """
    Update an existing task with supported Microsoft Graph API properties.
    """
    list_id = _get_default_todo_list_id(client)
    if not list_id:
        return {"status": "error", "message": "No default todo list found"}

    tz = "UTC"
    payload = {}
    
    if 'title' in kwargs:
        payload["title"] = kwargs['title']
    if 'body' in kwargs:
        payload["body"] = {"content": kwargs['body'], "contentType": kwargs.get('body_type', 'text')}
    if 'due_date' in kwargs:
        payload["dueDateTime"] = {"dateTime": kwargs['due_date'], "timeZone": tz}
    if 'start_date' in kwargs:
        payload["startDateTime"] = {"dateTime": kwargs['start_date'], "timeZone": tz}
    if 'reminder_date' in kwargs:
        payload["reminderDateTime"] = {"dateTime": kwargs['reminder_date'], "timeZone": tz}
        payload["isReminderOn"] = True
    if 'importance' in kwargs:
        payload["importance"] = kwargs['importance'].lower()
    if 'status' in kwargs:
        payload["status"] = kwargs['status']
    if 'completed_date' in kwargs:
        payload["completedDateTime"] = {"dateTime": kwargs['completed_date'], "timeZone": tz}
    if 'categories' in kwargs:
        payload["categories"] = kwargs['categories']
    if 'completed' in kwargs: # Legacy helper
        payload["status"] = "completed" if kwargs['completed'] else "notStarted"
        
    if not payload:
        return {"status": "error", "message": "No fields to update provided"}

    client.request("PATCH", f"/me/todo/lists/{list_id}/tasks/{task_id}", json=payload)
    return {"status": "success"}

def delete_task(client, task_id):
    """Delete a task."""
    list_id = _get_default_todo_list_id(client)
    client.request("DELETE", f"/me/todo/lists/{list_id}/tasks/{task_id}")
    return {"status": "success"}