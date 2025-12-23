def list_events(client, start_date=None, end_date=None):
    """List events from the primary calendar."""
    from datetime import datetime, timedelta
    
    if not start_date:
        start_date = datetime.now().isoformat()
    if not end_date:
        end_date = (datetime.fromisoformat(start_date) + timedelta(days=7)).isoformat()
    
    # Use calendarView for expanded recurring events
    endpoint = f"/me/calendar/calendarView?startDateTime={start_date}&endDateTime={end_date}"
    response = client.request("GET", endpoint)
    data = response.json()
    
    return [
        {
            "id": event.get("id"),
            "subject": event.get("subject"),
            "start": event.get("start", {}).get("dateTime"),
            "end": event.get("end", {}).get("dateTime"),
            "location": event.get("location", {}).get("displayName"),
            "body": event.get("bodyPreview"),
            "attendees": [a.get("emailAddress", {}).get("address") for a in event.get("attendees", [])]
        }
        for event in data.get("value", [])
    ]

def create_event(client, subject, start, end, body=None, body_type="HTML", location=None, attendees=None, is_all_day=False, is_online_meeting=False, importance="normal", categories=None):
    """
    Create a new event with all supported properties.
    :param attendees: List of email addresses.
    """
    payload = {
        "subject": subject,
        "start": {"dateTime": start, "timeZone": "China Standard Time"},
        "end": {"dateTime": end, "timeZone": "China Standard Time"},
        "isAllDay": is_all_day,
        "isOnlineMeeting": is_online_meeting,
        "importance": importance
    }
    
    if body:
        payload["body"] = {"contentType": body_type, "content": body}
    if location:
        payload["location"] = {"displayName": location}
    if categories:
        payload["categories"] = categories
    if attendees:
        payload["attendees"] = [
            {"emailAddress": {"address": addr}, "type": "required"} for addr in attendees
        ]
        
    response = client.request("POST", "/me/events", json=payload)
    data = response.json()
    return {"status": "success", "id": data.get("id")}

def update_event(client, event_id, **kwargs):
    """Update an existing event."""
    payload = {}
    
    # Mapping tool arguments to Graph API fields
    field_map = {
        'subject': 'subject',
        'is_all_day': 'isAllDay',
        'is_online_meeting': 'isOnlineMeeting',
        'importance': 'importance',
        'categories': 'categories'
    }
    
    for arg, api_field in field_map.items():
        if arg in kwargs:
            payload[api_field] = kwargs[arg]
            
    if 'start' in kwargs:
        payload["start"] = {"dateTime": kwargs['start'], "timeZone": "China Standard Time"}
    if 'end' in kwargs:
        payload["end"] = {"dateTime": kwargs['end'], "timeZone": "China Standard Time"}
    if 'location' in kwargs:
        payload["location"] = {"displayName": kwargs['location']}
    if 'body' in kwargs:
        payload["body"] = {"contentType": kwargs.get('body_type', 'HTML'), "content": kwargs['body']}
    if 'attendees' in kwargs:
        payload["attendees"] = [
            {"emailAddress": {"address": addr}, "type": "required"} for addr in kwargs['attendees']
        ]
        
    if not payload:
        return {"status": "error", "message": "No fields to update provided"}
        
    client.request("PATCH", f"/me/events/{event_id}", json=payload)
    return {"status": "success"}

def delete_event(client, event_id):
    """Delete an event."""
    client.request("DELETE", f"/me/events/{event_id}")
    return {"status": "success"}

def get_user_schedules(client, schedules, start, end, availability_view_interval=30):
    """
    Get free/busy availability for a collection of users.
    :param schedules: List of email addresses.
    :param start: ISO format UTC datetime string.
    :param end: ISO format UTC datetime string.
    :param availability_view_interval: Duration of each time slot in minutes.
    """
    payload = {
        "schedules": schedules,
        "startTime": {"dateTime": start, "timeZone": "China Standard Time"},
        "endTime": {"dateTime": end, "timeZone": "China Standard Time"},
        "availabilityViewInterval": availability_view_interval
    }
    
    response = client.request("POST", "/me/calendar/getSchedule", json=payload)
    return response.json()
