def list_events(client, start_date=None, end_date=None):
    """列出主日历中的事件。"""
    from datetime import datetime, timedelta
    
    if not start_date:
        start_date = datetime.now().isoformat()
    if not end_date:
        end_date = (datetime.fromisoformat(start_date) + timedelta(days=7)).isoformat()
    
    # 使用 calendarView 以获取展开后的循环事件
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
            "body": event.get("bodyPreview")
        }
        for event in data.get("value", [])
    ]

def create_event(client, subject, start, end, body=None, body_type="HTML", location=None, is_all_day=False, importance="normal", categories=None, is_reminder_on=True, reminder_minutes=15):
    """
    创建具有支持属性的新日程。
    """
    payload = {
        "subject": subject,
        "start": {"dateTime": start, "timeZone": "China Standard Time"},
        "end": {"dateTime": end, "timeZone": "China Standard Time"},
        "isAllDay": is_all_day,
        "importance": importance,
        "isReminderOn": is_reminder_on,
        "reminderMinutesBeforeStart": reminder_minutes
    }
    
    if body:
        payload["body"] = {"contentType": body_type, "content": body}
    if location:
        payload["location"] = {"displayName": location}
    if categories:
        payload["categories"] = categories
        
    response = client.request("POST", "/me/events", json=payload)
    data = response.json()
    return {"status": "success", "id": data.get("id")}

def update_event(client, event_id, **kwargs):
    """更新现有日程。"""
    payload = {}
    
    # 将工具参数映射到 Graph API 字段
    field_map = {
        'subject': 'subject',
        'is_all_day': 'isAllDay',
        'importance': 'importance',
        'categories': 'categories',
        'is_reminder_on': 'isReminderOn',
        'reminder_minutes': 'reminderMinutesBeforeStart'
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
        
    if not payload:
        return {"status": "error", "message": "未提供需要更新的字段"}
        
    client.request("PATCH", f"/me/events/{event_id}", json=payload)
    return {"status": "success"}

def delete_event(client, event_id):
    """删除日程。"""
    client.request("DELETE", f"/me/events/{event_id}")
    return {"status": "success"}

def get_user_schedules(client, schedules, start, end, availability_view_interval=30):
    """
    获取一组用户的忙闲日程。
    :param schedules: 邮箱地址列表。
    :param start: ISO 格式的本地时间字符串。
    :param end: ISO 格式的本地时间字符串。
    :param availability_view_interval: 每个时间槽的分钟数。
    """
    payload = {
        "schedules": schedules,
        "startTime": {"dateTime": start, "timeZone": "China Standard Time"},
        "endTime": {"dateTime": end, "timeZone": "China Standard Time"},
        "availabilityViewInterval": availability_view_interval
    }
    
    response = client.request("POST", "/me/calendar/getSchedule", json=payload)
    return response.json()
