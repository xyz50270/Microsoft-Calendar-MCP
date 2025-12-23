import os
import ssl
import sys
import re
from datetime import datetime
from typing import Optional, List

# Windows OpenSSL Applink Fix
try:
    if sys.platform == "win32":
        ssl.create_default_context()
except Exception:
    pass

from fastmcp import FastMCP
from .auth import get_client
from .capabilities import calendar_tools, tasks_tools, email_tools, system_tools

# Initialize FastMCP server
mcp = FastMCP("Microsoft-365")

# Module Toggles (Default to enabled)
def is_enabled(var_name):
    val = os.getenv(var_name, "true").lower()
    return val in ("true", "1", "yes")

ENABLE_CALENDAR = is_enabled("ENABLE_CALENDAR")
ENABLE_TASKS = is_enabled("ENABLE_TASKS")
ENABLE_EMAIL = is_enabled("ENABLE_EMAIL")

# --- Validation Helpers ---

def validate_iso_datetime(dt_str: Optional[str], name: str):
    if not dt_str:
        return
    try:
        # Check if it follows a basic ISO-like pattern first to give better error
        if not re.match(r"^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2})?.*$", dt_str):
            raise ValueError(f"Parameter '{name}' must be an ISO format date string (e.g., YYYY-MM-DD or YYYY-MM-DDTHH:MM:SS). Received: {dt_str}")
        datetime.fromisoformat(dt_str.replace('Z', '+00:00'))
    except ValueError as e:
        raise ValueError(f"Invalid ISO datetime for '{name}': {str(e)}")

def validate_email(email: str, name: str):
    if not re.match(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$", email):
        raise ValueError(f"Invalid email address for '{name}': {email}")

def validate_enum(value: Optional[str], valid_values: List[str], name: str):
    if value is not None and value.lower() not in [v.lower() for v in valid_values]:
        raise ValueError(f"Invalid value for '{name}'. Must be one of {valid_values}. Received: {value}")

# Helper to get authenticated client
def get_authenticated_client():
    client = get_client()
    if not client.is_authenticated:
        raise RuntimeError("Account not authenticated. Please run m365-auth first.")
    return client

# --- Calendar Tools ---
if ENABLE_CALENDAR:
    @mcp.tool()
    def list_calendar_events(start_date: str = None, end_date: str = None):
        """
        List events from the user's primary calendar.
        [MANDATORY] CALL `get_current_time` first.
        TIMEZONE NOTE: All date strings should be in LOCAL time (UTC+8).
        """
        validate_iso_datetime(start_date, "start_date")
        validate_iso_datetime(end_date, "end_date")
        client = get_authenticated_client()
        return calendar_tools.list_events(client, start_date, end_date)

    @mcp.tool()
    def create_calendar_event(
        subject: str, 
        start: str, 
        end: str, 
        body: str = None, 
        body_type: str = "HTML",
        location: str = None, 
        attendees: List[str] = None,
        is_all_day: bool = False,
        is_online_meeting: bool = False,
        importance: str = "normal",
        categories: List[str] = None,
        is_reminder_on: bool = True,
        reminder_minutes: int = 15
    ):
        """Create a new event in the primary calendar (UTC+8)."""
        validate_iso_datetime(start, "start")
        validate_iso_datetime(end, "end")
        validate_enum(body_type, ["Text", "HTML"], "body_type")
        validate_enum(importance, ["low", "normal", "high"], "importance")
        if attendees:
            for addr in attendees:
                validate_email(addr, "attendees")
        
        client = get_authenticated_client()
        return calendar_tools.create_event(
            client, subject, start, end, 
            body=body, body_type=body_type, location=location, 
            attendees=attendees, is_all_day=is_all_day, 
            is_online_meeting=is_online_meeting, importance=importance, 
            categories=categories, is_reminder_on=is_reminder_on, 
            reminder_minutes=reminder_minutes
        )

    @mcp.tool()
    def update_calendar_event(
        event_id: str, 
        subject: Optional[str] = None, 
        start: Optional[str] = None, 
        end: Optional[str] = None, 
        body: Optional[str] = None,
        body_type: str = "HTML",
        location: Optional[str] = None,
        attendees: Optional[List[str]] = None,
        is_all_day: Optional[bool] = None,
        is_online_meeting: Optional[bool] = None,
        importance: Optional[str] = None,
        categories: Optional[List[str]] = None,
        is_reminder_on: Optional[bool] = None,
        reminder_minutes: Optional[int] = None
    ):
        """Update an existing calendar event (UTC+8)."""
        validate_iso_datetime(start, "start")
        validate_iso_datetime(end, "end")
        validate_enum(body_type, ["Text", "HTML"], "body_type")
        validate_enum(importance, ["low", "normal", "high"], "importance")
        if attendees:
            for addr in attendees:
                validate_email(addr, "attendees")

        client = get_authenticated_client()
        # Collect provided arguments
        kwargs = {}
        if subject is not None: kwargs['subject'] = subject
        if start is not None: kwargs['start'] = start
        if end is not None: kwargs['end'] = end
        if body is not None: kwargs['body'] = body
        if body_type != "HTML": kwargs['body_type'] = body_type
        if location is not None: kwargs['location'] = location
        if attendees is not None: kwargs['attendees'] = attendees
        if is_all_day is not None: kwargs['is_all_day'] = is_all_day
        if is_online_meeting is not None: kwargs['is_online_meeting'] = is_online_meeting
        if importance is not None: kwargs['importance'] = importance
        if categories is not None: kwargs['categories'] = categories
        if is_reminder_on is not None: kwargs['is_reminder_on'] = is_reminder_on
        if reminder_minutes is not None: kwargs['reminder_minutes'] = reminder_minutes
        
        return calendar_tools.update_event(client, event_id, **kwargs)

    @mcp.tool()
    def delete_calendar_event(event_id: str):
        """Delete a calendar event."""
        client = get_authenticated_client()
        return calendar_tools.delete_event(client, event_id)

    @mcp.tool()
    def get_user_schedules(schedules: List[str], start: str, end: str, availability_view_interval: int = 30):
        """Get free/busy availability (UTC+8)."""
        for addr in schedules:
            validate_email(addr, "schedules")
        validate_iso_datetime(start, "start")
        validate_iso_datetime(end, "end")

        client = get_authenticated_client()
        return calendar_tools.get_user_schedules(client, schedules, start, end, availability_view_interval)

# --- Tasks Tools ---
if ENABLE_TASKS:
    @mcp.tool()
    def list_tasks():
        """List tasks from the user's default To Do list."""
        client = get_authenticated_client()
        return tasks_tools.list_tasks(client)

    @mcp.tool()
    def create_task(
        title: str, 
        body: Optional[str] = None, 
        body_type: str = "text",
        categories: Optional[List[str]] = None,
        due_date: Optional[str] = None, 
        start_date: Optional[str] = None,
        reminder_date: Optional[str] = None, 
        importance: Optional[str] = None, 
        status: Optional[str] = None,
        completed_date: Optional[str] = None
    ):
        """Create a new task in Microsoft To Do (UTC+8)."""
        validate_enum(body_type, ["text", "html"], "body_type")
        validate_iso_datetime(due_date, "due_date")
        validate_iso_datetime(start_date, "start_date")
        validate_iso_datetime(reminder_date, "reminder_date")
        validate_iso_datetime(completed_date, "completed_date")
        validate_enum(importance, ["low", "normal", "high"], "importance")
        validate_enum(status, ["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"], "status")

        client = get_authenticated_client()
        return tasks_tools.create_task(
            client, title, body=body, body_type=body_type, 
            categories=categories, due_date=due_date, start_date=start_date,
            reminder_date=reminder_date, importance=importance, 
            status=status, completed_date=completed_date
        )

    @mcp.tool()
    def update_task(
        task_id: str,
        title: Optional[str] = None, 
        body: Optional[str] = None, 
        body_type: str = "text",
        categories: Optional[List[str]] = None,
        due_date: Optional[str] = None, 
        start_date: Optional[str] = None,
        reminder_date: Optional[str] = None, 
        importance: Optional[str] = None, 
        status: Optional[str] = None,
        completed_date: Optional[str] = None
    ):
        """Update an existing task in Microsoft To Do (UTC+8)."""
        validate_enum(body_type, ["text", "html"], "body_type")
        validate_iso_datetime(due_date, "due_date")
        validate_iso_datetime(start_date, "start_date")
        validate_iso_datetime(reminder_date, "reminder_date")
        validate_iso_datetime(completed_date, "completed_date")
        validate_enum(importance, ["low", "normal", "high"], "importance")
        validate_enum(status, ["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"], "status")

        client = get_authenticated_client()
        # Collect provided arguments
        kwargs = {}
        if title is not None: kwargs['title'] = title
        if body is not None: kwargs['body'] = body
        if body_type != "text": kwargs['body_type'] = body_type
        if categories is not None: kwargs['categories'] = categories
        if due_date is not None: kwargs['due_date'] = due_date
        if start_date is not None: kwargs['start_date'] = start_date
        if reminder_date is not None: kwargs['reminder_date'] = reminder_date
        if importance is not None: kwargs['importance'] = importance
        if status is not None: kwargs['status'] = status
        if completed_date is not None: kwargs['completed_date'] = completed_date

        return tasks_tools.update_task(client, task_id, **kwargs)

    @mcp.tool()
    def complete_task(task_id: str):
        """Mark a task as completed."""
        client = get_authenticated_client()
        return tasks_tools.update_task(client, task_id, completed=True)

    @mcp.tool()
    def delete_task(task_id: str):
        """Delete a task."""
        client = get_authenticated_client()
        return tasks_tools.delete_task(client, task_id)

# --- Email Tools ---
if ENABLE_EMAIL:
    @mcp.tool()
    def list_emails(limit: int = 10):
        """List recent emails from the inbox (UTC+8)."""
        client = get_authenticated_client()
        return email_tools.list_emails(client, limit)

    @mcp.tool()
    def send_email(to: str, subject: str, body: str):
        """Send an email."""
        validate_email(to, "to")
        client = get_authenticated_client()
        return email_tools.send_email(client, to, subject, body)

    @mcp.tool()
    def delete_email(message_id: str):
        """Delete an email."""
        client = get_authenticated_client()
        return email_tools.delete_email(client, message_id)

# --- System Tools ---
@mcp.tool()
def get_current_time():
    """Get current local time (UTC+8). MANDATORY for relative time handling."""
    return system_tools.get_current_time()

# --- Resources ---
@mcp.resource("context://now")
def get_time_resource() -> str:
    """Current system time and timezone (UTC+8)."""
    data = system_tools.get_current_time()
    return f"Current Local Time: {data['current_time']}\nTimezone: {data['timezone']}\nDay: {data['day_of_week']}"

# --- Dynamic Prompt ---
@mcp.prompt("m365-assistant")
def m365_assistant_prompt():
    enabled_features = []
    if ENABLE_CALENDAR: enabled_features.append("Calendar (日历)")
    if ENABLE_TASKS: enabled_features.append("To Do Tasks (待办)")
    if ENABLE_EMAIL: enabled_features.append("Email (邮件)")
    
    features_str = ", ".join(enabled_features)
    
    prompt = f"""You are a professional Microsoft 365 productivity assistant.
Current enabled features: {features_str}.

TIMEZONE HANDLING:
1. You MUST call `get_current_time` or read `context://now` first.
2. The server handles LOCAL time (UTC+8) automatically. Do NOT perform UTC conversions.

指令：你是一个专业的 Microsoft 365 办公助手任务。
当前启用的模块：{features_str}。

时区处理提示：
1. 你必须先通过 `get_current_time` 或 `context://now` 获取时间。
2. 系统自动处理本地时间 (UTC+8)，严禁进行手动的 UTC 转换。"""

    if ENABLE_CALENDAR:
        prompt += "\n- 你可以管理日历事件，支持查看、创建、更新和删除日程。"
    if ENABLE_TASKS:
        prompt += "\n- 你可以管理待办事项，支持任务的完整生命周期管理。"
    if ENABLE_EMAIL:
        prompt += "\n- 你可以管理电子邮件，支持收件箱列表查询、发送和删除邮件。"
        
    return prompt

if __name__ == "__main__":
    mcp.run()