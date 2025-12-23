import os
import ssl
import sys

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

# Helper to get authenticated client
def get_authenticated_client():
    client = get_client()
    if not client.is_authenticated:
        # In a real MCP server context, we expect the token to be present.
        # If not, the user needs to run auth.py first.
        raise RuntimeError("Account not authenticated. Please run src/auth.py first.")
    return client

# --- Calendar Tools ---

@mcp.tool()
def list_calendar_events(start_date: str = None, end_date: str = None):
    """
    List events from the user's primary calendar.
    [MANDATORY] CALL `get_current_time` first.
    TIMEZONE NOTE: `get_current_time` returns LOCAL time. However, this API expects UTC strings.
    You MUST convert the local time to UTC before passing `start_date` and `end_date`.
    :param start_date: ISO format UTC datetime string.
    :param end_date: ISO format UTC datetime string.
    """
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
    attendees: list[str] = None,
    is_all_day: bool = False,
    is_online_meeting: bool = False,
    importance: str = "normal",
    categories: list[str] = None
):
    """
    Create a new event in the primary calendar with all supported properties.
    [MANDATORY] CALL `get_current_time` first.
    TIMEZONE NOTE: Convert local time to UTC before passing `start` and `end`.
    :param subject: Event subject (Required).
    :param start: Start time (ISO format UTC).
    :param end: End time (ISO format UTC).
    :param body: Event description.
    :param body_type: Content type of the body ('Text' or 'HTML'). Defaults to 'HTML'.
    :param location: Event location name.
    :param attendees: List of email addresses for attendees.
    :param is_all_day: Set to True if this is an all-day event.
    :param is_online_meeting: Set to True to make this an online meeting (e.g. Teams).
    :param importance: Importance level ('low', 'normal', 'high').
    :param categories: List of categories for the event.
    """
    client = get_authenticated_client()
    return calendar_tools.create_event(
        client, subject, start, end, 
        body=body, body_type=body_type, location=location, 
        attendees=attendees, is_all_day=is_all_day, 
        is_online_meeting=is_online_meeting, importance=importance, 
        categories=categories
    )

@mcp.tool()
def update_calendar_event(
    event_id: str, 
    subject: str = None, 
    start: str = None, 
    end: str = None, 
    body: str = None,
    body_type: str = "HTML",
    location: str = None,
    attendees: list[str] = None,
    is_all_day: bool = None,
    is_online_meeting: bool = None,
    importance: str = None,
    categories: list[str] = None
):
    """
    Update an existing calendar event with supported properties.
    [MANDATORY] CALL `get_current_time` first.
    TIMEZONE NOTE: Convert local time to UTC if updating `start` or `end`.
    :param event_id: The ID of the event to update (Required).
    :param subject: New event subject.
    :param start: New start time (ISO format UTC).
    :param end: New end time (ISO format UTC).
    :param body: New event description.
    :param body_type: Content type ('Text' or 'HTML').
    :param location: New location name.
    :param attendees: New list of attendee email addresses.
    :param is_all_day: Update all-day status.
    :param is_online_meeting: Update online meeting status.
    :param importance: Update importance ('low', 'normal', 'high').
    :param categories: New list of categories.
    """
    client = get_authenticated_client()
    kwargs = {k: v for k, v in locals().items() if v is not None and k not in ['client', 'event_id']}
    return calendar_tools.update_event(client, event_id, **kwargs)

@mcp.tool()
def delete_calendar_event(event_id: str):
    """
    Delete a calendar event.
    """
    client = get_authenticated_client()
    return calendar_tools.delete_event(client, event_id)

@mcp.tool()
def get_user_schedules(schedules: list[str], start: str, end: str, availability_view_interval: int = 30):
    """
    Get free/busy availability for a collection of users or resources.
    [MANDATORY] CALL `get_current_time` first.
    TIMEZONE NOTE: Convert local time to UTC before passing `start` and `end`.
    :param schedules: List of email addresses to check.
    :param start: Start time (ISO format UTC).
    :param end: End time (ISO format UTC).
    :param availability_view_interval: Optional. Slot size in minutes (e.g. 30, 60).
    """
    client = get_authenticated_client()
    return calendar_tools.get_user_schedules(client, schedules, start, end, availability_view_interval)

# --- Tasks Tools ---

@mcp.tool()
def list_tasks():
    """
    List tasks from the user's default To Do list.
    """
    client = get_authenticated_client()
    return tasks_tools.list_tasks(client)

@mcp.tool()
def create_task(
    title: str, 
    body: str = None, 
    body_type: str = "text",
    categories: list[str] = None,
    due_date: str = None, 
    start_date: str = None,
    reminder_date: str = None, 
    importance: str = None, 
    status: str = None,
    completed_date: str = None
):
    """
    Create a new task in Microsoft To Do with all supported properties.
    [MANDATORY] CALL `get_current_time` first.
    TIMEZONE NOTE: Microsoft Graph API for Tasks expects UTC for all date fields.
    :param title: Task title (Required).
    :param body: Task description/content.
    :param body_type: Content type of the body ('text' or 'html'). Defaults to 'text'.
    :param categories: List of categories for the task.
    :param due_date: Due date and time (ISO format UTC).
    :param start_date: Start date and time (ISO format UTC).
    :param reminder_date: Reminder date and time (ISO format UTC).
    :param importance: Importance level ('low', 'normal', 'high').
    :param status: Task status ('notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred').
    :param completed_date: Date and time when the task was completed (ISO format UTC).
    """
    client = get_authenticated_client()
    return tasks_tools.create_task(
        client, 
        title, 
        body=body, 
        body_type=body_type,
        categories=categories,
        due_date=due_date, 
        start_date=start_date,
        reminder_date=reminder_date, 
        importance=importance, 
        status=status,
        completed_date=completed_date
    )

@mcp.tool()
def update_task(
    task_id: str,
    title: str = None, 
    body: str = None, 
    body_type: str = "text",
    categories: list[str] = None,
    due_date: str = None, 
    start_date: str = None,
    reminder_date: str = None, 
    importance: str = None, 
    status: str = None,
    completed_date: str = None
):
    """
    Update an existing task in Microsoft To Do.
    [MANDATORY] CALL `get_current_time` first.
    TIMEZONE NOTE: Convert local time to UTC before passing date parameters.
    :param task_id: The ID of the task to update (Required).
    :param title: New task title.
    :param body: New task description.
    :param body_type: Content type of the body ('text' or 'html').
    :param categories: New list of categories.
    :param due_date: New due date (ISO format UTC).
    :param start_date: New start date (ISO format UTC).
    :param reminder_date: New reminder date (ISO format UTC).
    :param importance: Importance level ('low', 'normal', 'high').
    :param status: Task status ('notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred').
    :param completed_date: Completion date (ISO format UTC).
    """
    client = get_authenticated_client()
    kwargs = {k: v for k, v in locals().items() if v is not None and k not in ['client', 'task_id']}
    return tasks_tools.update_task(client, task_id, **kwargs)

@mcp.tool()
def complete_task(task_id: str):
    """
    Mark a task as completed.
    """
    client = get_authenticated_client()
    return tasks_tools.update_task(client, task_id, completed=True)

@mcp.tool()
def delete_task(task_id: str):
    """
    Delete a task.
    """
    client = get_authenticated_client()
    return tasks_tools.delete_task(client, task_id)

# --- Email Tools ---

@mcp.tool()
def list_emails(limit: int = 10):
    """
    List recent emails from the inbox.
    [MANDATORY] CALL `get_current_time` first.
    TIMEZONE NOTE: `get_current_time` returns LOCAL time. Use this to orient yourself, but remember that Graph API timestamp filters usually use UTC.
    """
    client = get_authenticated_client()
    return email_tools.list_emails(client, limit)

@mcp.tool()
def send_email(to: str, subject: str, body: str):
    """
    Send an email.
    """
    client = get_authenticated_client()
    return email_tools.send_email(client, to, subject, body)

@mcp.tool()
def delete_email(message_id: str):
    """
    Delete an email.
    """
    client = get_authenticated_client()
    return email_tools.delete_email(client, message_id)

# --- System Tools ---

@mcp.tool()
def get_current_time():
    """
    Get the current system time and timezone. 
    MANDATORY: Call this tool BEFORE any other tools if the user uses relative time.
    NOTE: This returns the user's CURRENT LOCAL TIME. 
    Microsoft Graph API calls generally require UTC. You are responsible for converting this local time to UTC when calling other tools.
    """
    return system_tools.get_current_time()

# --- Resources ---

@mcp.resource("context://now")
def get_time_resource() -> str:
    """
    CURRENT_SYSTEM_TIME: READ THIS or call get_current_time for Today's Date and Time.
    NOTE: This is LOCAL time. Microsoft Graph API requires UTC.
    """
    data = system_tools.get_current_time()
    return f"Current Local Time: {data['current_time']}\nTimezone: {data['timezone']}\nDay: {data['day_of_week']}"

# --- Prompts ---

@mcp.prompt("m365-assistant")
def m365_assistant_prompt():
    """
    M365 Assistant: MANDATORY TIMEZONE CONVERSION REQUIRED.
    Local time from `get_current_time` MUST be converted to UTC for Graph API tools.
    """
    return """You are a professional Microsoft 365 productivity assistant using the Microsoft Graph API. 

CRITICAL TIMEZONE INSTRUCTION:
1. You do NOT know the current time. You MUST call `get_current_time` or read `context://now` first.
2. `get_current_time` returns the user's LOCAL time and timezone (e.g., Asia/Shanghai).
3. The Microsoft Graph API (Calendar, Tasks, etc.) expects time strings in UTC.
4. You MUST manually calculate the UTC time based on the local time and timezone offset provided before calling any tool that accepts a date or time.

Example: If local time is 2025-12-22 09:00 AM in Asia/Shanghai (UTC+8), you must pass 2025-12-22T01:00:00 to the tool.

指令：你是一个专业的 Microsoft 365 办公助手（基于 Microsoft Graph API）。
时区处理重要提示：
1. 你必须先通过 `get_current_time` 或 `context://now` 获取时间。
2. `get_current_time` 返回的是用户的【本地时间】和【时区】。
3. Microsoft Graph API 所有的工具（日历、待办等）都接收【UTC 时间】。
4. 你必须根据获取到的本地时间和时区偏移量，自行计算出对应的 UTC 时间，然后再调用相关工具。
严禁直接将本地时间作为参数传递给日历或待办工具。"""

if __name__ == "__main__":
    mcp.run()
