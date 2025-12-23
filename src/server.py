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
from .utils.validation import validate_iso_datetime, validate_email, validate_enum

# Initialize FastMCP server
mcp = FastMCP("Microsoft-365", version="0.1.0")

# Module Toggles (Default to enabled)
def is_enabled(var_name):
    val = os.getenv(var_name, "true").lower()
    return val in ("true", "1", "yes")

ENABLE_CALENDAR = is_enabled("ENABLE_CALENDAR")
ENABLE_TASKS = is_enabled("ENABLE_TASKS")
ENABLE_EMAIL = is_enabled("ENABLE_EMAIL")

# Helper to get authenticated client
def get_authenticated_client():
    client = get_client()
    if not client.is_authenticated:
        raise RuntimeError("账号未认证。请先运行 m365-auth 进行登录。")
    return client

# --- Calendar Tools ---
if ENABLE_CALENDAR:
    @mcp.tool()
    def list_calendar_events(start_date: str = None, end_date: str = None):
        """
        列出用户主日历中的事件。
        [注意] 调用前请务必先执行 `get_current_time` 获取当前时间。
        [时区] 所有日期字符串必须使用本地时间 (UTC+8)。

        参数:
            start_date (str, 可选): 查询范围的开始时间。ISO 8601 格式 (如 '2025-12-23T00:00:00')。必须是本地时间。
            end_date (str, 可选): 查询范围的结束时间。ISO 8601 格式 (如 '2025-12-23T23:59:59')。必须是本地时间。
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
        is_all_day: bool = False,
        importance: str = "normal",
        categories: List[str] = None,
        is_reminder_on: bool = True,
        reminder_minutes: int = 15
    ):
        """
        在主日历中创建新日程 (UTC+8)。

        参数:
            subject (str): 日程标题。
            start (str): 开始时间。ISO 8601 格式 (如 '2025-12-25T09:00:00')。必须是本地时间。
            end (str): 结束时间。ISO 8601 格式 (如 '2025-12-25T10:00:00')。必须是本地时间。
            body (str, 可选): 日程内容描述。
            body_type (str, 可选): 内容类型，可选 'Text' 或 'HTML'。默认为 'HTML'。
            location (str, 可选): 地点名称。
            is_all_day (bool, optional): 是否为全天事件。默认为 False。
            importance (str, optional): 重要程度：'low' (低), 'normal' (普通), 'high' (高)。默认为 'normal'。
            categories (List[str], optional): 关联的分类名称列表。
            is_reminder_on (bool, optional): 是否设置提醒。默认为 True。
            reminder_minutes (int, optional): 开始前多少分钟发出提醒。默认为 15。
        """
        validate_iso_datetime(start, "start")
        validate_iso_datetime(end, "end")
        validate_enum(body_type, ["Text", "HTML"], "body_type")
        validate_enum(importance, ["low", "normal", "high"], "importance")
        
        client = get_authenticated_client()
        return calendar_tools.create_event(
            client, subject, start, end, 
            body=body, body_type=body_type, location=location, 
            is_all_day=is_all_day, 
            importance=importance, 
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
        is_all_day: Optional[bool] = None,
        importance: Optional[str] = None,
        categories: Optional[List[str]] = None,
        is_reminder_on: Optional[bool] = None,
        reminder_minutes: Optional[int] = None
    ):
        """
        更新现有的日历事件 (UTC+8)。仅更新提供的字段。

        参数:
            event_id (str): 待更新事件的唯一 ID。
            subject (str, 可选): 新标题。
            start (str, 可选): 新开始时间 (ISO 8601)。必须是本地时间。
            end (str, 可选): 新结束时间 (ISO 8601)。必须是本地时间。
            body (str, 可选): 新内容描述。
            body_type (str, 可选): 'Text' 或 'HTML'。
            location (str, 可选): 新地点。
            is_all_day (bool, 可选): 是否更新为全天事件。
            importance (str, 可选): 重要程度：'low', 'normal', 'high'。
            categories (List[str], 可选): 新的分类列表。
            is_reminder_on (bool, 可选): 是否开启提醒。
            reminder_minutes (int, 可选): 提醒提前分钟数。
        """
        validate_iso_datetime(start, "start")
        validate_iso_datetime(end, "end")
        validate_enum(body_type, ["Text", "HTML"], "body_type")
        validate_enum(importance, ["low", "normal", "high"], "importance")

        client = get_authenticated_client()
        # Collect provided arguments
        kwargs = {}
        if subject is not None: kwargs['subject'] = subject
        if start is not None: kwargs['start'] = start
        if end is not None: kwargs['end'] = end
        if body is not None: kwargs['body'] = body
        if body_type != "HTML": kwargs['body_type'] = body_type
        if location is not None: kwargs['location'] = location
        if is_all_day is not None: kwargs['is_all_day'] = is_all_day
        if importance is not None: kwargs['importance'] = importance
        if categories is not None: kwargs['categories'] = categories
        if is_reminder_on is not None: kwargs['is_reminder_on'] = is_reminder_on
        if reminder_minutes is not None: kwargs['reminder_minutes'] = reminder_minutes
        
        return calendar_tools.update_event(client, event_id, **kwargs)

    @mcp.tool()
    def delete_calendar_event(event_id: str):
        """
        删除日历事件。

        参数:
            event_id (str): 待删除事件的唯一 ID。
        """
        client = get_authenticated_client()
        return calendar_tools.delete_event(client, event_id)

    @mcp.tool()
    def get_user_schedules(start: str, end: str, availability_view_interval: int = 30):
        """
        [首选] 检查当前用户在特定时间段内是否有空 (UTC+8)。
        当用户询问“我是否有空？”、“是否有冲突？”或“检查我的忙闲”时，请务必【优先】使用此工具而非 list_calendar_events。
        它能更高效、更直观地提供时间段的占用情况。

        参数:
            start (str): 查询范围的开始时间。ISO 8601 格式 (如 '2025-12-23T00:00:00')。必须是本地时间。
            end (str): 查询范围的结束时间。ISO 8601 格式 (如 '2025-12-23T23:59:59')。必须是本地时间。
            availability_view_interval (int, 可选): 响应中每个时间槽的持续分钟数。默认为 30。
        """
        client = get_authenticated_client()
        
        # Always use the current user
        me_info = client.request("GET", "/me").json()
        my_email = me_info.get('mail') or me_info.get('userPrincipalName')
        final_schedules = [my_email] if my_email else ["me"]
                
        validate_iso_datetime(start, "start")
        validate_iso_datetime(end, "end")

        return calendar_tools.get_user_schedules(client, final_schedules, start, end, availability_view_interval)

# --- Tasks Tools ---
if ENABLE_TASKS:
    @mcp.tool()
    def list_tasks():
        """列出用户默认待办事项列表中的任务。"""
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
        """
        在 Microsoft To Do 中创建新任务 (UTC+8)。

        参数:
            title (str): 任务标题。
            body (str, 可选): 任务内容描述。
            body_type (str, 可选): 'text' (纯文本) 或 'html'。默认为 'text'。
            categories (List[str], 可选): 任务关联的分类。
            due_date (str, 可选): 截止日期。ISO 8601 格式 (如 '2025-12-31T23:59:59')。必须是东八区本地时间 (UTC+8)。
            start_date (str, 可选): 开始日期。ISO 8601 格式。必须是东八区本地时间 (UTC+8)。
            reminder_date (str, 可选): 提醒日期/时间。ISO 8601 格式。必须是东八区本地时间 (UTC+8)。
            importance (str, 可选): 重要程度：'low' (低), 'normal' (普通), 'high' (高)。
            status (str, 可选): 任务状态：'notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'。
            completed_date (str, 可选): 任务完成日期。ISO 8601 格式。必须是东八区本地时间 (UTC+8)。
        """
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
        """
        更新 Microsoft To Do 中现有的任务 (UTC+8)。

        参数:
            task_id (str): 待更新任务的唯一 ID。
            title (str, 可选): 新标题。
            body (str, 可选): 新内容描述。
            body_type (str, 可选): 'text' 或 'html'。
            categories (List[str], 可选): 新分类列表。
            due_date (str, 可选): 新截止日期 (ISO 8601)。必须是东八区本地时间 (UTC+8)。
            start_date (str, 可选): 新开始日期 (ISO 8601)。必须是东八区本地时间 (UTC+8)。
            reminder_date (str, 可选): 新提醒日期 (ISO 8601)。必须是东八区本地时间 (UTC+8)。
            importance (str, 可选): 'low', 'normal', 'high'。
            status (str, 可选): 'notStarted', 'inProgress', 'completed' 等。
            completed_date (str, 可选): 新完成日期 (ISO 8601)。必须是东八区本地时间 (UTC+8)。
        """
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
        """
        将任务标记为已完成。

        参数:
            task_id (str): 待完成任务的唯一 ID。
        """
        client = get_authenticated_client()
        return tasks_tools.update_task(client, task_id, completed=True)

    @mcp.tool()
    def delete_task(task_id: str):
        """
        删除任务。

        参数:
            task_id (str): 待删除任务的唯一 ID。
        """
        client = get_authenticated_client()
        return tasks_tools.delete_task(client, task_id)

# --- Email Tools ---
if ENABLE_EMAIL:
    @mcp.tool()
    def list_emails(limit: int = 10):
        """
        列出收件箱中的最近邮件 (UTC+8)。

        参数:
            limit (int, 可选): 返回邮件的最大数量。默认为 10。
        """
        client = get_authenticated_client()
        return email_tools.list_emails(client, limit)

    @mcp.tool()
    def send_email(to: str, subject: str, body: str):
        """
        发送电子邮件。

        参数:
            to (str): 收件人邮箱地址。
            subject (str): 邮件主题。
            body (str): 邮件正文内容。
        """
        validate_email(to, "to")
        client = get_authenticated_client()
        return email_tools.send_email(client, to, subject, body)

    @mcp.tool()
    def delete_email(message_id: str):
        """
        删除电子邮件。

        参数:
            message_id (str): 待删除邮件的唯一 ID。
        """
        client = get_authenticated_client()
        return email_tools.delete_email(client, message_id)

# --- System Tools ---
@mcp.tool()
def get_current_time():
    """获取当前本地时间 (UTC+8)。处理相对时间请求时，请【必须】先调用此工具以获取参考时间。"""
    return system_tools.get_current_time()

# --- Resources ---
@mcp.resource("context://now")
def get_time_resource() -> str:
    """获取当前的系统时间与时区信息 (UTC+8)。"""
    data = system_tools.get_current_time()
    return f"当前本地时间: {data['current_time']}\n时区: {data['timezone']}\n星期: {data['day_of_week']}"

# --- Dynamic Prompt ---
@mcp.prompt("m365-assistant")
def m365_assistant_prompt():
    """在调用本MCP的任何工具或资源前先阅读以下须知内容"""
    # 重新获取最新的启用状态
    cal_enabled = is_enabled("ENABLE_CALENDAR")
    tasks_enabled = is_enabled("ENABLE_TASKS")
    email_enabled = is_enabled("ENABLE_EMAIL")

    enabled_features = []
    if cal_enabled: enabled_features.append("日历 (Calendar)")
    if tasks_enabled: enabled_features.append("待办事项 (To Do)")
    if email_enabled: enabled_features.append("电子邮件 (Outlook)")
    
    features_str = "、".join(enabled_features) if enabled_features else "基础系统"
    
    # 构建核心指令
    instructions = [
        "你是一个专业的 Microsoft 365 办公助手。",
        f"当前启用的功能模块：{features_str}。",
        "",
        "【时区与时间处理核心原则】：",
        "1. 在处理任何与时间相关的请求前，你【必须】先调用 `get_current_time` 工具或读取 `context://now` resource。",
        "2. 本服务器自动处理【本地时间 (东八区 UTC+8)】。",
        "3. 【禁止转换】：请直接使用本地时间进行交互，严禁将时间转换为 UTC 或其他时区。",
        "4. 【格式要求】：所有输入时间必须符合 ISO 8601 格式（例如：2025-12-23T09:00:00）。",
        "",
        "【功能使用指南】："
    ]

    if cal_enabled:
        instructions.append("- 日历：管理日程安排。当用户询问“是否有空”、“是否有冲突”或“查看忙闲”时，【必须首选】使用 `get_user_schedules` 而非 `list_calendar_events`。")
    if tasks_enabled:
        instructions.append("- 待办：管理任务清单。支持设置优先级、截止日期和提醒。")
    if email_enabled:
        instructions.append("- 邮件：处理 Outlook 邮件。支持查询收件箱、发送新邮件和删除邮件。")

    instructions.append("\n请始终以专业、高效、友好的语气为用户提供服务。")
    
    prompt_text = "\n".join(instructions)
    
    # 直接返回字符串内容
    return prompt_text

if __name__ == "__main__":
    mcp.run()