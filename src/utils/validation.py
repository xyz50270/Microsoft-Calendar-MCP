import re
from datetime import datetime
from typing import Optional, List

def validate_iso_datetime(dt_str: Optional[str], name: str):
    if not dt_str:
        return
    try:
        # Check if it follows a basic ISO-like pattern first to give better error
        if not re.match(r"^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2})?.*$", dt_str):
            raise ValueError(f"参数 '{name}' 必须是 ISO 格式的日期字符串 (例如：YYYY-MM-DD 或 YYYY-MM-DDTHH:MM:SS)。收到值: {dt_str}")
        datetime.fromisoformat(dt_str.replace('Z', '+00:00'))
    except ValueError as e:
        raise ValueError(f"参数 '{name}' 的 ISO 日期格式无效: {str(e)}")

def validate_email(email: str, name: str):
    if not re.match(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$", email):
        raise ValueError(f"参数 '{name}' 的邮箱地址格式无效: {email}")

def validate_enum(value: Optional[str], valid_values: List[str], name: str):
    if value is not None and value.lower() not in [v.lower() for v in valid_values]:
        raise ValueError(f"参数 '{name}' 的值无效。必须是 {valid_values} 之一。收到值: {value}")
