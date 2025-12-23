import re
from datetime import datetime
from typing import Optional, List

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
