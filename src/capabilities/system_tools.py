from datetime import datetime
import tzlocal

def get_current_time():
    """Returns the current local time and timezone."""
    now = datetime.now()
    local_tz = tzlocal.get_localzone_name()
    return {
        "current_time": now.isoformat(),
        "timezone": local_tz,
        "day_of_week": now.strftime("%A")
    }
