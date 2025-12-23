import sys
import os
from datetime import datetime, timedelta

# 将 src 目录添加到 sys.path
sys.path.append(os.path.join(os.getcwd(), 'src'))

from auth import get_client
from capabilities import calendar_tools
from utils.validation import validate_email, validate_iso_datetime

def verify_tool_logic():
    print("=== 验证 get_user_schedules 工具修复情况 ===")
    
    # 1. 获取认证客户端
    try:
        client = get_client()
        if not client.is_authenticated:
            print("❌ 错误: 未认证。请先运行 'uv run m365-auth'。")
            return
    except Exception as e:
        print(f"❌ 认证初始化失败: {e}")
        return

    # 2. 准备测试参数 (模拟大模型的输入)
    schedules_input = ["me"]
    start_time = datetime.utcnow().isoformat()
    end_time = (datetime.utcnow() + timedelta(days=1)).isoformat()
    
    print(f"输入参数: schedules={schedules_input}, start={start_time}")

    # 3. 执行工具内部的解析逻辑 (我们在 src/server.py 中添加的逻辑)
    try:
        print("正在解析 'me' 标识符...")
        my_email = None
        final_schedules = []
        for addr in schedules_input:
            if addr.lower() == "me":
                if not my_email:
                    me_info = client.request("GET", "/me").json()
                    my_email = me_info.get('mail') or me_info.get('userPrincipalName')
                
                if my_email:
                    print(f"✅ 'me' 成功解析为: {my_email}")
                    final_schedules.append(my_email)
                else:
                    print("⚠️ 未能获取用户邮箱，回退到 'me'")
                    final_schedules.append(addr)
            else:
                validate_email(addr, "schedules")
                final_schedules.append(addr)
        
        # 验证解析后的列表
        assert len(final_schedules) > 0
        assert final_schedules[0] != "me"
        
        # 4. 调用实际的能力层工具
        print("正在调用 Graph API 获取日程...")
        result = calendar_tools.get_user_schedules(client, final_schedules, start_time, end_time)
        
        if 'value' in result:
            print("✅ 成功! 已收到日程信息。")
            for item in result['value']:
                schedule_id = item.get('scheduleId')
                availability = item.get('availabilityView', '')
                print(f"   日程 ID: {schedule_id}")
                print(f"   空闲状态预览: {availability[:50]}...")
        else:
            print(f"❌ 响应异常: {result}")

    except Exception as e:
        print(f"❌ 测试过程中发生异常: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    verify_tool_logic()
