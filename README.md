# Microsoft 365 MCP Service

这是一个基于 **Model Context Protocol (MCP)** 标准开发的 Microsoft 365 助手服务。它允许大语言模型（LLM）通过直接调用 **Microsoft Graph API** 来管理您的日历（Calendar）、待办事项（To Do）和电子邮件（Outlook）。

## 核心特性

- **全功能集成**:
  - **日历 (Calendar)**: 增删改查日程、支持在线会议 (Teams) 创建、空闲/繁忙状态查询。
  - **待办 (Tasks)**: 全属性支持（优先级、截止日期、提醒、正文类型、分类等）。
  - **邮件 (Email)**: 发送邮件、列表查询、删除及移动操作。
- **架构优势**:
  - **直接交互**: 舍弃了第三方库封装，直接与 Microsoft Graph REST API 通信，确保全字段支持。
  - **时区感知**: 内置 `context://now` 资源和 `get_current_time` 工具，强制 LLM 进行时区转换以确保时间准确。
  - **安全认证**: 使用 `msal` 实现 OAuth2 流程，支持令牌自动刷新。
- **开发者友好**: 提供完整的交互式认证脚本和功能测试工具。

## 快速开始

### 1. 环境准备
- Python 3.10+
- [Azure App Registration](https://portal.azure.com/):
  - 创建一个 "Mobile and desktop applications" 类型的应用。
  - 添加重定向 URI: `https://login.microsoftonline.com/common/oauth2/nativeclient`。
  - 在 API Permissions 中添加: `User.Read`, `Calendars.ReadWrite`, `Tasks.ReadWrite`, `Mail.ReadWrite`, `Mail.Send`。

### 2. 本地安装与认证
```bash
# 安装项目
pip install -e .

# 完成 OAuth2 认证
m365-auth
```

## 作为 MCP 服务运行

您可以将此服务接入支持 MCP 的客户端（如 Claude Desktop）。在您的配置文件中添加以下内容：

### 方式 A: 远程启动 (推荐)
无需手动下载代码，直接通过 `uv` 运行。

```json
{
  "mcpServers": {
    "m365": {
      "command": "uv",
      "args": [
        "run",
        "--with", "git+https://github.com/xyz50270/Microsoft-Calendar-MCP.git",
        "m365-mcp"
      ],
      "env": {
        "MS_GRAPH_CLIENT_ID": "您的客户端ID",
        "MS_GRAPH_TOKEN_PATH": "C:/path/to/your/graph_token.json"
      }
    }
  }
}
```

### 方式 B: 本地启动
如果您已克隆仓库并在本地完成了认证：

```json
{
  "mcpServers": {
    "m365": {
      "command": "uv",
      "args": [
        "--directory", "E:/开发工具/Microsoft-Calendar-MCP",
        "run",
        "m365-mcp"
      ]
    }
  }
}
```

## 可用工具 (Tools)

- **系统**: `get_current_time` (必须先调用以处理相对时间)。
- **日历**: `list_calendar_events`, `create_calendar_event`, `update_calendar_event`, `delete_calendar_event`, `get_user_schedules`。
- **待办**: `list_tasks`, `create_task`, `update_task`, `delete_task`。
- **邮件**: `list_emails`, `send_email`, `delete_email`。

## 可用资源 (Resources)

- `context://now`: 提供当前系统本地时间、时区和星期。

## 测试

项目包含了一系列手动测试脚本，用于验证各模块功能：
- `python tests/test_calendar_manual.py`
- `python tests/test_tasks_manual.py`
- `python tests/test_email_manual.py`
- `python tests/test_get_schedule.py`

## 开源协议
MIT
