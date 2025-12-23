# Microsoft 365 MCP Service

这是一个基于 **Model Context Protocol (MCP)** 标准开发的 Microsoft 365 助手服务。它允许大语言模型（LLM）通过直接调用 **Microsoft Graph API** 来管理您的日历（Calendar）、待办事项（To Do）和电子邮件（Outlook）。

## 核心特性

- **全功能集成**:
  - **日历 (Calendar)**: 增删改查日程、支持在线会议 (Teams) 创建、空闲/繁忙状态查询、**自定义提醒 (Reminders)**。
  - **待办 (Tasks)**: 全属性支持（优先级、截止日期、提醒、状态管理等）。
  - **邮件 (Email)**: 发送邮件、列表查询、删除操作。
- **架构优势**:
  - **自动时区处理**: 服务器自动处理 **东八区 (UTC+8)** 时间。LLM 无需进行复杂的 UTC 转换，直接使用本地时间交互。
  - **模块化控制**: 可通过环境变量独立启用/禁用日历、待办或邮件模块，系统提示词将根据启用模块动态调整。
  - **直接交互**: 舍弃第三方库封装，直接与 REST API 通信，支持全字段配置。
  - **安全认证**: 使用 `msal` 实现 OAuth2 流程，支持令牌自动刷新。
- **Windows 优化**: 内置 OpenSSL Applink 补丁，完美解决 Windows 隔离环境下运行的兼容性问题。

## 快速开始

### 1. 环境准备
- Python 3.10+
- [Azure App Registration](https://portal.azure.com/):
  - 创建 "Mobile and desktop applications" 类型的应用。
  - 添加重定向 URI: `https://login.microsoftonline.com/common/oauth2/nativeclient`。
  - 添加 API 权限: `User.Read`, `Calendars.ReadWrite`, `Tasks.ReadWrite`, `Mail.ReadWrite`, `Mail.Send`。

### 2. 本地安装与认证
```bash
# 安装项目
pip install -e .

# 完成首次 OAuth2 认证
m365-auth
```

## 作为 MCP 服务运行

在您的 MCP 客户端（如 Claude Desktop）配置中添加以下内容：

### 方式 A: 远程启动 (推荐)
无需手动下载代码，直接通过 `uv` 远程拉取并运行。

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
        "MS_GRAPH_TOKEN_PATH": "C:/path/to/your/graph_token.json",
        "ENABLE_CALENDAR": "true",
        "ENABLE_TASKS": "true",
        "ENABLE_EMAIL": "true"
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

## 环境变量配置

| 变量名 | 说明 | 默认值 |
| :--- | :--- | :--- |
| `MS_GRAPH_CLIENT_ID` | Azure 应用客户端 ID | 必填 |
| `MS_GRAPH_TOKEN_PATH` | Token 存储路径 | `graph_token.json` |
| `ENABLE_CALENDAR` | 是否启用日历模块 | `true` |
| `ENABLE_TASKS` | 是否启用待办模块 | `true` |
| `ENABLE_EMAIL` | 是否启用邮件模块 | `true` |

## 可用工具 (Tools)

- **系统**: `get_current_time` (用于定位当前时间)。
- **日历**: `list_calendar_events`, `create_calendar_event` (支持提醒), `update_calendar_event`, `delete_calendar_event`, `get_user_schedules`。
- **待办**: `list_tasks`, `create_task`, `update_task`, `complete_task`, `delete_task`。
- **邮件**: `list_emails`, `send_email`, `delete_email`。

## 开源协议
MIT