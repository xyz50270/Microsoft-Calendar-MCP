# Microsoft 365 MCP Service

这是一个基于 **Model Context Protocol (MCP)** 标准开发的个人 Microsoft 365 助手服务。它允许大语言模型（LLM）通过直接调用 **Microsoft Graph API** 来管理您的个人日历（Calendar）、待办事项（To Do）和电子邮件（Outlook）。

## 核心特性

- **个人化集成 (Private & Secure)**:
  - **日历 (Calendar)**: 专为个人设计，移除邀请他人功能，杜绝误发会议风险。支持增删改查日程、空闲/繁忙状态查询、**自定义提醒 (Reminders)**。
  - **待办 (Tasks)**: 全属性支持（优先级、截止日期、提醒、状态管理等）。
  - **邮件 (Email)**: 支持收件箱列表查询、发送邮件和删除操作。
- **架构优势**:
  - **自动时区处理**: 服务器原生支持 **东八区 (UTC+8)** 时间。LLM 直接使用本地时间交互，严禁进行手动的 UTC 转换。
  - **参数优化**: 简化工具参数（如日程查询自动锁定当前用户），降低 LLM 调用出错率。
  - **规范化文档**: 所有工具均提供 Google 风格详细文档，明确 ISO 8601 格式要求。
  - **安全认证**: 使用 `msal` 实现 OAuth2 流程，支持令牌自动刷新。
- **Windows 优化**: 内置 OpenSSL Applink 补丁，完美解决 Windows 环境下的运行兼容性问题。

## 快速开始

### 1. 环境准备
- Python 3.10+
- [Azure App Registration](https://portal.azure.com/):
  - 创建 "Mobile and desktop applications" 类型的应用。
  - 添加重定向 URI: `https://login.microsoftonline.com/common/oauth2/nativeclient`。
  - 添加 API 权限 (Delegated): `User.Read`, `Calendars.ReadWrite`, `Tasks.ReadWrite`, `Mail.ReadWrite`, `Mail.Send`。

### 2. 本地安装与认证
```bash
# 使用 uv 安装依赖 (推荐)
uv pip install -e .

# 完成首次 OAuth2 交互式认证
uv run m365-auth
```

## 作为 MCP 服务运行

在您的 MCP 客户端（如 Claude Desktop）配置中添加以下内容：

### 方式 A: 远程启动 (推荐)
通过 `uvx` 直接从 GitHub 运行（需确保已完成本地认证）：

```json
{
  "mcpServers": {
    "m365": {
      "command": "uvx",
      "args": [
        "--with", "git+https://github.com/xyz50270/Microsoft-Calendar-MCP.git",
        "m365-mcp"
      ],
      "env": {
        "MS_GRAPH_CLIENT_ID": "您的客户端ID",
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
      "command": "python",
      "args": [
        "E:/path/to/Microsoft-Calendar-MCP/src/server.py"
      ],
      "env": {
        "MS_GRAPH_CLIENT_ID": "您的客户端ID"
      }
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

- **系统**: `get_current_time` (定位当前时间，调用其他工具前的必备动作)。
- **日历**: 
  - `list_calendar_events`: 获取日程列表。
  - `create_calendar_event`: 创建日程（仅限个人，不发邀请）。
  - `update_calendar_event`: 修改日程。
  - `delete_calendar_event`: 删除日程。
  - `get_user_schedules`: 查询当前用户的忙闲状态。
- **待办**: `list_tasks`, `create_task`, `update_task`, `complete_task`, `delete_task`。
- **邮件**: `list_emails`, `send_email`, `delete_email`。

## 注意事项
1. **时间格式**: 所有输入时间必须符合 **ISO 8601** 格式（例如：`2025-12-23T09:00:00`）。
2. **时区**: 系统强制使用 **UTC+8**。请确保 LLM 获取的参考时间正确。
3. **安全**: `secrets.dat` 为加密的开发密钥包。生产环境请使用自己的 `MS_GRAPH_CLIENT_ID`。

## 开源协议
MIT
