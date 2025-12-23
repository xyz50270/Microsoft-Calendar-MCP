# Microsoft 365 MCP 个人助理服务 (v0.1.0)

这是一个基于 **Model Context Protocol (MCP)** 标准构建的个人生产力助手。通过直接集成 **Microsoft Graph API**，它赋予大语言模型（LLM）管理您 Microsoft 365 账户（日历、待办事项、电子邮件）的能力。

## 🎯 定位：私密、安全、高效
本服务专为**个人使用**优化，移除了所有可能导致意外向外部发送信息的复杂功能（如会议邀请、参与者管理、线上会议创建），确保 LLM 作为一个纯粹的“个人秘书”在您的私有域内工作。

---

## 🌟 核心特性

### 1. 深度个人化集成
- **日历 (Calendar)**: 增删改查个人日程。支持查询忙闲状态（自动锁定当前用户），支持自定义提醒设置。
- **待办 (Tasks)**: 完整支持 Microsoft To Do。可管理截止日期、优先级、分类及提醒。
- **邮件 (Email)**: 快速查阅收件箱、发送个人邮件、清理过期邮件。

### 2. 完美的时区方案 (UTC+8)
- **原生东八区支持**: 服务器在请求头中硬编码 `Prefer: outlook.timezone="China Standard Time"`，确保所有返回的时间数据均为北京时间。
- **时区禁转原则**: 通过系统提示词严格限制 LLM 进行 UTC 转换，杜绝因时区计算错误导致的日程偏移。

### 3. 智能任务规划
- **提示词优化**: 针对中文大模型优化的全中文指令集。
- **工具引导**: 强制 LLM 在查询“是否有空”时优先使用 `get_user_schedules` 而非低效的列表扫描。
- **动态适配**: 根据您启用的模块（日历/任务/邮件），自动调整 LLM 的系统指令。

---

## 🚀 快速开始

### 1. Azure 应用注册 (必选)
1. 访问 [Azure Portal](https://portal.azure.com/) 并创建一个新应用。
2. 应用类型选择：**Mobile and desktop applications**。
3. 添加重定向 URI：`https://login.microsoftonline.com/common/oauth2/nativeclient`。
4. 在 **API 权限** 中添加以下 **Delegated** 权限：
   - `User.Read`
   - `Calendars.ReadWrite`
   - `Tasks.ReadWrite`
   - `Mail.ReadWrite`
   - `Mail.Send`

### 2. 本地安装与认证
```bash
# 使用 uv 安装 (推荐)
uv pip install -e .

# 执行交互式认证 (根据提示在浏览器登录)
uv run m365-auth
```

### 3. 配置 MCP 客户端 (以 Claude Desktop 为例)
修改您的配置文件（通常在 `%AppData%/Roaming/Instructions/claude_desktop_config.json`）：

```json
{
  "mcpServers": {
    "m365": {
      "command": "uv",
      "args": [
        "--directory", "F:/your/path/Microsoft-Calendar-MCP",
        "run",
        "m365-mcp"
      ],
      "env": {
        "MS_GRAPH_CLIENT_ID": "您的 Azure 客户端 ID",
        "ENABLE_CALENDAR": "true",
        "ENABLE_TASKS": "true",
        "ENABLE_EMAIL": "true"
      }
    }
  }
}
```

---

## 🛠️ 工具箱 (Tools)

### 📅 日历
- `list_calendar_events`: 列出日程。
- `create_calendar_event`: 创建日程（支持设置 `reminder_minutes`）。
- `update_calendar_event`: 修改日程。
- `delete_calendar_event`: 删除日程。
- `get_user_schedules`: **[推荐]** 查询自己是否有空。

### ✅ 待办 (To Do)
- `list_tasks`: 查看待办列表。
- `create_task`: 新建任务（支持 `due_date`, `importance`, `reminder_date`）。
- `update_task`: 更新任务状态或内容。
- `complete_task`: 快速完成任务。
- `delete_task`: 删除任务。

### 📧 邮件
- `list_emails`: 查看最近邮件。
- `send_email`: 发送邮件。
- `delete_email`: 删除邮件。

### ⚙️ 系统
- `get_current_time`: 获取当前精确的本地时间（LLM 处理相对时间的前提）。

---

## 💡 使用建议
为了获得最佳体验，建议在与 LLM 对话开始时输入：
> “请先获取 m365-assistant 提示词，并检查我当前的时间。”

这会让 LLM 明确：
1. 哪些工具当前可用。
2. 必须遵守 ISO 8601 格式和 UTC+8 时区。
3. 在处理日程冲突时优先调用忙闲查询工具。

---

## 🔒 安全说明
- **secrets.dat**: 该文件包含加密的开发环境配置，仅供内部开发使用。
- **Token 存储**: 认证后的 Token 默认存储在本地 `graph_token.json` 中，请妥善保管。

## 📄 开源协议
MIT