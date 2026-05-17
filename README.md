# Microsoft 365 Hermes Skill

这是一个为 [Hermes Agent](https://hermes-agent.nousresearch.com/) 量身定做的 **Microsoft 365 (Office 365)** 效率技能。
它能够通过直接调用 Microsoft Graph API，为大语言模型赋予管理 Outlook 日历、Microsoft To-Do 以及邮件的能力。

## 项目结构
```text
├── SKILL.md                 # Hermes 技能描述文档 (元数据声明)
├── README.md                # 本说明文件
└── scripts/
    ├── m365_client.py       # 技能核心调用脚本 (Python)
    └── .env.example         # 配置文件模板
```

## 安装与配置

### 1. 下载技能
将此仓库分支下载到你 Hermes 网关的技能库目录中：

*   **Windows**: `%LOCALAPPDATA%\hermes\skills\productivity\microsoft-365\`
*   **Unix/macOS**: `~/.hermes/skills/productivity/microsoft-365/`

```bash
git clone -b hermes_skill https://github.com/xyz50270/Microsoft-Calendar-MCP.git microsoft-365
```

### 2. 环境准备
确保你的环境中已安装必要的 Python 依赖：
```bash
pip install msal httpx python-dotenv
```

### 3. 配置 Client ID
1.  在 `scripts/` 目录下，将 `.env.example` 复制为 `.env`。
2.  在 `.env` 中填入你的 **Azure Application (client) ID**：
    ```env
    MS_GRAPH_CLIENT_ID=你的_CLIENT_ID_在这里
    ```

### 4. 首次身份认证
在终端运行认证脚本以获取访问令牌：
```bash
python scripts/m365_client.py auth
```
按照屏幕提示访问 [microsoft.com/devicelogin](https://microsoft.com/devicelogin)，输入显示的 9 位代码并完成登录。认证成功后，Token 会加密保存在 `scripts/graph_token.json`。

## 如何在 Hermes 中使用

一旦安装完成，你可以直接通过自然语言指令要求 Hermes 处理 365 事务：

### 日历日程
*   “帮我查一下下周三下午有什么安排？”
*   “在明天上午 10 点帮我创建一个‘周工作汇报’的会议。”
*   “取消本周五下午的那个测试会议。”

### 待办事项 (To-Do)
*   “我今天有哪些待办任务？”
*   “帮我新建一个任务：‘准备下周的出差报告’，截止日期是明天。”
*   “把刚才那个报告任务标记为已完成。”

### 电子邮件 (Mail)
*   “列出我最近收到的 5 封邮件。”
*   “发邮件给 boss@example.com，主题是‘项目进度汇报’，正文写‘一切顺利，请查收附件。’”

## 注意事项
*   **时区**: 脚本默认使用 `China Standard Time` (UTC+8)。
*   **安全**: `scripts/graph_token.json` 包含你的访问凭证，请勿将其上传到公开的代码仓库。
*   **权限**: 确保你的 Azure 应用注册时已授予以下权限（Delegated permissions）：
    *   `User.Read`
    *   `Calendars.ReadWrite`
    *   `Tasks.ReadWrite`
    *   `Mail.ReadWrite`
    *   `Mail.Send`
