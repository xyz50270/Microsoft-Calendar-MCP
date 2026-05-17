# Microsoft 365 Hermes Skill

这是一个为 [Hermes Agent](https://hermes-agent.nousresearch.com/) 量身定做的 **Microsoft 365 (Office 365)** 效率技能。
它能够通过直接调用 Microsoft Graph API，为大语言模型赋予管理 Outlook 日历、Microsoft To-Do 以及邮件的能力。

## 项目结构
```text
├── SKILL.md                 # Hermes 技能描述文档 (元数据声明)
└── scripts/
    ├── m365_client.py       # 技能核心调用脚本 (Python)
    └── .env.example         # 配置文件模板 (请在此填入 Client ID)
```

## 快速使用

### 1. 配置
复制 `scripts/.env.example` 为 `scripts/.env` 并填写你的 Azure 应用客户端 ID (`MS_GRAPH_CLIENT_ID`)。

### 2. 身份认证
在终端运行认证脚本：
```bash
python scripts/m365_client.py auth
```
根据控制台提示在浏览器中登录你的 Microsoft 账户完成握手。

### 3. 加入技能库
将整个文件夹放入你的 Hermes 技能路径：
*   **Windows**: `%LOCALAPPDATA%\hermes\skills\productivity\microsoft-365\`
*   **Unix**: `~/.hermes/skills/productivity/microsoft-365/`
