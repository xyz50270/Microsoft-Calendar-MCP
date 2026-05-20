---
name: microsoft-365
description: "管理 Microsoft 365 资源，包括 Outlook 日历日程、Microsoft To-Do 待办任务和电子邮件。使用前需要设置 MS_GRAPH_CLIENT_ID 环境变量并运行认证命令。"
version: 1.0.0
author: xyz50270
license: MIT
metadata:
  hermes:
    tags: [m365, office365, calendar, todo, outlook, productivity]
    related_skills: []
---

# Microsoft 365 Skill

## 概述
本 Skill 允许 Hermes 直接连接到用户的 Microsoft 365 账户，处理日历、待办事项和邮件管理。它集成了 Microsoft Graph API 的核心功能。

## 使用条件
- **日程管理 (最高优先级)**: 用户 (bingyu) 明确要求：以后所有日程记录、查询工作或会议安排，必须优先使用本 Skill。
- **日历管理**: 用户要求查看、创建、修改或删除 Outlook 日历行程时。
- **待办任务**: 用户要求管理 Microsoft To-Do 任务列表时。
- **邮件处理**: 用户要求列出、发送或管理 Outlook 邮件时。

## 秘书角色指南 (Secretary Role)
当作为秘书处理日程时：
- **主动提醒**: 在查询日程后，主动分析并指出时间冲突或紧凑的行程。
- **格式化输出**: 使用清晰的列表（时间、地点、主题、持续时间）呈现日程。
- **多日事件提示**: 特别标注出跨越多日的持续性事件。

## 环境配置
使用此 Skill 前，必须确保：
1. **客户端 ID**: 在环境变量中设置 `MS_GRAPH_CLIENT_ID` (从 Azure Portal 获取)。
2. **依赖安装**: 确保安装了 `msal` 和 `httpx`。在 Linux 环境下，如果依赖缺失，可使用：
   ```bash
   uv pip install msal httpx
   ```
3. **认证**: 首次使用需运行认证脚本。
4. **虚拟环境**: 在本系统中，推荐使用 `/opt/hermes/.venv/bin/python` 运行脚本。

## 常用操作脚本

脚本路径: `/opt/data/skills/productivity/microsoft-365/scripts/m365_client.py`
API 笔记: [references/graph-api-notes.md](references/graph-api-notes.md)

### 1. 身份认证 (首次运行)
```bash
/opt/hermes/.venv/bin/python /opt/data/skills/productivity/microsoft-365/scripts/m365_client.py auth
```
运行后按照提示在浏览器中输入代码并登录 Microsoft 账户。

### 2. 日历操作
- **列出近期日程**:
  ```bash
  /opt/hermes/.venv/bin/python /opt/data/skills/productivity/microsoft-365/scripts/m365_client.py calendar-list [start_date] [end_date]
  ```
- **创建日程**:
  ```bash
  /opt/hermes/.venv/bin/python /opt/data/skills/productivity/microsoft-365/scripts/m365_client.py calendar-create "主题" "开始时间(ISO)" "结束时间(ISO)" [地点]
  ```

... (其余命令类推)

## 常见坑点
1. **认证过期**: 如果脚本报错 `Authentication required`，请重新运行 `auth` 命令。
2. **权限问题 (Permission Denied)**: 如果 `graph_token.json` 报错权限拒绝，需检查文件所有者及权限。必要时执行 `chmod o+rw` (注意安全风险)。
3. **依赖缺失**: 如果报错 `ModuleNotFoundError`，需在正确的虚拟环境下运行或重新安装依赖。
4. **时区**: 脚本默认使用 `China Standard Time` (UTC+8)。

- **查询忙闲**:
  ```bash
  python ~/.hermes/skills/productivity/microsoft-365/scripts/m365_client.py calendar-freebusy "邮箱1,邮箱2" "开始时间" "结束时间"
  ```

### 3. 待办事项
- **列出任务**:
  ```bash
  python ~/.hermes/skills/productivity/microsoft-365/scripts/m365_client.py tasks-list
  ```
- **创建任务**:
  ```bash
  python ~/.hermes/skills/productivity/microsoft-365/scripts/m365_client.py tasks-create "任务标题" [截止日期(ISO)]
  ```
- **更新任务**:
  ```bash
  python ~/.hermes/skills/productivity/microsoft-365/scripts/m365_client.py tasks-update "任务ID" '{"status": "completed"}'
  ```
- **删除任务**:
  ```bash
  python ~/.hermes/skills/productivity/microsoft-365/scripts/m365_client.py tasks-delete "任务ID"
  ```

### 4. 邮件管理
- **列出最近邮件**:
  ```bash
  python ~/.hermes/skills/productivity/microsoft-365/scripts/m365_client.py mail-list [limit]
  ```
- **发送邮件**:
  ```bash
  python ~/.hermes/skills/productivity/microsoft-365/scripts/m365_client.py mail-send "收件人1,收件人2" "主题" "正文"
  ```
- **移动邮件**:
  ```bash
  python ~/.hermes/skills/productivity/microsoft-365/scripts/m365_client.py mail-move "邮件ID" "文件夹ID"
  ```
- **删除邮件**:
  ```bash
  python ~/.hermes/skills/productivity/microsoft-365/scripts/m365_client.py mail-delete "邮件ID"
  ```

## 常见坑点
1. **认证过期**: 如果脚本报错 `Authentication required`，请重新运行 `auth` 命令。
2. **环境变量**: 如果未设置 `MS_GRAPH_CLIENT_ID`，脚本将无法启动。
3. **时区**: 脚本默认使用 `China Standard Time` (UTC+8)。
4. **Windows 路径与目录创建 (WinError 3)**: 在 Windows 环境下，使用未完全解析为绝对路径的文件名进行 `os.makedirs(os.path.dirname(path))` 时，如果 `dirname` 提取出的父目录是空字符串 `''` 或包含未展开的 `~`，会导致 `WinError 3` 系统找不到指定路径的错误。应始终使用 `os.path.abspath` 或 `os.path.isabs` 将其转为绝对路径，并在创建目录前加上 `if dir_name:` 的防空保护。
4. **Windows 下的路径解析问题**: 在 Windows Git-bash 中执行 Python 脚本时，如果 token_path 仅包含相对路径或文件名（如 `graph_token.json`），`os.path.dirname(token_path)` 会返回空字符串。若直接调用 `os.makedirs('', exist_ok=True)` 会触发 `WinError 3` 错误。因此，脚本中在创建父目录前必须先判断目录名是否为空，且强烈推荐使用绝对路径。
5. **在 Windows 命令行下传参**: 使用修改日程或任务（`calendar-update` 或 `tasks-update`）时，更新参数需以 JSON 格式传递。在 Git-bash/Windows 下注意使用单双引号的嵌套，例如：
   ```bash
   python m365_client.py calendar-update "事件ID" '{"subject": "新主题", "location": "新地点"}'
   ```
4. **时区**: 脚本默认使用 `China Standard Time` (UTC+8)。

## 验证清单
- [ ] 确保 `MS_GRAPH_CLIENT_ID` 已设置。
- [ ] 运行 `m365_client.py auth` 并成功获取令牌。
- [ ] 运行 `calendar-list` 能够输出 JSON 格式的日程数据。
