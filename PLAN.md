# M365 CLI - Project Plan

## 项目概述

M365 CLI 是一个现代化的 Microsoft 365 命令行工具，用于管理邮件、日历、OneDrive 和 SharePoint 服务。基于 Node.js ESM 和 Microsoft Graph API 构建，专为个人用户设计。

## 目标

- 提供简洁易用的 CLI 界面操作 M365 服务
- 支持 AI 友好的输出格式（简洁文本 + JSON 选项）
- 最小化依赖，保持轻量级
- 完善的错误处理和自动 token 刷新

## 开发阶段

### Phase 1: 框架 + 邮件功能 ✅ (已完成)

**目标**：搭建 CLI 框架，实现认证和邮件功能

**功能清单**：
- [x] 项目结构搭建
- [x] CLI 框架 (commander.js)
- [x] Token 管理（Device Code Flow + 自动刷新）
- [x] Graph API 统一调用层
- [x] 邮件功能
  - [x] 列出邮件 `m365 mail list`
  - [x] 读取邮件 `m365 mail read <id>`
  - [x] 发送邮件 `m365 mail send` (支持附件)
  - [x] 搜索邮件 `m365 mail search`

**交付物**：
- ✅ 可全局使用的 `m365` 命令
- ✅ 完整的邮件操作功能
- ✅ 文档和测试

**完成时间**：2026-02-16

### Phase 2: 日历功能 ✅ (已完成)

**目标**：实现日历管理功能

**功能清单**：
- [x] `m365 calendar list [--days N]` - 列出日程（默认 7 天）
- [x] `m365 calendar get <id>` - 查看日程详情
- [x] `m365 calendar create <title> --start --end [options]` - 创建日程
- [x] `m365 calendar update <id> [options]` - 更新日程
- [x] `m365 calendar delete <id>` - 删除日程

**已实现功能**：
- ✅ 时区处理（Asia/Shanghai）
- ✅ 日期时间格式转换
- ✅ 全天事件支持（--allday）
- ✅ 参会人支持（--attendees）
- ✅ 地点和描述（--location, --body）
- ✅ JSON 输出格式（--json）
- ✅ 友好的文本输出格式
- ✅ 命令别名 `cal`

**技术要点**：
- 使用 Graph API `/me/calendarView` 获取日期范围内的事件
- 使用 `Prefer: outlook.timezone="Asia/Shanghai"` 处理时区
- 支持 CRUD 操作（Create, Read, Update, Delete）
- 全天事件需要至少跨越 24 小时（如 2026-02-18 到 2026-02-19）

**完成时间**：2026-02-16

### Phase 3: OneDrive 文件管理 ✅ (已完成)

**目标**：实现文件和文件夹操作

**功能清单**：
- [x] `m365 onedrive ls [path] [--top N] [--json]` - 列出文件/文件夹
- [x] `m365 onedrive get <path> [--json]` - 获取文件/文件夹元数据
- [x] `m365 onedrive download <remote-path> [local-path]` - 下载文件
- [x] `m365 onedrive upload <local-path> [remote-path]` - 上传文件
- [x] `m365 onedrive search <query> [--top N] [--json]` - 搜索文件
- [x] `m365 onedrive share <path> [--type view|edit] [--json]` - 分享文件
- [x] `m365 onedrive mkdir <path>` - 创建文件夹
- [x] `m365 onedrive rm <path> [--force]` - 删除文件/文件夹

**已实现功能**：
- ✅ 列出指定路径下的文件和文件夹
- ✅ 获取文件/文件夹详细元数据
- ✅ 下载文件到本地（带进度显示）
- ✅ 上传文件到 OneDrive（小文件直接上传，大文件分片上传）
- ✅ 搜索文件
- ✅ 创建分享链接（只读/编辑）
- ✅ 创建文件夹
- ✅ 删除文件/文件夹（带确认提示，支持 --force）
- ✅ JSON 输出格式（--json）
- ✅ 友好的文本输出格式
- ✅ 文件大小人性化显示（B/KB/MB/GB/TB）
- ✅ 命令别名 `od`

**技术要点**：
- 使用 Graph API `/me/drive/root/children` 和 `/me/drive/root:/{path}:/children` 获取文件列表
- 小文件上传（<4MB）使用 `PUT /me/drive/root:/{path}:/content`
- 大文件上传（≥4MB）使用 upload session（分片上传，每片 3.2MB）
- 下载文件显示实时进度
- 路径编码处理（支持中文和特殊字符）
- 删除操作需要用户确认（除非使用 --force）

**完成时间**：2026-02-16

### Phase 3.5: SharePoint 文档协作 ✅ (已完成)

**目标**：实现 SharePoint 站点和文档操作

**功能清单**：
- [x] `m365 sharepoint sites [--search query] [--top N] [--json]` - 列出/搜索站点
- [x] `m365 sharepoint lists <site> [--top N] [--json]` - 列出站点列表
- [x] `m365 sharepoint items <site> <list> [--top N] [--json]` - 列出列表项目
- [x] `m365 sharepoint files <site> [path] [--top N] [--json]` - 列出文档库文件
- [x] `m365 sharepoint download <site> <file-path> [local-path]` - 下载文件
- [x] `m365 sharepoint upload <site> <local-path> [remote-path]` - 上传文件
- [x] `m365 sharepoint search <query> [--top N] [--json]` - 搜索内容

**已实现功能**：
- ✅ 列出可访问的 SharePoint 站点（已关注站点）
- ✅ 搜索站点（使用 `GET /sites?search={query}`）
- ✅ 列出站点的列表和文档库
- ✅ 列出列表中的项目（展开 fields）
- ✅ 列出文档库文件
- ✅ 下载文件（支持进度显示）
- ✅ 上传文件（小文件直接上传，大文件分片上传）
- ✅ 搜索 SharePoint 内容（driveItem、listItem、site）
- ✅ 站点解析：支持 URL 格式（hostname:/path）和 site-id
- ✅ 命令别名 `sp`
- ✅ JSON 输出格式（--json）
- ✅ 友好的文本输出格式
- ✅ 复用 OneDrive 文件操作逻辑

**技术要点**：
- 使用 Graph API `/sites/{hostname}:/{path}` 解析站点 URL 获取 site-id
- 使用 `GET /me/followedSites` 列出已关注站点
- 使用 `GET /sites?search={query}` 搜索站点
- 使用 `GET /sites/{site-id}/lists` 获取列表
- 使用 `GET /sites/{site-id}/lists/{list-id}/items?expand=fields` 获取列表项
- 使用 `GET /sites/{site-id}/drive/root/children` 获取文档库文件
- 使用 `POST /search/query` 搜索 SharePoint 内容
- 文件上传/下载复用 OneDrive 逻辑（upload session 支持）
- 站点参数支持两种格式：URL（contoso.sharepoint.com:/sites/team）或 site-id

**权限要求**：
- ⚠️ 需要在 Azure AD 应用中添加权限：`Sites.ReadWrite.All` 或 `Sites.Read.All`
- 用户首次使用需要重新登录：`m365 logout` → `m365 login`
- 配置文件已更新（`config/default.json` 包含 `Sites.ReadWrite.All` scope）

**完成时间**：2026-02-16

### Phase 4: 联系人和高级功能 (待开发)

**目标**：扩展功能和优化体验

**功能清单**：
- [ ] 联系人管理 `m365 contacts`
  - [ ] `m365 contacts list` - 列出联系人
  - [ ] `m365 contacts get <id>` - 查看联系人详情
  - [ ] `m365 contacts create` - 创建联系人
  - [ ] `m365 contacts update <id>` - 更新联系人
  - [ ] `m365 contacts delete <id>` - 删除联系人
  - [ ] `m365 contacts search <query>` - 搜索联系人
- [ ] ~~Teams 消息（已跳过 - 不需要）~~
- [ ] 配置文件管理 `m365 config`
  - [ ] `m365 config show` - 显示当前配置
  - [ ] `m365 config set <key> <value>` - 设置配置项
  - [ ] `m365 config reset` - 重置为默认配置
- [ ] 多账号支持
  - [ ] `m365 login --profile <name>` - 使用指定配置
  - [ ] `m365 profile list` - 列出所有配置
  - [ ] `m365 profile switch <name>` - 切换配置
- [ ] 交互式模式
  - [ ] `m365 interactive` - 启动交互式 shell

**说明**：
- Teams 功能已跳过：主要面向企业用户，个人用户通常不需要
- 联系人管理优先级较高，适合个人用户
- 配置管理和多账号支持可以提升使用体验

### Phase 5: 性能优化和发布 (待开发)

**目标**：优化性能，准备发布

**任务清单**：
- [ ] 性能优化
  - [ ] 实现缓存机制（token、常用数据）
  - [ ] 批量请求支持（Graph API batch）
  - [ ] 并发请求优化
- [ ] 完善错误处理和重试机制
  - [ ] 指数退避重试（429 Too Many Requests）
  - [ ] 更友好的错误提示
  - [ ] 网络超时处理
- [ ] 编写完整测试
  - [ ] 单元测试（Jest）
  - [ ] 集成测试
  - [ ] 端到端测试
- [ ] 完善文档
  - [ ] API 文档（JSDoc）
  - [ ] 更多使用示例
  - [ ] 常见问题 FAQ
  - [ ] 贡献指南
- [ ] CI/CD 配置
  - [ ] GitHub Actions workflow
  - [ ] 自动化测试
  - [ ] 自动发布
- [ ] NPM 发布准备
  - [ ] 包名注册
  - [ ] 版本管理策略
  - [ ] LICENSE 文件
  - [ ] CHANGELOG.md

## 架构设计

### 核心模块

```
m365-cli/
├── bin/m365.js              # CLI 入口，解析命令
├── src/
│   ├── auth/                # 认证模块
│   │   ├── token-manager.js # Token 存储、刷新
│   │   └── device-flow.js   # Device Code Flow 实现
│   ├── graph/               # Graph API 调用层
│   │   └── client.js        # 统一 API 客户端
│   ├── commands/            # 命令实现
│   │   ├── mail.js          # 邮件命令
│   │   ├── calendar.js      # 日历命令
│   │   ├── onedrive.js      # OneDrive 命令
│   │   └── sharepoint.js    # SharePoint 命令
│   └── utils/               # 工具函数
│       ├── config.js        # 配置管理
│       ├── output.js        # 输出格式化
│       └── error.js         # 错误处理
└── config/
    └── default.json         # 默认配置
```

### 数据流

```
用户输入 (CLI)
    ↓
bin/m365.js (commander.js)
    ↓
commands/* (业务逻辑)
    ↓
graph/client.js (HTTP 请求)
    ↓
auth/token-manager.js (Token 管理)
    ↓
Microsoft Graph API
```

### 认证流程

1. **初次登录**：
   - 用户执行 `m365 login`
   - Device Code Flow：显示登录 URL 和 code
   - 轮询获取 token
   - 保存 access_token + refresh_token 到 `~/.openclaw/workspace/creds/.m365-creds`

2. **后续请求**：
   - 检查 token 是否过期
   - 未过期：直接使用
   - 已过期：自动使用 refresh_token 刷新
   - 刷新失败：提示用户重新登录

### 错误处理策略

- **401 Unauthorized**：Token 过期，尝试刷新
- **403 Forbidden**：权限不足，提示用户
- **404 Not Found**：资源不存在，返回友好提示
- **429 Too Many Requests**：速率限制，自动重试（exponential backoff）
- **5xx Server Error**：服务端错误，重试最多 3 次

## 技术选型

### 核心依赖

| 依赖 | 版本 | 用途 |
|------|------|------|
| Node.js | ≥18 | 运行时（支持 ESM 和内置 fetch） |
| commander.js | ^12.0.0 | CLI 参数解析 |

### 技术决策

1. **ESM vs CommonJS**：选择 ESM（现代化，更好的 tree-shaking）
2. **HTTP 客户端**：使用 Node.js 内置 `fetch`（Node.js 18+），无需额外依赖
3. **JSON 处理**：内置 `JSON.parse/stringify`
4. **文件操作**：内置 `fs/promises`
5. **命令行参数**：`commander.js`（轻量、功能完善）

### 输出格式设计

**默认输出（文本）**：
```
📧 Mail List (top 5)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[1] 📩 Meeting Reminder
    From: alice@example.com
    Date: 2026-02-16 09:30
    
[2] ✅ Project Update
    From: bob@example.com
    Date: 2026-02-15 18:45
```

**JSON 输出（`--json`）**：
```json
[
  {
    "id": "AAMkAG...",
    "subject": "Meeting Reminder",
    "from": "alice@example.com",
    "receivedDateTime": "2026-02-16T09:30:00Z",
    "isRead": false
  }
]
```

## 配置管理

### 配置文件位置

- **默认配置**：`config/default.json`
- **用户配置**：`~/.m365-cli/config.json`（可选，Phase 4 实现）
- **环境变量**：`M365_*` 前缀

### 配置优先级

环境变量 > 用户配置 > 默认配置

### 凭据存储

- **路径**：`~/.openclaw/workspace/creds/.m365-creds`
- **格式**：JSON
- **内容**：
  ```json
  {
    "tenant_id": "...",
    "client_id": "...",
    "access_token": "...",
    "refresh_token": "...",
    "expires_at": 1234567890
  }
  ```
- **权限**：`600`（仅当前用户可读写）

## 安全考虑

1. **凭据保护**：
   - Token 文件权限 `600`
   - 不在日志中输出敏感信息
   - 不在错误信息中暴露 token

2. **输入验证**：
   - 邮件地址格式验证
   - 文件路径安全检查
   - 防止路径遍历攻击

3. **网络安全**：
   - 仅使用 HTTPS
   - 验证 SSL 证书
   - 设置合理的超时时间

## 性能指标

- **启动时间**：< 100ms（冷启动）
- **命令响应**：< 2s（含 API 请求）
- **Token 刷新**：< 1s
- **大文件上传**：支持进度显示，分片上传（3.2MB/片）
- **大文件下载**：支持进度显示，流式传输

## 发布计划

### 版本规划

- **v0.1.0** ✅：Phase 1（框架 + 邮件）
- **v0.2.0** ✅：Phase 2（日历）
- **v0.3.0** ✅：Phase 3（OneDrive + SharePoint）
- **v0.4.0**：Phase 4（联系人和高级功能）
- **v1.0.0**：Phase 5（优化和稳定版）

### 发布检查清单

- [x] Phase 1-3 功能测试通过
- [x] 基础文档完整（README + PLAN）
- [ ] 所有功能测试通过
- [ ] API 文档
- [ ] 无已知严重 bug
- [ ] 性能指标达标
- [ ] 安全审计通过
- [ ] LICENSE 文件
- [ ] CHANGELOG.md 更新

## 未来规划

- 🚀 插件系统（支持第三方扩展）
- 🎨 主题和输出格式自定义
- 📊 使用统计和分析（可选）
- 🌐 多语言支持
- 🔔 Webhook 和通知
- 📱 移动端支持（Termux）
- 🤝 团队协作功能增强

## 参考资料

- [Microsoft Graph API 文档](https://learn.microsoft.com/en-us/graph/overview)
- [OAuth 2.0 Device Code Flow](https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code)
- [Commander.js 文档](https://github.com/tj/commander.js)
- [SharePoint REST API](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-rest-endpoints)

---

**最后更新**：2026-02-16  
**当前版本**：0.3.0  
**当前阶段**：Phase 3 完成，Phase 4 待开发  
**状态**：Phases 1-3 ✅ 完成
