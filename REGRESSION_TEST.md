# M365 CLI 回归测试报告（含附件功能）

**测试时间**: 2026-02-16 18:47 CST  
**版本**: 0.1.0  
**环境**: Linux vmjasUbuntu  
**执行**: `bash test-all.sh` + 手动补充测试  
**日志**: `test-results.log`, `manual-tests.log`

---

## 1. 测试总览
- ✅ 通过: **38**
- ❌ 失败: **3**
- ⚠️ 跳过/受限: **2**
- **总计**: **43**

---

## 2. 详细测试结果

### TC-001: 基础帮助
- 命令: `m365 --help`
- 结果: ✅
- 输出摘要: 显示命令列表与说明
- 问题: 无

### TC-002: 版本号
- 命令: `m365 --version`
- 结果: ✅
- 输出摘要: `0.1.0`
- 问题: 无

### TC-003: 邮件列表 Top3
- 命令: `m365 mail list --top 3`
- 结果: ✅
- 输出摘要: 正确输出 3 封邮件
- 问题: 无

### TC-004: 邮件列表 JSON
- 命令: `m365 mail list --top 3 --json`
- 结果: ✅
- 输出摘要: JSON 正常返回
- 问题: 部分内部邮件 from.emailAddress.address 仍为 Exchange DN（详见“新发现问题”）

### TC-005: 已发送邮件 Top3
- 命令: `m365 mail list --folder sent --top 3`
- 结果: ✅
- 输出摘要: 正常列出 Sent Items
- 问题: 无

### TC-006: 草稿箱邮件 Top3
- 命令: `m365 mail list --folder drafts --top 3`
- 结果: ✅
- 输出摘要: 正常列出 Drafts
- 问题: 无

### TC-007: 已删除邮件 Top3
- 命令: `m365 mail list --folder deleted --top 3`
- 结果: ✅
- 输出摘要: 正常列出 Deleted Items
- 问题: 无

### TC-008: 读取邮件
- 命令: `m365 mail read <id>`
- 结果: ✅
- 输出摘要: 标题/发件人/正文正常
- 问题: 无

### TC-009: 读取邮件 JSON
- 命令: `m365 mail read <id> --json`
- 结果: ✅
- 输出摘要: JSON 正常
- 问题: 无

### TC-010: 搜索邮件
- 命令: `m365 mail search "test" --top 3`
- 结果: ✅
- 输出摘要: 返回包含 test 的结果
- 问题: 无

### TC-011: 发送邮件（回归测试）
- 命令: `m365 mail send your@email.com "回归测试" "这是回归测试邮件"`
- 结果: ✅
- 输出摘要: `✅ Email sent successfully`
- 问题: 无

### TC-012: 发件人显示
- 命令: `m365 mail list --top 10`
- 结果: ✅
- 输出摘要: From 显示为姓名/邮箱，不是 Exchange DN
- 问题: 无

### TC-013: 发送带附件邮件
- 命令: `m365 mail send your@email.com "附件测试" "这封邮件带有附件" --attach /tmp/attachment-test.txt`
- 结果: ✅
- 输出摘要: `Attachments: 1`
- 问题: 无

### TC-014: 读取附件邮件（附件列表）
- 命令: `m365 mail read <attach-id>`
- 结果: ❌
- 输出摘要: 仅显示正文，未显示附件列表
- 问题: CLI 输出未展示附件信息

### TC-015: 读取附件邮件 JSON
- 命令: `m365 mail read <attach-id> --json | jq '{subject, hasAttachments, attachments}'`
- 结果: ❌
- 输出摘要: `hasAttachments: true` 但 `attachments: null`
- 问题: JSON 未包含附件列表

### TC-016: 日历列表
- 命令: `m365 cal list`
- 结果: ✅
- 输出摘要: 未来 7 天事件正常
- 问题: 无

### TC-017: 日历 JSON（3 天）
- 命令: `m365 cal list --days 3 --json`
- 结果: ✅
- 输出摘要: JSON 返回正常
- 问题: 无

### TC-018: 创建日历事件
- 命令: `m365 cal create "回归测试事件" --start "2026-02-17T14:00" --end "2026-02-17T15:00"`
- 结果: ✅
- 输出摘要: 成功返回事件 ID
- 问题: 无

### TC-019: 获取日历事件
- 命令: `m365 cal get <id>`
- 结果: ✅
- 输出摘要: 事件详情正常
- 问题: 无

### TC-020: 更新日历事件
- 命令: `m365 cal update <id> --title "更新回归测试"`
- 结果: ✅
- 输出摘要: 标题更新成功
- 问题: 无

### TC-021: 时区验证
- 命令: `m365 cal get <id> --json`
- 结果: ✅
- 输出摘要: `timeZone: Asia/Shanghai`
- 问题: 无

### TC-022: 删除日历事件
- 命令: `m365 cal delete <id>`
- 结果: ✅
- 输出摘要: 成功删除
- 问题: 无

### TC-023: OneDrive 根目录
- 命令: `m365 od ls`
- 结果: ✅
- 输出摘要: 列表正常
- 问题: 无

### TC-024: OneDrive JSON
- 命令: `m365 od ls --json`
- 结果: ✅
- 输出摘要: type 字段正确（file/folder）
- 问题: 无

### TC-025: OneDrive 创建目录
- 命令: `m365 od mkdir "regression-test"`
- 结果: ✅
- 输出摘要: Created
- 问题: 无

### TC-026: OneDrive 上传文件
- 命令: `m365 od upload /tmp/regression-test.txt "regression-test/test.txt"`
- 结果: ✅
- 输出摘要: Uploaded
- 问题: 无

### TC-027: OneDrive 列目录
- 命令: `m365 od ls "regression-test"`
- 结果: ✅
- 输出摘要: test.txt 存在
- 问题: 无

### TC-028: OneDrive 获取文件
- 命令: `m365 od get "regression-test/test.txt"`
- 结果: ✅
- 输出摘要: size/webUrl 正常
- 问题: 无

### TC-029: OneDrive 下载文件并校验
- 命令: `m365 od download "regression-test/test.txt" /tmp/regression-download.txt`
- 结果: ❌
- 输出摘要: 下载成功，但文件内容为 `{"type":"Buffer","data":[...]}`
- 问题: 下载内容与原始文本不一致

### TC-030: OneDrive 搜索
- 命令: `m365 od search "regression"`
- 结果: ⚠️
- 输出摘要: `No results found`
- 问题: 可能为索引延迟，待确认

### TC-031: OneDrive 分享
- 命令: `m365 od share "regression-test/test.txt"`
- 结果: ⚠️
- 输出摘要: `sharingDisabled`
- 问题: 站点禁用分享（环境限制）

### TC-032: OneDrive 删除文件
- 命令: `m365 od rm "regression-test/test.txt" --force`
- 结果: ✅
- 输出摘要: Deleted
- 问题: 无

### TC-033: OneDrive 删除目录
- 命令: `m365 od rm "regression-test" --force`
- 结果: ✅
- 输出摘要: Deleted
- 问题: 无

### TC-034: SharePoint 站点列表
- 命令: `m365 sp sites`
- 结果: ✅
- 输出摘要: 正常列出站点
- 问题: 无

### TC-035: SharePoint JSON
- 命令: `m365 sp sites --json`
- 结果: ✅
- 输出摘要: name 字段存在
- 问题: 无

### TC-036: SharePoint 站点搜索
- 命令: `m365 sp sites --search "migration"`
- 结果: ✅
- 输出摘要: 返回 Lifescan Migration
- 问题: 无

### TC-037: SharePoint 列表（site-id）
- 命令: `m365 sp lists <site-id>`
- 结果: ✅
- 输出摘要: 返回 Documents 列表
- 问题: 无

### TC-038: SharePoint 文件
- 命令: `m365 sp files <site-id>`
- 结果: ✅
- 输出摘要: 无文件（空列表）
- 问题: 无

### TC-039: SharePoint 搜索
- 命令: `m365 sp search "project"`
- 结果: ✅
- 输出摘要: 返回 3 条结果
- 问题: 无

### TC-040: 边界测试 - 无效邮件 ID
- 命令: `m365 mail read "invalid-id-12345"`
- 结果: ✅
- 输出摘要: 友好错误提示
- 问题: 无

### TC-041: 边界测试 - 不存在路径
- 命令: `m365 od ls "path-that-does-not-exist"`
- 结果: ✅
- 输出摘要: 友好错误提示
- 问题: 无

### TC-042: 边界测试 - 无效日历 ID
- 命令: `m365 cal get "invalid-id"`
- 结果: ✅
- 输出摘要: 友好错误提示
- 问题: 无

### TC-043: 边界测试 - 无效站点 ID
- 命令: `m365 sp lists "invalid-site"`
- 结果: ✅
- 输出摘要: 友好错误提示
- 问题: 无

---

## 3. ISSUE-001 ~ ISSUE-011 回归验证
- **ISSUE-001**（sent 文件夹）: ✅ `mail list --folder sent` 正常
- **ISSUE-002**（SharePoint lists site-id）: ✅ `sp lists <site-id>` 正常
- **ISSUE-003**（JSON 字段缺失）: ✅ `od ls --json` 有 type，`sp sites --json` 有 name
- **ISSUE-004**（友好错误信息）: ✅ 边界测试提示友好
- **ISSUE-005**（日历时区显示）: ✅ 更新后 JSON 显示 Asia/Shanghai
- **ISSUE-006**（测试脚本 jq 语法）: ✅ `test-all.sh` 正常执行
- **ISSUE-007**（mail --folder 文档）: ✅ README 已补充（未回归失败）
- **ISSUE-008**（SharePoint site-id 文档）: ✅ README 已补充（未回归失败）
- **ISSUE-009**（OneDrive JSON type）: ✅ type 字段正确
- **ISSUE-010**（SharePoint JSON name）: ✅ name 字段正确
- **ISSUE-011**（发件人显示 DN）: ✅ `mail list --top 10` 显示姓名/邮箱

---

## 4. 附件功能测试结果
- ✅ **发送带附件邮件**成功（Attachments: 1）
- ❌ **读取邮件未显示附件列表**（CLI 纯文本输出无附件信息）
- ❌ **JSON 未返回附件列表**（`hasAttachments: true` 但 `attachments: null`）
- ⚠️ **无附件下载子命令**（源码 `mail.js` 仅 list/read/send/search）

---

## 5. 新发现的问题
1. **OneDrive 下载内容不一致**
   - `m365 od download` 生成的文件内容为 `{"type":"Buffer","data":[...]}`，非原始文本
   - 影响: 下载内容无法直接使用

2. **邮件列表 JSON 中内部发件人仍显示 Exchange DN**
   - `mail list --json` 中 `from.emailAddress.address` 返回 Exchange DN
   - 影响: JSON 数据可读性差（文本列表已修复）

3. **附件信息缺失**
   - `mail read` 文本输出不显示附件列表
   - JSON 输出 `attachments` 为 `null`
   - 无附件下载命令

4. **OneDrive search 可能存在索引延迟**
   - 上传后立即搜索 `regression` 无结果

5. **OneDrive share 受限**
   - 返回 `sharingDisabled`（环境限制，非功能性问题）

---

## 6. 总结与建议
- 主流程功能基本可用，ISSUE-001 ~ ISSUE-011 均已验证通过。
- **需优先修复**: OneDrive 下载内容不一致、邮件附件读取/下载功能缺失、JSON 中 Exchange DN。
- **建议**:
  1. 检查 OneDrive download 实现（Buffer 写入方式）。
  2. `mail read` 时获取附件列表（Graph: `/messages/{id}/attachments`）。
  3. 增加 `mail attachments` / `mail attach download` 子命令或扩展 `mail read` 输出。
  4. JSON 中 `from.emailAddress.address` 应优先展示真实邮箱（或返回 name/address 两个字段）。

---

> 备注：详细命令输出见 `test-results.log` 与 `manual-tests.log`。
