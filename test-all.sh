#!/bin/bash
# M365 CLI 全功能自动化测试脚本

set +e  # 不因错误而退出，继续测试其他命令

LOG_FILE="test-results.log"
> "$LOG_FILE"  # 清空日志文件

echo "========== M365 CLI 全功能测试 ==========" | tee -a "$LOG_FILE"
echo "测试时间: $(date)" | tee -a "$LOG_FILE"
echo "" | tee -a "$LOG_FILE"

# 辅助函数：测试命令并记录结果
test_cmd() {
    local test_id="$1"
    local desc="$2"
    local cmd="$3"
    
    echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" | tee -a "$LOG_FILE"
    echo "[$test_id] $desc" | tee -a "$LOG_FILE"
    echo "命令: $cmd" | tee -a "$LOG_FILE"
    echo "" | tee -a "$LOG_FILE"
    
    # 执行命令并捕获输出和返回码
    output=$(eval "$cmd" 2>&1)
    exit_code=$?
    
    echo "$output" | tee -a "$LOG_FILE"
    echo "" | tee -a "$LOG_FILE"
    echo "返回码: $exit_code" | tee -a "$LOG_FILE"
    echo "" | tee -a "$LOG_FILE"
    
    if [ $exit_code -eq 0 ]; then
        echo "结果: ✅ PASS" | tee -a "$LOG_FILE"
    else
        echo "结果: ❌ FAIL" | tee -a "$LOG_FILE"
    fi
    echo "" | tee -a "$LOG_FILE"
}

# 1. 基础命令测试
echo "========== 1. 基础命令 ==========" | tee -a "$LOG_FILE"
test_cmd "TC-001" "测试 --help" "m365 --help"
test_cmd "TC-002" "测试 --version" "m365 --version"
test_cmd "TC-003" "测试 logout --help" "m365 logout --help"

# 2. 邮件功能测试
echo "========== 2. 邮件 (mail) ==========" | tee -a "$LOG_FILE"
test_cmd "TC-004" "默认列出邮件" "m365 mail list"
test_cmd "TC-005" "限制邮件数量 --top 3" "m365 mail list --top 3"
test_cmd "TC-006" "JSON 输出" "m365 mail list --top 3 --json | jq '.[] | {subject, from: .from.emailAddress.address}'"
test_cmd "TC-007" "列出已发送文件夹 (预期失败)" "m365 mail list --folder sent"

# 获取第一封邮件 ID 用于测试
MAIL_ID=$(m365 mail list --top 1 --json 2>/dev/null | jq -r '.[0].id')
if [ -n "$MAIL_ID" ]; then
    test_cmd "TC-008" "读取邮件" "m365 mail read '$MAIL_ID' | head -30"
    test_cmd "TC-009" "JSON 格式读取邮件" "m365 mail read '$MAIL_ID' --json | jq '{subject, from: .from.emailAddress}'"
fi

test_cmd "TC-010" "搜索邮件" "m365 mail search 'test' --top 3"
test_cmd "TC-011" "搜索邮件 JSON 输出" "m365 mail search 'test' --json | jq '.[] | .subject'"

echo "发送测试邮件..." | tee -a "$LOG_FILE"
test_cmd "TC-012" "发送测试邮件" "m365 mail send eleven@qzitech.cn 'CLI Test $(date +%H:%M:%S)' 'This is a test from m365 cli'"

# 3. 日历功能测试
echo "========== 3. 日历 (calendar) ==========" | tee -a "$LOG_FILE"
test_cmd "TC-013" "列出未来7天事件" "m365 cal list"
test_cmd "TC-014" "列出未来30天事件" "m365 cal list --days 30"
test_cmd "TC-015" "日历 JSON 输出" "m365 cal list --json | jq '.[] | {subject, start: .start.dateTime}'"

echo "创建测试事件..." | tee -a "$LOG_FILE"
test_cmd "TC-016" "创建事件" "m365 cal create 'CLI Test Event' --start '2026-02-17T10:00' --end '2026-02-17T11:00'"

# 获取刚创建的事件 ID
EVENT_ID=$(m365 cal list --days 1 --json 2>/dev/null | jq -r '.[] | select(.subject == "CLI Test Event") | .id' | head -1)
if [ -n "$EVENT_ID" ]; then
    test_cmd "TC-017" "获取事件" "m365 cal get '$EVENT_ID'"
    test_cmd "TC-018" "获取事件 JSON" "m365 cal get '$EVENT_ID' --json | jq '{subject: .subject, start: .start, end: .end}'"
    test_cmd "TC-019" "更新事件" "m365 cal update '$EVENT_ID' --title 'Updated CLI Test'"
    test_cmd "TC-020" "删除事件" "m365 cal delete '$EVENT_ID'"
fi

test_cmd "TC-021" "创建全天事件" "m365 cal create '全天测试' --start '2026-02-18' --end '2026-02-19' --allday"

# 获取并删除全天事件
ALLDAY_EVENT_ID=$(m365 cal list --days 3 --json 2>/dev/null | jq -r '.[] | select(.subject == "全天测试") | .id' | head -1)
if [ -n "$ALLDAY_EVENT_ID" ]; then
    test_cmd "TC-022" "删除全天事件" "m365 cal delete '$ALLDAY_EVENT_ID'"
fi

# 4. OneDrive 功能测试
echo "========== 4. OneDrive (onedrive) ==========" | tee -a "$LOG_FILE"
test_cmd "TC-023" "列出根目录" "m365 od ls"
test_cmd "TC-024" "OneDrive JSON 输出" "m365 od ls --json | jq '.[] | {name, type}'"
test_cmd "TC-025" "列出子目录 (Documents)" "m365 od ls 'Documents'"
test_cmd "TC-026" "创建文件夹" "m365 od mkdir 'cli-test-folder'"

# 创建本地测试文件
echo "This is a test file for m365 cli - $(date)" > /tmp/m365-test.txt

test_cmd "TC-027" "上传文件" "m365 od upload /tmp/m365-test.txt 'cli-test-folder/test.txt'"
test_cmd "TC-028" "验证上传" "m365 od ls 'cli-test-folder'"
test_cmd "TC-029" "获取文件信息" "m365 od get 'cli-test-folder/test.txt'"
test_cmd "TC-030" "获取文件信息 JSON" "m365 od get 'cli-test-folder/test.txt' --json | jq '{name, size, webUrl}'"
test_cmd "TC-031" "下载文件" "m365 od download 'cli-test-folder/test.txt' /tmp/m365-download-test.txt"

# 验证下载内容
if [ -f /tmp/m365-download-test.txt ]; then
    echo "验证下载文件内容:" | tee -a "$LOG_FILE"
    diff /tmp/m365-test.txt /tmp/m365-download-test.txt >> "$LOG_FILE" 2>&1
    if [ $? -eq 0 ]; then
        echo "✅ 下载文件内容一致" | tee -a "$LOG_FILE"
    else
        echo "❌ 下载文件内容不一致" | tee -a "$LOG_FILE"
    fi
fi

test_cmd "TC-032" "搜索文件" "m365 od search 'm365-test'"
test_cmd "TC-033" "创建分享链接 (view)" "m365 od share 'cli-test-folder/test.txt'"
test_cmd "TC-034" "创建编辑权限链接" "m365 od share 'cli-test-folder/test.txt' --type edit"
test_cmd "TC-035" "删除文件" "m365 od rm 'cli-test-folder/test.txt' --force"
test_cmd "TC-036" "删除文件夹" "m365 od rm 'cli-test-folder' --force"

# 5. SharePoint 功能测试
echo "========== 5. SharePoint (sharepoint) ==========" | tee -a "$LOG_FILE"
test_cmd "TC-037" "列出站点" "m365 sp sites"
test_cmd "TC-038" "站点 JSON 输出" "m365 sp sites --json | jq '.[] | {name, webUrl}'"
test_cmd "TC-039" "搜索站点" "m365 sp sites --search 'migration'"

# 获取第一个站点 ID
SITE_ID=$(m365 sp sites --json 2>/dev/null | jq -r '.[0].id')
if [ -n "$SITE_ID" ]; then
    test_cmd "TC-040" "列出站点列表" "m365 sp lists '$SITE_ID'"
    test_cmd "TC-041" "列出站点文件" "m365 sp files '$SITE_ID'"
fi

test_cmd "TC-042" "搜索内容" "m365 sp search 'migration'"
test_cmd "TC-043" "SharePoint 帮助" "m365 sp --help"

# 6. 边界测试
echo "========== 6. 边界测试 ==========" | tee -a "$LOG_FILE"
test_cmd "TC-044" "零结果 (预期失败/空)" "m365 mail list --top 0"
test_cmd "TC-045" "无效邮件 ID" "m365 mail read 'invalid-id'"
test_cmd "TC-046" "不存在的路径" "m365 od ls 'nonexistent-path'"
test_cmd "TC-047" "下载不存在的文件" "m365 od download 'nonexistent-file' /tmp/test.txt"
test_cmd "TC-048" "无效日历 ID" "m365 cal get 'invalid-id'"
test_cmd "TC-049" "无效站点 ID" "m365 sp lists 'invalid-site'"

echo "========================================" | tee -a "$LOG_FILE"
echo "测试完成！详细日志: $LOG_FILE" | tee -a "$LOG_FILE"
echo "========================================" | tee -a "$LOG_FILE"
