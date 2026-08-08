# CHANGELOG 自动生成脚本
# 从上次 tag 以来的 commit 自动提取变更，按 Conventional Commits 分类
# 用法: .\scripts\generate-changelog.ps1 [-Version x.y.z] [-Date YYYY-MM-DD]
#
# 分类规则:
#   feat  -> 新增 (Added)
#   fix   -> 修复 (Fixed)
#   perf  -> 性能 (Performance)
#   refactor -> 改进 (Changed)
#   BREAKING CHANGE / ! -> 破坏性变更 (Breaking Changes)

param(
    [string]$Version,
    [string]$Date = (Get-Date -Format "yyyy-MM-dd")
)

$ErrorActionPreference = "Stop"

# 获取上次 tag
$lastTag = git describe --tags --abbrev=0 2>$null
if (-not $lastTag) {
    Write-Host "未找到历史 tag，将提取全部 commit"
    $range = ""
} else {
    Write-Host "上次 tag: $lastTag"
    $range = "$lastTag..HEAD"
}

# 获取 commit 列表
$commits = git log $range --pretty=format:"%s" 2>$null
if (-not $commits) {
    Write-Host "没有新的 commit"
    exit 0
}

# 分类容器
$added = @()
$fixed = @()
$perf = @()
$changed = @()
$breaking = @()

foreach ($line in $commits) {
    # 跳过 release/chore 提交
    if ($line -match '^chore\(release\)') { continue }

    # 检测破坏性变更
    if ($line -match 'BREAKING CHANGE|!') {
        $breaking += $line
        continue
    }

    if ($line -match '^feat') { $added += $line }
    elseif ($line -match '^fix') { $fixed += $line }
    elseif ($line -match '^perf') { $perf += $line }
    elseif ($line -match '^refactor') { $changed += $line }
    # docs/style/test/build/ci/chore 忽略
}

# 生成 Markdown
$output = @()
$output += "## [$Version] - $Date"
$output += ""

if ($breaking.Count -gt 0) {
    $output += "### 破坏性变更"
    $output += ""
    foreach ($c in $breaking) { $output += "- $c" }
    $output += ""
}
if ($added.Count -gt 0) {
    $output += "### 新增"
    $output += ""
    foreach ($c in $added) { $output += "- $c" }
    $output += ""
}
if ($fixed.Count -gt 0) {
    $output += "### 修复"
    $output += ""
    foreach ($c in $fixed) { $output += "- $c" }
    $output += ""
}
if ($perf.Count -gt 0) {
    $output += "### 性能"
    $output += ""
    foreach ($c in $perf) { $output += "- $c" }
    $output += ""
}
if ($changed.Count -gt 0) {
    $output += "### 改进"
    $output += ""
    foreach ($c in $changed) { $output += "- $c" }
    $output += ""
}

$result = $output -join "`n"
Write-Host ""
Write-Host "===== 生成的 CHANGELOG 条目 ====="
Write-Host $result
Write-Host "================================"

# 写入 CHANGELOG.md
$changelogPath = Join-Path (Get-Location) "CHANGELOG.md"
if (Test-Path $changelogPath) {
    $content = Get-Content $changelogPath -Raw -Encoding UTF8
    # 在 [Unreleased] 之后插入新条目
    if ($content -match '## \[Unreleased\]') {
        $content = $content -replace '## \[Unreleased\]', "## [Unreleased]`n`n$result"
    } else {
        $content = "$result`n`n$content"
    }
    Set-Content $changelogPath $content -Encoding UTF8
    Write-Host "✅ CHANGELOG.md 已更新"
} else {
    Write-Host "⚠️ 未找到 CHANGELOG.md，请手动创建"
}
