# 统一发布脚本
# 当用户说"发布新版本"时，AI 按此流程执行，确保 3 个项目行为一致
# 用法: .\scripts\release.ps1 [-Version x.y.z] [-Type major|minor|patch] [-CommitOnly]
#
# 流程:
#   1. 检查工作区干净
#   2. 分析 commit 确定版本号（若未指定）
#   3. 更新版本文件（Cargo.toml / tauri.conf.json / .csproj）
#   4. 生成 CHANGELOG 条目
#   5. 提交版本更新
#   6. 打 tag 并推送（触发 CI）

param(
    [string]$Version,
    [ValidateSet("major", "minor", "patch")]
    [string]$Type = "patch",
    [switch]$CommitOnly
)

$ErrorActionPreference = "Stop"

# ---------- 0. 辅助函数 ----------
function Get-DefaultBranch {
    $ref = git symbolic-ref refs/remotes/origin/HEAD 2>$null
    if ($ref -match 'refs/remotes/origin/(.+)') { return $matches[1] }
    $branch = git branch --show-current 2>$null
    if ($branch) { return $branch }
    return "main"
}

# ---------- 1. 检查工作区 ----------
Write-Host "=== 检查工作区 ==="
$status = git status --porcelain
if ($status) {
    Write-Host "⚠️ 工作区有未提交变更："
    Write-Host $status
    Write-Host "请先提交代码，再执行发布。"
    exit 1
}
Write-Host "✅ 工作区干净"

# ---------- 2. 确定版本号 ----------
if (-not $Version) {
    # 从 Cargo.toml 或 .csproj 读取当前版本
    $current = $null
    if (Test-Path "Cargo.toml") {
        $current = (Select-String -Path "Cargo.toml" -Pattern '^version = "([^"]+)"' | Select-Object -First 1).Matches[0].Groups[1].Value
    } elseif (Test-Path "*.csproj") {
        $csproj = Get-ChildItem "*.csproj" | Select-Object -First 1
        $current = (Select-String -Path $csproj.FullName -Pattern '<Version>([^<]+)</Version>' | Select-Object -First 1).Matches[0].Groups[1].Value
    }
    if (-not $current) {
        Write-Host "❌ 无法读取当前版本，请用 -Version 指定"
        exit 1
    }
    Write-Host "当前版本: $current"

    # 分析 commit 类型
    $lastTag = git describe --tags --abbrev=0 2>$null
    $range = if ($lastTag) { "$lastTag..HEAD" } else { "" }
    $commits = git log $range --pretty=format:"%s" 2>$null

    $hasBreaking = $commits | Where-Object { $_ -match 'BREAKING CHANGE|!' }
    $hasFeat = $commits | Where-Object { $_ -match '^feat' }
    $hasFix = $commits | Where-Object { $_ -match '^fix|^perf' }

    $parts = $current -split '\.'
    $major = [int]$parts[0]; $minor = [int]$parts[1]; $patch = [int]$parts[2]

    if ($hasBreaking) {
        $major++; $minor = 0; $patch = 0
        $Type = "major"
    } elseif ($hasFeat) {
        $minor++; $patch = 0
        $Type = "minor"
    } elseif ($hasFix) {
        $patch++
        $Type = "patch"
    } else {
        Write-Host "⚠️ 无功能/修复变更，版本号不变"
        exit 0
    }
    $Version = "$major.$minor.$patch"
}
Write-Host "✅ 新版本号: v$Version"

# ---------- 3. 更新版本文件 ----------
Write-Host "=== 更新版本文件 ==="
if (Test-Path "Cargo.toml") {
    (Get-Content "Cargo.toml" -Raw) -replace '^version = "[^"]+"', "version = `"$Version`"" | Set-Content "Cargo.toml" -Encoding UTF8
    Write-Host "✅ Cargo.toml -> $Version"
}
if (Test-Path "tauri.conf.json") {
    (Get-Content "tauri.conf.json" -Raw) -replace '"version": "[^"]+"', "`"version`": `"$Version`"" | Set-Content "tauri.conf.json" -Encoding UTF8
    Write-Host "✅ tauri.conf.json -> $Version"
}
Get-ChildItem "*.csproj" -ErrorAction SilentlyContinue | ForEach-Object {
    (Get-Content $_.FullName -Raw) -replace '<Version>[^<]+</Version>', "<Version>$Version</Version>" | Set-Content $_.FullName -Encoding UTF8
    Write-Host "✅ $($_.Name) -> $Version"
}

# ---------- 4. 生成 CHANGELOG ----------
Write-Host "=== 生成 CHANGELOG ==="
if (Test-Path "scripts\generate-changelog.ps1") {
    & ".\scripts\generate-changelog.ps1" -Version $Version
}

# ---------- 5. 提交 ----------
Write-Host "=== 提交版本更新 ==="
$filesToStage = @()
if (Test-Path "Cargo.toml") { $filesToStage += "Cargo.toml" }
if (Test-Path "tauri.conf.json") { $filesToStage += "tauri.conf.json" }
Get-ChildItem "*.csproj" -ErrorAction SilentlyContinue | ForEach-Object { $filesToStage += $_.FullName }
if (Test-Path "CHANGELOG.md") { $filesToStage += "CHANGELOG.md" }
git add @filesToStage
git commit -m "chore(release): bump version to v$Version"

if ($CommitOnly) {
    Write-Host "✅ CommitOnly 模式：已提交但尚未打 tag"
    exit 0
}

$defaultBranch = Get-DefaultBranch
Write-Host "目标分支: $defaultBranch"
git push origin $defaultBranch

# ---------- 6. 打 tag 并推送 ----------
Write-Host "=== 打 tag 并推送 ==="
$tagRef = "v$Version"
if (git tag -l $tagRef) {
    Write-Host "❌ tag $tagRef 已存在，停止发布以避免覆盖远端 tag"
    Write-Host "如需重新发布，请先删除本地 tag：git tag -d $tagRef"
    exit 1
}
git tag -a $tagRef -m "Release $tagRef"
git push origin $tagRef

Write-Host ""
Write-Host "🎉 发布流程完成！CI 已触发，等待构建完成。"
$repo = git remote get-url origin | ForEach-Object { $_ -replace '.*github.com[:/]', '' -replace '\.git$', '' }
Write-Host "查看: https://github.com/$repo/releases"
