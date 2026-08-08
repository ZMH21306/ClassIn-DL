# README 下载链接更新脚本
# 从最新 GitHub Release 读取资产，更新 README.md 中的下载表格
# 用法: .\scripts\update-readme.ps1 [-Owner owner] [-Repo repo]
#
# 说明:
#   - 需要 gh CLI 已登录
#   - 自动匹配 README 中 "## 下载" 表格的每行平台/架构
#   - 更新 GitHub 直链和 CDN 镜像链接

param(
    [string]$Owner,
    [string]$Repo
)

$ErrorActionPreference = "Stop"

# 从 git remote 推断 owner/repo
if (-not $Owner -or -not $Repo) {
    $remote = git remote get-url origin
    if ($remote -match 'github\.com[:/]([^/]+)/([^/]+?)(\.git)?$') {
        $Owner = $Matches[1]
        $Repo = $Matches[2]
    } else {
        Write-Host "❌ 无法推断仓库，请用 -Owner 和 -Repo 指定"
        exit 1
    }
}
Write-Host "仓库: $Owner/$Repo"

# 获取最新 release
$latest = gh release view --repo "$Owner/$Repo" --json tagName,assets 2>$null
if (-not $latest) {
    Write-Host "❌ 无法获取最新 release"
    exit 1
}
$json = $latest | ConvertFrom-Json
$tag = $json.tagName
$assets = $json.assets | ForEach-Object { $_.name }
Write-Host "最新 tag: $tag"
Write-Host "资产文件: $($assets.Count) 个"

# 读取 README
$readmePath = Join-Path (Get-Location) "README.md"
if (-not (Test-Path $readmePath)) {
    Write-Host "❌ 未找到 README.md"
    exit 1
}
$content = Get-Content $readmePath -Raw -Encoding UTF8

# 定位下载表格
$tableMatch = [regex]::Match($content, '(## 下载\n\n.*?)(\n## |\z)', [System.Text.RegularExpressions.RegexOptions]::Singleline)
if (-not $tableMatch.Success) {
    Write-Host "⚠️ 未找到下载表格区域"
    exit 0
}
$oldTable = $tableMatch.Groups[1].Value
$lines = $oldTable -split "`n"
$newLines = @()

foreach ($line in $lines) {
    if ($line -match '^\|\s*(\w+)\s*\|\s*([\w_]+)\s*\|') {
        $platform = $Matches[1]
        $arch = $Matches[2]
        # 在资产中查找匹配文件
        $normPlatform = $platform.ToLower()
        $archVariants = @($arch.ToLower())
        if ($arch -eq 'x86_64') { $archVariants += 'x64' }
        if ($arch -eq 'arm64') { $archVariants += 'arm64' }
        if ($arch -eq 'x86') { $archVariants += 'x86' }

        $matched = $null
        foreach ($asset in $assets) {
            $lower = $asset.ToLower()
            if ($lower -match [regex]::Escape($normPlatform) -and $archVariants | Where-Object { $lower -match [regex]::Escape($_) }) {
                $matched = $asset
                break
            }
        }
        if ($matched) {
            $newGithub = "https://github.com/$Owner/$Repo/releases/download/$tag/$matched"
            $newCdn = "https://gh-proxy.org/$newGithub"
            $line = $line -replace 'https://gh-proxy\.org/https://github\.com/[^\s\)]+', $newCdn
            $line = $line -replace 'https://github\.com/[^\s\)]+', $newGithub
            Write-Host "  ✅ $platform $arch -> $matched"
        } else {
            Write-Host "  ⚠️ $platform $arch 未找到匹配资产"
        }
    }
    $newLines += $line
}

$newTable = $newLines -join "`n"
$newContent = $content.Replace($oldTable, $newTable)

if ($newContent -ne $content) {
    Set-Content $readmePath $newContent -Encoding UTF8
    Write-Host "✅ README.md 已更新"
} else {
    Write-Host "📭 README.md 无变化"
}
