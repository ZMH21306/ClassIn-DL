# 启动脚本
# 统一启动流程，参考 AI-IDE scripts/ 结构

param(
    [string]$Action = "start"
)

switch ($Action) {
    "start" { dotnet run --project . }
    "build" { .\build\build.ps1 }
    "test" { dotnet test }
    "release" { Write-Host "触发 CI 发布流程：推送 v* 标签" }
    default { Write-Host "未知操作: $Action" }
}
