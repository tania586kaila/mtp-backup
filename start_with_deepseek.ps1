# start_with_deepseek.ps1
# 一键启动：设置环境变量 + 启动代理 + 打开 VS Code
# 用法: .\start_with_deepseek.ps1 -ApiKey "sk-your-deepseek-key"

param(
    [string]$ApiKey = $env:DEEPSEEK_API_KEY,
    [int]$Port = 8742,
    [string]$ProjectDir = "C:\Users\Administrator\mtp-backup"
)

if (-not $ApiKey) {
    Write-Host "错误：请提供 DeepSeek API Key" -ForegroundColor Red
    Write-Host "用法: .\start_with_deepseek.ps1 -ApiKey 'sk-your-key'"
    Write-Host "或先设置: `$env:DEEPSEEK_API_KEY='sk-your-key'"
    exit 1
}

# 1. 设置本次会话的环境变量
$env:DEEPSEEK_API_KEY    = $ApiKey
$env:ANTHROPIC_BASE_URL  = "http://127.0.0.1:$Port"
$env:ANTHROPIC_API_KEY   = "sk-proxy-placeholder"  # Claude Code 需要此变量存在

Write-Host "环境变量已设置:" -ForegroundColor Green
Write-Host "  ANTHROPIC_BASE_URL = $env:ANTHROPIC_BASE_URL"
Write-Host "  DEEPSEEK_API_KEY   = $($ApiKey.Substring(0,8))..."

# 2. 后台启动代理
Write-Host "`n启动 DeepSeek 代理..." -ForegroundColor Cyan
$proxy = Start-Process python -ArgumentList "$ProjectDir\deepseek_proxy.py" `
    -PassThru -WindowStyle Minimized `
    -Environment @{ DEEPSEEK_API_KEY = $ApiKey }

Write-Host "代理 PID: $($proxy.Id)"
Start-Sleep -Seconds 2

# 3. 验证代理是否启动
try {
    $health = Invoke-RestMethod "http://127.0.0.1:$Port/health" -TimeoutSec 3
    Write-Host "代理健康检查: OK (backend=$($health.backend), model=$($health.model))" -ForegroundColor Green
} catch {
    Write-Host "代理启动失败，请检查日志" -ForegroundColor Red
    exit 1
}

# 4. 打开 VS Code（继承当前环境变量）
Write-Host "`n打开 VS Code..." -ForegroundColor Cyan
code $ProjectDir

Write-Host "`n完成！Claude Code 现在将使用 DeepSeek 模型" -ForegroundColor Green
Write-Host "关闭此窗口将停止代理服务 (PID: $($proxy.Id))"
Write-Host "按 Ctrl+C 停止代理..."

try {
    $proxy.WaitForExit()
} finally {
    if (-not $proxy.HasExited) {
        $proxy.Kill()
        Write-Host "代理已停止"
    }
}
