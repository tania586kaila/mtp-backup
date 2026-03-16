# start_proxy.ps1
# 启动 DeepSeek 代理服务（供任务计划程序调用）
$env:DEEPSEEK_API_KEY = [System.Environment]::GetEnvironmentVariable("DEEPSEEK_API_KEY", "User")
Set-Location "C:\Users\Administrator\mtp-backup"
python deepseek_proxy.py
