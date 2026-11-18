# Wrapper so the scheduled task always has OPENAI_BASE_URL set correctly,
# without relying on a durable env var that may not have propagated.
$env:OPENAI_BASE_URL = "https://us.api.openai.com/v1"
& "C:\Users\mivogel\AppData\Local\Programs\Python\Python312\python.exe" "$PSScriptRoot\daily_news_summary.py"
