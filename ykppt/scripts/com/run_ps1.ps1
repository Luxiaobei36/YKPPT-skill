# run_ps1.ps1 — 把目标 .ps1 重写为 UTF-8 BOM 再执行
# 用法: powershell -NoProfile -ExecutionPolicy Bypass -File run_ps1.ps1 <目标脚本.ps1>
# 原因: Windows PowerShell 5.1 读无 BOM 的 UTF-8 .ps1 会按 ANSI 解码 -> 中文乱码。
#       多数编辑器/工具写出的是无 BOM UTF-8，故每次执行前强制补 BOM。
param([Parameter(Mandatory=$true)][string]$Path)
$ErrorActionPreference = 'Stop'
if(-not (Test-Path $Path)){ throw "not found: $Path" }
$text = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
$utf8bom = New-Object System.Text.UTF8Encoding($true)   # $true = 带 BOM
[System.IO.File]::WriteAllText($Path, $text, $utf8bom)
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $Path
