# export_png.ps1 — 用 PowerPoint COM 把每一页导出成 PNG
# 被 export_png.py 调用；也可单独运行。
#
# 为什么用英文文件名：中文 Windows 下 PowerPoint 的默认导出名是「幻灯片1.PNG」，
# 脚本里按 Slide1.png 找会扑空。这里显式指定 s01.png。
#
# 用法:
#   powershell -NoProfile -ExecutionPolicy Bypass -File export_png.ps1 `
#       -PptPath out.pptx -OutDir renders -Width 1600 -Height 900

param(
    [Parameter(Mandatory = $true)][string]$PptPath,
    [Parameter(Mandatory = $true)][string]$OutDir,
    [int]$Width = 1600,
    [int]$Height = 900
)

$ErrorActionPreference = "Stop"

$PptPath = (Resolve-Path $PptPath).Path
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Force $OutDir | Out-Null }
$OutDir = (Resolve-Path $OutDir).Path

Get-ChildItem -Path $OutDir -Filter "s*.png" -ErrorAction SilentlyContinue | Remove-Item -Force

$app = $null
$pres = $null
try {
    $app = New-Object -ComObject PowerPoint.Application
    # Open(path, ReadOnly, Untitled, WithWindow)
    $pres = $app.Presentations.Open($PptPath, $true, $false, $false)
    $count = $pres.Slides.Count
    for ($i = 1; $i -le $count; $i++) {
        $target = Join-Path $OutDir ("s{0:00}.png" -f $i)
        $pres.Slides.Item($i).Export($target, "PNG", $Width, $Height)
    }
    Write-Output "EXPORTED $count slide(s) to $OutDir"
}
finally {
    if ($null -ne $pres) { try { $pres.Close() } catch {} }
    if ($null -ne $app) {
        try { $app.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}
    }
    [GC]::Collect()
}
