# export_check.ps1 — 导出为 PNG 供逐页目视检查
# SaveAs 成功 != 排版正确；溢出/重叠/乱码只有看图才知道。
$ErrorActionPreference = 'Stop'
$pptPath = 'E:\path\to\your.pptx'
$outDir  = 'E:\path\to\out_png'
$Only    = 0     # 0=导出全部；否则只导该页号

if(Test-Path $outDir){ Remove-Item $outDir -Recurse -Force }
New-Item -ItemType Directory -Path $outDir | Out-Null

$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open($pptPath, $true, $false, $false)   # 只读
"slides = $($pres.Slides.Count)"
if($Only -gt 0){
  $pres.Slides.Item($Only).Export(("{0}\s{1}.png" -f $outDir,$Only), "PNG", 1600, 900)
} else {
  for($i=1;$i -le $pres.Slides.Count;$i++){
    $pres.Slides.Item($i).Export(("{0}\s{1}.png" -f $outDir,$i), "PNG", 1600, 900)
  }
}
$pres.Close(); $app.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
"EXPORTED -> $outDir"
