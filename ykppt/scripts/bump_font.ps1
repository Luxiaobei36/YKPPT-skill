# bump_font.ps1 — 全篇字号 +N（含组内子形状 与 表格每个单元格）
# 改前请先备份 pptx。改后用 export_check.ps1 逐页看有无溢出。
$ErrorActionPreference = 'Stop'
$pptPath = 'E:\path\to\your.pptx'
$Delta   = 2      # 每处字号增量（负数即缩小）

$script:bumped = 0
function Process($shape){
  $ty = $shape.Type
  if($ty -eq 6){ foreach($gi in $shape.GroupItems){ Process $gi }; return }   # 6=组
  if($ty -eq 19){                                                            # 19=表格
    try {
      $tbl = $shape.Table
      for($r=1;$r -le $tbl.Rows.Count;$r++){
        for($c=1;$c -le $tbl.Columns.Count;$c++){
          $rng = $tbl.Cell($r,$c).Shape.TextFrame.TextRange
          $cur = [double]$rng.Font.Size
          if($cur -gt 0){ $rng.Font.Size = [single]($cur + $Delta); $script:bumped++ }
        }
      }
    } catch {}
    return
  }
  try {
    if($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1){
      $rng = $shape.TextFrame.TextRange
      $cur = [double]$rng.Font.Size
      if($cur -gt 0){ $rng.Font.Size = [single]($cur + $Delta); $script:bumped++ }
    }
  } catch {}
}

$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open($pptPath, $false, $false, $false)
for($i=1;$i -le $pres.Slides.Count;$i++){
  foreach($sh in $pres.Slides.Item($i).Shapes){ Process $sh }
}
"fonts bumped = $($script:bumped)"
$pres.Save(); $pres.Close(); $app.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
"DONE"
