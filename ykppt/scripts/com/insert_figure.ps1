# insert_figure.ps1 — 把"整页框架图"搬进目标模板页（术语替换 + 缩放 + 同步缩字号 + 删原图页）
#
# 场景：图在源页 $FigSlide，目标模板页 $DstSlide（已含标题/页眉/页脚）。
# 做法：改图内术语 -> 删目标页中间旧内容 -> 组合搬运 -> 等比缩放并同步缩字号 -> 删原图页。
# 运行：先备份 pptx，再 powershell -File run_ps1.ps1 insert_figure.ps1
#
# ==== 按需修改的参数 ====
$ErrorActionPreference = 'Stop'
$pptPath  = 'E:\path\to\your.pptx'
$DstSlide = 5      # 目标模板页
$FigSlide = 6      # 图所在页（搬完会删）
# 目标页中间"要删的旧内容"用 Top 区间界定（保留标题在上、页脚在下）
$DelTopMin = 125.0
$DelTopMax = 500.0
# 图的可用区（960x540 画布，占满整页只留标题时的推荐值）
$availL = 45.0; $availT = 142.0; $availW = 870.0; $availH = 356.0
# 术语替换表（务必先 dump 源图真实文本再按实际写；键=旧词，值=新词）
$repl = @(
  @('旧术语A','新术语A'),
  @('旧术语B','新术语B')
)
$FONT = 'Microsoft YaHei'
# ========================

function RGBv([int]$r,[int]$g,[int]$b){ return ($r + $g*256 + $b*65536) }

# 术语替换：改文字后立刻回写字号与中文字体（Text= 会重置 run 字号）
function Apply-Repl($shape){
  try {
    if($shape.Type -eq 6){ foreach($gi in $shape.GroupItems){ Apply-Repl $gi }; return }  # 6=组
    if($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1){
      $r = $shape.TextFrame.TextRange
      $sz = $null; try { $sz = [double]$r.Font.Size } catch {}
      $t = $r.Text; $orig = $t
      foreach($p in $repl){ $t = $t.Replace($p[0],$p[1]) }
      if($t -ne $orig){
        $r.Text = $t
        if($sz -ne $null -and $sz -gt 0){ $r.Font.Size = [single]$sz }
        $r.Font.Name = $FONT; $r.Font.NameFarEast = $FONT
      }
    }
  } catch {}
}

# 关键：Group 等比缩放不缩字号 -> 递归把组内每个子形状字号 ×k
function ScaleFonts($shape,$k){
  if($shape.Type -eq 6){ foreach($gi in $shape.GroupItems){ ScaleFonts $gi $k }; return }
  try {
    if($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1){
      $cur = [double]$shape.TextFrame.TextRange.Font.Size
      if($cur -gt 0){ $shape.TextFrame.TextRange.Font.Size = [single]($cur * $k) }
    }
  } catch {}
}

$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open($pptPath, $false, $false, $false)
$dst = $pres.Slides.Item($DstSlide)
$fig = $pres.Slides.Item($FigSlide)

# 1) 图内术语替换
foreach($sh in $fig.Shapes){ Apply-Repl $sh }

# 2) 删目标页中间旧内容
$toDel = @()
foreach($sh in $dst.Shapes){ if($sh.Top -ge $DelTopMin -and $sh.Top -lt $DelTopMax){ $toDel += $sh } }
"deleting $($toDel.Count) shapes from dst slide"
foreach($d in $toDel){ $d.Delete() }

# 3) 组合 -> 复制 -> 粘贴到目标页
$g6 = $fig.Shapes.Range().Group()
$g6.Copy()
$pasted = $dst.Shapes.Paste()
$g = $pasted.Item(1)

# 4) 等比缩放 + 居中 + 同步缩字号
$g.LockAspectRatio = -1
$W0 = [double]$g.Width; $H0 = [double]$g.Height
$scale = [math]::Min($availW/$W0, $availH/$H0)
$g.Width = [single]($W0 * $scale)
$g.Left = [single]($availL + ($availW - [double]$g.Width)/2)
$g.Top  = [single]($availT + ($availH - [double]$g.Height)/2)
ScaleFonts $g $scale
"figure: scale=$([math]::Round($scale,3)) W=$([int]$g.Width) H=$([int]$g.Height) L=$([int]$g.Left) T=$([int]$g.Top)"

# 5) 删原图页
$pres.Slides.Item($FigSlide).Delete()

$pres.Save()
$pres.Close()
$app.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
"DONE"
