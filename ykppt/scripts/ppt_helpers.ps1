# ppt_helpers.ps1 — 可 dot-source 的公共函数库
# 用法: . "$PSScriptRoot\ppt_helpers.ps1"
# 约定画布 960x540（16:9，单位 point）。所有 COM 赋值均已显式强转。

# 颜色：PowerPoint 的 .RGB 是 BGR 打包整数 = R + G*256 + B*65536
function RGBv([int]$r,[int]$g,[int]$b){ return ($r + $g*256 + $b*65536) }

$script:FONT = 'Microsoft YaHei'   # 中文字体；Name 与 NameFarEast 都要设

# 多行文本：COM 段落分隔是回车符 [char]13，不是 `n
function NL($arr){ return ($arr -join [char]13) }

# 加文本框。anchor: 1=Top 3=Middle 4=Bottom（注意没有 2，此处把 2 也映射到 3）
# align: 1=左 2=中 3=右
function Add-Text {
  param($slide,$l,$t,$w,$h,$text,$size=14,$bold=$false,$color=0,$align=1,$anchor=1)
  $tb = $slide.Shapes.AddTextbox(1,$l,$t,$w,$h)   # 1 = 横排
  $tb.TextFrame.WordWrap = -1
  $va = switch($anchor){ 2 {3} 3 {3} 4 {4} default {1} }
  $tb.TextFrame.VerticalAnchor = $va
  $tb.TextFrame.MarginLeft = 3; $tb.TextFrame.MarginRight = 3
  $tb.TextFrame.MarginTop = 2; $tb.TextFrame.MarginBottom = 2
  $tr = $tb.TextFrame.TextRange
  $tr.Text = $text
  $tr.Font.Size = [single]$size
  $tr.Font.Name = $script:FONT
  $tr.Font.NameFarEast = $script:FONT
  $tr.Font.Bold = $(if($bold){-1}else{0})
  $tr.Font.Color.RGB = [int]$color
  $tr.ParagraphFormat.Alignment = [int]$align
  return $tb
}

# 加矩形/圆角矩形卡片。line<0 表示无边框；round=$true 用圆角(5)否则矩形(1)
function Add-Box {
  param($slide,$l,$t,$w,$h,$fill,$line=-1,$lineW=1.0,$round=$true)
  $shape = if($round){5}else{1}
  $sh = $slide.Shapes.AddShape($shape,$l,$t,$w,$h)
  $sh.Fill.ForeColor.RGB = [int]$fill
  if($line -ge 0){ $sh.Line.Visible=-1; $sh.Line.ForeColor.RGB=[int]$line; $sh.Line.Weight=[single]$lineW }
  else { $sh.Line.Visible = 0 }
  $sh.Shadow.Visible = 0
  if($round){ try { $sh.Adjustments.Item(1) = 0.06 } catch {} }
  return $sh
}

# 细色条（实心矩形），常用作顶部强调条 / 标题下划线
function Add-Bar {
  param($slide,$l,$t,$w,$h,$fill)
  $sh = $slide.Shapes.AddShape(1,$l,$t,$w,$h)
  $sh.Fill.ForeColor.RGB = [int]$fill
  $sh.Line.Visible = 0
  $sh.Shadow.Visible = 0
  return $sh
}

# 新建空白页（版式 12 = ppLayoutBlank），白底
function New-Slide($pres){
  $s = $pres.Slides.Add($pres.Slides.Count+1,12)
  $s.FollowMasterBackground = 0
  $s.Background.Fill.Solid()
  $s.Background.Fill.ForeColor.RGB = (RGBv 255 255 255)
  return $s
}

# 启动 PowerPoint 并建一份 960x540 演示（0 = 不显示窗口）
function New-Deck {
  $app = New-Object -ComObject PowerPoint.Application
  $pres = $app.Presentations.Add(0)
  $pres.PageSetup.SlideWidth = 960
  $pres.PageSetup.SlideHeight = 540
  return @{ app = $app; pres = $pres }
}

# 收尾：保存为 .pptx(24) 并释放 COM
function Save-Close($ctx,$outPath){
  if(Test-Path $outPath){ Remove-Item $outPath -Force }
  $ctx.pres.SaveAs($outPath, 24)   # 24 = ppSaveAsOpenXMLPresentation
  $ctx.pres.Close()
  $ctx.app.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ctx.app) | Out-Null
}
