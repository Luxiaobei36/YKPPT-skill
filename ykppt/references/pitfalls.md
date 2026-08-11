# 用 PowerShell COM 生成中文 PPTX 的踩坑总结（交接文档）

> 场景：Windows + Windows PowerShell 5.1 + 已装 Microsoft PowerPoint（Office16），
> **无网络**（pip / pywin32 都装不了），需要用代码批量生成一份**含中文**的 .pptx。
> 本文件记录本次实际踩到的每一个坑与最终可用做法，供下次直接复用，避免重犯。

---

## 0. 环境探测结论（先确认再动手）

| 能力 | 本机结果 | 影响 |
|---|---|---|
| `python-pptx` | ❌ 无，且 pip 装不了（无网络 `No matching distribution`） | 不能走 python-pptx 路线 |
| `pywin32`（win32com） | ❌ 无 | 不能用 Python 驱动 COM |
| PowerPoint COM | ✅ 有（`POWERPNT.EXE` 存在） | **最终方案：PowerShell 直接 COM 自动化** |
| LibreOffice `soffice` | ❌ 无 | 不能用 headless 转换 |

**教训**：动手前先探测 `python-pptx / pywin32 / soffice / POWERPNT.EXE` 四者，别默认某条路可行。无网络环境几乎注定要走 PowerShell + PowerPoint COM。

---

## 1. 最大的坑：中文 .ps1 必须 UTF-8 **带 BOM**

- Windows PowerShell 5.1 读取 `.ps1` 时，**无 BOM 的 UTF-8 会被当成本地 ANSI 代码页解码**，脚本里的中文字符串全部变成乱码写进 PPT。
- 本仓库的 `Write` 工具写出的是**无 BOM** UTF-8 → 直接跑必乱码。
- **每次用 Write/Edit 改完脚本，都要重新写入 BOM 再执行**（Edit 也会把文件恢复成无 BOM）。

可复用的"编码+执行"片段：

```powershell
$p = "…\build_ppt.ps1"
$text = [System.IO.File]::ReadAllText($p, [System.Text.Encoding]::UTF8)
$utf8bom = New-Object System.Text.UTF8Encoding($true)   # $true = 带 BOM
[System.IO.File]::WriteAllText($p, $text, $utf8bom)
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $p
```

> 顺带：PowerShell 5.1 里 `Set-Content`/`Out-File` 的默认编码也不可靠，写给别的程序读的文件一律显式 `-Encoding utf8`（但注意 5.1 的 `utf8` 是带 BOM 的）。

---

## 2. COM 属性对 .NET 类型极其挑剔 → 显式强转

报错 `Specified cast is not valid.`（InvalidCastException）几乎都是**把 PowerShell 默认类型塞给 COM 属性**导致。逐个属性确认期望类型并强转：

| 属性 | 需要的类型 | 写法 |
|---|---|---|
| `Fill.ForeColor.RGB` / `Line.ForeColor.RGB` / `Font.Color.RGB` | `[int]`（32 位整数） | `= [int]$color` |
| `Line.Weight` | `[single]`（浮点磅值） | `= [single]$w` |
| `Font.Size` | `[single]` | `= [single]$size` |
| `ParagraphFormat.Alignment` | `[int]`（枚举值） | `= [int]$align` |

**教训**：不要依赖 PowerShell 的隐式转换。凡是给 COM 属性赋值，`RGB/Alignment/枚举 → [int]`，`Weight/Size/度量 → [single]`。

**排错方法**（有效）：报错信息只给行号不给具体调用时，给公共 helper 里的赋值语句包一层 `try{…}catch{ Write-Host "FAIL …$val ($($val.GetType().Name))"; throw }`，一次跑出到底是哪一次调用、哪个值、什么类型挂了。**不要靠猜**——本次就是靠这个把"以为是 ForeColor 的问题"纠正为"其实是下一句 Line.Weight"。

---

## 3. 枚举值不能乱用（VerticalAnchor 等）

- `TextFrame.VerticalAnchor` 合法值：`msoAnchorTop=1`、`msoAnchorMiddle=3`、`msoAnchorBottom=4`。**没有 2**，传 2 直接 `指定值已被否决 / ArgumentException`。
- 我原来用 2 表示"垂直居中"是想当然。做了个映射 `switch($anchor){2{3}3{3}4{4}default{1}}` 兜底。
- **教训**：Office 枚举不是连续的，用前查准官方常量值（`msoAnchor*`、`ppAlign*`、`msoShape*`）。常用：
  - 段落对齐 `ppParagraphAlignment`：左=1、居中=2、右=3。
  - 形状 `AddShape`：矩形=1、圆角矩形=5。
  - `Font.Bold`：真=-1、假=0（COM 布尔是 `-1/0`，不是 `$true/$false`）。

---

## 4. 颜色是 BGR 打包的整数

PowerPoint 的 `.RGB` 实际是 `R + G*256 + B*65536`（即低字节为 R，等价 BGR 排列）。封个函数别手算：

```powershell
function RGBv([int]$r,[int]$g,[int]$b){ return ($r + $g*256 + $b*65536) }
```

---

## 5. 生成后必须**导出 PNG 逐页目视检查**

- "SaveAs 成功" ≠ "排版正确"。文本溢出、卡片重叠、乱码只有看图才知道。
- 用 COM 导出整份为 PNG，再逐张 Read 检查：

```powershell
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open($pptPath, $true, $false, $false)  # ReadOnly, Untitled, WithWindow=false
$pres.Export($outDir, "PNG", 1280, 720)
$pres.Close(); $app.Quit()
```

- 导出的文件名是本地化的（中文系统下叫 `幻灯片1.PNG`…），别假设是 `Slide1.png`。
- **注意**：低分辨率缩略图里形近字容易看走眼（本次把正确的"硬拦截"误判成"硬兰截"）。要判定是不是错别字，**回源脚本 `Grep` 确认**，不要只凭缩略图下结论。

---

## 5.5 组合（Group）等比缩放**不会自动缩字号** → 必须手动同步

- 把一张图里的所有形状 `Shapes.Range().Group()` 组合后，改 `group.Width/Height`（哪怕 `LockAspectRatio=-1` 等比缩）**只缩形状外框，组内文本框的 `Font.Size` 原封不动**。
- 后果：文本框被缩窄缩矮，但字还是原字号 → 在小框里被迫换行、溢出、层层重叠（本次把一张整页图缩到 0.75 倍后，18pt 的字挤成一团、完全不可读）。
- **正确做法**：缩放组之后，递归遍历组内每个子形状，把字号乘以同一缩放系数：

```powershell
function ScaleFonts($shape,$k){
  if($shape.Type -eq 6){ foreach($gi in $shape.GroupItems){ ScaleFonts $gi $k }; return }  # 6 = msoGroup
  try {
    if($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1){
      $cur = [double]$shape.TextFrame.TextRange.Font.Size
      if($cur -gt 0){ $shape.TextFrame.TextRange.Font.Size = [single]($cur * $k) }
    }
  } catch {}
}
$scale = [math]::Min($availW/$W0, $availH/$H0)
$g.Width = [single]($W0 * $scale)   # 等比
ScaleFonts $g $scale                 # ← 关键：字号同步缩
```

- 排错要点：拿不准就先**探针**一个子形状 `$_.GroupItems.Item(1).TextFrame.TextRange.Font.Size` 打印出来，会看到缩放后它还是 18 —— 印证"外框缩了、字号没缩"。
- 反过来同理：整体放大（如"全篇字体 +2 号"）也要递归到组内子形状 + **表格每个单元格**（`Type=19` 的表格要走 `Table.Cell(r,c).Shape.TextFrame`，普通遍历碰不到单元格），否则组里/表里的字被漏掉。

---

## 5.6 "整页设计"的图**塞不进"保留标题+横幅"的窄条**

- 别人给的框架图，往往是按**占满一整页**画的（本次实测：形状从 L58 铺到 887、T38 铺到 512，字号统一 18pt）。
- 若目标页要"保留顶部大标题 + 底部深蓝横幅"，留给图的中间区只剩 ~300pt 高的窄条 → 图被迫缩到 0.6 倍、字掉到 ~11pt，**必糊**。
- **权衡要先跟用户确认**，别默认"保标题+横幅"：
  1. **图占满整页**（推荐）：只留顶部 section 标签 + 大标题，把横幅那句话精简成标题下一行小字，图从标题下铺到页脚上方（~640×360，字号回到 ~14pt，清晰）。
  2. 保标题+横幅、图缩中间：最省事但图明显变小。
  3. 图单独成一页：效果最好但总页数 +1。
- 经验：整页图 + 只留标题（做法 1）时，可用区约 `x 45..915 / y 142..498`（960×540 画布），等比缩放系数约 0.75，配合 5.5 的字号同步缩，最终清晰不溢出。

---

## 5.7 改图内术语：**改完文字后字号会被 COM 重置** → 立刻回写

- 对组内文本做 `TextRange.Text = 新文本` 替换后，PowerPoint 有时会把该 run 的 `Font.Size` / 字体名重置为默认。
- 做法：替换前先读旧字号 `$sz=[double]$r.Font.Size`，`Text=` 之后立刻 `$r.Font.Size=[single]$sz` 并重设 `Font.Name/NameFarEast`（中文字体两者都要设，见 §6）。
- 另外：**替换用的旧词要对准图里的真实文本**——先把组内所有子形状的文本原样 dump 出来（`GroupItems` 逐个打印 `TextRange.Text`），按 dump 结果写 `Replace` 键，别照记忆猜（本次"验证层→物理执行层""模块 A Reward Coder→受限奖励接口"等都是先 dump 再逐条替换才对上的）。

---

## 6. 其他实务约定

- 画布尺寸用 `PageSetup.SlideWidth/Height` 设成 `960×540`（16:9，单位 point），所有坐标按此排。
- `Presentations.Add(0)` 的 0 = 不显示窗口（`msoFalse`）；跑批量时更稳。
- `SaveAs($path, 24)`：`24 = ppSaveAsOpenXMLPresentation`（.pptx）。
- 收尾释放 COM：`[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)`，否则可能残留 `POWERPNT` 进程。
- 多行文本：COM 段落分隔符是 `[char]13`（回车），不是 `` `n ``。封 `NL(@('行1','行2'))` join 一下。
- `AddTextbox(1,…)` 首参 1 = 横排文本框（`msoTextOrientationHorizontal`）。
- 中文字体同时设 `Font.Name` 和 `Font.NameFarEast`（本次用 `Microsoft YaHei`），否则中文可能回退到默认字体。
- helper 函数返回形状对象时，脚本里用 `| Out-Null` 吞掉，避免管道把 COM 对象打印一屏。

---

## 7. 一句话 checklist（下次照做）

1. 探测 python-pptx / pywin32 / soffice / POWERPNT，决定路线。
2. 用 PowerShell COM；脚本每次执行前**重写为 UTF-8 BOM**。
3. 所有 COM 赋值：颜色/对齐/枚举 `[int]`，磅值/字号 `[single]`，布尔 `-1/0`。
4. 枚举值查准（VerticalAnchor 无 2）。
5. 生成后**导出 PNG 逐页看**；疑似错别字回源 `Grep` 核对。
6. 结束 `ReleaseComObject` + `Quit`。
7. **搬别人给的整页图**：先 dump 组内所有文本再按实际文本写 `Replace`；`Group` 等比缩放后**必须递归同步缩字号**（外框缩、字号不缩）；改文字后**立刻回写字号**；"整页图塞进保留标题+横幅的窄条"会糊，放法先问用户。
8. **全局改字号**（如 +2 号）要递归进组内子形状**和表格每个单元格**（表格 `Type=19` 走 `Table.Cell().Shape`），普通遍历漏掉它们。改前先备份，改后逐页看有无溢出。
