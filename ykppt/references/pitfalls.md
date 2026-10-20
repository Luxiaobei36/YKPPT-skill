# 踩坑清单（交接文档）

分两部分：

- **A. python-pptx 路线**（当前主路线，`scripts/*.py`）—— 更新到 2026-10
- **B. PowerShell COM 路线**（旧版，`scripts/com/*.ps1`，仍在用）—— 原始记录，仍然有效

环境：Windows 10 中文版 + Python 3.11 + python-pptx 1.0.2 + PowerPoint 16.0。

---

# A. python-pptx 路线

> 这一节的坑有个共同点：**没有任何一条会抛异常**。文件能存盘、页数正确、
> 校验脚本也不报错，但 PowerPoint 打不开，或者内容静默错位。
> 所以每条都记了「怎么发现的」。

## A1. 删页后 `<p14:sectionLst>` 留下悬空 slide id → 文件打不开

现象：`python-pptx` 正常存盘，zip 结构合法、关系表完整、XML 全部可解析，
但 PowerPoint 弹「PowerPoint could not open the file」，不给任何细节。

原因：`presentation.xml` 里有一段 `<p:extLst>` 存着**分节**信息，
`<p14:sectionLst><p14:section><p14:sldIdLst>` 逐个列出 slide id。
只从 `<p:sldIdLst>` 删页、不清理分节列表，就会留下指向已删页的 id。

```python
P14 = "{http://schemas.microsoft.com/office/powerpoint/2010/main}"
for section in prs.part._element.iter(P14 + "section"):
    lst = section.find(P14 + "sldIdLst")
    live = {sid.get("id") for sid in prs.slides._sldIdLst}
    for sid in list(lst):
        if sid.get("id") not in live:
            lst.remove(sid)
```

**注意 `sldId.get("id")` 是 `str`，`Slide.slide_id` 是 `int`。**
两者直接比较永远为 False —— 我第一次修这个 bug 时正是踩在这上面：
代码"看起来"在清理，实际一个都没删，于是花了两轮才排掉。

怎么发现的：把原文件二分（只 round-trip → 只删页 → 删页+删形状），
定位到"删页"这一步；再直接读 `presentation.xml` 的 `extLst` 看到 27 个 id 而
`p:sldIdLst` 只剩 5 个。

## A2. 删形状后 `<p:timing>` 动画轴按 shape id 悬空 → 文件打不开

现象同上，只删页没问题，一删形状就打不开。

原因：`<p:timing>` 里的动画时间轴用 `<p:spTgt spid="N"/>` 引用形状。
形状删了，引用还在。

```python
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
for t in slide._element.findall(P_NS + "timing"):
    slide._element.remove(t)
```

## A3. 深拷贝形状时 `p:tags r:id` 撞上别的 rId → 文件打不开（最阴的一条）

现象：母版里的页型克隆到新页后，**目录页和所有正文页**都打不开，
**封面和结束页**却能开。看起来像"某些页型有问题"。

原因：原 PPT 的每个形状都绑了形状标签（tag part），XML 长这样：

```xml
<p:nvSpPr>
  <p:cNvPr id="5" name="矩形 4"/>
  <p:nvPr>
    <p:custDataLst><p:tags r:id="rId1"/></p:custDataLst>
  </p:nvPr>
</p:nvSpPr>
```

`r:id` 是**在新页里重新分配**的。拷贝过去的 `rId1` 在新页关系表里恰好是
「幻灯片版式」关系——所以**校验 rId 存在也查不出来**，但 `p:tags` 指向
的是 slideLayout 而不是 tags part，PowerPoint 直接判损坏。

更坑的是：`p:tags` 少了 `r:id` 是非法 XML，不能只删属性，得整块容器删。

```python
for cd in list(el.iter(qn("p:custDataLst"))):
    cd.getparent().remove(cd)
```

标签只是「形状分组/标记」的元数据，渲染无关，整块剥掉最稳。

封面/结束页没有 tag 关系所以能开 —— 这就是"只有某些页型坏"的真相。

怎么发现的：把 9 个页型各建一个单页文件分别试开，拿到 FAIL 集合
（toc + 全部正文页）；再比对 OK 文件与 FAIL 文件的 `p:nvPr` 内容，
看到只有 FAIL 那边有 `custDataLst/p:tags`。

## A4. 克隆页时把页码占位符也过滤掉了 → 所有正文页没页码

现象：生成成功、没报错，但每页右下角空白。

原因：`add_slide()` 只会实例化版式里**有 `type` 的**占位符；
页码占位符在版式里存在，但不会被带进新页的 shape 树。
于是"过滤掉所有占位符"这个思路把页码一起杀了。

```python
ph = el.find(".//" + qn("p:ph"))
if ph is not None and ph.get("type") != "sldNum":
    continue      # 内容/标题占位符不复制（会撞 idx），页码要复制
```

## A5. 组合内子形状的名字改不到 → 内容静默不写入

现象：封面「汇报人」永远是模板默认的"某某某"，
`build_deck.py` 里 `presenter` 字段填了也无效，**且没有任何警告**。

原因：`slide.shapes` 只遍历顶层；「汇报人」文本框在「汇报人组合」这个
group 里面。给母版打角色名时漏了递归，渲染时按名字找也就找不到。

```python
def iter_all(shapes):          # 顺序遍历必须穿透 group
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:      # GROUP
            yield from iter_all(sh.shapes)
```

**并且**：找不到形状时不要静默 `return None`，要报警。
静默是这里最危险的失败模式 —— 生成"成功"、页数正确、无异常，
但那一块其实是模板默认文案。`ykrender._require()` 就是为此而写。

## A6. `<a:ea>` 缺 `charset="-122"` → 中文少一个字就多折一行

现象：结尾页「请各位老师同学批评指正！」的 `！` 单独掉到第二行，
正好压在下面的分隔线上，看起来像一根竖线加一个点。

原因：`<a:ea typeface="微软雅黑" charset="-122"/>` 里的 `charset`
是 GB2312 的中文字符集标识。只写 `typeface`，PowerPoint 会挑到另一款
**更宽**的 CJK 字形，同样 48pt 的文字就多折一行。

另外 `rPr` 上要带 `lang="zh-CN"`，中文标点悬挂/压缩规则才会生效：

```python
rPr.set("lang", "zh-CN")
rPr.set("altLang", "en-US")
ea.set("typeface", "微软雅黑")
ea.set("charset", "-122")
```

顺带记一条 schema 约束：`a:rPr` 的子元素是**有序**的
（`a:solidFill` → `a:latin` → `a:ea` → `a:cs`），
`a:ea` 必须紧跟 `a:latin`，`append` 到末尾会被判非法。

## A7. 用 `**文字**` 做强调标记 → YAML 直接扫描失败

```
- **异构协同性**：一个任务通常由多类资源……
  ^
ScannerError: while scanning an alias, expected alphabetic or numeric character
```

YAML 把行首 `**` 当成别名（anchor/alias）符号。
中文稿里"列表项以加粗词开头"是极常见的写法，所以这个坑必踩。

换成 `〔文字〕`（中文方头括号，输入法可直接打），在 YAML 任何位置都安全。
渲染器同时兼容 `**` 和 `==`，但文档只推荐 `〔〕`。

## A8. 控制台 GBK 打不出 `✓`

现象：文件已经生成好了，最后一行 `print("✓ 已生成")` 抛
`UnicodeEncodeError: 'gbk' codec can't encode character '✓'`，
整个进程非零退出，看起来像"生成失败"。

对策：统一用 ASCII 标记 `[OK]` / `[!]`，并把 stdout/stderr 设成
`errors="replace"` 兜底（见 `scripts/_console.py`）。

## A9. 单行文字被按段落行距误判溢出

校验器用"行数 × 字号 × 行距"估高度。但单行文本的实际高度是字体自身行高
（约 1.2em），和 `line_spacing` 无关。54pt 的「目录」按 1.5 倍算就
被误报成溢出。

```python
eff = spacing if n > 1 else 1.2
```

## A10. 组合缩放系数没算，坐标校验全错

组合内子形状的 `left/width` 是**子坐标系**的值。封面「汇报人」在母版里
只有 2.31in 宽，但所在组合横向缩放 2.87 → 实际画出来 6.78in。
不乘这个系数，会把本来放得下的文字全判成溢出。

```python
xfrm = group._element.find(".//" + qn("a:xfrm"))
sx = int(xfrm.find(qn("a:ext")).get("cx")) / int(xfrm.find(qn("a:chExt")).get("cx"))
```

## A11. `Slide` 没有 `.rels`，是 `.part.rels`

小事，但报错信息（`'Slide' object has no attribute 'rels'`）会把人往
"是不是要用 `slide.part.rels`"以外的地方带。

---

# B. PowerShell COM 路线（旧版，仍然有效）

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
  **新版 `export_png.ps1` 显式指定 `s01.png`** 就是为了绕开这个。
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

## 7. 一句话 checklist

**python-pptx 路线（主）**
1. 删页 → 清 `p14:sectionLst`；删形状 → 先删 `p:timing`。
2. 克隆形状 → 剥 `p:custDataLst`（tag 引用会撞 rId，且校验查不出）。
3. 页码占位符 `type="sldNum"` 要显式复制；内容占位符不要复制。
4. 找形状要**递归穿透 group**；找不到要**报警**，别静默返回 None。
5. 中文 run 要带 `a:ea charset="-122"` 和 `rPr lang="zh-CN"`。
6. YAML 里用 `〔〕` 标强调，别用 `**`。
7. 生成后必导出 PNG 逐页看。

**PowerShell COM 路线（旧）**
8. 探测 python-pptx / pywin32 / soffice / POWERPNT，决定路线。
9. 脚本每次执行前**重写为 UTF-8 BOM**。
10. 所有 COM 赋值：颜色/对齐/枚举 `[int]`，磅值/字号 `[single]`，布尔 `-1/0`。
11. 枚举值查准（VerticalAnchor 无 2）。
12. **搬别人给的整页图**：先 dump 组内所有文本再按实际文本写 `Replace`；
    `Group` 等比缩放后**必须递归同步缩字号**；改文字后**立刻回写字号**。
13. **全局改字号**要递归进组内子形状**和表格每个单元格**（表格 `Type=19`
    走 `Table.Cell().Shape`）。
