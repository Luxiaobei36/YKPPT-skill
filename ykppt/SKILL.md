---
name: ykppt
description: >-
  用 Windows PowerShell + PowerPoint COM 在无网络环境批量生成/编辑含中文的学术风格
  .pptx（封面、分节页眉、卡片、表格、框架图搬运与术语替换、全局改字号）。当需要在
  Windows 上不依赖 python-pptx / pywin32 / LibreOffice 生成或修改中文 PPT、
  或把外来整页框架图塞进已有模板页时，使用本 skill。
---

# ykppt — 无网络中文学术 PPT 生成 skill

> 沉淀自一次真实任务：把本科答辩 PPT 改造成会议口头报告，并把用户自制的整页框架图
> 搬进模板页、统一术语、全局调字号。核心手段是 **Windows PowerShell 直接驱动
> PowerPoint COM**，因为目标机器 **无网络**（`pip` 装不了 python-pptx / pywin32），
> 也没有 LibreOffice `soffice`。

## 何时用

- 目标平台是 **Windows + 已装 Microsoft PowerPoint**，但 **无网络 / 无 python-pptx / 无 pywin32 / 无 soffice**。
- 要生成或修改**含中文**的 `.pptx`：封面、分节页、卡片、表格、KPI 数字块。
- 要把**别人给的整页框架图**搬进你的模板页、替换其中术语、按目标版面缩放。
- 要**全局调整字号**（如"所有字 +2 号"）而不破坏排版。

## 铁律（每次照做，顺序不能省）

1. **先探测环境**：`python-pptx / pywin32 / soffice / POWERPNT.EXE` 四者，决定路线。无网络几乎注定走 PowerShell + PowerPoint COM。
2. **中文 .ps1 必须 UTF-8 带 BOM**：本仓库/多数编辑器写出的是无 BOM UTF-8，PowerShell 5.1 会按 ANSI 解码 → 中文全乱码。**每次改完脚本、执行前，重写为 BOM**（见 `scripts/run_ps1.ps1`）。
3. **COM 属性显式强转**：颜色/对齐/枚举 → `[int]`；磅值/字号 → `[single]`；布尔 → `-1/0`（不是 `$true/$false`）。
4. **枚举值查准**：`VerticalAnchor` 只有 Top=1 / Middle=3 / Bottom=4（**没有 2**）；段落对齐左=1/中=2/右=3；`AddShape` 矩形=1/圆角=5；`SaveAs(...,24)`=.pptx。
5. **颜色是 BGR 打包整数**：`RGB = R + G*256 + B*65536`，用 `RGBv` 函数别手算。
6. **生成/改完必须导出 PNG 逐页目视检查**：SaveAs 成功 ≠ 排版正确。溢出、重叠、乱码只有看图才知道。导出文件名本地化（中文系统是 `幻灯片1.PNG`）。
7. **收尾** `ReleaseComObject` + `Quit`，否则残留 `POWERPNT` 进程。

## 搬运"整页框架图"的专门流程（本 skill 的重头）

别人给的框架图常按**占满整页**设计（字号统一、形状铺满 960×540）。直接缩进"保留标题+横幅"的窄条会**糊成一团**。按下面来：

1. **先 dump** 目标图（通常是单独一页）里所有形状的文本与几何，按**实际文本**写替换键，别照记忆猜。
2. **术语替换**：读旧字号 → `TextRange.Text=新词` → **立刻回写字号和中文字体**（`Text=` 后 COM 会重置 run 字号/字体）。
3. **确认放法**（先问用户，别默认）：
   - 图占满整页（推荐）：只留 section 标签 + 大标题，横幅精简成标题下一行小字，图铺 `x45..915 / y142..498`。
   - 保标题+横幅、图缩中间：最省事但图明显变小。
   - 图单独成一页：效果最好但总页数 +1。
4. **组合 → 复制 → 粘贴到目标页**：`Shapes.Range().Group()` → `Copy()` → 目标页 `Paste()`。
5. **等比缩放 + 同步缩字号**：`Group` 缩 `Width/Height` 只缩外框，**组内文本字号不变** → 必须递归把每个子形状字号 ×同一系数（见 `scripts/insert_figure.ps1` 的 `ScaleFonts`）。不同步就会换行溢出重叠。
6. **删原图页**，导出目标页 PNG 核对。

## 全局改字号

- "所有字 +N 号"要**递归进组内子形状**，还要**遍历表格每个单元格**（表格 `Type=19`，走 `Table.Cell(r,c).Shape.TextFrame`，普通遍历碰不到单元格），否则组里/表里的字被漏。
- 改前**备份**，改后**逐页导出核对**有无溢出。见 `scripts/bump_font.ps1`。

## 文件

- `scripts/run_ps1.ps1` — 把任意 .ps1 重写为 UTF-8 BOM 再执行（每次跑脚本的入口）。
- `scripts/ppt_helpers.ps1` — 可 dot-source 的公共函数：`RGBv / Add-Text / Add-Box / Add-Bar / NL / New-Slide`。
- `scripts/insert_figure.ps1` — 搬运整页图模板：术语替换 + 删旧中间区 + 组合搬运 + 等比缩放 + **同步缩字号** + 删原图页。
- `scripts/bump_font.ps1` — 全局字号 +N（含组内与表格单元格）。
- `scripts/export_check.ps1` — 导出指定页/全篇为 PNG 供目视检查。
- `references/pitfalls.md` — 完整踩坑清单（含最小复现与排错法）。
