# ykppt

> 无网络环境下，用 **Windows PowerShell + PowerPoint COM** 批量生成 / 编辑
> **含中文**的学术风格 `.pptx` 的 skill。

沉淀自一次真实任务：把一份本科答辩 PPT 改造成学术会议口头报告，并把用户自制的
**整页框架图**搬进模板页、统一术语、全局调字号。目标机器无网络（`pip` 装不了
python-pptx / pywin32），也没有 LibreOffice，因此走 PowerShell 直接驱动
PowerPoint COM 这条路。

## 适用前提

- Windows + 已安装 Microsoft PowerPoint（提供 `POWERPNT.EXE` / COM）。
- 典型场景：**无网络 / 无 python-pptx / 无 pywin32 / 无 soffice**。
- Windows PowerShell 5.1（脚本按 5.1 的坑写，如 UTF-8 BOM、无 `&&`）。

## 目录

```
ykppt/
├─ SKILL.md                  # skill 主文件：何时用 + 铁律 + 整页图搬运流程
├─ README.md                 # 本文件
├─ scripts/
│  ├─ run_ps1.ps1            # 把 .ps1 重写为 UTF-8 BOM 再执行（每次跑脚本的入口）
│  ├─ ppt_helpers.ps1        # 公共函数：RGBv/Add-Text/Add-Box/Add-Bar/NL/New-Slide...
│  ├─ insert_figure.ps1      # 搬整页图：术语替换+删旧区+组合搬运+等比缩放+同步缩字号+删原页
│  ├─ bump_font.ps1          # 全局字号 +N（含组内子形状 与 表格单元格）
│  └─ export_check.ps1       # 导出 PNG 逐页目视检查
└─ references/
   └─ pitfalls.md            # 完整踩坑清单（含最小复现与排错法）
```

## 快速上手

1. 打开 `scripts/` 下要用的脚本，改顶部参数（`$pptPath` 等）。
2. 用 BOM 入口执行，避免中文乱码：
   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\run_ps1.ps1 scripts\insert_figure.ps1
   ```
3. 跑 `export_check.ps1` 导出 PNG，**逐页目视检查**排版。

## 七条铁律（详见 SKILL.md / pitfalls.md）

1. 动手前先探测 python-pptx / pywin32 / soffice / POWERPNT，决定路线。
2. 中文 `.ps1` **必须 UTF-8 带 BOM**，否则 PS 5.1 按 ANSI 解码 → 全乱码。
3. COM 赋值显式强转：颜色/对齐/枚举 `[int]`，磅值/字号 `[single]`，布尔 `-1/0`。
4. 枚举值查准：`VerticalAnchor` 无 2（Top1/Mid3/Bot4）；`SaveAs(...,24)`=.pptx。
5. 颜色是 BGR 打包整数：`RGB = R + G*256 + B*65536`。
6. 生成/改完**必须导出 PNG 逐页看**；SaveAs 成功 ≠ 排版正确。
7. 收尾 `ReleaseComObject` + `Quit`，别残留 `POWERPNT` 进程。

搬整页框架图额外注意：**先 dump 真实文本再写替换键**；**Group 等比缩放不缩字号，
必须递归同步缩**；改文字后**立刻回写字号**；整页图塞进"保留标题+横幅"的窄条会糊，
放法先与用户确认。
