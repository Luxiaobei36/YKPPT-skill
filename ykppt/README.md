# ykppt

> 用 **python-pptx 从 YAML 大纲**批量生成西工大组会/答辩风格的中文 `.pptx`，
> 再用 **PowerPoint COM 导出 PNG 逐页目视校验**。

沉淀自 4 份真实组会稿（共 105 页），把其中反复出现却互相矛盾的写法
**收敛成一套自洽规范**，并固化成「净化母版 + 渲染引擎 + 静态校验 + 目视校验」。

---

## 它解决什么问题

手工做一份 20 页组会 PPT，最耗时的不是写字，而是：

- 每页页眉的五边形/校徽/斜切横幅要摆对齐 —— 本 skill 由母版继承，不用画；
- 页与页之间的横幅字号、正文行距、结论条位置不知不觉就走样 —— 本 skill 用固定 token；
- 改了一处字号，别的页没跟着改，全篇看起来不齐 —— 本 skill 只有 9 级字号阶梯；
- 生成完不知道排得对不对，只能一页页翻 —— 本 skill 有静态校验 + PNG 导出。

## 环境

```bash
pip install python-pptx pyyaml pillow
```

- **必需**：Python 3.8+、python-pptx ≥ 0.6.21、PyYAML。
- **可选**：Pillow（更准的字宽估算）、pywin32（更快的导出）。
  没有 Pillow 会退回启发式估算；没有 pywin32 会走自带的 PowerShell 导出。
- **导出 PNG 需要 Windows + 已装 Microsoft PowerPoint**；只生成不导出则跨平台可用。

## 快速上手

```bash
python scripts/build_deck.py examples/deck.yaml -o out.pptx --verify --export renders/
```

跑完会打印校验结果，并把每页导出成 `renders/s01.png …`。

## 目录

```
ykppt/
├─ SKILL.md                  铁律与工作流程（先读这个）
├─ assets/ykppt-template.pptx  净化母版：5 张页型样板页
├─ scripts/
│  ├─ build_deck.py          YAML → pptx（主入口）
│  ├─ ykrender.py            渲染引擎
│  ├─ verify_deck.py         静态校验
│  ├─ export_png.py / .ps1   导出 PNG
│  ├─ _console.py            控制台编码兜底
│  └─ com/                   旧版纯 COM 脚本（搬图 / 改字号）
├─ references/
│  ├─ style.md               设计 token（色板、字号、坐标）
│  ├─ archetypes.md          10 种页型的配方
│  └─ pitfalls.md            踩坑清单
└─ examples/deck.yaml        完整示例
```

## 十种页型

`cover` `toc` `lead` `bullets` `cards` `flow` `flow_h` `table` `image` `close`

详见 `references/archetypes.md`。

## 这套模板的视觉规范

| 项 | 值 |
|---|---|
| 画布 | 1333 × 750（16:9，即 13.33in × 7.5in） |
| 主色 | 深蓝 `#142C76`、品牌蓝 `#014DA1`、浅蓝 `#5BB0E6`、浅底 `#EAF3FA` |
| 强调 | 正文强调 `#C00000`，结论/警示 `#FF0000` |
| 字体 | 微软雅黑（中）+ Arial（西） |
| 字号 | 48 / 54 / 28 / 24 / 20 / 16 / 14 / 12 / 10 |
| 正文行距 | 1.5 倍 |

完整数值见 `references/style.md`。

## 一条重要的排版原则

**一页一个论点**。每页必须有：

```
section   章标签           一、研究背景
title     横幅标题（断言）  亲和性判断能不能配
lead      定性句 20pt 粗    亲和性回答"哪些资源适合一起干"
conclusion 结论条           现有运行时缺少对二者的持续对象化管理
```

缺 `lead` 或 `conclusion`，`verify_deck.py` 会报警。
因为这套模板的说服力全在"一页一句结论"上，不是靠堆内容。

## 与旧版（纯 PowerShell COM）的关系

旧版是在**无网络、装不了 python-pptx** 的机器上被迫走的 COM 路线，
能力只有"搬整页图"和"全局改字号"。这两个脚本仍保留在 `scripts/com/`，
在"只动一处、其余不碰"的场景下比重新生成更快。

新增的生成能力走 python-pptx，因为它的坐标控制和版本管理都比 COM 可靠。
