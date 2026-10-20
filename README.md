# YKPPT-skill

> 西工大（NPU）组会 / 答辩学术 PPT 的生成 skill。
> **YAML 大纲 → python-pptx 出稿 → PowerPoint 导出 PNG 逐页校验**。

沉淀自 4 份真实组会稿（20260415 / 20260716 / 20260731 / 多主体跨域协同，共 105 页），
把其中反复出现却互相矛盾的写法**收敛成一套自洽规范**，并固化为
「净化母版 + 渲染引擎 + 静态校验 + 目视校验」四件套。

## 这套模板长什么样

- 深蓝五边形 + 校徽 + 斜切深蓝横幅构成的页眉，由 slideLayout 提供，**不用手画**
- 目录页兼作章节过渡页（当前章标红）
- 正文页 = 20pt 粗的定性句 + 分栏卡片 / 流程 / 表格 + 底部一句话结论条
- 统一 9 级字号阶梯、微软雅黑 + Arial 双字体、1.5 倍行距

## 快速开始

```bash
pip install python-pptx pyyaml pillow

python ykppt/scripts/build_deck.py ykppt/examples/deck.yaml \
    -o out.pptx --verify --export renders/
```

导出 PNG 需要 Windows + Microsoft PowerPoint；只生成不导出则跨平台可用。

## 目录

```
ykppt/
├─ SKILL.md                  iron laws + 工作流程（先读这个）
├─ assets/ykppt-template.pptx  净化母版：5 张页型样板页（2.7 MB，无姓名无正文）
├─ scripts/
│  ├─ build_deck.py          YAML → pptx（主入口）
│  ├─ ykrender.py            渲染引擎
│  ├─ verify_deck.py         静态校验
│  ├─ export_png.py / .ps1   导出 PNG
│  ├─ _console.py            控制台编码兜底
│  └─ com/                   旧版纯 COM 脚本（搬整页图 / 全局改字号）
├─ references/
│  ├─ style.md               设计 token：画布、色板、字号阶梯、行距
│  ├─ archetypes.md          10 种页型的精确坐标与容量
│  └─ pitfalls.md            踩坑清单（含最小复现）
└─ examples/deck.yaml        完整示例大纲
```

## 十种页型

| type | 用途 |
|---|---|
| `cover` | 封面 |
| `toc` | 目录 / 章节过渡页 |
| `lead` | 定性句页（只有一句结论） |
| `bullets` | 3–5 条并列要点 |
| `cards` | 2–3 张并排卡片 |
| `flow` | 竖向 3 步流程 |
| `flow_h` | 横向 4–5 步流程 |
| `table` | 系统 / 指标对比表 |
| `image` | 整页自绘框架图 |
| `close` | 结束页 |

## 两条路线的分工

2026-10 起主路线改为 **python-pptx**：坐标可控、可版本管理、可跨平台生成。
旧版**纯 PowerShell + PowerPoint COM** 的脚本保留在 `scripts/com/`，
在"只改一处、其余不碰"（搬一张整页图、全篇字号 +2）时依然更快。

两边的坑都记在 `references/pitfalls.md`：
A 部分是 python-pptx 的（全部是**不抛异常的静默失败**，最值得看），
B 部分是 COM 的（UTF-8 BOM、COM 类型强转、枚举值、Group 缩放不缩字号）。
