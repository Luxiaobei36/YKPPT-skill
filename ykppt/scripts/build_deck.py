# -*- coding: utf-8 -*-
"""
build_deck.py — 从 YAML 大纲生成整份组会/答辩 PPT。

用法
----
    python build_deck.py deck.yaml
    python build_deck.py deck.yaml -o out.pptx --verify

YAML 结构见 references/archetypes.md；最小示例如下：

    meta:
      template: ../assets/ykppt-template.pptx   # 可选，默认用仓库自带母版
    pages:
      - type: cover
        title: 面向泛在人机物资源协同运行时编程模型
        presenter: 孙愉祥
        advisor: 陈亚兴老师
        date: 2026年7月31日
      - type: toc
      - type: lead
        section: 一、研究背景
        title: 人机物融合系统
        lead: 软件系统正由传统 CPS 向人机物融合系统（HCPS）转变。
        conclusion: 资源间的关联特性正成为影响执行效率的关键因素。
"""

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import _console  # noqa: F401  （必须在任何 print 之前）
import ykrender

DEFAULT_TEMPLATE = HERE.parent / "assets" / "ykppt-template.pptx"

REQUIRED_FIELDS = {
    "cover": ["title"],
    "toc": [],
    "lead": ["section", "title", "lead"],
    "bullets": ["section", "title", "lead", "bullets"],
    "cards": ["section", "title", "lead", "cards"],
    "flow": ["section", "title", "lead", "steps"],
    "flow_h": ["section", "title", "lead", "steps"],
    "table": ["section", "title", "lead", "table"],
    "image": ["section", "title", "lead", "image"],
    "close": [],
}


def load_spec(path):
    try:
        import yaml
    except ImportError:
        sys.exit("缺少 PyYAML：pip install pyyaml   （或改用 build_deck.py 的 JSON 分支）")
    text = Path(path).read_text(encoding="utf-8")
    if str(path).lower().endswith(".json"):
        import json
        return json.loads(text)
    return yaml.safe_load(text)


def validate(spec):
    """结构校验。返回 (pages, errors)。"""
    errors = []
    if not isinstance(spec, dict):
        return [], ["顶层必须是 mapping，含 meta / pages"]
    pages = spec.get("pages")
    if not isinstance(pages, list) or not pages:
        return [], ["缺少 pages 列表或为空"]

    content_types = {"lead", "bullets", "cards", "flow", "flow_h", "table", "image"}
    seen = set()
    n_content = 0
    for i, page in enumerate(pages, 1):
        if not isinstance(page, dict):
            errors.append(f"第 {i} 页不是 mapping")
            continue
        kind = page.get("type", "lead")
        if kind not in ykrender.RENDERERS:
            errors.append(f"第 {i} 页未知 type：{kind}")
            continue
        for field in REQUIRED_FIELDS.get(kind, []):
            if not page.get(field):
                errors.append(f"第 {i} 页（{kind}）缺少字段 {field}")
        # 结构规则：toc 可以反复出现（它是章节过渡页），cover/close 唯一
        if kind in ("cover", "close"):
            if kind in seen:
                errors.append(f"第 {i} 页：{kind} 只允许出现一次")
            seen.add(kind)
        if kind in content_types:
            n_content += 1
            title = ykrender.plain(page.get("title", ""))
            if len(title) > 26:
                errors.append(
                    f"第 {i} 页标题 {len(title)} 字，超过 26 字上限（28pt 横幅会溢出）：{title}")
            concl = page.get("conclusion")
            if concl and len(ykrender.plain(concl)) > 52:
                errors.append(
                    f"第 {i} 页结论条 {len(ykrender.plain(concl))} 字，超过 52 字上限")
            if kind == "flow" and len(page.get("steps") or []) > 3:
                errors.append(f"第 {i} 页 flow 竖排最多 3 步，请改用 flow_h")
            if kind == "flow_h" and not (4 <= len(page.get("steps") or []) <= 5):
                errors.append(f"第 {i} 页 flow_h 应为 4–5 步")
            if kind == "cards" and not (2 <= len(page.get("cards") or []) <= 3):
                errors.append(f"第 {i} 页 cards 应为 2–3 张卡")
            if kind == "table":
                t = page.get("table") or {}
                if len(t.get("head") or []) > 3:
                    errors.append(f"第 {i} 页 table 最多 3 列")
                if len(t.get("rows") or []) > 6:
                    errors.append(f"第 {i} 页 table 最多 6 行")
        if kind == "image":
            img = page.get("image")
            if img and not Path(img).exists():
                # 相对 spec 文件所在目录再试一次
                errors.append(f"第 {i} 页图片不存在：{img}")

    if pages[0].get("type") != "cover":
        errors.append("第一页必须是 cover")
    if n_content == 0:
        errors.append("至少需要一页正文")
    return pages, errors


def resolve_paths(pages, base_dir):
    """图片路径按 spec 文件所在目录解析，便于把 spec 和素材放一起。"""
    for page in pages:
        for key in ("image",):
            p = page.get(key)
            if p and not Path(p).is_absolute():
                cand = base_dir / p
                if cand.exists():
                    page[key] = str(cand)
        for card in page.get("cards") or []:
            p = card.get("image") if isinstance(card, dict) else None
            if p and not Path(p).is_absolute():
                cand = base_dir / p
                if cand.exists():
                    card["image"] = str(cand)
    return pages


def main(argv=None):
    ap = argparse.ArgumentParser(description="从 YAML 大纲生成学术风格 PPTX")
    ap.add_argument("spec", help="YAML / JSON 大纲路径")
    ap.add_argument("-o", "--out", help="输出 .pptx 路径（默认同名 .pptx）")
    ap.add_argument("-t", "--template", default=str(DEFAULT_TEMPLATE),
                    help="母版 .pptx 路径")
    ap.add_argument("--verify", action="store_true", help="生成后跑静态校验")
    ap.add_argument("--export", metavar="DIR", help="生成后用 PowerPoint 导出 PNG 到 DIR")
    ap.add_argument("--strict", action="store_true", help="有任何警告即以非零码退出")
    args = ap.parse_args(argv)

    spec_path = Path(args.spec).resolve()
    out = Path(args.out) if args.out else spec_path.with_suffix(".pptx")
    if not Path(args.template).exists():
        sys.exit(f"母版不存在：{args.template}")

    spec = load_spec(spec_path)
    pages, errors = validate(spec)
    if errors:
        print("大纲校验失败：", file=sys.stderr)
        for e in errors:
            print("  ✗", e, file=sys.stderr)
        return 2
    pages = resolve_paths(pages, spec_path.parent)

    out_path, warnings = ykrender.render_deck(Path(args.template), pages, out)
    print(f"[OK] 已生成：{out_path}")
    print(f"     共 {len(pages)} 页")

    for w in warnings:
        print("  ! " + w)

    if args.verify:
        import verify_deck
        problems = verify_deck.check(out_path)
        for p in problems:
            print("  ! " + p)
        warnings += problems

    if args.export:
        import export_png
        export_png.export(out_path, args.export)
        print(f"✓ 已导出 PNG：{args.export}")

    if args.strict and warnings:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
