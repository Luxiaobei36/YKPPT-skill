# -*- coding: utf-8 -*-
"""
export_png.py — 把每页导出成 PNG，供逐页目视检查。

为什么必须做这一步
------------------
`SaveAs` 成功 ≠ 排版正确。溢出、重叠、字体回退成宋体、图片丢关系、
渐变糊成纯色——这些都只有**看图**才能发现。本脚本是
「生成 → 导出 → 目视 → 修正」闭环里不可省的一环。

两条实现路径
------------
1. **PowerShell + COM**（默认，零额外依赖）：调用 `export_png.ps1`。
   已验证在无 pywin32 的机器上可用。
2. **pywin32**（有则自动用）：直接 Dispatch，比拉起 powershell 快一点。

导出文件名用英文 `s01.png`，避开中文系统「幻灯片1.PNG」的本地化坑。

用法
----
    python export_png.py out.pptx renders/
    python export_png.py out.pptx renders/ --width 1920 --height 1080
"""

import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
PS1 = HERE / "export_png.ps1"


def _export_pywin32(pptx_path, out_dir, width, height):
    import win32com.client as win32
    app = pres = None
    try:
        app = win32.Dispatch("PowerPoint.Application")
        pres = app.Presentations.Open(str(pptx_path), True, False, False)
        for i, slide in enumerate(pres.Slides, 1):
            slide.Export(str(out_dir / f"s{i:02d}.png"), "PNG", width, height)
    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass


def _export_powershell(pptx_path, out_dir, width, height):
    cmd = [
        "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", str(PS1),
        "-PptPath", str(pptx_path),
        "-OutDir", str(out_dir),
        "-Width", str(width),
        "-Height", str(height),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout or "").strip()
        raise RuntimeError(f"PowerShell 导出失败：\n{detail}")


def export(pptx_path, out_dir, width=1600, height=900):
    pptx_path = Path(pptx_path).resolve()
    out_dir = Path(out_dir).resolve()
    if not pptx_path.exists():
        raise FileNotFoundError(f"找不到 {pptx_path}")
    out_dir.mkdir(parents=True, exist_ok=True)
    for old in out_dir.glob("s*.png"):
        old.unlink()

    try:
        import win32com.client  # noqa: F401
        _export_pywin32(pptx_path, out_dir, width, height)
    except ImportError:
        if not PS1.exists():
            raise RuntimeError(
                f"既没有 pywin32，也找不到 {PS1.name}。\n"
                "  装 pywin32：pip install pywin32\n"
                "  或手动用 PowerPoint 导出 PNG")
        _export_powershell(pptx_path, out_dir, width, height)

    files = sorted(out_dir.glob("s*.png"))
    if not files:
        raise RuntimeError("导出结束但没有产生 PNG，检查 PowerPoint 是否可用")
    return [str(f) for f in files]


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser(description="导出每页 PNG 供目视检查")
    ap.add_argument("pptx")
    ap.add_argument("out_dir")
    ap.add_argument("--width", type=int, default=1600)
    ap.add_argument("--height", type=int, default=900)
    args = ap.parse_args(argv)

    try:
        files = export(args.pptx, args.out_dir, args.width, args.height)
    except (RuntimeError, FileNotFoundError) as e:
        print(e, file=sys.stderr)
        return 2
    print(f"✓ 导出 {len(files)} 张 → {args.out_dir}")
    for f in files:
        print("  " + f)
    return 0


if __name__ == "__main__":
    sys.exit(main())
