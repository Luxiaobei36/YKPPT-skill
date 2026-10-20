# -*- coding: utf-8 -*-
"""
_console.py — 控制台输出兜底（Windows 中文版默认代码页是 GBK）

坑：在 `chcp 936` 的终端里 print("✓ 已生成") 会抛
    UnicodeEncodeError: 'gbk' codec can't encode character '✓'
生成已经成功、文件也写好了，却因为打印一个对勾而整体失败——非常难查。

对策：
1. 统一用 ASCII 标记 `[OK]` / `[!]`，不用 ✓ ✗ ─ 等符号；
2. 仍然把 stdout/stderr 的 errors 设成 replace，兜住用户自己填的中文/emoji。
"""

import sys


def setup():
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass


setup()
