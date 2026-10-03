# -*- coding: utf-8 -*-
"""生成 docweave 的演示样例文档（薄封装，实体在 docweave/samples.py）。

运行：<python> demo/make_fixtures.py
输出目录：demo/fixtures/
"""
from __future__ import annotations

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from docweave.samples import build_all

OUT = os.path.join(ROOT, "demo", "fixtures")
ASSETS = os.path.join(ROOT, "demo", "assets")


def main():
    files = build_all(OUT, ASSETS, log=lambda m: print("  " + m))
    print(f"已生成 {len(files)} 个样例文件 → {OUT}")
    for f in files:
        print("  -", f)
    return 0


if __name__ == "__main__":
    sys.exit(main())
