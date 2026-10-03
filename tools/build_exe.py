# -*- coding: utf-8 -*-
"""构建单文件 exe。

    <python> tools/build_exe.py            # 生成图标 → PyInstaller 打包 → 自检
    <python> tools/build_exe.py --no-verify

产物：dist/docweave.exe（GUI 单文件；`docweave.exe --selftest <目录>` 可无界面自检）
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS = os.path.join(ROOT, "demo", "assets")
ICO = os.path.join(ASSETS, "docweave.ico")
ENTRY = os.path.join(ROOT, "tools", "exe_entry.py")
DIST = os.path.join(ROOT, "dist")
BUILD = os.path.join(ROOT, "build")

ENTRY_CODE = '''# -*- coding: utf-8 -*-
"""PyInstaller 打包入口。"""
import sys

from docweave.gui import main

if __name__ == "__main__":
    sys.exit(main())
'''


def make_icon() -> str:
    """画一个图标：两张交叠的文档 + 黄/青标注（呼应冲突配色）。"""
    from PIL import Image, ImageDraw

    S = 256
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([6, 6, S - 6, S - 6], radius=46, fill=(37, 86, 168, 255))
    d.rounded_rectangle([14, 14, S - 14, S - 14], radius=40,
                        outline=(255, 255, 255, 60), width=3)

    def page(x, y, w, h, band):
        d.rounded_rectangle([x, y, x + w, y + h], radius=12,
                            fill=(255, 255, 255, 255), outline=(206, 214, 226, 255),
                            width=3)
        for i in range(3):
            d.rounded_rectangle([x + 18, y + 26 + i * 22, x + w - 18, y + 34 + i * 22],
                                radius=4, fill=(198, 208, 224, 255))
        d.rounded_rectangle([x + 18, y + h - 40, x + w - 18, y + h - 22], radius=5,
                            fill=band)

    page(38, 34, 108, 150, (255, 214, 0, 255))       # 后面的文档：黄色标注
    page(110, 72, 108, 150, (0, 214, 235, 255))      # 前面的文档：青色标注
    d.line([(146, 40), (146, 24)], fill=(255, 255, 255, 200), width=6)
    img.save(ICO, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64),
                         (128, 128), (256, 256)])
    return ICO


def build(verify: bool = True) -> int:
    os.makedirs(ASSETS, exist_ok=True)
    ico = make_icon()
    print(f"[1/3] 图标已生成：{ico}")

    with open(ENTRY, "w", encoding="utf-8") as f:
        f.write(ENTRY_CODE)

    args = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean", "--onefile", "--windowed",
        "--name", "docweave",
        "--icon", ico,
        "--add-data", f"{ico}{os.pathsep}assets",
        "--paths", ROOT,
        "--distpath", DIST,
        "--workpath", BUILD,
        "--specpath", BUILD,
        # 让 PyInstaller 明确收集这些包（python-docx 需要自带模板文件）
        "--collect-data", "docx",
        "--hidden-import", "docx",
        "--hidden-import", "openpyxl",
        "--hidden-import", "lxml._elementpath",
        "--hidden-import", "PIL.Image",
        # 瘦身：这些都没用到
        "--exclude-module", "numpy",
        "--exclude-module", "pandas",
        "--exclude-module", "matplotlib",
        "--exclude-module", "scipy",
        "--exclude-module", "PyQt5",
        "--exclude-module", "PySide2",
        "--exclude-module", "IPython",
        "--exclude-module", "pytest",
        "--exclude-module", "setuptools",
        ENTRY,
    ]
    print("[2/3] 正在打包（首次约 1-3 分钟）……")
    t0 = time.time()
    proc = subprocess.run(args, cwd=ROOT)
    if proc.returncode != 0:
        print("打包失败。")
        return proc.returncode
    exe = os.path.join(DIST, "docweave.exe")
    if not os.path.exists(exe):
        print("未找到产物：", exe)
        return 1
    size = os.path.getsize(exe) / 1024 / 1024
    print(f"      打包完成，用时 {time.time() - t0:.0f}s，大小 {size:.1f} MB")
    print(f"[3/3] 产物：{exe}")

    if verify:
        work = os.path.join(ROOT, "demo", "out", "exe_selftest")
        if os.path.exists(work):
            shutil.rmtree(work, ignore_errors=True)
        print("      正在用 exe 自身做无界面自检 ……")
        p = subprocess.run([exe, "--selftest", work], timeout=600)
        report = os.path.join(work, "_selftest.json")
        if os.path.exists(report):
            import json
            with open(report, encoding="utf-8") as f:
                data = json.load(f)
            print("      自检结果：", "OK" if data.get("ok") else "失败")
            for k, v in (data.get("checks") or {}).items():
                print(f"        · {k} = {v}")
            if not data.get("ok"):
                for s in data.get("steps", []):
                    if not s.get("ok"):
                        print("        失败步骤：", s["step"], s.get("error"))
                        print(s.get("trace", ""))
                return 1
        return p.returncode
    return 0


if __name__ == "__main__":
    sys.exit(build("--no-verify" not in sys.argv))
