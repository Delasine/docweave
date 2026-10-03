# -*- coding: utf-8 -*-
"""给 GUI 窗口截图（用于验证界面真的画出来了）。

    <python> tools/gui_shot.py --exe dist/docweave.exe --out demo/out/gui.png
    <python> tools/gui_shot.py --py                     # 用源码启动
"""
from __future__ import annotations

import argparse
import ctypes
import ctypes.wintypes as wt
import os
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TITLE_KEY = "docweave"

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32
PW_RENDERFULLCONTENT = 0x00000002
SRCCOPY = 0x00CC0020


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [("biSize", wt.DWORD), ("biWidth", wt.LONG), ("biHeight", wt.LONG),
                ("biPlanes", wt.WORD), ("biBitCount", wt.WORD),
                ("biCompression", wt.DWORD), ("biSizeImage", wt.DWORD),
                ("biXPelsPerMeter", wt.LONG), ("biYPelsPerMeter", wt.LONG),
                ("biClrUsed", wt.DWORD), ("biClrImportant", wt.DWORD)]


class BITMAPINFO(ctypes.Structure):
    _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", wt.DWORD * 3)]


def list_windows(substr: str):
    out = []
    proc_type = ctypes.WINFUNCTYPE(ctypes.c_bool, wt.HWND, wt.LPARAM)

    def cb(hwnd, _lparam):
        if not user32.IsWindowVisible(hwnd):
            return True
        n = user32.GetWindowTextLengthW(hwnd)
        if n <= 0:
            return True
        buf = ctypes.create_unicode_buffer(n + 1)
        user32.GetWindowTextW(hwnd, buf, n + 1)
        if substr in buf.value:
            pid = wt.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            out.append((hwnd, buf.value, pid.value))
        return True

    user32.EnumWindows(proc_type(cb), 0)
    return out


def kill_by_title(substr: str):
    """清理同名窗口所属进程树（onefile 引导进程会留子进程，必须整棵树杀）。"""
    for hwnd, title, pid in list_windows(substr):
        subprocess.run(["taskkill", "/F", "/T", "/PID", str(pid)],
                       capture_output=True)


def find_window(substr: str, timeout: float = 60.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        found = list_windows(substr)
        if found:
            # 取面积最大的那个（真正的主窗口）
            best = None
            for hwnd, title, pid in found:
                rect = wt.RECT()
                user32.GetWindowRect(hwnd, ctypes.byref(rect))
                area = (rect.right - rect.left) * (rect.bottom - rect.top)
                if best is None or area > best[0]:
                    best = (area, hwnd, title)
            if best and best[0] > 10000:
                return best[1], best[2]
        time.sleep(0.4)
    return None


def _bits(hdc, bmp, width, height):
    bmi = BITMAPINFO()
    bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bmi.bmiHeader.biWidth = width
    bmi.bmiHeader.biHeight = -height
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 32
    bmi.bmiHeader.biCompression = 0
    buf = ctypes.create_string_buffer(width * height * 4)
    gdi32.GetDIBits(hdc, bmp, 0, height, buf, ctypes.byref(bmi), 0)
    return buf.raw


def capture(hwnd, path: str) -> str:
    from PIL import Image

    user32.ShowWindow(hwnd, 9)          # SW_RESTORE
    user32.SetForegroundWindow(hwnd)
    time.sleep(0.8)
    rect = wt.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    w = max(rect.right - rect.left, 1)
    h = max(rect.bottom - rect.top, 1)

    src = user32.GetWindowDC(hwnd)
    mem = gdi32.CreateCompatibleDC(src)
    bmp = gdi32.CreateCompatibleBitmap(src, w, h)
    gdi32.SelectObject(mem, bmp)
    user32.PrintWindow(hwnd, mem, PW_RENDERFULLCONTENT)
    raw = _bits(mem, bmp, w, h)
    img = Image.frombuffer("RGBA", (w, h), raw, "raw", "BGRA", 0, 1).convert("RGB")

    # PrintWindow 对个别窗口会返回空白：检测到就改用屏幕抓取
    if _is_blank(img):
        user32.SetForegroundWindow(hwnd)
        time.sleep(0.6)
        screen = user32.GetDC(0)
        mem2 = gdi32.CreateCompatibleDC(screen)
        bmp2 = gdi32.CreateCompatibleBitmap(screen, w, h)
        gdi32.SelectObject(mem2, bmp2)
        gdi32.BitBlt(mem2, 0, 0, w, h, screen, rect.left, rect.top, SRCCOPY)
        raw2 = _bits(mem2, bmp2, w, h)
        img = Image.frombuffer("RGBA", (w, h), raw2, "raw", "BGRA", 0, 1).convert("RGB")
        gdi32.DeleteObject(bmp2)
        gdi32.DeleteDC(mem2)
        user32.ReleaseDC(0, screen)

    gdi32.DeleteObject(bmp)
    gdi32.DeleteDC(mem)
    user32.ReleaseDC(hwnd, src)
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    img.save(path)
    return path


def _is_blank(img) -> bool:
    small = img.resize((64, 48))
    colors = small.getcolors(maxcolors=4096)
    if not colors:
        return True
    total = sum(c for c, _ in colors)
    top = max(c for c, _ in colors)
    return top / total > 0.985


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--exe", help="要启动的 exe")
    ap.add_argument("--exe-args", default="", help="传给 exe 的额外参数")
    ap.add_argument("--py", action="store_true", help="用源码启动 GUI")
    ap.add_argument("--out", default=os.path.join(ROOT, "demo", "out", "gui.png"))
    ap.add_argument("--wait", type=float, default=25.0, help="窗口出现后等待秒数")
    args = ap.parse_args()

    kill_by_title(TITLE_KEY)          # 先清掉残留窗口，避免抓到旧的那个
    time.sleep(0.5)
    if args.exe:
        cmd = [args.exe] + [a for a in args.exe_args.split() if a]
        proc = subprocess.Popen(cmd, cwd=ROOT)
    else:
        proc = subprocess.Popen([sys.executable, "-m", "docweave.gui"], cwd=ROOT)
    try:
        hit = find_window(TITLE_KEY, timeout=90)
        if not hit:
            print("未找到窗口")
            return 1
        hwnd, title = hit
        print(f"窗口：{title}")
        time.sleep(args.wait)
        path = capture(hwnd, args.out)
        print(f"截图：{path}（{os.path.getsize(path)} 字节）")
        return 0
    finally:
        kill_by_title(TITLE_KEY)
        try:
            proc.terminate()
        except Exception:
            pass


if __name__ == "__main__":
    sys.exit(main())
