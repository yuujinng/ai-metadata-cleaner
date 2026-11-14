# -*- coding: utf-8 -*-
"""窗口抓图工具（_make_screenshot.py / _make_gif.py 共用）。

优先 PIL.ImageGrab 抓屏幕（与用户桌面所见完全一致）；在没有交互桌面
的会话（后台任务、远程 CI、锁屏）里抓屏会抛 OSError，此时回退
PrintWindow 直接抓本进程窗口——不经过屏幕，窗口只要创建成功就能出图。
两种后端都返回客户区图像，构图一致。
"""
import ctypes
import ctypes.wintypes as wt


class _BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [('biSize', wt.DWORD), ('biWidth', wt.LONG), ('biHeight', wt.LONG),
                ('biPlanes', wt.WORD), ('biBitCount', wt.WORD),
                ('biCompression', wt.DWORD), ('biSizeImage', wt.DWORD),
                ('biXPelsPerMeter', wt.LONG), ('biYPelsPerMeter', wt.LONG),
                ('biClrUsed', wt.DWORD), ('biClrImportant', wt.DWORD)]


class _BITMAPINFO(ctypes.Structure):
    _fields_ = [('bmiHeader', _BITMAPINFOHEADER), ('bmiColors', wt.DWORD * 3)]


def grab_window(win):
    """抓取 Tk 窗口的客户区，返回 PIL RGB Image。"""
    try:
        from PIL import ImageGrab
        x, y = win.winfo_rootx(), win.winfo_rooty()
        w, h = win.winfo_width(), win.winfo_height()
        img = ImageGrab.grab(bbox=(x, y, x + w, y + h))
        img.load()   # 强制立即读取，失败落到 PrintWindow 回退
        return img
    except Exception:
        return _print_window(win)


def _print_window(win):
    """PrintWindow 抓全窗口位图，再裁出客户区。"""
    user32, gdi32 = ctypes.windll.user32, ctypes.windll.gdi32
    win.update_idletasks()
    hwnd = user32.GetParent(win.winfo_id()) or win.winfo_id()

    wr, cr = wt.RECT(), wt.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(wr))
    user32.GetClientRect(hwnd, ctypes.byref(cr))
    pt = wt.POINT(0, 0)
    user32.ClientToScreen(hwnd, ctypes.byref(pt))
    off_x, off_y = pt.x - wr.left, pt.y - wr.top          # 客户区在全窗口里的偏移
    W, H = wr.right - wr.left, wr.bottom - wr.top
    CW, CH = cr.right - cr.left, cr.bottom - cr.top

    hdc = user32.GetWindowDC(hwnd)
    mem = gdi32.CreateCompatibleDC(hdc)
    bmi = _BITMAPINFO()
    bmi.bmiHeader.biSize = ctypes.sizeof(_BITMAPINFOHEADER)
    bmi.bmiHeader.biWidth = W
    bmi.bmiHeader.biHeight = -H                           # 负高：自上而下
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 32
    bits = ctypes.c_void_p()
    hbmp = gdi32.CreateDIBSection(mem, ctypes.byref(bmi), 0,
                                  ctypes.byref(bits), None, 0)
    old = gdi32.SelectObject(mem, hbmp)
    try:
        # PW_RENDERFULLCONTENT = 2：含 DWM 合成内容（Win 8.1+）
        if not user32.PrintWindow(hwnd, mem, 2):
            raise OSError('PrintWindow 失败')
        from PIL import Image
        img = Image.frombytes('RGB', (W, H),
                              ctypes.string_at(bits, W * 4 * H), 'raw', 'BGRX')
        return img.crop((off_x, off_y, off_x + CW, off_y + CH))
    finally:
        gdi32.SelectObject(mem, old)
        gdi32.DeleteObject(hbmp)
        gdi32.DeleteDC(mem)
        user32.ReleaseDC(hwnd, hdc)
