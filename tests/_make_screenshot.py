# -*- coding: utf-8 -*-
"""生成 README 用的界面截图：把示例文件装入界面后截取窗口。

用法：python tests/_make_screenshot.py [输出路径]
输出默认为 docs/screenshot.png。
"""
import ctypes
import os
import sys
import types

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

try:  # 与主程序一致的 DPI 感知，保证截图坐标与物理像素对齐
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

import doc_meta_cleaner as dmc  # noqa: E402
from PIL import ImageGrab  # noqa: E402

out_path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, 'docs', 'screenshot.png')
os.makedirs(os.path.dirname(out_path), exist_ok=True)

samples = os.path.join(ROOT, '示例文件')
paths = [os.path.join(samples, n)
         for n in ('sample.docx', 'sample.xlsx', 'sample.pptx', 'sample.pdf')]
drop_data = '{%s}' % '} {'.join(paths)

app = dmc.CleanerApp()
app.geometry('1080x680+40+40')
app.attributes('-topmost', True)
app.update()
app.on_drop(types.SimpleNamespace(data=drop_data))
app.wait_for_idle()  # 拖入后扫描在后台线程进行，等扫描完成再截图
app.update()


def snap():
    try:
        hwnd = ctypes.windll.user32.GetParent(app.winfo_id())
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        app.update()
        x, y = app.winfo_rootx(), app.winfo_rooty()
        w, h = app.winfo_width(), app.winfo_height()
        img = ImageGrab.grab(bbox=(x, y, x + w, y + h))
        img.save(out_path)
        print('saved:', out_path, img.size)
    finally:
        app.destroy()


app.after(1500, snap)  # 等界面完全绘制完成
app.mainloop()
