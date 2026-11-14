# -*- coding: utf-8 -*-
"""生成 README 用的演示 GIF：拖入 → AI 痕迹标红勾选 → 一键清理。

用法：python tests/_make_gif.py [输出路径]
输出默认为 docs/demo.gif。样例先复制到临时目录再清理，
产生的「原名_已清理」副本不会进仓库。
"""
import ctypes
import os
import shutil
import sys
import tempfile
import types

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

try:  # 与主程序一致的 DPI 感知，保证截图坐标与物理像素对齐
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

import doc_meta_cleaner as dmc  # noqa: E402
from PIL import Image, ImageDraw, ImageFont, ImageGrab  # noqa: E402

out_path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, 'docs', 'demo.gif')
os.makedirs(os.path.dirname(out_path), exist_ok=True)


def _load_font(size):
    for name in ('msyh.ttc', 'msyhbd.ttc', 'simhei.ttf'):
        try:
            return ImageFont.truetype('C:/Windows/Fonts/' + name, size)
        except OSError:
            continue
    return ImageFont.load_default()


CAPTION_FONT = _load_font(22)


def add_caption(img, text):
    """在画面底部加一条深色字幕带，标注当前步骤。"""
    bar_h = 46
    w, h = img.size
    out = Image.new('RGB', (w, h + bar_h), (24, 26, 32))
    out.paste(img, (0, 0))
    d = ImageDraw.Draw(out)
    tw = d.textlength(text, font=CAPTION_FONT)
    d.text(((w - tw) / 2, h + 10), text, font=CAPTION_FONT, fill=(235, 235, 235))
    return out


# 样例复制到临时目录，清理产物不进仓库
tmp = tempfile.mkdtemp(prefix='dmc_gif_')
samples = os.path.join(ROOT, '示例文件')
names = ('sample.docx', 'sample.xlsx', 'sample.pptx', 'sample.pdf')
paths = []
for n in names:
    dest = os.path.join(tmp, n)
    shutil.copy(os.path.join(samples, n), dest)
    paths.append(dest)

app = dmc.CleanerApp()
app.geometry('1080x680+40+40')
app.attributes('-topmost', True)
app.update()

frames = []  # (PIL.Image RGB, 持续毫秒)


def snap(caption=None, dur=700):
    try:
        hwnd = ctypes.windll.user32.GetParent(app.winfo_id())
        ctypes.windll.user32.SetForegroundWindow(hwnd)
    except Exception:
        pass
    app.update()
    x, y = app.winfo_rootx(), app.winfo_rooty()
    w, h = app.winfo_width(), app.winfo_height()
    img = ImageGrab.grab(bbox=(x, y, x + w, y + h))
    if caption:
        img = add_caption(img, caption)
    frames.append((img, dur))


try:
    # ① 空窗口
    snap('① 把文件或整个文件夹直接拖进窗口', 1000)

    # ② 拖入后：疑似 AI 痕迹标红并自动勾选
    app.on_drop(types.SimpleNamespace(data='{%s}' % '} {'.join(paths)))
    app.wait_for_idle()  # 扫描在后台线程进行，等完成后再截「拖入后」帧
    app.update()
    snap('② 疑似 AI 痕迹自动标红并勾选', 1800)

    # ③ 复刻 on_clean 的逐文件循环，真实清理临时副本，日志逐行出现；
    #    与真实流程一致：进度条步进、文件列表逐个打上「已另存 ✓」状态标记
    app.start_btn.config(state='disabled')
    total = len(app.files)
    app.progress.pack(side='right', padx=(0, 10))
    app.progress.config(maximum=max(total, 1), value=0)
    snap('③ 点「开始清理」，进度条走起，日志逐个文件报告', 700)
    checked = [k for k, v in app.fields.items() if v['checked']]
    mode = app.save_var.get()  # 'copy' 另存副本
    ok = fail = 0
    for i, f in enumerate(list(app.files), 1):
        name = os.path.basename(f['path'])
        _d, _s, skipped = dmc.collect_ops(f, checked, app.edits)
        for k, why in skipped:
            app.log('· %s：%s 跳过（%s）' % (name, dmc.field_display_name(k), why))
        try:
            dest, removed, edited = dmc.clean_file(f, checked, app.edits, mode)
        except Exception as e:
            fail += 1
            app.log('✗ %s：%s' % (name, e), 'err')
            continue
        parts = []
        if removed:
            parts.append('清除 %d 项：%s' % (
                len(removed), '、'.join(dmc.field_display_name(k) for k in removed)))
        if edited:
            parts.append('修改 %d 项：%s' % (
                len(edited), '、'.join('%s→%s' % (dmc.field_display_name(k), dmc._truncate(v, 30))
                                      for k, v in edited)))
        app.log('✓ %s → 已另存为 %s（%s）'
                % (name, os.path.basename(dest), '；'.join(parts)), 'ok')
        f['status'] = ('已另存 ✓', 'ok')
        app._update_file_row(f['path'])
        app.progress.config(value=i)
        app.status.config(text='正在清理 %d/%d…' % (i, total))
        ok += 1
        snap('③ 点「开始清理」，进度条走起，日志逐个文件报告', 550)
    app.start_btn.config(state='normal')
    app.progress.pack_forget()
    app.update_status()
    app.edits.clear()
    app.log('—— 完成：成功 %d 个，失败 %d 个 ——' % (ok, fail), 'ok')

    # ④ 完成（不重扫，保留绿标状态与干净日志的最终画面）
    app.update()
    snap('④ 清理完成：列表标「已另存 ✓」，文档属性已干净', 1800)
finally:
    app.destroy()
    shutil.rmtree(tmp, ignore_errors=True)

# 缩到 810 宽、以「拖入后」帧为基准做统一调色板，避免逐帧调色板抖动闪烁
WIDTH = 810
resized = []
for img, _dur in frames:
    w, h = img.size
    resized.append(img.resize((WIDTH, round(h * WIDTH / w)), Image.LANCZOS))
palette = resized[1].quantize(colors=128, method=Image.MEDIANCUT)
gif_frames = [img.quantize(palette=palette, dither=Image.FLOYDSTEINBERG) for img in resized]
durations = [dur for _img, dur in frames]

gif_frames[0].save(out_path, save_all=True, append_images=gif_frames[1:],
                   duration=durations, loop=0, optimize=True)
print('saved:', out_path, os.path.getsize(out_path), 'bytes,',
      len(gif_frames), 'frames')
