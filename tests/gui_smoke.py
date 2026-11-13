# -*- coding: utf-8 -*-
"""GUI 冒烟测试：构建窗口、模拟拖入、模拟编辑、验证互斥与清理，然后自动关闭。"""
import os
import shutil
import sys
import types

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

import doc_meta_cleaner as dmc  # noqa: E402

out = os.path.join(HERE, 'out')
sample = os.path.join(out, 'sample.docx')
assert os.path.exists(sample), '先运行 run_tests.py 生成样例'
# 带空格的路径副本，用于验证拖拽数据的花括号解析
spaced = os.path.join(out, 'sample copy.docx')
shutil.copy(sample, spaced)
# 文件夹拖入用例
folder = os.path.join(out, 'subdir')
os.makedirs(folder, exist_ok=True)
shutil.copy(os.path.join(out, 'sample.pdf'), os.path.join(folder, 'inner.pdf'))

app = dmc.CleanerApp()
print('拖拽组件:', '已启用' if dmc.TkinterDnD is not None else '未安装（跳过拖入用例）')

if dmc.TkinterDnD is not None:
    # 模拟拖入：一个带空格路径 + 一个文件夹（应递归找到 inner.pdf）
    ev = types.SimpleNamespace(data='{%s} {%s}' % (spaced, folder))
    app.on_drop(ev)
    names = sorted(os.path.basename(f['path']) for f in app.files)
    print('拖入后文件列表:', names)
    assert len(app.files) == 2, '拖入文件数量不对: %d' % len(app.files)
    assert 'inner.pdf' in names, '文件夹未递归展开'
    assert 'sample copy.docx' in names, '带空格路径解析失败'

# 选中 docx 文件做后续操作
docx_entry = next(f for f in app.files if f['path'].endswith('.docx'))
app.current_path = docx_entry['path']
app._refresh_values()

# 勾选清除与编辑互斥
app._toggle('title')
assert app.fields['title']['checked'] is True
ok, err = app._apply_edit('creator', '张三', False)
assert ok and app.edits['creator']['value'] == '张三', (ok, err)
assert app.fields['creator']['checked'] is False, '编辑保存后应取消该字段清除勾选'
app._toggle('creator')  # 重新勾选清除 -> 待修改应被移除
assert 'creator' not in app.edits, '勾选清除应移除同字段的待修改'

# 日期自动规范化 + 应用到全部
ok, err = app._apply_edit('created', '2026-01-02 03:04:05', True)
assert ok and app.edits['created']['value'] == '2026-01-02T03:04:05Z', (ok, err)
# 非法输入被拦截
ok, err = app._apply_edit('created', '明天下午', False)
assert not ok and err, (ok, err)
# 留空 = 转为清除勾选
ok, err = app._apply_edit('lastModifiedBy', '', False)
assert ok and app.fields['lastModifiedBy']['checked'] is True, (ok, err)
print('编辑互斥/规范化/留空删除: OK, 待修改', len(app.edits), '项')

# 撤销全部修改
app.undo_edits()
assert app.edits == {}
assert app.fields['created']['checked'] is False, '撤销修改后不应误回清除勾选（互斥时已取消）'
assert app.fields['lastModifiedBy']['checked'] is True, '留空转为的清除勾选不受撤销影响'

# 清除 + 修改混合处理（副本模式）
app.edits['creator'] = {'value': '张三', 'all': False, 'path': docx_entry['path']}
app.fields['creator']['checked'] = False
dest, removed, edited = dmc.clean_file(docx_entry, ['title'], app.edits, 'copy')
assert os.path.basename(dest).startswith('sample copy_已清理'), dest
assert 'title' in removed and edited == [('creator', '张三')], (removed, edited)
after = dmc.scan_file(dest)['props']
assert after.get('creator') == '张三' and 'title' not in after, after
print('混合处理输出:', os.path.basename(dest), '清除', len(removed), '修改', len(edited))

app.after(600, app.destroy)
app.mainloop()
print('GUI OK')
