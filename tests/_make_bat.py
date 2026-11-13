# -*- coding: utf-8 -*-
"""把 UTF-8 模板转成 GBK + CRLF 的 .bat，并自检。"""
import os

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

src = open(os.path.join(HERE, '_bat_template.txt'), encoding='utf-8').read()
src = src.replace('\r\n', '\n').replace('\n', '\r\n')
data = src.encode('gbk')
bat = os.path.join(ROOT, '文档属性清理工具.bat')
with open(bat, 'wb') as f:
    f.write(data)

# 自检
back = data.decode('gbk')
lines = back.split('\r\n')
assert back.count('\r\n') >= 20, 'CRLF missing'
assert 'doc_meta_cleaner.py' in back
print('GBK bytes:', len(data), '| lines:', len(lines))
print('first line:', repr(lines[0]))
print('python line:', [l for l in lines if 'doc_meta_cleaner' in l])
