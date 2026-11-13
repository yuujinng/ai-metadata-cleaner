# -*- coding: utf-8 -*-
"""自动化测试：生成带 AI 痕迹的样例文档 -> 验证扫描与清理逻辑。"""
import os
import shutil
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

import doc_meta_cleaner as dmc  # noqa: E402

OUT = os.path.join(HERE, 'out')

FAILURES = []


def check(name, cond, extra=''):
    mark = '[OK]  ' if cond else '[FAIL]'
    print(mark, name, ('-- ' + str(extra)) if (extra and not cond) else '')
    if not cond:
        FAILURES.append(name)


# ---------------------------------------------------------------------------
# 样例文件
# ---------------------------------------------------------------------------

CORE_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
    ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"'
    ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    '<dc:title>AI 起草的测试报告</dc:title>'
    '<dc:subject>测试</dc:subject>'
    '<dc:creator>ChatGPT (OpenAI)</dc:creator>'
    '<cp:lastModifiedBy>Claude (Anthropic)</cp:lastModifiedBy>'
    '<cp:keywords>AI, 测试</cp:keywords>'
    '<cp:revision>3</cp:revision>'
    '<dcterms:created xsi:type="dcterms:W3CDTF">2026-10-20T10:00:00Z</dcterms:created>'
    '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-10-20T10:00:00Z</dcterms:modified>'
    '</cp:coreProperties>'
)

APP_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
    ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    '<Application>python-docx</Application>'
    '<AppVersion>1.1.2</AppVersion>'
    '<Company>OpenAI</Company>'
    '<TotalTime>0</TotalTime>'
    '</Properties>'
)

CUSTOM_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Properties xmlns="http://schemas.openxmlformats.org/officedocument/2006/custom-properties"'
    ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    '<Property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Source">'
    '<vt:lpstr>z.ai GLM-4</vt:lpstr></Property>'
    '<Property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Reviewer">'
    '<vt:lpstr>张三</vt:lpstr></Property>'
    '</Properties>'
)

CONTENT_TYPES_DOCX = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>'
    '</Types>'
)

RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
    '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officedocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>'
    '</Relationships>'
)

DOCUMENT_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:body><w:p><w:r><w:t>这是正文内容，清理后必须原样保留。</w:t></w:r></w:p></w:body>'
    '</w:document>'
)


def build_ooxml(path, main_part, main_target):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        ct = CONTENT_TYPES_DOCX.replace('/word/document.xml', '/' + main_target)
        z.writestr('[Content_Types].xml', ct)
        z.writestr('_rels/.rels', RELS.replace('word/document.xml', main_target))
        z.writestr(main_target, main_part)
        z.writestr('docProps/core.xml', CORE_XML)
        z.writestr('docProps/app.xml', APP_XML)
        z.writestr('docProps/custom.xml', CUSTOM_XML)


def make_pdf(path):
    from pypdf import PdfWriter
    w = PdfWriter()
    w.add_blank_page(300, 400)
    w.add_metadata({
        '/Title': '保留的标题',
        '/Author': 'Claude (Anthropic)',
        '/Creator': 'ChatGPT',
        '/Producer': 'Skia/PDF m123',
        '/CreationDate': "D:20261020100000Z",
        '/ModDate': "D:20261020100000Z",
    })
    with open(path, 'wb') as f:
        w.write(f)


def make_samples():
    if os.path.isdir(OUT):
        shutil.rmtree(OUT)
    os.makedirs(OUT)
    build_ooxml(os.path.join(OUT, 'sample.docx'), DOCUMENT_XML, 'word/document.xml')
    build_ooxml(os.path.join(OUT, 'sample.xlsx'),
                '<?xml version="1.0" encoding="UTF-8"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
                'xl/workbook.xml')
    build_ooxml(os.path.join(OUT, 'sample.pptx'),
                '<?xml version="1.0" encoding="UTF-8"?>'
                '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
                'ppt/presentation.xml')
    make_pdf(os.path.join(OUT, 'sample.pdf'))


# ---------------------------------------------------------------------------
# 测试
# ---------------------------------------------------------------------------

def test_scan():
    entry = dmc.scan_file(os.path.join(OUT, 'sample.docx'))
    p = entry['props']
    check('docx 读取 creator', p.get('creator') == 'ChatGPT (OpenAI)')
    check('docx 读取 lastModifiedBy', p.get('lastModifiedBy') == 'Claude (Anthropic)')
    check('docx 读取 Application', p.get('Application') == 'python-docx')
    check('docx 读取自定义属性', p.get('custom:Source') == 'z.ai GLM-4')
    for k in ('creator', 'lastModifiedBy', 'Application', 'Company',
              'created', 'modified', 'TotalTime', 'custom:Source'):
        check('docx 疑似AI: ' + k, k in entry['suspects'], entry['suspects'])
    check('docx 标题不算AI痕迹', 'title' not in entry['suspects'])
    check('docx 关键词裸AI不算痕迹', 'keywords' not in entry['suspects'])
    check('docx 正常自定义属性不算AI', 'custom:Reviewer' not in entry['suspects'])
    check('识别 AIGC', dmc.looks_like_ai('本文由AIGC工具生成'))
    check('识别中文相邻的Claude', dmc.looks_like_ai('作者Claude修改'))
    check('识别 由AI生成', dmc.looks_like_ai('由 AI 生成的初稿'))
    check('识别 AI-generated', dmc.looks_like_ai('AI-generated summary'))
    check('不误报裸AI', not dmc.looks_like_ai('AI, 机器学习'))

    entry = dmc.scan_file(os.path.join(OUT, 'sample.pdf'))
    p = entry['props']
    check('pdf 读取 Author', 'Claude (Anthropic)' in p.get('pdf:Author', ''))
    check('pdf 读取 Producer', 'Skia' in p.get('pdf:Producer', ''))
    for k in ('pdf:Author', 'pdf:Creator', 'pdf:Producer'):
        check('pdf 疑似AI: ' + k, k in entry['suspects'], entry['suspects'])
    check('pdf 标题不算AI', 'pdf:Title' not in entry['suspects'])

    for fn in ('sample.xlsx', 'sample.pptx'):
        entry = dmc.scan_file(os.path.join(OUT, fn))
        check(fn + ' 读取', entry['props'].get('creator') == 'ChatGPT (OpenAI)')


def test_clean_copy():
    src = os.path.join(OUT, 'sample.docx')
    entry = dmc.scan_file(src)
    keys = ['creator', 'lastModifiedBy', 'Application', 'Company', 'TotalTime',
            'created', 'modified', 'custom:Source']   # 不勾 title/subject/revision/keywords
    dest, removed, edited = dmc.clean_file(entry, keys, {}, 'copy')
    check('副本命名', os.path.basename(dest) == 'sample_已清理.docx', dest)
    check('清除项数', len(removed) == len(keys), removed)
    after = dmc.scan_file(dest)['props']
    check('creator 已删除', 'creator' not in after)
    check('lastModifiedBy 已删除', 'lastModifiedBy' not in after)
    check('Application 已删除', 'Application' not in after)
    check('自定义属性 Source 已删除', 'custom:Source' not in after)
    check('title 保留', after.get('title') == 'AI 起草的测试报告')
    check('revision 保留', after.get('revision') == '3')
    check('自定义属性 Reviewer 保留', after.get('custom:Reviewer') == '张三')

    # 非属性部件必须逐字节一致
    with zipfile.ZipFile(src) as za, zipfile.ZipFile(dest) as zb:
        sa, sb = set(za.namelist()), set(zb.namelist())
        check('zip 部件清单一致', sa == sb, (sa, sb))
        same = all(za.read(n) == zb.read(n) for n in sa
                   if n not in ('docProps/core.xml', 'docProps/app.xml', 'docProps/custom.xml'))
        check('正文等部件逐字节保留', same)
        core = zb.read('docProps/core.xml').decode('utf-8')
        check('core.xml 命名空间前缀保留', 'cp:coreProperties' in core and 'xmlns:dc=' in core, core[:200])
        import xml.etree.ElementTree as ET
        ET.fromstring(core)  # 结构合法
        ET.fromstring(zb.read('docProps/app.xml'))
        ET.fromstring(zb.read('docProps/custom.xml'))
    check('原文件未被改动', dmc.scan_file(src)['props'].get('creator') == 'ChatGPT (OpenAI)')

    # 第二次副本不覆盖第一次
    dest2, _, _ = dmc.clean_file(dmc.scan_file(src), keys, {}, 'copy')
    check('副本重名自动编号', dest2 != dest, dest2)


def test_clean_overwrite():
    import shutil as sh
    src = os.path.join(OUT, 'sample.xlsx')
    ow = os.path.join(OUT, 'ow.xlsx')
    sh.copy(src, ow)
    entry = dmc.scan_file(ow)
    dest, removed, edited = dmc.clean_file(entry, ['creator', 'lastModifiedBy'], {}, 'overwrite')
    check('覆盖模式返回原路径', dest == ow, dest)
    after = dmc.scan_file(ow)['props']
    check('覆盖后 creator 已删除', 'creator' not in after)
    check('覆盖后 title 保留', after.get('title') == 'AI 起草的测试报告')
    check('无临时文件残留', not os.path.exists(ow + '.meta_tmp'))


def test_clean_pdf():
    src = os.path.join(OUT, 'sample.pdf')
    entry = dmc.scan_file(src)
    keys = ['pdf:Author', 'pdf:Creator', 'pdf:Producer', 'pdf:CreationDate', 'pdf:ModDate']
    dest, removed, edited = dmc.clean_file(entry, keys, {}, 'copy')
    check('pdf 副本命名', os.path.basename(dest) == 'sample_已清理.pdf', dest)
    after = dmc.scan_file(dest)['props']
    for k in keys:
        check('pdf 已删除: ' + k, k not in after, after.keys())
    check('pdf Title 保留', after.get('pdf:Title') == '保留的标题')
    # 页面还在
    from pypdf import PdfReader
    check('pdf 页数不变', len(PdfReader(dest).pages) == 1)


# ---------------------------------------------------------------------------
# 编辑属性
# ---------------------------------------------------------------------------

def test_normalize():
    err, v = dmc.normalize_edit_value('ooxml', 'created', '2026-01-02 03:04:05')
    check('ooxml 日期转换', err is None and v == '2026-01-02T03:04:05Z', (err, v))
    err, v = dmc.normalize_edit_value('ooxml', 'created', '2026-01-02')
    check('纯日期转换', err is None and v == '2026-01-02T00:00:00Z', (err, v))
    err, v = dmc.normalize_edit_value('ooxml', 'modified', '2026年1月2日 3:04')
    check('非法日期报错', err is not None and v is None, (err, v))
    err, v = dmc.normalize_edit_value('pdf', 'pdf:CreationDate', '2026-01-02 03:04:05')
    check('pdf 日期转换', err is None and v == 'D:20260102030405Z', (err, v))
    err, v = dmc.normalize_edit_value('pdf', 'pdf:ModDate', 'D:20260102030405')
    check('pdf D:格式输入', err is None and v == 'D:20260102030405Z', (err, v))
    err, v = dmc.normalize_edit_value('ooxml', 'TotalTime', '90')
    check('整数通过', err is None and v == '90', (err, v))
    err, v = dmc.normalize_edit_value('ooxml', 'TotalTime', 'abc')
    check('非整数报错', err is not None, (err, v))
    err, v = dmc.normalize_edit_value('ooxml', 'creator', '  张三 ')
    check('自由文本', err is None and v == '张三', (err, v))
    err, v = dmc.normalize_edit_value('ooxml', 'creator', '')
    check('留空=删除', err is None and v == '', (err, v))


def test_edit_ooxml():
    src = os.path.join(OUT, 'sample.docx')
    entry = dmc.scan_file(src)
    edits = {
        'creator':  {'value': '张三', 'all': False, 'path': src},
        'TotalTime': {'value': '120', 'all': False, 'path': src},
        'category': {'value': '年度报告', 'all': False, 'path': src},   # 原文件没有 -> 新建
        'created':  {'value': '2026-01-02T03:04:05Z', 'all': False, 'path': src},  # GUI 已规范化
        'custom:Source': {'value': '人工撰写', 'all': False, 'path': src},
    }
    dest, removed, edited = dmc.clean_file(entry, [], edits, 'copy')
    check('纯编辑也有输出副本',
          os.path.basename(dest).startswith('sample_已清理') and dest.endswith('.docx'), dest)
    check('编辑项数', len(edited) == len(edits), edited)
    check('无删除项', removed == [], removed)
    after = dmc.scan_file(dest)['props']
    check('creator 已修改', after.get('creator') == '张三', after.get('creator'))
    check('TotalTime 已修改', after.get('TotalTime') == '120', after.get('TotalTime'))
    check('category 已新建', after.get('category') == '年度报告', after.get('category'))
    check('created 已规范化', after.get('created') == '2026-01-02T03:04:05Z', after.get('created'))
    check('自定义属性已修改', after.get('custom:Source') == '人工撰写', after.get('custom:Source'))
    check('未编辑字段不受影响', after.get('title') == 'AI 起草的测试报告')
    check('keywords 未受影响', after.get('keywords') == 'AI, 测试')

    # 新建元素的位置与命名空间要符合 schema（保证 Word 正常打开）
    import xml.etree.ElementTree as ET
    with zipfile.ZipFile(dest) as z:
        core = z.read('docProps/core.xml')
        app = z.read('docProps/app.xml')
    root = ET.fromstring(core)
    order = [dmc._local(c.tag) for c in root]
    check('新建元素符合 schema 顺序',
          order.index('category') > order.index('modified'), order)
    created_el = next(c for c in root if dmc._local(c.tag) == 'created')
    check('created 保留 xsi:type',
          created_el.get('{%s}type' % dmc.XSI_NS) == 'dcterms:W3CDTF',
          created_el.attrib)
    approot = ET.fromstring(app)
    check('app.xml TotalTime 已修改',
          next(c for c in approot if dmc._local(c.tag) == 'TotalTime').text == '120')
    # 正文逐字节不变
    with zipfile.ZipFile(src) as za, zipfile.ZipFile(dest) as zb:
        same = all(za.read(n) == zb.read(n) for n in za.namelist()
                   if not n.startswith('docProps/'))
        check('编辑模式正文逐字节保留', same)


def test_edit_apply_all_and_skip():
    import shutil as sh
    x1 = os.path.join(OUT, 'all1.xlsx')
    p1 = os.path.join(OUT, 'all2.pptx')
    sh.copy(os.path.join(OUT, 'sample.xlsx'), x1)
    sh.copy(os.path.join(OUT, 'sample.pptx'), p1)
    e1, e2 = dmc.scan_file(x1), dmc.scan_file(p1)
    edits = {'creator': {'value': '王五', 'all': True, 'path': x1}}
    d1, s1, sk1 = dmc.collect_ops(e1, [], edits)
    check('collect: 应用到全部-命中', s1 == {'creator': '王五'}, (s1, sk1))
    d2, s2, sk2 = dmc.collect_ops(e2, [], edits)
    check('collect: 应用到全部-第二文件', s2 == {'creator': '王五'}, (s2, sk2))
    for e in (e1, e2):
        dmc.clean_file(e, [], edits, 'overwrite')
    check('应用全部后 xlsx creator', dmc.scan_file(x1)['props'].get('creator') == '王五')
    check('应用全部后 pptx creator', dmc.scan_file(p1)['props'].get('creator') == '王五')

    # 单文件编辑不串到别的文件
    edits2 = {'creator': {'value': '只改 xlsx', 'all': False, 'path': x1}}
    _d, s3, _sk = dmc.collect_ops(e2, [], edits2)
    check('collect: 单文件编辑不影响其他文件', s3 == {}, (s3,))

    # PDF 字段在 OOXML 文件上跳过；未识别字段跳过
    edits3 = {'pdf:Author': {'value': 'x', 'all': True, 'path': ''},
              'core:whatever': {'value': 'y', 'all': True, 'path': ''}}
    _d, _s, sk = dmc.collect_ops(e1, [], edits3)
    check('collect: 不适用字段被跳过', len(sk) == 2, sk)

    # 勾选删除与编辑同字段时，编辑优先（不被删除）
    _d, s4, _sk = dmc.collect_ops(e1, ['creator'], edits)
    check('collect: 编辑优先于删除', s4 == {'creator': '王五'}, (s4,))


def test_edit_pdf():
    src = os.path.join(OUT, 'sample.pdf')
    entry = dmc.scan_file(src)
    edits = {
        'pdf:Author':      {'value': '李四', 'all': False, 'path': src},
        'pdf:Subject':     {'value': '测试主题', 'all': False, 'path': src},   # 新建
        'pdf:ModDate':     {'value': 'D:20260102030405Z', 'all': False, 'path': src},  # GUI 已规范化
    }
    dest, removed, edited = dmc.clean_file(entry, [], edits, 'copy')
    check('pdf 编辑项数', len(edited) == 3, edited)
    after = dmc.scan_file(dest)['props']
    check('pdf Author 已修改', after.get('pdf:Author') == '李四', after.get('pdf:Author'))
    check('pdf Subject 已新建', after.get('pdf:Subject') == '测试主题', after.get('pdf:Subject'))
    check('pdf ModDate 已规范化', after.get('pdf:ModDate') == 'D:20260102030405Z', after.get('pdf:ModDate'))
    check('pdf Title 不受影响', after.get('pdf:Title') == '保留的标题')
    from pypdf import PdfReader
    check('pdf 页数不变', len(PdfReader(dest).pages) == 1)


def main():
    make_samples()
    test_scan()
    test_clean_copy()
    test_clean_overwrite()
    test_clean_pdf()
    test_normalize()
    test_edit_ooxml()
    test_edit_apply_all_and_skip()
    test_edit_pdf()
    print()
    if FAILURES:
        print('FAILED:', len(FAILURES), '项未通过 ->', FAILURES)
        sys.exit(1)
    print('ALL TESTS PASSED')


if __name__ == '__main__':
    main()
