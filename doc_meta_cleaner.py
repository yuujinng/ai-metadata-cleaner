# -*- coding: utf-8 -*-
"""
文档属性清理工具（Document Metadata Cleaner）
================================================
删除 Word / Excel / PPT / PDF 文档属性（元数据）中泄露生成来源的信息，
重点针对 AI 生成留下的痕迹（作者名为 ChatGPT/Claude、生成程序为脚本库等）。

用法：双击「文档属性清理工具.bat」启动，或 `python doc_meta_cleaner.py`。

流程：添加文件 -> 预览属性并勾选要清除的字段（疑似 AI 痕迹自动勾选）
      -> 选择「另存副本 / 覆盖原文件」-> 开始清理。
      也可以双击属性值直接修改字段内容（如把作者改成自己的名字）。

原理：
  * .docx/.xlsx/.pptx（含 .docm/.xlsm/.pptm）是 ZIP+XML 包，
    直接改写其中的 docProps/core.xml、app.xml、custom.xml，其余部件原样保留。
  * .pdf 通过 pypdf 克隆后删除 Info 字典字段及 /Metadata(XMP)。
"""

import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timezone

import tkinter as tk
from tkinter import ttk

# 拖拽放入文件支持（可选依赖，缺失时程序其余功能不受影响）
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    TkinterDnD = None
    DND_FILES = None

# ---------------------------------------------------------------------------
# 字段目录
# ---------------------------------------------------------------------------

# docProps/core.xml（按元素 local name 识别，兼容过渡/严格两套命名空间）
CORE_FIELDS = {
    'title':          '标题',
    'subject':        '主题',
    'creator':        '创建者（作者）',
    'keywords':       '关键词（标记）',
    'description':    '备注（描述）',
    'lastModifiedBy': '最后修改者',
    'revision':       '修订版本号',
    'created':        '创建时间',
    'modified':       '修改时间',
    'category':       '类别',
    'contentStatus':  '内容状态',
    'lastPrinted':    '最后打印时间',
}

# docProps/app.xml
APP_FIELDS = {
    'Application':   '应用程序',
    'AppVersion':    '应用程序版本',
    'Company':       '公司',
    'Manager':       '经理',
    'Template':      '模板',
    'HyperlinkBase': '超链接基础（可能含本机路径）',
    'TotalTime':     '总编辑时间（分钟）',
}

# PDF Info 字典
PDF_FIELDS = {
    'Title':        '标题',
    'Author':       '作者',
    'Subject':      '主题',
    'Keywords':     '关键词',
    'Creator':      '创建程序',
    'Producer':     '生成程序（Producer）',
    'CreationDate': '创建时间',
    'ModDate':      '修改时间',
    'Trapped':      'Trap 状态',
}

XMP_KEY = 'pdf:xmp'
XMP_NAME = 'XMP 元数据包（内嵌 XML，可能重复包含以上信息）'

# 列表展示顺序
FIELD_ORDER = [
    'title', 'subject', 'creator', 'keywords', 'description',
    'lastModifiedBy', 'revision', 'created', 'modified', 'category',
    'contentStatus', 'lastPrinted',
    'Application', 'AppVersion', 'Company', 'Manager', 'Template',
    'HyperlinkBase', 'TotalTime',
]
PDF_ORDER = ['Title', 'Author', 'Subject', 'Keywords', 'Creator',
             'Producer', 'CreationDate', 'ModDate', 'Trapped']

OOXML_EXTS = {'.docx', '.docm', '.xlsx', '.xlsm', '.pptx', '.pptm', '.ppsx', '.ppsm'}
PDF_EXTS = {'.pdf'}
SUPPORTED_EXTS = OOXML_EXTS | PDF_EXTS

# ---------------------------------------------------------------------------
# 疑似 AI / 自动化痕迹识别
# ---------------------------------------------------------------------------

# 中英混排时 \b 会把中文当单词字符导致漏配（如"由AIGC生成"），
# 故用 ASCII 字母数字边界代替 \b。
_L = r'(?<![a-z0-9])'
_R = r'(?![a-z0-9])'

_AI_VALUE_PATTERN = re.compile(
    r"chatgpt|openai|"
    + _L + r"gpt[\s-]?\d|"
    + _L + r"claude" + _R + r"|anthropic|"
    + _L + r"gemini" + _R + r"|copilot|deepseek|"
    + _L + r"qwen" + _R + r"|通义|千问|"
    + _L + r"kimi" + _R + r"|moonshot|doubao|豆包|文心|"
    + _L + r"ernie" + _R + r"|"
    + _L + r"llama" + _R + r"|"
    + _L + r"grok" + _R + r"|"
    + _L + r"xai" + _R + r"|"
    + _L + r"x\.ai" + _R + r"|mistral|zhipu|智谱|z\.ai|bigmodel|"
    + _L + r"glm" + _R + r"|hunyuan|混元|星火|sparkdesk|讯飞|灵犀|"
    r"python-docx|python-pptx|openpyxl|xlsxwriter|reportlab|"
    + _L + r"fpdf" + _R + r"|weasyprint|pandoc|apache\s+poi|skia/?pdf|"
    r"notion\.so|gamma\.app|beautiful\.ai|tome\.app|wps\s*ai" + _R + r"|"
    + _L + r"aigc" + _R + r"|ai[-\s]?generated|written\s+by\s+(?:ai|chatgpt|gpt)"
    r"|由\s*ai\s*生成|ai\s*生成",
    re.IGNORECASE,
)


def looks_like_ai(value):
    """值里是否带有 AI 工具或脚本生成库的痕迹。"""
    if not value:
        return False
    return bool(_AI_VALUE_PATTERN.search(value))


# ---------------------------------------------------------------------------
# 编辑值校验 / 日期格式转换
# ---------------------------------------------------------------------------

OOXML_DATE_FIELDS = ('created', 'modified', 'lastPrinted')
PDF_DATE_FIELDS = ('pdf:CreationDate', 'pdf:ModDate')
INT_FIELDS = ('revision', 'TotalTime')
_DT_COMPACT = re.compile(r'^(?:D:)?(\d{4})(\d{1,2})(\d{1,2})(\d{1,2})?:?(\d{1,2})?:?(\d{1,2})?:?$')


def _parse_dt(s):
    """把用户输入解析为 datetime；兼容 ISO、带空格、纯数字日期、PDF D: 前缀。"""
    s = (s or '').strip()
    if not s:
        return None
    ends_utc = s.endswith(('Z', 'z', '+00:00'))
    m = _DT_COMPACT.match(s)
    if m:
        y, mo, d, h, mi, se = ((m.group(i) or 0) for i in range(1, 7))
        try:
            dt = datetime(int(y), int(mo), int(d), int(h), int(mi), int(se))
        except ValueError:
            return None
        return dt.replace(tzinfo=timezone.utc) if ends_utc else dt
    t = s[:-1] if ends_utc and s.endswith(('Z', 'z')) else s
    try:
        dt = datetime.fromisoformat(t.replace(' ', 'T'))
    except ValueError:
        return None
    if ends_utc:
        return dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc) if dt.tzinfo else dt


def normalize_edit_value(kind, key, raw):
    """校验并规范化编辑输入。返回 (错误信息, 规范化值)；raw 为空表示删除该字段。"""
    raw = (raw or '').strip()
    if not raw:
        return None, ''
    is_date = (key in OOXML_DATE_FIELDS if kind == 'ooxml'
               else key in PDF_DATE_FIELDS)
    if is_date:
        dt = _parse_dt(raw)
        if dt is None:
            return '日期无法识别，请类似输入：2026-10-20 10:00:05', None
        if kind == 'pdf':
            return None, 'D:' + dt.strftime('%Y%m%d%H%M%SZ')
        return None, dt.strftime('%Y-%m-%dT%H:%M:%SZ')
    if key in INT_FIELDS and not re.fullmatch(r'\d+', raw):
        return '该字段需要非负整数', None
    return None, raw


# ---------------------------------------------------------------------------
# OOXML（Word / Excel / PPT）
# ---------------------------------------------------------------------------

CP_NS = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
DC_NS = 'http://purl.org/dc/elements/1.1/'
DCTERMS_NS = 'http://purl.org/dc/terms/'
XSI_NS = 'http://www.w3.org/2001/XMLSchema-instance'
APP_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
VT_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'

# 新建 core.xml 元素时的命名空间与 schema 顺序（插入到正确位置，保证 Word 正常打开）
CORE_CREATE_NS = {
    'title': DC_NS, 'subject': DC_NS, 'creator': DC_NS, 'description': DC_NS,
    'keywords': CP_NS, 'lastModifiedBy': CP_NS, 'revision': CP_NS,
    'created': DCTERMS_NS, 'modified': DCTERMS_NS, 'category': CP_NS,
    'contentStatus': CP_NS, 'lastPrinted': CP_NS,
}
CORE_ORDER = ['title', 'subject', 'creator', 'keywords', 'description',
              'lastModifiedBy', 'revision', 'created', 'modified', 'category',
              'contentStatus', 'lastPrinted']
APP_ORDER = ['Template', 'Manager', 'Company', 'Pages', 'Words', 'Characters',
             'PresentationFormat', 'Lines', 'Paragraphs', 'Slides', 'Notes',
             'TotalTime', 'HiddenSlides', 'MMClips', 'ScaleCrop', 'HeadingPairs',
             'TitlesOfParts', 'LinksUpToDate', 'CharactersWithSpaces', 'SharedDoc',
             'HyperlinkBase', 'HyperlinksChanged', 'DigSig', 'Application',
             'AppVersion', 'DocSecurity']


def _local(tag):
    """'{namespace}localname' -> 'localname'"""
    return tag.rsplit('}', 1)[-1]


def _elem_text(elem):
    return ''.join(elem.itertext()).strip()


def read_ooxml_props(path):
    """返回 (props: {key: 值}, suspects: {key, ...})，key 见 CORE/APP/custom 规则。"""
    props, suspects = {}, set()
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        if 'docProps/core.xml' in names:
            root = ET.fromstring(z.read('docProps/core.xml'))
            for child in root:
                name = _local(child.tag)
                key = name if name in CORE_FIELDS else 'core:' + name
                props[key] = _elem_text(child)
        if 'docProps/app.xml' in names:
            root = ET.fromstring(z.read('docProps/app.xml'))
            for child in root:
                name = _local(child.tag)
                key = name if name in APP_FIELDS else 'app:' + name
                props[key] = _elem_text(child)
        if 'docProps/custom.xml' in names:
            root = ET.fromstring(z.read('docProps/custom.xml'))
            for child in root:
                cname = child.get('name')
                if cname:
                    props['custom:' + cname] = _elem_text(child)

    # 规则化判断疑似痕迹
    for key, val in props.items():
        if looks_like_ai(val):
            suspects.add(key)
    # 创建时间 == 修改时间（精确到秒）：典型的机器一次性生成
    if props.get('created') and props.get('created') == props.get('modified'):
        suspects.update(('created', 'modified'))
    # 总编辑时间为 0 分钟
    if props.get('TotalTime') in ('0', '0分钟'):
        suspects.add('TotalTime')
    return props, suspects


_XMLNS_RE = re.compile(rb'xmlns:([\w.-]+)\s*=\s*"([^"]+)"')
_DEFAULT_NS_RE = re.compile(rb'xmlns\s*=\s*"([^"]+)"')


def _serialize(root, original_bytes, extra_ns=None):
    """尽量按原文件的命名空间前缀与 XML 声明重新序列化。

    extra_ns: {uri: 前缀}，为新建元素提供命名空间前缀（原文件没有该 uri 时才生效）。
    """
    seen = {}
    for m in _XMLNS_RE.finditer(original_bytes):
        pfx, uri = m.group(1).decode('ascii'), m.group(2).decode('ascii')
        seen[uri] = pfx
        ET.register_namespace(pfx, uri)
    for m in _DEFAULT_NS_RE.finditer(original_bytes):
        seen[m.group(1).decode('ascii')] = ''
        ET.register_namespace('', m.group(1).decode('ascii'))
    for uri, pfx in (extra_ns or {}).items():
        if uri not in seen:
            ET.register_namespace(pfx, uri)

    body = ET.tostring(root, encoding='unicode')
    decl_end = original_bytes.find(b'?>')
    if decl_end != -1:
        out = original_bytes[:decl_end + 2].decode('utf-8', 'replace') + body
    else:
        out = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + body
    return out.encode('utf-8')


def _child_key(child):
    """按读取规则推导子元素的属性 key（与 read_ooxml_props 一致）。"""
    cname = child.get('name')
    name = _local(child.tag)
    if cname is not None:
        return 'custom:' + cname
    if name in CORE_FIELDS or name in APP_FIELDS:
        return name
    return ('core:' if name[0].islower() else 'app:') + name


def _set_or_create(root, key, value, part, extra_changed):
    """把 key 的值设为 value；元素不存在时按 schema 顺序新建（仅 core/app）。"""
    existing = None
    for child in root:
        if child.get('name') is None and _local(child.tag) == key:
            existing = child
            break
    if existing is not None:
        for sub in list(existing):
            existing.remove(sub)
        existing.text = value
        return True
    if part == 'core':
        ns = CORE_CREATE_NS.get(key)
        if ns is None:
            return extra_changed
        order = CORE_ORDER
    else:
        ns, order = APP_NS, APP_ORDER
    elem = ET.Element('{%s}%s' % (ns, key))
    elem.text = value
    if part == 'core' and key in ('created', 'modified'):
        elem.set('{%s}type' % XSI_NS, 'dcterms:W3CDTF')
    idx = order.index(key)
    for i, child in enumerate(root):
        if child.get('name') is not None:
            continue
        ci = order.index(_local(child.tag)) if _local(child.tag) in order else None
        if ci is not None and ci > idx:
            root.insert(i, elem)
            return True
    root.append(elem)
    return True


def _rewrite_part(xml_bytes, wanted, part_ops, part):
    """对单个属性部件执行删除（wanted）与修改（part_ops: key -> value）。

    返回 (新字节或原字节, 删除的 key 列表, 修改的 key 列表)。
    """
    root = ET.fromstring(xml_bytes)
    removed, edited = [], []
    to_drop = []
    for child in list(root):
        key = _child_key(child)
        if key in wanted:
            to_drop.append(child)
            removed.append(key)
    for child in to_drop:
        root.remove(child)

    if part == 'custom':
        # 自定义属性：只改已有 Property 的值文本，不支持新建
        for key, value in part_ops.items():
            cname = key[len('custom:'):]
            target = next((c for c in root if c.get('name') == cname), None)
            if target is None:
                continue
            vt = list(target)[0] if len(target) else None
            if vt is None:
                vt = ET.SubElement(target, '{%s}lpstr' % VT_NS)
            for sub in list(vt):
                vt.remove(sub)
            vt.text = value
            edited.append(key)
    else:
        for key, value in part_ops.items():
            if _set_or_create(root, key, value, part, False):
                edited.append(key)

    if not removed and not edited:
        return xml_bytes, removed, edited
    extra_ns = {CP_NS: 'cp', DC_NS: 'dc', DCTERMS_NS: 'dcterms', XSI_NS: 'xsi',
                APP_NS: '', VT_NS: 'vt'}
    return _serialize(root, xml_bytes, extra_ns), removed, edited


def clean_ooxml(src, dst, delete_keys, set_ops):
    """改写 OOXML 包：删除 delete_keys 字段并把 set_ops 字段设为新值。

    其余部件逐字节保留。返回 (删除的 key 列表, 修改的 key 列表)。
    """
    wanted = set(delete_keys)
    core_ops = {k: v for k, v in set_ops.items() if k in CORE_FIELDS}
    app_ops = {k: v for k, v in set_ops.items() if k in APP_FIELDS}
    custom_ops = {k: v for k, v in set_ops.items() if k.startswith('custom:')}
    removed, edited = [], []
    with zipfile.ZipFile(src) as zin:
        with zipfile.ZipFile(dst, 'w') as zout:
            zout.comment = zin.comment
            for info in zin.infolist():
                data = zin.read(info.filename)
                fname = info.filename.replace('\\', '/')
                rm, ed = [], []
                if fname == 'docProps/core.xml':
                    data, rm, ed = _rewrite_part(data, wanted, core_ops, 'core')
                elif fname == 'docProps/app.xml':
                    data, rm, ed = _rewrite_part(data, wanted, app_ops, 'app')
                elif fname == 'docProps/custom.xml':
                    data, rm, ed = _rewrite_part(data, wanted, custom_ops, 'custom')
                removed.extend(rm)
                edited.extend(ed)
                zi = zipfile.ZipInfo(fname, date_time=info.date_time)
                zi.compress_type = info.compress_type
                zi.external_attr = info.external_attr
                zi.comment = info.comment
                zout.writestr(zi, data)
    return removed, edited


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------

def read_pdf_props(path):
    try:
        from pypdf import PdfReader
    except ImportError:
        raise ValueError('PDF 支持组件未安装：请在命令行运行  pip install pypdf  后重试')

    props, suspects = {}, set()
    reader = PdfReader(path)
    if reader.is_encrypted:
        raise ValueError('PDF 已加密，无法读取属性')
    meta = reader.metadata
    if meta:
        for k, v in meta.items():
            key = 'pdf:' + str(k).lstrip('/')
            props[key] = str(v)
    try:
        xmp = reader.xmp_metadata
        if xmp is not None:
            raw = xmp.stream.get_data()
            props[XMP_KEY] = '存在（约 %d 字节，含内嵌 XML 描述）' % len(raw)
    except Exception:
        pass

    for key, val in props.items():
        if looks_like_ai(val):
            suspects.add(key)
    try:
        xmp = reader.xmp_metadata
        if xmp is not None and looks_like_ai(xmp.stream.get_data().decode('utf-8', 'replace')):
            suspects.add(XMP_KEY)
    except Exception:
        pass
    return props, suspects


def clean_pdf(src, dst, delete_keys, set_ops):
    """删除 delete_keys 并把 set_ops（key -> value）写入 PDF Info 字典。

    返回 (删除的 key 列表, 修改的 key 列表)。
    """
    try:
        from pypdf import PdfReader, PdfWriter
        from pypdf.generic import NameObject
    except ImportError:
        raise ValueError('PDF 支持组件未安装：请在命令行运行  pip install pypdf  后重试')

    wanted = set(delete_keys)
    removed, edited = [], []
    writer = PdfWriter(clone_from=src)
    info = getattr(writer, '_info', None)
    if info is None:
        info = writer._info_obj
    for k in list(info.keys()):
        key = str(k).lstrip('/')
        if ('pdf:' + key) in wanted:
            del info[NameObject(str(k))]
            removed.append('pdf:' + key)
    if XMP_KEY in wanted and '/Metadata' in writer.root_object:
        del writer.root_object[NameObject('/Metadata')]
        removed.append(XMP_KEY)

    updates = {}
    for key, value in set_ops.items():
        if not key.startswith('pdf:'):
            continue
        updates[NameObject('/' + key[4:])] = value
        edited.append(key)
    if updates:
        writer.add_metadata(updates)
    with open(dst, 'wb') as f:
        writer.write(f)
    return removed, edited


# ---------------------------------------------------------------------------
# 统一入口
# ---------------------------------------------------------------------------

def field_display_name(key):
    if key in CORE_FIELDS:
        return CORE_FIELDS[key]
    if key in APP_FIELDS:
        return APP_FIELDS[key]
    if key == XMP_KEY:
        return XMP_NAME
    if key.startswith('pdf:'):
        return 'PDF ' + PDF_FIELDS.get(key[4:], key[4:])
    if key.startswith('custom:'):
        return '自定义属性：' + key[7:]
    if key.startswith('core:') or key.startswith('app:'):
        return key.split(':', 1)[1] + '（未识别字段）'
    return key


def field_sort_key(key):
    if key in FIELD_ORDER:
        return FIELD_ORDER.index(key)
    if key.startswith('pdf:'):
        local = key[4:]
        return 500 + (PDF_ORDER.index(local) if local in PDF_ORDER else 50)
    if key.startswith('custom:'):
        return 700
    if key == XMP_KEY:
        return 900
    return 950


def scan_file(path):
    """读取单个文件的属性。返回 dict(path/kind/props/suspects)。"""
    ext = os.path.splitext(path)[1].lower()
    if ext in OOXML_EXTS:
        kind = 'ooxml'
        props, suspects = read_ooxml_props(path)
    elif ext in PDF_EXTS:
        kind = 'pdf'
        props, suspects = read_pdf_props(path)
    else:
        raise ValueError('不支持的格式：' + ext)
    return {'path': path, 'kind': kind, 'props': props, 'suspects': suspects}


def collect_ops(entry, checked_keys, edits):
    """为一个文件汇总要执行的删除与修改。

    edits: {key: {'value': 非空新值, 'all': 是否应用到全部, 'path': 编辑时选中文件}}
    返回 (删除 key 列表, 修改 {key: value}, 跳过 [(key, 原因)])。
    """
    delete = [k for k in checked_keys if k in entry['props'] and k not in edits]
    sets, skipped = {}, []
    for key, e in edits.items():
        if not (e['all'] or e['path'] == entry['path']):
            continue
        if key == XMP_KEY:
            skipped.append((key, 'XMP 不支持直接编辑'))
            continue
        if entry['kind'] == 'pdf':
            if not key.startswith('pdf:'):
                skipped.append((key, 'PDF 文件不适用'))
                continue
            sets[key] = e['value']
            continue
        if key.startswith('custom:'):
            if key in entry['props']:
                sets[key] = e['value']
            else:
                skipped.append((key, '不支持新建自定义属性'))
            continue
        if key in CORE_FIELDS or key in APP_FIELDS:
            sets[key] = e['value']
        else:
            skipped.append((key, '未识别字段'))
    return delete, sets, skipped


def _sorted_edited(edited_keys, set_ops):
    """把实际修改的 key 列表转为 [(key, 新值)] 并按展示顺序排序。"""
    return sorted(((k, set_ops[k]) for k in edited_keys if k in set_ops),
                  key=lambda kv: field_sort_key(kv[0]))


def clean_file(entry, checked_keys, edits, mode):
    """按勾选字段与编辑项处理文件。mode: 'copy' 另存副本 / 'overwrite' 覆盖。

    返回 (输出路径, 删除的 key 列表, [(key, 新值)])。
    """
    path = entry['path']
    delete_keys, set_ops, _skipped = collect_ops(entry, checked_keys, edits)
    if not delete_keys and not set_ops:
        return path, [], []

    def run(a, b):
        if entry['kind'] == 'ooxml':
            return clean_ooxml(a, b, delete_keys, set_ops)
        return clean_pdf(a, b, delete_keys, set_ops)

    if mode == 'overwrite':
        tmp = path + '.meta_tmp'
        try:
            removed, edited = run(path, tmp)
            os.replace(tmp, path)
        finally:
            if os.path.exists(path + '.meta_tmp'):
                try:
                    os.remove(path + '.meta_tmp')
                except OSError:
                    pass
        return path, removed, _sorted_edited(edited, set_ops)

    # 另存副本
    base, ext = os.path.splitext(path)
    dest = base + '_已清理' + ext
    n = 1
    while os.path.exists(dest):
        dest = base + '_已清理%d' % n + ext
        n += 1
    removed, edited = run(path, dest)
    return dest, removed, _sorted_edited(edited, set_ops)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

_Base = TkinterDnD.Tk if TkinterDnD else tk.Tk


class CleanerApp(_Base):
    def __init__(self):
        super().__init__()
        self.title('文档属性清理工具 — 去除 AI 生成痕迹')
        self.geometry('1080x680')
        self.minsize(880, 560)

        self.files = []          # [ {path, kind, props, suspects} ]
        self.fields = {}         # key -> {name, suspect, checked}
        self.edits = {}          # key -> {'value', 'all', 'path'}（待应用的修改）
        self.current_path = None

        self._build_style()
        self._build_ui()
        self.log('提示：清理前请先关闭正在 Word/WPS 中编辑的文档，否则文件被占用会失败。')
        self.log('步骤：1) 添加或拖入文件 → 2) 在右侧勾选要清除的属性 → '
                 '3) 选择保存方式 → 4) 开始清理。')
        try:
            import pypdf  # noqa: F401
            self.log('PDF 支持：已就绪。')
        except ImportError:
            self.log('PDF 支持：未安装（.pdf 文件将无法读取）。'
                     '启动器会自动安装，也可手动运行：pip install pypdf', 'err')
        if TkinterDnD is not None:
            self._enable_dragdrop()
            self.log('已启用拖拽：把文件或文件夹直接拖进本窗口即可。')
        else:
            self.log('拖拽功能未启用（可选组件 tkinterdnd2 未安装，其余功能不受影响）')

    # ---- 界面 -------------------------------------------------------------
    def _build_style(self):
        style = ttk.Style(self)
        for theme in ('vista', 'winnative', 'clam'):
            if theme in style.theme_names():
                style.theme_use(theme)
                break
        style.configure('Treeview', rowheight=26)
        style.configure('Start.TButton', font=('Microsoft YaHei UI', 10, 'bold'))

    def _build_ui(self):
        # 顶部工具栏
        bar = ttk.Frame(self, padding=(8, 6))
        bar.pack(fill='x')
        ttk.Button(bar, text='添加文件…', command=self.add_files).pack(side='left')
        ttk.Button(bar, text='添加文件夹…', command=self.add_folder).pack(side='left', padx=(6, 0))
        ttk.Button(bar, text='移除所选文件', command=self.remove_selected).pack(side='left', padx=(6, 0))
        ttk.Button(bar, text='清空列表', command=self.clear_files).pack(side='left', padx=(6, 0))
        self.auto_check_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            bar, text='添加时自动勾选疑似 AI 痕迹',
            variable=self.auto_check_var).pack(side='right')

        # 主体：左文件列表 / 右属性表
        body = ttk.Panedwindow(self, orient='horizontal')
        body.pack(fill='both', expand=True, padx=8, pady=(0, 4))

        left = ttk.Frame(body)
        ttk.Label(left, text='文件列表').pack(anchor='w')
        listframe = ttk.Frame(left)
        listframe.pack(fill='both', expand=True)
        self.file_list = tk.Listbox(listframe, selectmode='single', activestyle='dotbox')
        lsb = ttk.Scrollbar(listframe, orient='vertical', command=self.file_list.yview)
        self.file_list.config(yscrollcommand=lsb.set)
        self.file_list.pack(side='left', fill='both', expand=True)
        lsb.pack(side='right', fill='y')
        self.file_list.bind('<<ListboxSelect>>', self.on_file_selected)
        body.add(left, weight=1)

        right = ttk.Frame(body)
        ttk.Label(right, text='属性预览（勾选=清除该字段；双击“当前值”一列可直接编辑修改）').pack(anchor='w')
        treeframe = ttk.Frame(right)
        treeframe.pack(fill='both', expand=True)
        cols = ('chk', 'name', 'value', 'flag')
        self.tree = ttk.Treeview(treeframe, columns=cols, show='headings', selectmode='browse')
        self.tree.heading('chk', text='清除')
        self.tree.column('chk', width=48, anchor='center', stretch=False)
        self.tree.heading('name', text='字段')
        self.tree.column('name', width=185, anchor='w')
        self.tree.heading('value', text='当前值（所选文件）')
        self.tree.column('value', width=330, anchor='w')
        self.tree.heading('flag', text='标记')
        self.tree.column('flag', width=100, anchor='center', stretch=False)
        tsb = ttk.Scrollbar(treeframe, orient='vertical', command=self.tree.yview)
        self.tree.config(yscrollcommand=tsb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        tsb.pack(side='right', fill='y')
        self.tree.tag_configure('suspect', foreground='#c0392b')
        self.tree.tag_configure('checked', background='#eef6ee')
        self.tree.tag_configure('edited', foreground='#1a5fb4')
        self.tree.bind('<Button-1>', self.on_tree_click)
        self.tree.bind('<Double-1>', self.on_tree_double)
        self.tree.bind('<space>', self.on_tree_space)
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_selected)

        ops = ttk.Frame(right)
        ops.pack(fill='x', pady=3)
        ttk.Button(ops, text='全选', command=lambda: self.set_all(True)).pack(side='left')
        ttk.Button(ops, text='全不选', command=lambda: self.set_all(False)).pack(side='left', padx=(6, 0))
        ttk.Button(ops, text='仅勾选疑似 AI 痕迹', command=self.check_suspects_only).pack(side='left', padx=(6, 0))
        ttk.Button(ops, text='编辑所选…', command=self.open_edit_dialog).pack(side='left', padx=(6, 0))
        ttk.Button(ops, text='撤销全部修改', command=self.undo_edits).pack(side='left', padx=(6, 0))
        self.detail = ttk.Label(right, text='', wraplength=560, foreground='#555')
        self.detail.pack(anchor='w', pady=(2, 0))
        body.add(right, weight=3)

        # 底部：保存方式 + 开始
        bottom = ttk.Frame(self, padding=(8, 4))
        bottom.pack(fill='x')
        ttk.Label(bottom, text='保存方式：').pack(side='left')
        self.save_var = tk.StringVar(value='copy')
        ttk.Radiobutton(bottom, text='另存副本（推荐，原名_已清理）',
                        variable=self.save_var, value='copy').pack(side='left')
        ttk.Radiobutton(bottom, text='覆盖原文件',
                        variable=self.save_var, value='overwrite').pack(side='left', padx=(10, 0))
        self.start_btn = ttk.Button(bottom, text='开始清理', style='Start.TButton', command=self.on_clean)
        self.start_btn.pack(side='right')
        self.status = ttk.Label(self, text='共 0 个文件', padding=(10, 2), foreground='#666')
        self.status.pack(anchor='w')

        # 日志
        logframe = ttk.Frame(self, padding=(8, 0))
        logframe.pack(fill='both')
        ttk.Label(logframe, text='处理日志').pack(anchor='w')
        from tkinter.scrolledtext import ScrolledText
        self.log_text = ScrolledText(logframe, height=7, state='disabled',
                                     font=('Consolas', 9), wrap='word')
        self.log_text.pack(fill='both', expand=True)
        for tag, color in (('ok', '#1a7f37'), ('err', '#c0392b'), ('info', '#555')):
            self.log_text.tag_configure(tag, foreground=color)

    # ---- 日志 / 状态 ------------------------------------------------------
    def log(self, msg, tag='info'):
        self.log_text.config(state='normal')
        self.log_text.insert('end', msg + '\n', tag)
        self.log_text.see('end')
        self.log_text.config(state='disabled')
        self.update_idletasks()

    def update_status(self):
        checked = sum(1 for v in self.fields.values() if v['checked'])
        self.status.config(text='共 %d 个文件，已勾选 %d 个字段，待修改 %d 项'
                           % (len(self.files), checked, len(self.edits)))

    # ---- 拖拽放入 ---------------------------------------------------------
    def _enable_dragdrop(self):
        # Windows 上每个 tk 子窗口都是独立的 HWND，需要分别注册为拖放目标
        for w in (self, self.file_list, self.tree, self.log_text):
            try:
                w.drop_target_register(DND_FILES)
                w.dnd_bind('<<Drop>>', self.on_drop)
            except Exception:
                pass

    def on_drop(self, event):
        """接收拖入的文件/文件夹，文件夹会被递归展开。"""
        try:
            raw = self.tk.splitlist(event.data)
        except Exception:
            return 'copy'
        files = []
        for p in raw:
            if os.path.isdir(p):
                for dirpath, _dirs, fns in os.walk(p):
                    files.extend(
                        os.path.join(dirpath, fn) for fn in fns
                        if os.path.splitext(fn)[1].lower() in SUPPORTED_EXTS)
            elif os.path.isfile(p):
                files.append(p)
        if not files:
            self.log('拖入内容里没有支持的文档（docx/xlsx/pptx/pdf 及对应带宏格式）', 'err')
            return 'copy'
        files.sort()
        before = len(self.files)
        self._add_paths(files)
        added = len(self.files) - before
        if added:
            self.log('—— 拖入新增 %d 个文件 ——' % added, 'ok')
        return 'copy'

    # ---- 文件列表 ---------------------------------------------------------
    def add_files(self):
        from tkinter import filedialog
        paths = filedialog.askopenfilenames(
            parent=self, title='选择文档',
            filetypes=[('支持的文档', '*.docx *.docm *.xlsx *.xlsm *.pptx *.pptm *.ppsx *.ppsm *.pdf'),
                       ('所有文件', '*.*')])
        self._add_paths(paths)

    def add_folder(self):
        from tkinter import filedialog
        folder = filedialog.askdirectory(parent=self, title='选择文件夹（会递归查找支持的文档）')
        if not folder:
            return
        found = []
        for dirpath, _dirnames, filenames in os.walk(folder):
            for fn in filenames:
                if os.path.splitext(fn)[1].lower() in SUPPORTED_EXTS:
                    found.append(os.path.join(dirpath, fn))
        if not found:
            self.log('文件夹中没有找到支持的文档：' + folder)
            return
        self._add_paths(sorted(found))

    def _add_paths(self, paths):
        known = {f['path'].lower() for f in self.files}
        added = failed = 0
        for p in paths:
            if p.lower() in known:
                continue
            try:
                entry = scan_file(p)
            except Exception as e:
                failed += 1
                self.log('✗ 无法读取 %s：%s' % (os.path.basename(p), e), 'err')
                continue
            self.files.append(entry)
            known.add(p.lower())
            added += 1
            n_sus = len(entry['suspects'])
            if n_sus:
                self.log('＋ %s（发现 %d 项疑似 AI 痕迹）' % (os.path.basename(p), n_sus), 'ok')
            else:
                self.log('＋ %s' % os.path.basename(p))
        self._rebuild_fields()
        self._refresh_list()
        if self.files and self.current_path is None:
            self.file_list.selection_set(0)
            self.current_path = self.files[0]['path']
        self._refresh_tree()
        self.update_status()

    def remove_selected(self):
        sel = self.file_list.curselection()
        if not sel:
            return
        del self.files[sel[0]]
        self.current_path = None
        self.detail.config(text='')
        self._rebuild_fields()
        self._refresh_list()
        self._refresh_tree()
        self.update_status()

    def clear_files(self):
        self.files.clear()
        self.current_path = None
        self.detail.config(text='')
        self._rebuild_fields()
        self._refresh_list()
        self._refresh_tree()
        self.update_status()

    def _refresh_list(self):
        self.file_list.delete(0, 'end')
        for f in self.files:
            self.file_list.insert('end', os.path.basename(f['path']))

    def on_file_selected(self, _event):
        sel = self.file_list.curselection()
        if sel and sel[0] < len(self.files):
            self.current_path = self.files[sel[0]]['path']
        self._refresh_values()

    # ---- 属性表 -----------------------------------------------------------
    def _rebuild_fields(self):
        prev = self.fields
        new_fields = {}
        auto = self.auto_check_var.get()
        for f in self.files:
            for key in f['props']:
                entry = new_fields.get(key)
                if entry is None:
                    was = prev.get(key)
                    suspect = key in f['suspects']
                    checked = was['checked'] if was else (suspect and auto)
                    new_fields[key] = {'name': field_display_name(key),
                                       'suspect': suspect, 'checked': checked}
                elif key in f['suspects']:
                    entry['suspect'] = True
        self.fields = new_fields

    def _current_entry(self):
        for f in self.files:
            if f['path'] == self.current_path:
                return f
        return None

    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        entry = self._current_entry() or (self.files[0] if self.files else None)
        for key in sorted(self.fields, key=field_sort_key):
            meta = self.fields[key]
            value, flag, tags = self._row_values(key, meta, entry)
            self.tree.insert('', 'end', iid=key, values=(
                self._glyph(meta['checked']), meta['name'], value, flag), tags=tags)

    def _row_values(self, key, meta, entry):
        """计算一行的 值/标记/样式，包含待应用修改的展示。"""
        val = entry['props'].get(key) if entry is not None else None
        disp = '—' if val is None else str(val)
        flag = '⚠ AI痕迹' if meta['suspect'] else ''
        tags = []
        edit = self.edits.get(key)
        if edit and entry is not None and self._edit_applies(entry, key):
            disp = '%s  →  %s' % (val if val is not None else '（无）', edit['value'])
            flag = (flag + ' ' if flag else '') + '✎ 待修改'
            tags.append('edited')
        if meta['suspect']:
            tags.append('suspect')
        if meta['checked']:
            tags.append('checked')
        return _truncate(disp, 90), flag, tags

    def _edit_applies(self, entry, key):
        edit = self.edits.get(key)
        if not edit or entry is None:
            return False
        return edit['all'] or edit['path'] == entry['path']

    def _refresh_values(self):
        entry = self._current_entry()
        for key in self.fields:
            value, flag, _tags = self._row_values(key, self.fields[key], entry)
            self.tree.set(key, 'value', value)
            self.tree.set(key, 'flag', flag)
            self.tree.item(key, tags=_tags)

    @staticmethod
    def _glyph(checked):
        return '☑' if checked else '☐'

    def _set_checked(self, key, checked):
        """设置勾选状态；勾选清除与待应用修改互斥。"""
        meta = self.fields.get(key)
        if meta is None:
            return
        meta['checked'] = checked
        if checked:
            self.edits.pop(key, None)
        value, flag, tags = self._row_values(key, meta, self._current_entry())
        self.tree.set(key, 'chk', self._glyph(checked))
        self.tree.set(key, 'value', value)
        self.tree.set(key, 'flag', flag)
        self.tree.item(key, tags=tags)
        self.update_status()

    def _toggle(self, iid):
        if iid and iid in self.fields:
            self._set_checked(iid, not self.fields[iid]['checked'])

    def on_tree_click(self, event):
        if self.tree.identify('region', event.x, event.y) != 'cell':
            return
        if self.tree.identify_column(event.x) == '#1':
            self._toggle(self.tree.identify_row(event.y))
            return 'break'

    def on_tree_double(self, event):
        if self.tree.identify('region', event.x, event.y) != 'cell':
            return
        if self.tree.identify_column(event.x) == '#3':
            self.open_edit_dialog()
            return 'break'
        iid = self.tree.identify_row(event.y)
        if iid:
            self._toggle(iid)
            return 'break'

    def on_tree_space(self, _event):
        sel = self.tree.selection()
        if sel:
            self._toggle(sel[0])
            return 'break'

    def on_tree_selected(self, _event):
        sel = self.tree.selection()
        if not sel:
            return
        key = sel[0]
        entry = self._current_entry()
        val = entry['props'].get(key, '（该文件无此字段）') if entry else '（未知文件）'
        text = '%s：%s' % (self.fields[key]['name'], val)
        if self._edit_applies(entry or {}, key):
            text += '；将改为：%s' % self.edits[key]['value']
        self.detail.config(text=text)

    def set_all(self, checked):
        for key in self.fields:
            self._set_checked(key, checked)

    def check_suspects_only(self):
        self.set_all(False)
        for key, meta in self.fields.items():
            if meta['suspect']:
                self._set_checked(key, True)

    # ---- 编辑属性 ---------------------------------------------------------
    def _edit_hint(self, kind, key):
        if kind == 'ooxml' and key in OOXML_DATE_FIELDS:
            return '日期字段，格式如：2026-10-20 10:00:05（保存时自动转为文档标准格式）'
        if kind == 'pdf' and key in PDF_DATE_FIELDS:
            return '日期字段，格式如：2026-10-20 10:00:05（保存时自动转为 PDF 标准格式）'
        if key in INT_FIELDS:
            return '该字段需要非负整数（修订次数 / 编辑分钟数）'
        return '留空保存 = 清除该字段；否则将写入所输入的内容'

    def _apply_edit(self, key, raw, apply_all):
        """校验并保存一个待应用修改。返回 (是否成功, 错误信息)。"""
        entry = self._current_entry()
        if entry is None or key not in self.fields:
            return False, '没有选中文件或属性行'
        err, value = normalize_edit_value(entry['kind'], key, raw)
        if err:
            return False, err
        if value == '':
            # 留空 = 清除：交给已有的勾选删除流程
            self.edits.pop(key, None)
            self._set_checked(key, True)
        else:
            self.edits[key] = {'value': value, 'all': bool(apply_all),
                               'path': entry['path']}
            self._set_checked(key, False)
        self._refresh_tree()
        self.update_status()
        return True, ''

    def undo_edits(self):
        if not self.edits:
            self.log('当前没有待应用的修改。')
            return
        self.edits.clear()
        self._refresh_tree()
        self.update_status()
        self.log('已撤销全部待应用的修改。')

    def open_edit_dialog(self):
        from tkinter import messagebox
        entry = self._current_entry()
        sel = self.tree.selection()
        if entry is None or not sel:
            messagebox.showinfo('提示', '请先在左侧选中文件，再在右侧选中要编辑的属性行。')
            return
        key = sel[0]
        if key == XMP_KEY:
            messagebox.showinfo('提示', 'XMP 元数据包不支持直接编辑，可勾选清除整个包。')
            return
        if key.startswith('custom:') and key not in entry['props']:
            messagebox.showinfo('提示', '该文件没有这个自定义属性，不支持新建自定义属性。')
            return
        meta = self.fields[key]
        current = entry['props'].get(key)
        pending = self.edits.get(key)

        dlg = tk.Toplevel(self)
        dlg.title('编辑属性：%s' % meta['name'])
        dlg.transient(self)
        dlg.resizable(False, False)
        frm = ttk.Frame(dlg, padding=14)
        frm.pack(fill='both', expand=True)
        ttk.Label(frm, text='文件：%s' % os.path.basename(entry['path'])).pack(anchor='w')
        cur_text = '（该文件无此字段，保存后将新建）' if current is None else current
        ttk.Label(frm, text='当前值：%s' % _truncate(cur_text, 80),
                  wraplength=520).pack(anchor='w', pady=(2, 8))
        var = tk.StringVar(value=pending['value'] if pending else (current if current is not None else ''))
        ent = ttk.Entry(frm, textvariable=var, width=70)
        ent.pack(fill='x')
        ent.focus_set()
        ent.select_range(0, 'end')
        ttk.Label(frm, text=self._edit_hint(entry['kind'], key),
                  foreground='#777', wraplength=520).pack(anchor='w', pady=(4, 6))
        all_var = tk.BooleanVar(value=pending['all'] if pending else False)
        if len(self.files) > 1:
            ttk.Checkbutton(frm, text='应用到列表中的所有文件（共 %d 个）' % len(self.files),
                            variable=all_var).pack(anchor='w')

        btns = ttk.Frame(frm)
        btns.pack(fill='x', pady=(12, 0))

        def on_ok():
            ok, err = self._apply_edit(key, var.get(), all_var.get())
            if not ok:
                messagebox.showerror('无法保存', err, parent=dlg)
                return
            dlg.destroy()

        ttk.Button(btns, text='保存', command=on_ok).pack(side='right')
        ttk.Button(btns, text='取消', command=dlg.destroy).pack(side='right', padx=(0, 8))
        dlg.bind('<Return>', lambda _e: on_ok())
        dlg.bind('<Escape>', lambda _e: dlg.destroy())
        dlg.grab_set()
        self.wait_window(dlg)

    # ---- 清理 -------------------------------------------------------------
    def on_clean(self):
        from tkinter import messagebox
        if not self.files:
            messagebox.showinfo('提示', '请先添加文件。')
            return
        checked = [k for k, v in self.fields.items() if v['checked']]
        if not checked and not self.edits:
            messagebox.showinfo('提示', '请先在右侧勾选要清除的字段，或双击属性值进行修改。')
            return
        mode = self.save_var.get()
        if mode == 'overwrite':
            if not messagebox.askyesno(
                    '确认覆盖',
                    '覆盖模式会直接修改原文件（建议先备份）。\n确定继续吗？'):
                return

        self.start_btn.config(state='disabled')
        ok = fail = 0
        try:
            for f in list(self.files):
                name = os.path.basename(f['path'])
                _d, _s, skipped = collect_ops(f, checked, self.edits)
                for k, why in skipped:
                    self.log('· %s：%s 跳过（%s）' % (name, field_display_name(k), why))
                try:
                    dest, removed, edited = clean_file(f, checked, self.edits, mode)
                except PermissionError:
                    fail += 1
                    self.log('✗ %s：文件被占用，请先关闭正在编辑它的程序' % name, 'err')
                    continue
                except Exception as e:
                    fail += 1
                    self.log('✗ %s：%s' % (name, e), 'err')
                    continue
                if removed or edited:
                    parts = []
                    if removed:
                        parts.append('清除 %d 项：%s' % (
                            len(removed), '、'.join(field_display_name(k) for k in removed)))
                    if edited:
                        parts.append('修改 %d 项：%s' % (
                            len(edited), '、'.join('%s→%s' % (field_display_name(k), _truncate(v, 30))
                                                  for k, v in edited)))
                    detail = '；'.join(parts)
                    if mode == 'overwrite':
                        self.log('✓ %s → 已覆盖原文件（%s）' % (name, detail), 'ok')
                    else:
                        self.log('✓ %s → 已另存为 %s（%s）'
                                 % (name, os.path.basename(dest), detail), 'ok')
                    ok += 1
                else:
                    self.log('· %s：勾选/修改的字段在该文件中不适用，未改动' % name)
                    ok += 1
                self.update_idletasks()
        finally:
            self.start_btn.config(state='normal')
        self.edits.clear()
        self.log('—— 完成：成功 %d 个，失败 %d 个 ——' % (ok, fail), 'ok')
        # 重新扫描，刷新预览
        self._rescan_all()
        self.update_status()

    def _rescan_all(self):
        for f in self.files:
            try:
                fresh = scan_file(f['path'])
                f['props'], f['suspects'] = fresh['props'], fresh['suspects']
            except Exception:
                pass
        self._rebuild_fields()
        self._refresh_tree()


def _truncate(s, n):
    s = str(s)
    return s if len(s) <= n else s[:n - 1] + '…'


def main():
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    app = CleanerApp()
    app.mainloop()


if __name__ == '__main__':
    main()
