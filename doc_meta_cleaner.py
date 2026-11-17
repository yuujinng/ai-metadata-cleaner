# -*- coding: utf-8 -*-
"""
文档AI属性清理工具（AI Metadata Cleaner）
================================================
删除 Word / Excel / PPT / PDF 文档属性（元数据）中泄露生成来源的信息，
重点针对 AI 生成留下的痕迹（作者名为 ChatGPT/Claude、生成程序为脚本库等）。

用法：双击「文档AI属性清理工具.bat」启动，或 `python doc_meta_cleaner.py`。

流程：添加文件 -> 预览属性并勾选要清除的字段（疑似 AI 痕迹自动勾选）
      -> 选择「另存副本 / 覆盖原文件」-> 开始清理。
      也可以双击属性值直接修改字段内容（如把作者改成自己的名字）。

原理：
  * .docx/.xlsx/.pptx（含 .docm/.xlsm/.pptm）是 ZIP+XML 包，
    直接改写其中的 docProps/core.xml、app.xml、custom.xml，其余部件原样保留；
    文件缺少属性部件时会按规范新建并登记，保证 Word/Excel/PPT 正常打开。
  * .pdf 通过 pypdf 克隆后改写 Info 字典字段与 /Metadata(XMP)；修改/删除
    Info 字段时会同步改写 XMP 包，避免旧值残留；挂在页面对象上的内嵌
    元数据同样支持检测与清除。
"""

import json
import os
import queue
import re
import sys
import threading
import time
import traceback
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone

import tkinter as tk
from tkinter import ttk

try:
    from version import __version__
except ImportError:  # 单独拷贝本文件时兜底，正常随仓库分发不会走到
    __version__ = 'dev'

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
PAGEMETA_KEY = 'pdf:pagemeta'
PAGEMETA_NAME = '页面级元数据（挂在页面对象上的内嵌 XML）'
THUMB_KEY = 'ooxml:thumbnail'
THUMB_NAME = '文档缩略图（docProps/thumbnail.jpeg，Word/Excel 默认不生成，多为生成库模板残留）'
THUMB_PREFIX = 'docprops/thumbnail.'   # 小写常量：统一配合 n.lower() 做前缀比较
THUMB_REL_SUFFIX = '/metadata/thumbnail'

# app.xml 里 Office 的统计字段（Office 保存时自动计算；脚本生成的文档
# 往往保留模板默认值——正文有字而统计为 0，属性面板里一眼可见）
APP_STAT_KEYS = ('app:Pages', 'app:Words', 'app:Characters',
                 'app:Lines', 'app:Paragraphs', 'app:CharactersWithSpaces')
APP_STAT_NAMES = {
    'Pages': '页数', 'Words': '字数', 'Characters': '字符数',
    'Lines': '行数', 'Paragraphs': '段落数', 'CharactersWithSpaces': '字符数（计空格）',
    'Slides': '幻灯片数', 'Notes': '备注页数', 'HiddenSlides': '隐藏幻灯片数',
    'MMClips': '多媒体剪辑数',
}

# WPS 在保存/云端同步时写入的自定义属性：ICV 是云文档指纹（GUID）、
# KSOProductBuildVer 是 WPS 版本号，值本身不含任何 AI 关键词，需按属性名识别
WPS_CUSTOM_PROPS = {'ICV', 'KSOProductBuildVer'}

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
# PowerPoint 默认随文件保存缩略图；Word/Excel 默认不保存，出现即多为生成库模板残留
PPTX_EXTS = {'.pptx', '.pptm', '.ppsx', '.ppsm'}
OOXML_NO_PPT_EXTS = OOXML_EXTS - PPTX_EXTS


def is_lock_file(path):
    """Word/WPS 打开文档时会在同目录生成 ~$ 开头的锁文件，不是文档本身，
    读取必然失败，目录扫描时应跳过。"""
    return os.path.basename(path).startswith('~$')

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
    + _L + r"qwen" + _R + r"|"
    + _L + r"kimi" + _R + r"|moonshot|doubao|"
    + _L + r"ernie" + _R + r"|"
    + _L + r"llama" + _R + r"|"
    + _L + r"grok" + _R + r"|"
    + _L + r"xai" + _R + r"|"
    + _L + r"x\.ai" + _R + r"|mistral|zhipu|z\.ai|bigmodel|"
    + _L + r"glm" + _R + r"|sparkdesk|"
    r"python-docx|python-pptx|openpyxl|xlsxwriter|reportlab|"
    + _L + r"fpdf" + _R + r"|weasyprint|pandoc|apache\s+poi|skia/?pdf|"
    r"libre\s*office|open\s*office|wkhtmltopdf|matplotlib|playwright|puppeteer|"
    + _L + r"(?:latex|luatex|pdftex|xetex|miktex|typst)" + _R + r"|"
    r"notion\.so|gamma\.app|beautiful\.ai|tome\.app|wps\s*ai" + _R + r"|"
    + _L + r"aigc" + _R + r"|ai[-\s]?generated|written\s+by\s+(?:ai|chatgpt|gpt)"
    r"|由\s*ai\s*生成|ai\s*生成",
    re.IGNORECASE,
)

# 弱信号：中文产品名与普通词同形（星火计划、心有灵犀、文心雕龙、豆包……），
# 在标题/关键词里直接匹配会大量误报，只在身份字段参与判断
_WEAK_AI_PATTERN = re.compile(r"通义|千问|豆包|文心|混元|星火|灵犀|智谱|讯飞")


# 身份字段：值里出现生成器名称即视为痕迹
AI_IDENTITY_FIELDS = {
    'creator', 'lastModifiedBy', 'Application', 'AppVersion', 'Company', 'Manager',
    'pdf:Author', 'pdf:Creator', 'pdf:Producer',
}


def is_identity_field(key):
    return key in AI_IDENTITY_FIELDS or key.startswith('custom:')


def is_auto_checkable(key):
    """自动勾选白名单：身份字段 + 时间/统计类规则字段 + PDF 内嵌元数据包 +
    OOXML 缩略图 + 统计字段（只在矛盾规则命中时才会进 suspects）。

    标题、关键词等正文类字段即便命中关键词也只标红提示，避免误删。
    """
    return (is_identity_field(key)
            or key in ('created', 'modified', 'revision', 'TotalTime',
                       'pdf:CreationDate', 'pdf:ModDate',
                       XMP_KEY, PAGEMETA_KEY, THUMB_KEY)
            or key in APP_STAT_KEYS)


def auto_clean_keys(entry):
    """与命令行共用的默认清理范围；内容字段只提示。"""
    return {key for key in entry['suspects'] if is_auto_checkable(key)}


def looks_like_ai(value, strict=False):
    """值里是否带有 AI 工具或脚本生成库的痕迹。

    strict=True 用于作者/应用程序等身份字段：中文产品名等弱信号也算数。
    """
    if not value:
        return False
    if _AI_VALUE_PATTERN.search(value):
        return True
    return strict and bool(_WEAK_AI_PATTERN.search(value))


# ---------------------------------------------------------------------------
# 编辑值校验 / 日期格式转换
# ---------------------------------------------------------------------------

OOXML_DATE_FIELDS = ('created', 'modified', 'lastPrinted')
PDF_DATE_FIELDS = ('pdf:CreationDate', 'pdf:ModDate')
INT_FIELDS = ('revision', 'TotalTime')
_DT_COMPACT = re.compile(r'^(?:D:)?(\d{4})(\d{1,2})(\d{1,2})(\d{1,2})?:?(\d{1,2})?:?(\d{1,2})?:?[Zz]?$')


def _parse_dt(s):
    """把用户输入解析为 UTC datetime；兼容 ISO、带空格、纯数字日期、PDF D: 前缀。

    不带时区的输入按本地墙钟时间理解，转换成 UTC 存储，保证 Word/PDF 里
    显示的时间与用户输入一致；PDF 的 D: 前缀无时区后缀时按 PDF 规范
    （缺省偏移即 UT）直接视为 UTC。
    """
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
        if s[:2].upper() == 'D:':
            return dt.replace(tzinfo=timezone.utc)
        return dt.replace(tzinfo=timezone.utc) if ends_utc else dt.astimezone(timezone.utc)
    t = s[:-1] if ends_utc and s.endswith(('Z', 'z')) else s
    try:
        dt = datetime.fromisoformat(t.replace(' ', 'T'))
    except ValueError:
        return None
    if dt.tzinfo:
        return dt.astimezone(timezone.utc)
    return dt.replace(tzinfo=timezone.utc) if ends_utc else dt.astimezone(timezone.utc)


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


def _flag_stale_dates(props, suspects, path, date_keys):
    """文档内部时间与文件系统时间的矛盾。

    * created/modified 比文件的修改时间早一年以上：生成库模板的冻结日期
      残留（python-docx 模板把日期冻在 2013 年）。两个时间都偏旧（或
      modified 缺失）才判——只有 created 偏旧是"多年前的旧稿仍在维护"
      的正常形态，不标红。
    * 时间晚于当前时间：异常的未来时间。
    """
    try:
        mtime = datetime.fromtimestamp(os.path.getmtime(path), tz=timezone.utc)
    except (OSError, OverflowError, ValueError):
        return
    parsed = {}
    for k in date_keys:
        dt = _parse_dt(props.get(k) or '')
        if dt is not None:
            parsed[k] = dt
    if not parsed:
        return
    created_key, modified_key = date_keys[0], date_keys[1]
    stale = {k for k, dt in parsed.items() if dt < mtime - timedelta(days=365)}
    if created_key in stale and (modified_key not in parsed or modified_key in stale):
        suspects.update(stale)
    suspects.update(k for k, dt in parsed.items()
                    if dt > datetime.now(timezone.utc) + timedelta(days=1))


def read_ooxml_props(path):
    """返回 (props: {key: 值}, suspects: {key, ...})，key 见 CORE/APP/custom 规则。"""
    props, suspects = {}, set()
    body_text = ''
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
        # 缩略图：整体作为一个"字段"暴露，供勾选清除（部件 + 关系 + 登记）
        thumb_parts = sorted(n for n in names if n.lower().startswith(THUMB_PREFIX))
        if thumb_parts:
            props[THUMB_KEY] = '存在（%s）' % thumb_parts[0]
        # 主文档故事的可见文本（含表格），用于"统计字段为 0"矛盾判断
        if 'word/document.xml' in names:
            try:
                root = ET.fromstring(z.read('word/document.xml'))
                body_text = ''.join(e.text for e in root.iter()
                                    if _local(e.tag) == 't' and e.text)
            except ET.ParseError:
                body_text = ''

    # 规则化判断疑似痕迹
    for key, val in props.items():
        if looks_like_ai(val, strict=is_identity_field(key)):
            suspects.add(key)
    # 创建时间 == 修改时间（精确到秒）：典型的机器一次性生成
    if props.get('created') and props.get('created') == props.get('modified'):
        suspects.update(('created', 'modified'))
    # 总编辑时间为 0 分钟
    if props.get('TotalTime') in ('0', '0分钟'):
        suspects.add('TotalTime')
    # 正文有字而统计字段为 0：Office 保存时会自动重算，脚本生成的文档
    # 保留模板默认值（python-docx 模板 Words=0），属性面板里一眼可见
    if body_text:
        for key in APP_STAT_KEYS:
            if props.get(key) == '0':
                suspects.add(key)
    # 模板残留的 "Microsoft Macintosh Word"：在 Windows 上提交 Mac 版 Word 的痕迹
    if 'macintosh' in (props.get('Application') or '').lower():
        suspects.add('Application')
    # AppVersion 不是 Office 的 NN.NNNN 样式（openpyxl 的 3.1 等脚本库版本号）
    if props.get('AppVersion') and not re.fullmatch(r'\d{2}\.\d{4}', props['AppVersion']):
        suspects.add('AppVersion')
    # 修订号 1 且编辑时间 0：一次性写出、从未再次保存
    if props.get('revision') == '1' and props.get('TotalTime') == '0':
        suspects.add('revision')
    # WPS 云文档指纹与版本键：值是 GUID/版本号，关键词匹配打不中，按属性名识别
    suspects.update('custom:' + cname for cname in WPS_CUSTOM_PROPS
                    if props.get('custom:' + cname))
    # Word/Excel 文档不该有缩略图（PowerPoint 默认有，不判）
    if THUMB_KEY in props and os.path.splitext(path)[1].lower() in OOXML_NO_PPT_EXTS:
        suspects.add(THUMB_KEY)
    # 模板冻结日期 / 未来时间（created 与 modified 一起看，避免误伤旧稿）
    _flag_stale_dates(props, suspects, path, ('created', 'modified'))
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


# 缺失属性部件的新建模板与登记信息（保证 Word/Excel/PPT 正常打开）
CORE_PART_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<cp:coreProperties xmlns:cp="%s" xmlns:dc="%s" xmlns:dcterms="%s" xmlns:xsi="%s"/>'
) % (CP_NS, DC_NS, DCTERMS_NS, XSI_NS)
APP_PART_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Properties xmlns="%s" xmlns:vt="%s"/>'
) % (APP_NS, VT_NS)
CORE_CONTENT_TYPE = 'application/vnd.openxmlformats-package.core-properties+xml'
APP_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
CORE_REL_TYPE = ('http://schemas.openxmlformats.org/package/2006/relationships'
                 '/metadata/core-properties')
APP_REL_TYPE = ('http://schemas.openxmlformats.org/officeDocument/2006'
                '/relationships/extended-properties')
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'


def _build_props_part(template, part, ops):
    """按模板新建属性部件（core.xml / app.xml）并写入 ops 里的字段。"""
    root = ET.fromstring(template)
    for key, value in ops.items():
        _set_or_create(root, key, value, part, False)
    return _serialize(root, template.encode('utf-8'))


def _register_part(ct_bytes, rels_bytes, part_name, content_type, rel_type):
    """把新建部件登记进 [Content_Types].xml 与 _rels/.rels。返回 (ct, rels)。"""
    ct_root = ET.fromstring(ct_bytes)
    if not any(c.get('PartName') == '/' + part_name for c in ct_root):
        ET.SubElement(ct_root, '{%s}Override' % CT_NS,
                      {'PartName': '/' + part_name, 'ContentType': content_type})
        ct_bytes = _serialize(ct_root, ct_bytes)
    rel_root = ET.fromstring(rels_bytes)
    if not any((c.get('Target') or '').lstrip('/') == part_name for c in rel_root):
        used = {c.get('Id') for c in rel_root}
        n = 1
        while 'rId%d' % n in used:
            n += 1
        ET.SubElement(rel_root, '{%s}Relationship' % RELS_NS,
                      {'Id': 'rId%d' % n, 'Type': rel_type, 'Target': part_name})
        rels_bytes = _serialize(rel_root, rels_bytes)
    return ct_bytes, rels_bytes


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


def _remove_thumbnail_refs(ct_bytes, rels_bytes):
    """从 [Content_Types].xml 与 _rels/.rels 里摘除缩略图的登记项。

    返回 (新 ct, 新 rels, 是否确实存在缩略图登记)。
    """
    found = False
    if ct_bytes is not None:
        root = ET.fromstring(ct_bytes)
        drop = [c for c in root
                if (c.get('PartName') or '').lower().startswith('/' + THUMB_PREFIX)]
        if drop:
            found = True
            for c in drop:
                root.remove(c)
            ct_bytes = _serialize(root, ct_bytes)
    if rels_bytes is not None:
        root = ET.fromstring(rels_bytes)
        drop = [c for c in root
                if (c.get('Type') or '').lower().endswith(THUMB_REL_SUFFIX)]
        if drop:
            found = True
            for c in drop:
                root.remove(c)
            rels_bytes = _serialize(root, rels_bytes)
    return ct_bytes, rels_bytes, found


def clean_ooxml(src, dst, delete_keys, set_ops):
    """改写 OOXML 包：删除 delete_keys 字段并把 set_ops 字段设为新值。

    其余部件逐字节保留。目标属性部件缺失但有待写入字段时，按规范新建
    部件并登记到 [Content_Types].xml 与 _rels/.rels。
    返回 (删除的 key 列表, 修改的 key 列表)。
    """
    wanted = set(delete_keys)
    core_ops = {k: v for k, v in set_ops.items() if k in CORE_FIELDS}
    app_ops = {k: v for k, v in set_ops.items() if k in APP_FIELDS}
    custom_ops = {k: v for k, v in set_ops.items() if k.startswith('custom:')}
    removed, edited = [], []
    drop_thumb = THUMB_KEY in wanted
    thumb_present = False

    # 包内 entry 时间戳：Word 随编辑过程渐进保存，时间戳是分布开的；
    # python-docx/openpyxl 等脚本一次性写出，所有 entry 时间完全相同，
    # 是一眼可辨的指纹。呈脚本特征（全部相同，或出现 2000 年前/未来年份）
    # 时统一归一化为文件的修改时间，原本就正常的包保持原样。
    try:
        pkg_time = datetime.fromtimestamp(os.path.getmtime(src)).timetuple()[:6]
        if pkg_time[0] < 1980:      # DOS 时间戳下限，异常时钟兜底
            pkg_time = None
    except (OSError, OverflowError, ValueError):
        pkg_time = None
    with zipfile.ZipFile(src) as zin:
        names = {n.replace('\\', '/') for n in zin.namelist()}
        infolist = zin.infolist()
        times = {i.date_time for i in infolist}
        normalize_times = (
            pkg_time is not None and bool(infolist)
            and (len(times) == 1
                 or any(t[0] < 2000 or t[0] > datetime.now().year + 1 for t in times)))

        # 部件缺失但有待写入字段时新建（包结构不完整则放弃，如实报告）
        new_parts = []   # (部件名, 数据, ContentType, 关系Type, 涉及的 key)
        if (core_ops or app_ops) and '[Content_Types].xml' in names \
                and '_rels/.rels' in names:
            if core_ops and 'docProps/core.xml' not in names:
                data = _build_props_part(CORE_PART_TEMPLATE, 'core', core_ops)
                new_parts.append(('docProps/core.xml', data, CORE_CONTENT_TYPE,
                                  CORE_REL_TYPE, list(core_ops)))
            if app_ops and 'docProps/app.xml' not in names:
                data = _build_props_part(APP_PART_TEMPLATE, 'app', app_ops)
                new_parts.append(('docProps/app.xml', data, APP_CONTENT_TYPE,
                                  APP_REL_TYPE, list(app_ops)))
        ct_bytes = zin.read('[Content_Types].xml') if '[Content_Types].xml' in names else None
        rels_bytes = zin.read('_rels/.rels') if '_rels/.rels' in names else None
        for part_name, _data, ctype, rtype, keys in new_parts:
            ct_bytes, rels_bytes = _register_part(ct_bytes, rels_bytes,
                                                  part_name, ctype, rtype)
            edited.extend(keys)
        thumb_present = drop_thumb and any(
            n.lower().startswith(THUMB_PREFIX) for n in names)
        if thumb_present:
            # 缩略图部件连同关系/登记一起摘除，避免悬空引用导致打开报"损坏"
            ct_bytes, rels_bytes, _found = _remove_thumbnail_refs(ct_bytes, rels_bytes)

        with zipfile.ZipFile(dst, 'w') as zout:
            zout.comment = zin.comment
            for info in infolist:
                fname = info.filename.replace('\\', '/')
                if thumb_present and fname.lower().startswith(THUMB_PREFIX):
                    continue    # 缩略图部件整体剔除
                data = zin.read(info.filename)
                rm, ed = [], []
                if fname == '[Content_Types].xml' and ct_bytes is not None:
                    data = ct_bytes
                elif fname == '_rels/.rels' and rels_bytes is not None:
                    data = rels_bytes
                elif fname == 'docProps/core.xml':
                    data, rm, ed = _rewrite_part(data, wanted, core_ops, 'core')
                elif fname == 'docProps/app.xml':
                    data, rm, ed = _rewrite_part(data, wanted, app_ops, 'app')
                elif fname == 'docProps/custom.xml':
                    data, rm, ed = _rewrite_part(data, wanted, custom_ops, 'custom')
                removed.extend(rm)
                edited.extend(ed)
                zi = zipfile.ZipInfo(fname, date_time=(pkg_time if normalize_times
                                                       else info.date_time))
                zi.compress_type = info.compress_type
                zi.external_attr = info.external_attr
                zi.comment = info.comment
                zout.writestr(zi, data)
            for part_name, data, _ct, _rt, _keys in new_parts:
                zi = zipfile.ZipInfo(part_name, date_time=pkg_time or (1980, 1, 1, 0, 0, 0))
                zi.compress_type = zipfile.ZIP_DEFLATED
                zout.writestr(zi, data)
    if thumb_present:
        removed.append(THUMB_KEY)
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

    # 页面级 /Metadata（挂在页面对象上，目录级读不到）
    try:
        page_meta_count = 0
        for pg in reader.pages:
            ref = pg.get('/Metadata')
            if ref is None:
                continue
            page_meta_count += 1
            try:
                raw = ref.get_object().get_data()
            except Exception:
                raw = b''
            if looks_like_ai(raw.decode('utf-8', 'replace')):
                suspects.add(PAGEMETA_KEY)
        if page_meta_count:
            props[PAGEMETA_KEY] = '存在（%d 页含内嵌元数据）' % page_meta_count
    except Exception:
        pass

    for key, val in props.items():
        if looks_like_ai(val, strict=is_identity_field(key)):
            suspects.add(key)
    # 脚本库缺省值指纹：reportlab 等生成库不做设置时留下的占位值
    for key, val in (('pdf:Author', 'anonymous'), ('pdf:Creator', 'anonymous'),
                     ('pdf:Title', 'untitled')):
        if (props.get(key) or '').strip().lower() == val:
            suspects.add(key)
    # 模板冻结日期 / 未来时间（与 OOXML 同一套规则）
    _flag_stale_dates(props, suspects, path, ('pdf:CreationDate', 'pdf:ModDate'))
    try:
        xmp = reader.xmp_metadata
        if xmp is not None and looks_like_ai(xmp.stream.get_data().decode('utf-8', 'replace')):
            suspects.add(XMP_KEY)
    except Exception:
        pass
    return props, suspects


# XMP 与 Info 字典的对应关系：(命名空间, 属性名, 容器类型 None/'Alt'/'Seq')
RDF_NS = 'http://www.w3.org/1999/02/22-rdf-syntax-ns#'
XMP_META_NS = 'adobe:ns:meta/'
XMP_NS = 'http://ns.adobe.com/xap/1.0/'
PDF_NS = 'http://ns.adobe.com/pdf/1.3/'
INFO_XMP_MAP = {
    'Title':        (DC_NS, 'title', 'Alt'),
    'Author':       (DC_NS, 'creator', 'Seq'),
    'Subject':      (DC_NS, 'description', 'Alt'),
    'Keywords':     (PDF_NS, 'Keywords', None),
    'Creator':      (XMP_NS, 'CreatorTool', None),
    'Producer':     (PDF_NS, 'Producer', None),
    'CreationDate': (XMP_NS, 'CreateDate', None),
    'ModDate':      (XMP_NS, 'ModifyDate', None),
    'Trapped':      (PDF_NS, 'Trapped', None),
}

_XPACKET_BEGIN = '<?xpacket begin="" id="W5M0MpCehiHzreSzNTczkc9d"?>'
_XPACKET_END = '<?xpacket end="w"?>'


def _xmp_value(info_key, value):
    """Info 格式的值转成 XMP 格式（日期 D:... -> ISO 8601）。"""
    if info_key in ('CreationDate', 'ModDate'):
        dt = _parse_dt(value)
        if dt is not None:
            return dt.strftime('%Y-%m-%dT%H:%M:%SZ')
    return str(value)


def _build_xmp_prop(ns, local, kind, value):
    """构造 XMP 属性元素；Alt/Seq 容器内包一层 rdf:li。"""
    elem = ET.Element('{%s}%s' % (ns, local))
    if kind in ('Alt', 'Seq'):
        cont = ET.SubElement(elem, '{%s}%s' % (RDF_NS, kind))
        li = ET.SubElement(cont, '{%s}li' % RDF_NS)
        li.text = value
    else:
        elem.text = value
    return elem


def _serialize_xmp(root, original_bytes):
    """序列化 XMP：保留原文件命名空间前缀，统一 UTF-8，补全 xpacket 包裹
    （ElementTree 解析时会丢弃 PI，需重新补上）。"""
    for pfx, uri in (('x', XMP_META_NS), ('rdf', RDF_NS), ('dc', DC_NS),
                     ('xmp', XMP_NS), ('pdf', PDF_NS)):
        ET.register_namespace(pfx, uri)
    for m in _XMLNS_RE.finditer(original_bytes):
        ET.register_namespace(m.group(1).decode('ascii'), m.group(2).decode('ascii'))
    for m in _DEFAULT_NS_RE.finditer(original_bytes):
        ET.register_namespace('', m.group(1).decode('ascii'))
    body = ET.tostring(root, encoding='unicode')
    decl = '<?xml version="1.0" encoding="UTF-8"?>' \
        if original_bytes.startswith(b'<?xml') else ''
    return (decl + _XPACKET_BEGIN + body + _XPACKET_END).encode('utf-8')


def _rewrite_xmp(xml_bytes, del_props, set_props, *, strict=False):
    """把 Info 字段的删除/修改同步进 XMP 包。

    del_props: Info 字段名列表（如 ['Producer']）；set_props: {Info 字段名: 值}。
    同时支持元素形式（<pdf:Producer>x</pdf:Producer>）与属性形式
    （<rdf:Description pdf:Producer="x">）。返回 (新字节, 删除的字段,
    写入的字段)；XMP 无法解析时返回 None，strict 模式下报告失败。
    """
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as exc:
        if strict:
            raise ValueError('无法解析 PDF 的 XMP，无法同步属性修改') from exc
        return None
    removed_keys, edited_keys = [], []

    for name in del_props:
        mapping = INFO_XMP_MAP.get(name)
        if mapping is None:
            continue
        ns, local, _kind = mapping
        tag = attr = '{%s}%s' % (ns, local)
        hit = False
        for parent in list(root.iter()):
            for child in list(parent):
                if child.tag == tag:
                    parent.remove(child)
                    hit = True
        for elem in root.iter():
            if attr in elem.attrib:
                del elem.attrib[attr]
                hit = True
        if hit:
            removed_keys.append(name)

    for name, raw_value in set_props.items():
        mapping = INFO_XMP_MAP.get(name)
        if mapping is None:
            continue
        ns, local, kind = mapping
        tag = '{%s}%s' % (ns, local)
        value = _xmp_value(name, raw_value)
        nodes = list(root.iter())
        elements = [e for e in nodes if e.tag == tag]
        attributes = [e for e in nodes if tag in e.attrib]
        for e in attributes:
            e.set(tag, value)
        for elem in elements:
            subs = list(elem)
            for s in subs:
                elem.remove(s)
            if subs and _local(subs[0].tag) in ('Alt', 'Seq', 'Bag'):
                cont = ET.SubElement(elem, subs[0].tag)
                li = ET.SubElement(cont, '{%s}li' % RDF_NS)
                li.text = value
            else:
                elem.text = value
        if not elements and not attributes:
            desc = next((e for e in nodes if e.tag == '{%s}Description' % RDF_NS), None)
            if desc is None:
                rdf = next((e for e in nodes if e.tag == '{%s}RDF' % RDF_NS), None)
                if rdf is None:
                    if strict:
                        raise ValueError('PDF 的 XMP 缺少 RDF 结构，无法同步属性修改')
                    continue
                desc = ET.SubElement(rdf, '{%s}Description' % RDF_NS)
                desc.set('{%s}about' % RDF_NS, '')
            desc.append(_build_xmp_prop(ns, local, kind, value))
        edited_keys.append(name)

    if not removed_keys and not edited_keys:
        return None
    return _serialize_xmp(root, xml_bytes), removed_keys, edited_keys


def _replace_xmp_stream(writer, data):
    """用未压缩的新流替换目录级 /Metadata（规避原流的过滤器/长度问题）。"""
    from pypdf.generic import NameObject, NumberObject, StreamObject
    stream = StreamObject()
    stream[NameObject('/Type')] = NameObject('/Metadata')
    stream[NameObject('/Subtype')] = NameObject('/XML')
    stream.set_data(data)
    stream[NameObject('/Length')] = NumberObject(len(data))
    writer.root_object[NameObject('/Metadata')] = writer._add_object(stream)


def clean_pdf(src, dst, delete_keys, set_ops):
    """删除 delete_keys 并把 set_ops（key -> value）写入 PDF Info 字典。

    Info 的修改/删除会同步改写 XMP 包（存在时），避免旧值残留；
    页面级 /Metadata 通过勾选 pdf:pagemeta 整体删除。
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
    info_dels = []
    for k in list(info.keys()):
        key = str(k).lstrip('/')
        if ('pdf:' + key) in wanted:
            del info[NameObject(str(k))]
            removed.append('pdf:' + key)
            info_dels.append(key)
    if XMP_KEY in wanted and '/Metadata' in writer.root_object:
        del writer.root_object[NameObject('/Metadata')]
        removed.append(XMP_KEY)
    if PAGEMETA_KEY in wanted:
        page_hits = 0
        for page in writer.pages:
            if '/Metadata' in page:
                del page[NameObject('/Metadata')]
                page_hits += 1
        if page_hits:
            removed.append(PAGEMETA_KEY)

    updates = {}
    for key, value in set_ops.items():
        if not key.startswith('pdf:') or key in (XMP_KEY, PAGEMETA_KEY):
            continue
        updates[NameObject('/' + key[4:])] = value
        edited.append(key)
    if updates:
        writer.add_metadata(updates)

    # Info 的删除/修改同步进 XMP，防止旧值残留在内嵌 XML 里
    if '/Metadata' in writer.root_object:
        info_sets = {k[4:]: v for k, v in set_ops.items()
                     if k.startswith('pdf:') and k not in (XMP_KEY, PAGEMETA_KEY)}
        if any(k in INFO_XMP_MAP for k in (*info_dels, *info_sets)):
            xmp_stream = writer.root_object[NameObject('/Metadata')].get_object()
            try:
                raw = xmp_stream.get_data()
            except Exception as exc:
                raise ValueError('无法读取 PDF 的 XMP，无法同步属性修改') from exc
            result = _rewrite_xmp(raw, info_dels, info_sets, strict=True)
            if result is not None and (result[1] or result[2]):
                _replace_xmp_stream(writer, result[0])

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
    if key == PAGEMETA_KEY:
        return PAGEMETA_NAME
    if key == XMP_KEY:
        return XMP_NAME
    if key == THUMB_KEY:
        return THUMB_NAME
    if key.startswith('pdf:'):
        return 'PDF ' + PDF_FIELDS.get(key[4:], key[4:])
    if key.startswith('custom:'):
        return '自定义属性：' + key[7:]
    if key.startswith('core:') or key.startswith('app:'):
        local = key.split(':', 1)[1]
        if key.startswith('app:') and local in APP_STAT_NAMES:
            return APP_STAT_NAMES[local] + '（应用统计）'
        return local + '（未识别字段）'
    return key


def field_sort_key(key):
    if key in FIELD_ORDER:
        return FIELD_ORDER.index(key)
    if key.startswith('pdf:'):
        local = key[4:]
        return 500 + (PDF_ORDER.index(local) if local in PDF_ORDER else 50)
    if key.startswith('custom:'):
        return 700
    if key == THUMB_KEY:
        return 880
    if key == PAGEMETA_KEY:
        return 890
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
    applicable_edits = {k: e for k, e in edits.items()
                        if e['all'] or e['path'] == entry['path']}
    delete = [k for k in checked_keys if k in entry['props'] and k not in applicable_edits]
    sets, skipped = {}, []
    for key, e in applicable_edits.items():
        if key in (XMP_KEY, PAGEMETA_KEY):
            skipped.append((key, 'XMP / 页面级元数据不支持直接编辑'))
            continue
        if key == THUMB_KEY:
            skipped.append((key, '缩略图只能勾选整体清除，不支持逐项编辑'))
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


# ---------------------------------------------------------------------------
# 文件系统时间保护
# ---------------------------------------------------------------------------

def _set_win_creation_time(path, ctime_ns):
    """Windows 下设置文件创建时间（os.stat 的 st_ctime 即创建时间）。
    最佳努力，失败返回 False。"""
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        kernel32.SetFileTime.argtypes = (ctypes.c_void_p,) * 4
        kernel32.CreateFileW.restype = ctypes.c_void_p
        # FILE_WRITE_ATTRIBUTES / 共享读写删 / OPEN_EXISTING
        handle = kernel32.CreateFileW(str(path), 0x0100, 0x7, None, 3, 0, None)
        if not handle or handle == ctypes.c_void_p(-1).value:
            return False
        try:
            # FILETIME：1601-01-01 起的 100 纳秒数
            ft = ctypes.c_ulonglong(int(ctime_ns) // 100 + 116444736000000000)
            return bool(kernel32.SetFileTime(handle, ctypes.byref(ft), None, None))
        finally:
            kernel32.CloseHandle(handle)
    except Exception:
        return False


def _restore_file_times(path, st):
    """覆盖模式：恢复访问/修改时间；Windows 连创建时间一并恢复，
    让文件在系统层面看起来未被改动。"""
    try:
        os.utime(path, ns=(st.st_atime_ns, st.st_mtime_ns))
    except OSError:
        pass
    if os.name == 'nt':
        _set_win_creation_time(path, st.st_ctime_ns)


def _copy_source_times(src, dst):
    """副本模式：副本沿用原文件的访问/修改时间（与资源管理器复制行为一致），
    避免副本以"刚刚生成"的时间戳暴露处理痕迹。"""
    try:
        st = os.stat(src)
        os.utime(dst, ns=(st.st_atime_ns, st.st_mtime_ns))
    except OSError:
        pass


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
        st = os.stat(path)
        tmp = path + '.meta_tmp'
        try:
            removed, edited = run(path, tmp)
            os.replace(tmp, path)
            _restore_file_times(path, st)
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
    _copy_source_times(path, dest)
    return dest, removed, _sorted_edited(edited, set_ops)


# ---------------------------------------------------------------------------
# 用户设置 / 崩溃日志（%APPDATA%/AIMetadataCleaner）
# ---------------------------------------------------------------------------

_GEOMETRY_RE = re.compile(r'^\d+x\d+(?:[+-]\d+){2}$')


def _settings_dir():
    """设置与崩溃日志的存放目录：优先 %APPDATA%，其余平台退 ~/.config。"""
    base = os.environ.get('APPDATA') or os.path.join(os.path.expanduser('~'), '.config')
    return os.path.join(base, 'AIMetadataCleaner')


def load_settings():
    """读取用户设置；文件缺失或损坏时返回空 dict（调用方用默认值兜底）。"""
    try:
        with open(os.path.join(_settings_dir(), 'settings.json'), encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_settings(data):
    """保存用户设置（临时文件 + 原子替换）；失败静默，设置丢失不影响主功能。"""
    try:
        d = _settings_dir()
        os.makedirs(d, exist_ok=True)
        tmp = os.path.join(d, 'settings.json.tmp')
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
        os.replace(tmp, os.path.join(d, 'settings.json'))
    except OSError:
        pass


def write_crash_log(exc, val, tb):
    """把未捕获异常追加写入 crash.log，返回日志路径；失败返回 None。"""
    try:
        d = _settings_dir()
        os.makedirs(d, exist_ok=True)
        path = os.path.join(d, 'crash.log')
        with open(path, 'a', encoding='utf-8', errors='replace') as f:
            f.write('\n[%s] %s: %s\n' % (datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                                         exc.__name__, val))
            f.write(''.join(traceback.format_exception(exc, val, tb)))
        return path
    except OSError:
        return None


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

_Base = TkinterDnD.Tk if TkinterDnD else tk.Tk


class CleanerApp(_Base):
    def __init__(self):
        super().__init__()
        self.title('文档AI属性清理工具 v' + __version__)
        self.geometry('1080x680')
        self.minsize(880, 560)

        settings = load_settings()

        self.files = []          # [ {path, kind, props, suspects, status} ]
        self.fields = {}         # key -> {name}
        self.batch_checks = {}   # 手动勾选应用到整批文件
        self.file_checks = {}    # (path, key) -> 留空编辑时只清除当前文件
        self.suspects_only = False
        self.edits = {}          # key -> {'value', 'all', 'path'}（待应用的修改）
        self.current_path = None

        # 后台任务状态：worker 线程只做 IO，控件操作全部经队列回到主线程
        self._queue = queue.Queue()
        self._busy = False
        self._job = None            # 'scan' | 'clean' | 'rescan'
        self._pending_paths = []    # 任务进行中拖入的路径，结束后自动接续
        self._stop_event = threading.Event()
        self._tip = None            # 文件路径悬浮提示
        self._last_crash_dialog = 0.0

        self._build_style()
        self._build_ui()
        self._apply_saved_prefs(settings)
        # 窗口几何要在 UI 构建完成后恢复：提前设置会被随后打包控件的
        # 布局重算覆盖（位置保留、尺寸回到内容请求值）
        geom = settings.get('geometry')
        if isinstance(geom, str) and _GEOMETRY_RE.match(geom):
            self.geometry(geom)
        self._start_pump()
        self.protocol('WM_DELETE_WINDOW', self._on_close)

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
        self.remove_btn = ttk.Button(bar, text='移除所选文件', command=self.remove_selected)
        self.remove_btn.pack(side='left', padx=(6, 0))
        self.clear_btn = ttk.Button(bar, text='清空列表', command=self.clear_files)
        self.clear_btn.pack(side='left', padx=(6, 0))
        self.auto_check_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            bar, text='添加时自动勾选身份类痕迹（作者/程序/时间）',
            variable=self.auto_check_var, command=self._on_auto_check).pack(side='right')

        # 主体：左文件列表 / 右属性表
        body = ttk.Panedwindow(self, orient='horizontal')
        body.pack(fill='both', expand=True, padx=8, pady=(0, 4))

        left = ttk.Frame(body)
        ttk.Label(left, text='文件列表（悬停查看完整路径）').pack(anchor='w')
        listframe = ttk.Frame(left)
        listframe.pack(fill='both', expand=True)
        self.file_tree = ttk.Treeview(listframe, columns=('name', 'status'),
                                      show='headings', selectmode='extended')
        self.file_tree.heading('name', text='文件')
        self.file_tree.column('name', width=170, anchor='w')
        self.file_tree.heading('status', text='状态')
        self.file_tree.column('status', width=80, anchor='center', stretch=False)
        fsb = ttk.Scrollbar(listframe, orient='vertical', command=self.file_tree.yview)
        self.file_tree.config(yscrollcommand=fsb.set)
        self.file_tree.pack(side='left', fill='both', expand=True)
        fsb.pack(side='right', fill='y')
        self.file_tree.tag_configure('ok', foreground='#1a7f37')
        self.file_tree.tag_configure('err', foreground='#c0392b')
        self.file_tree.bind('<<TreeviewSelect>>', self.on_file_selected)
        self.file_tree.bind('<Motion>', self._file_tooltip)
        self.file_tree.bind('<Leave>', self._hide_tooltip)
        body.add(left, weight=1)

        right = ttk.Frame(body)
        ttk.Label(right, text='属性预览（自动勾选按文件；手动勾选应用整批；双击当前值可编辑）').pack(anchor='w')
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
        self.progress = ttk.Progressbar(bottom, length=210, mode='determinate')
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
        checked = len(self._checked_keys(self._current_entry()))
        self.status.config(text='共 %d 个文件，当前文件勾选 %d 个字段，待修改 %d 项'
                           % (len(self.files), checked, len(self.edits)))

    # ---- 后台任务调度 ------------------------------------------------------
    # worker 线程只做文件 IO（展开目录 / scan_file / clean_file），绝不碰控件；
    # 结果经队列送回，主线程 _pump 每 80ms 取一次并分发。
    def _start_pump(self):
        self._pump_after = None
        self._pump()

    def _pump(self):
        try:
            try:
                while True:
                    msg = self._queue.get_nowait()
                    try:
                        self._dispatch(msg)
                    except tk.TclError:
                        return  # 窗口正在销毁，停止泵
                    except Exception:
                        # 单条消息处理出错不拖垮整个泵（否则后台 UI 更新会静默瘫痪）
                        self.report_callback_exception(*sys.exc_info())
            except queue.Empty:
                pass
        except tk.TclError:
            return  # 窗口正在销毁，停止泵
        self._pump_after = self.after(80, self._pump)

    def _dispatch(self, msg):
        kind = msg[0]
        if kind == 'log':
            self.log(msg[1], msg[2])
        elif kind == 'scanned':
            self._on_scanned(msg[1], msg[2])
        elif kind == 'progress':
            self._on_progress(msg[1], msg[2])
        elif kind == 'scan_done':
            self._on_scan_done()
        elif kind == 'cleaned':
            self._on_cleaned(msg[1], msg[2])
        elif kind == 'clean_done':
            self._on_clean_done(msg[1], msg[2], msg[3])
        elif kind == 'rescanned':
            self._on_rescanned(msg[1], msg[2])
        elif kind == 'rescan_done':
            self._on_rescan_done()

    def _set_busy(self, busy, job=None):
        self._busy = busy
        if job is not None:
            self._job = job
        state = 'disabled' if busy else 'normal'
        for btn in (self.start_btn, self.remove_btn, self.clear_btn):
            btn.config(state=state)
        if busy:
            self._stop_event.clear()
            self.progress.config(value=0)
            self.progress.pack(side='right', padx=(0, 10))
        else:
            self.progress.pack_forget()

    def _on_progress(self, done, total):
        self.progress.config(maximum=max(total, 1), value=done)
        verb = {'scan': '正在扫描', 'clean': '正在清理', 'rescan': '正在刷新'}[self._job]
        self.status.config(text='%s %d/%d…' % (verb, done, total))

    def wait_for_idle(self, timeout=60):
        """无 mainloop 的脚本（冒烟测试 / 截图）等待后台任务全部完成。"""
        deadline = time.time() + timeout
        while (self._busy or not self._queue.empty()) and time.time() < deadline:
            self.update()
            time.sleep(0.01)
        if self._busy or not self._queue.empty():
            raise RuntimeError('后台任务超时未完成')

    def _maybe_continue(self):
        """当前 job 结束后：有缓冲的拖入就接着扫描（返回 True），否则收尾。"""
        if self._pending_paths:
            pending, self._pending_paths = self._pending_paths, []
            self._launch_scan(pending)
            return True
        return False

    # ---- 拖拽放入 ---------------------------------------------------------
    def _enable_dragdrop(self):
        # Windows 上每个 tk 子窗口都是独立的 HWND，需要分别注册为拖放目标
        for w in (self, self.file_tree, self.tree, self.log_text):
            try:
                w.drop_target_register(DND_FILES)
                w.dnd_bind('<<Drop>>', self.on_drop)
            except Exception:
                pass

    def on_drop(self, event):
        """接收拖入的文件/文件夹；目录交给后台线程递归展开。"""
        try:
            raw = self.tk.splitlist(event.data)
        except Exception:
            return 'copy'
        if raw:
            self._add_paths(list(raw))
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
        if folder:
            self._add_paths([folder])

    def _add_paths(self, paths):
        """把文件/目录加入扫描队列；忙碌时排队缓冲，结束后自动接续。"""
        paths = [p for p in paths if p]
        if not paths:
            return
        if self._busy:
            self._pending_paths.extend(paths)
            self.log('（正在处理上一批，稍后自动加入新拖入的 %d 项）' % len(paths))
            return
        self._launch_scan(list(paths))

    def _launch_scan(self, paths):
        self._set_busy(True, 'scan')
        threading.Thread(target=self._scan_worker, args=(paths,), daemon=True).start()

    def _scan_worker(self, paths):
        """展开目录并逐个扫描，结果经队列送回主线程。"""
        q = self._queue
        files = []
        for p in paths:
            if os.path.isdir(p):
                for dirpath, _dirs, fns in os.walk(p):
                    files.extend(os.path.join(dirpath, fn) for fn in fns
                                 if os.path.splitext(fn)[1].lower() in SUPPORTED_EXTS
                                 and not is_lock_file(fn))
            elif os.path.isfile(p):
                files.append(p)
        files.sort()
        if not files:
            q.put(('log', '拖入/选择的内容里没有支持的文档'
                          '（docx/xlsx/pptx/pdf 及对应带宏格式）', 'err'))
            q.put(('scan_done',))
            return
        total = len(files)
        q.put(('progress', 0, total))
        for i, fp in enumerate(files, 1):
            if self._stop_event.is_set():
                break
            try:
                entry = scan_file(fp)
                entry['status'] = ('待处理', '')
                q.put(('scanned', entry, None))
            except Exception as e:
                q.put(('scanned', None, (os.path.basename(fp), str(e))))
            q.put(('progress', i, total))
        q.put(('scan_done',))

    def _on_scanned(self, entry, err):
        if entry is None:
            name, msg = err
            self.log('✗ 无法读取 %s：%s' % (name, msg), 'err')
            return
        if any(f['path'].lower() == entry['path'].lower() for f in self.files):
            return  # 已在列表里，跳过重复
        self.files.append(entry)
        self._append_file_row(entry)
        n_sus = len(entry['suspects'])
        if n_sus:
            self.log('＋ %s（发现 %d 项疑似 AI 痕迹）'
                     % (os.path.basename(entry['path']), n_sus), 'ok')
        else:
            self.log('＋ %s' % os.path.basename(entry['path']))

    def _on_scan_done(self):
        if self._maybe_continue():
            return
        self._set_busy(False)
        if self.files and self.current_path is None:
            self.current_path = self.files[0]['path']
            self.file_tree.selection_set(self.current_path)
        self._rebuild_fields()
        self._refresh_tree()
        self.update_status()

    def remove_selected(self):
        sel = self.file_tree.selection()
        if not sel:
            return
        removed = set(sel)
        self.files = [f for f in self.files if f['path'] not in removed]
        self.current_path = None
        self.detail.config(text='')
        self._rebuild_fields()
        self._refresh_list()
        if self.files:
            self.current_path = self.files[0]['path']
            self.file_tree.selection_set(self.current_path)
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

    # ---- 文件列表渲染 / 路径悬浮提示 ---------------------------------------
    def _refresh_list(self):
        self._hide_tooltip()
        self.file_tree.delete(*self.file_tree.get_children())
        for f in self.files:
            self._append_file_row(f)

    def _append_file_row(self, f):
        status, tag = f.get('status') or ('待处理', '')
        self.file_tree.insert('', 'end', iid=f['path'], tags=(tag,) if tag else (),
                              values=(os.path.basename(f['path']), status))

    def _update_file_row(self, path):
        f = next((f for f in self.files if f['path'] == path), None)
        if f is None:
            return
        status, tag = f.get('status') or ('待处理', '')
        self.file_tree.item(path, tags=(tag,) if tag else (),
                            values=(os.path.basename(path), status))

    def _show_tooltip(self, text, x, y):
        self._hide_tooltip()
        self._tip = tw = tk.Toplevel(self)
        tw.wm_overrideredirect(True)
        tw.wm_geometry('+%d+%d' % (x, y))
        tw._path = text
        tk.Label(tw, text=text, justify='left', background='#ffffe0',
                 relief='solid', borderwidth=1,
                 font=('Microsoft YaHei UI', 9)).pack(padx=1, pady=1)

    def _hide_tooltip(self, _event=None):
        if self._tip is not None:
            try:
                self._tip.destroy()
            except tk.TclError:
                pass
            self._tip = None

    def _file_tooltip(self, event):
        """文件名列只显示 basename，悬停显示完整路径，区分同名文件。"""
        row = self.file_tree.identify_row(event.y)
        if not row:
            self._hide_tooltip()
            return
        if self._tip is not None and getattr(self._tip, '_path', None) == row:
            return
        self._show_tooltip(row, event.x_root + 14, event.y_root + 14)

    def on_file_selected(self, _event=None):
        sel = self.file_tree.selection()
        if sel:
            self.current_path = sel[0]
        self._refresh_values()
        self.update_status()

    # ---- 属性表 -----------------------------------------------------------
    def _rebuild_fields(self):
        keys = dict.fromkeys(key for f in self.files for key in f['props'])
        self.fields = {key: {'name': field_display_name(key)} for key in keys}
        paths = {f['path'] for f in self.files}
        self.batch_checks = {k: v for k, v in self.batch_checks.items() if k in self.fields}
        self.file_checks = {(p, k): v for (p, k), v in self.file_checks.items() if p in paths}
        self.edits = {k: v for k, v in self.edits.items()
                      if paths and (v['all'] or v['path'] in paths)}
        if not paths:
            self.suspects_only = False

    def _on_auto_check(self):
        self.suspects_only = False
        self._refresh_values()
        self.update_status()

    def _checked_keys(self, entry):
        if entry is None:
            return set()
        defaults = (entry['suspects'] if self.suspects_only else
                    auto_clean_keys(entry) if self.auto_check_var.get() else set())
        return {k for k in entry['props']
                if self.file_checks.get((entry['path'], k),
                                        self.batch_checks.get(k, k in defaults))
                and not self._edit_applies(entry, k)}

    def _current_entry(self):
        for f in self.files:
            if f['path'] == self.current_path:
                return f
        return None

    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        entry = self._current_entry() or (self.files[0] if self.files else None)
        checked = self._checked_keys(entry)
        for key in sorted(self.fields, key=field_sort_key):
            meta = self.fields[key]
            value, flag, tags = self._row_values(key, entry, key in checked)
            self.tree.insert('', 'end', iid=key, values=(
                self._glyph(key in checked), meta['name'], value, flag), tags=tags)

    def _row_values(self, key, entry, checked):
        """计算一行的 值/标记/样式，包含待应用修改的展示。"""
        val = entry['props'].get(key) if entry is not None else None
        disp = '—' if val is None else str(val)
        suspect = entry is not None and key in entry['suspects']
        flag = '⚠ AI痕迹' if suspect else ''
        tags = []
        edit = self.edits.get(key)
        if edit and entry is not None and self._edit_applies(entry, key):
            disp = '%s  →  %s' % (val if val is not None else '（无）', edit['value'])
            flag = (flag + ' ' if flag else '') + '✎ 待修改'
            tags.append('edited')
        if suspect:
            tags.append('suspect')
        if checked:
            tags.append('checked')
        return _truncate(disp, 90), flag, tags

    def _edit_applies(self, entry, key):
        edit = self.edits.get(key)
        if not edit or entry is None:
            return False
        return edit['all'] or edit['path'] == entry['path']

    def _refresh_values(self):
        entry = self._current_entry()
        checked = self._checked_keys(entry)
        for key in self.fields:
            value, flag, _tags = self._row_values(key, entry, key in checked)
            self.tree.set(key, 'chk', self._glyph(key in checked))
            self.tree.set(key, 'value', value)
            self.tree.set(key, 'flag', flag)
            self.tree.item(key, tags=_tags)

    @staticmethod
    def _glyph(checked):
        return '☑' if checked else '☐'

    def _set_checked(self, key, checked):
        """设置勾选状态；勾选清除与待应用修改互斥。"""
        if key not in self.fields:
            return
        self.batch_checks[key] = checked
        self.file_checks = {pk: v for pk, v in self.file_checks.items() if pk[1] != key}
        if checked:
            self.edits.pop(key, None)
        self._refresh_values()
        self.update_status()

    def _toggle(self, iid):
        if iid and iid in self.fields:
            self._set_checked(iid, iid not in self._checked_keys(self._current_entry()))

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
        if self._edit_applies(entry, key):
            text += '；将改为：%s' % self.edits[key]['value']
        self.detail.config(text=text)

    def set_all(self, checked):
        self.suspects_only = False
        self.batch_checks = dict.fromkeys(self.fields, checked)
        self.file_checks.clear()
        if checked:
            self.edits.clear()
        self._refresh_values()
        self.update_status()

    def check_suspects_only(self):
        self.suspects_only = True
        self.batch_checks.clear()
        self.file_checks.clear()
        self._refresh_values()
        self.update_status()

    # ---- 编辑属性 ---------------------------------------------------------
    def _edit_hint(self, kind, key):
        if kind == 'ooxml' and key in OOXML_DATE_FIELDS:
            return '日期字段，格式如：2026-10-20 10:00:05（按本地时区理解，保存时自动转为文档标准 UTC 格式）'
        if kind == 'pdf' and key in PDF_DATE_FIELDS:
            return '日期字段，格式如：2026-10-20 10:00:05（按本地时区理解，保存时自动转为 PDF 标准 UTC 格式）'
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
            if apply_all:
                self._set_checked(key, True)
            else:
                self.file_checks[(entry['path'], key)] = True
        else:
            self.edits[key] = {'value': value, 'all': bool(apply_all),
                               'path': entry['path']}
            if apply_all:
                self.file_checks = {pk: v for pk, v in self.file_checks.items() if pk[1] != key}
            else:
                self.file_checks.pop((entry['path'], key), None)
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
        if key in (XMP_KEY, PAGEMETA_KEY):
            messagebox.showinfo('提示', 'XMP / 页面级元数据不支持逐项编辑，可勾选整体清除。')
            return
        if key == THUMB_KEY:
            messagebox.showinfo('提示', '缩略图不支持逐项编辑，可勾选后整体清除。')
            return
        if key.startswith('custom:') and key not in entry['props']:
            messagebox.showinfo('提示', '该文件没有这个自定义属性，不支持新建自定义属性。')
            return
        meta = self.fields[key]
        current = entry['props'].get(key)
        pending = self.edits.get(key) if self._edit_applies(entry, key) else None

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
        if self._busy:
            return
        if not self.files:
            messagebox.showinfo('提示', '请先添加文件。')
            return
        checked = {f['path']: self._checked_keys(f) for f in self.files}
        if not any(checked.values()) and not self.edits:
            messagebox.showinfo('提示', '请先在右侧勾选要清除的字段，或双击属性值进行修改。')
            return
        mode = self.save_var.get()
        if mode == 'overwrite':
            if not messagebox.askyesno(
                    '确认覆盖',
                    '覆盖模式会直接修改原文件（建议先备份）。\n确定继续吗？'):
                return
        # 主线程做快照，worker 只读快照，不再引用可变的界面状态
        snapshot = (list(self.files), checked, dict(self.edits), mode)
        self._set_busy(True, 'clean')
        threading.Thread(target=self._clean_worker, args=snapshot, daemon=True).start()

    def _clean_worker(self, files, checked_by_path, edits, mode):
        q = self._queue
        total = len(files)
        if total:
            q.put(('progress', 0, total))
        ok = fail = 0
        for i, f in enumerate(files, 1):
            if self._stop_event.is_set():
                break
            name = os.path.basename(f['path'])
            checked = checked_by_path[f['path']]
            _d, _s, skipped = collect_ops(f, checked, edits)
            for k, why in skipped:
                q.put(('log', '· %s：%s 跳过（%s）' % (name, field_display_name(k), why), 'info'))
            try:
                dest, removed, edited = clean_file(f, checked, edits, mode)
            except PermissionError:
                fail += 1
                q.put(('log', '✗ %s：文件被占用，请先关闭正在编辑它的程序' % name, 'err'))
                q.put(('cleaned', f['path'], ('失败 ✗', 'err')))
                q.put(('progress', i, total))
                continue
            except Exception as e:
                fail += 1
                q.put(('log', '✗ %s：%s' % (name, e), 'err'))
                q.put(('cleaned', f['path'], ('失败 ✗', 'err')))
                q.put(('progress', i, total))
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
                    q.put(('log', '✓ %s → 已覆盖原文件（%s）' % (name, detail), 'ok'))
                    status = ('已覆盖 ✓', 'ok')
                else:
                    q.put(('log', '✓ %s → 已另存为 %s（%s）'
                           % (name, os.path.basename(dest), detail), 'ok'))
                    status = ('已另存 ✓', 'ok')
                ok += 1
            else:
                q.put(('log', '· %s：勾选/修改的字段在该文件中不适用，未改动' % name, 'info'))
                status = ('无改动', '')
                ok += 1
            q.put(('cleaned', f['path'], status))
            if total:
                q.put(('progress', i, total))
        q.put(('clean_done', ok, fail, mode))

    def _on_cleaned(self, path, status):
        for f in self.files:
            if f['path'] == path:
                f['status'] = status
                break
        self._update_file_row(path)

    def _on_clean_done(self, ok, fail, mode):
        self.edits.clear()
        self.log('—— 完成：成功 %d 个，失败 %d 个 ——' % (ok, fail), 'ok')
        if mode == 'overwrite':
            # 覆盖模式：重扫刷新预览（保持忙碌，扫完自动解锁）
            self._launch_rescan()
            return
        # 副本模式：列表里是未改动的原文件，重扫没有意义，直接收尾
        if self._maybe_continue():
            return
        self._set_busy(False)
        self.update_status()

    def _launch_rescan(self):
        self._set_busy(True, 'rescan')
        threading.Thread(target=self._rescan_worker, daemon=True).start()

    def _rescan_worker(self):
        q = self._queue
        files = list(self.files)
        total = len(files)
        if total:
            q.put(('progress', 0, total))
        for i, f in enumerate(files, 1):
            if self._stop_event.is_set():
                break
            try:
                fresh = scan_file(f['path'])
                q.put(('rescanned', f['path'], fresh))
            except Exception:
                q.put(('rescanned', f['path'], None))
            if total:
                q.put(('progress', i, total))
        q.put(('rescan_done',))

    def _on_rescanned(self, path, fresh):
        for f in self.files:
            if f['path'] == path:
                if fresh is not None:
                    f['props'], f['suspects'] = fresh['props'], fresh['suspects']
                break

    def _on_rescan_done(self):
        if self._maybe_continue():
            return
        self._set_busy(False)
        self._rebuild_fields()
        self._refresh_tree()
        self.update_status()

    # ---- 设置记忆 / 退出 ----------------------------------------------------
    def _apply_saved_prefs(self, s):
        if s.get('save_mode') in ('copy', 'overwrite'):
            self.save_var.set(s['save_mode'])
        if isinstance(s.get('auto_check'), bool):
            self.auto_check_var.set(s['auto_check'])

    def _save_prefs(self):
        try:
            save_settings({'save_mode': self.save_var.get(),
                           'auto_check': self.auto_check_var.get(),
                           'geometry': self.geometry()})
        except tk.TclError:
            pass

    def _on_close(self):
        from tkinter import messagebox
        if self._busy and not messagebox.askyesno(
                '正在处理', '后台任务仍在进行，中途退出可能留下未写完的临时文件。\n确定退出吗？'):
            return
        self._stop_event.set()
        if self._pump_after is not None:  # 取消泵的定时回调，避免销毁后触发报错
            try:
                self.after_cancel(self._pump_after)
            except tk.TclError:
                pass
        self._save_prefs()
        self.destroy()

    # ---- 异常兜底 -----------------------------------------------------------
    def report_callback_exception(self, exc, val, tb):
        """tkinter 回调里的未捕获异常：写崩溃日志并弹窗（窗口版 exe 里 stderr 不可见）。"""
        from tkinter import messagebox
        path = write_crash_log(exc, val, tb)
        try:
            self.log('✗ 程序内部错误：%s（详见 %s）' % (val, path), 'err')
        except Exception:
            pass
        if time.time() - self._last_crash_dialog > 5:  # 连续出错时限流，避免弹窗刷屏
            self._last_crash_dialog = time.time()
            try:
                messagebox.showerror(
                    '程序内部错误',
                    '%s\n\n详细信息已写入：\n%s' % (val, path or '（日志写入失败）'),
                    parent=self)
            except Exception:
                pass


def _truncate(s, n):
    s = str(s)
    return s if len(s) <= n else s[:n - 1] + '…'


def main():
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    try:
        app = CleanerApp()
    except Exception:
        # 窗口还没建起来就出错：窗口版 exe 里 stderr 不可见，写日志并用系统弹窗报告
        info = sys.exc_info()
        write_crash_log(*info)
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(
                0, '启动失败：%s\n\n详细信息已写入 %%APPDATA%%\\AIMetadataCleaner\\crash.log'
                % info[1], '文档AI属性清理工具', 0x10)
        except Exception:
            pass
        raise
    app.mainloop()


if __name__ == '__main__':
    main()
