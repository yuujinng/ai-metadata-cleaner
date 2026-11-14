# -*- coding: utf-8 -*-
"""ai-metadata-cleaner 命令行入口：scan / clean / set。

与 doc_meta_cleaner.py（核心逻辑 + 图形界面）分离：
本文件只做参数解析与结果展示，所有能力都复用主模块的函数，
保证命令行与图形界面行为一致。
"""
import argparse
import os
import sys

import doc_meta_cleaner as dmc

try:
    from version import __version__
except ImportError:  # 单独拷贝本文件时兜底，正常随仓库分发不会走到
    __version__ = 'dev'


# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------

def _setup_stdio():
    """英文系统的控制台编码打不出中文时用替代符兜底，避免直接崩溃。"""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            try:
                stream.reconfigure(errors='replace')
            except Exception:
                pass


def _width(s):
    """终端显示宽度：CJK 字符按 2 列计。"""
    return sum(2 if ord(c) > 0x2E80 else 1 for c in s)


def _pad(s, w):
    return s + ' ' * max(0, w - _width(s))


def _short(value, limit=46):
    s = str(value).replace('\r', ' ').replace('\n', ' ')
    return s if _width(s) <= limit else s[:limit - 1] + '…'


def _kind_name(kind):
    return 'Office 文档' if kind == 'ooxml' else 'PDF 文档'


def _sorted_keys(props):
    return sorted(props, key=dmc.field_sort_key)


def iter_files(paths):
    """展开文件/目录参数，返回 (支持格式的文件列表, 缺失路径列表, 不支持的文件列表)。"""
    files, missing, unsupported = [], [], []
    for p in paths:
        if os.path.isdir(p):
            for root, _dirs, names in os.walk(p):
                for n in names:
                    fp = os.path.join(root, n)
                    if (os.path.splitext(n)[1].lower() in dmc.SUPPORTED_EXTS
                            and not dmc.is_lock_file(n)):
                        files.append(fp)
        elif os.path.isfile(p):
            if os.path.splitext(p)[1].lower() in dmc.SUPPORTED_EXTS:
                files.append(p)
            else:
                unsupported.append(p)
        else:
            missing.append(p)
    return files, missing, unsupported


def _warn_paths(missing, unsupported):
    for p in missing:
        print('⚠ 路径不存在：%s' % p, file=sys.stderr)
    for p in unsupported:
        print('⚠ 不支持的格式：%s' % p, file=sys.stderr)
    return bool(missing or unsupported)


def _summary(ok, total, failed):
    print()
    if total:
        print('完成：%d/%d 个文档处理成功%s'
              % (ok, total, '（部分文件失败，见上方 ⚠）' if failed else ''))
    return 1 if failed else 0


def _print_result(name, dest, removed, edited, mode):
    """打印 clean/set 的单文件结果。"""
    if not removed and not edited:
        print('· %s：没有需要处理的属性%s' % (name, '（未生成副本）' if mode == 'copy' else ''))
        return
    where = '已覆盖原文件' if mode == 'overwrite' else '→ 输出 %s' % os.path.basename(dest)
    print('✓ %s：清除 %d 项，修改 %d 项（%s）' % (name, len(removed), len(edited), where))
    if removed:
        print('    清除：' + '、'.join(dmc.field_display_name(k) for k in
                                      sorted(removed, key=dmc.field_sort_key)))
    for k, v in edited:
        print('    修改：%s = %s' % (dmc.field_display_name(k), _short(v, 36)))


def _run_clean(name, entry, checked, edits, mode):
    """执行清理并打印结果；单文件失败时给出警告而不中断整批。返回是否成功。"""
    try:
        dest, removed, edited = dmc.clean_file(entry, checked, edits, mode)
    except PermissionError:
        print('⚠ %s：文件被占用，请先关闭正在编辑它的程序' % name, file=sys.stderr)
        return False
    except Exception as exc:
        print('⚠ %s：%s' % (name, exc), file=sys.stderr)
        return False
    _print_result(name, dest, removed, edited, mode)
    return True


def resolve_fields(spec, entry):
    """把 --fields 规格解析为该文件实际存在的 key 列表。

    auto=仅疑似 AI 痕迹；all=全部属性；其余按逗号分隔，接受字段 key 或中文名称。
    """
    props = entry['props']
    if spec == 'auto':
        return list(entry['suspects'])
    if spec == 'all':
        return list(props)
    keys = []
    display = {dmc.field_display_name(k): k for k in props}
    for raw in [s.strip() for s in spec.split(',') if s.strip()]:
        if raw in props:
            keys.append(raw)
        elif raw in display:
            keys.append(display[raw])
        else:
            raise KeyError('%s（该文件没有这个字段；可用：%s）'
                           % (raw, '、'.join(_sorted_keys(props)) or '无'))
    return keys


def _canon_key(key, entry):
    """宽松解析字段名：PDF 字段允许省略 pdf: 前缀。"""
    kind, props = entry['kind'], entry['props']
    if kind == 'pdf':
        if key.startswith('pdf:'):
            return key
        if key in dmc.PDF_FIELDS:
            return 'pdf:' + key
        raise ValueError('未知字段 %s（PDF 可用字段见 --list-fields）' % key)
    if key.startswith('pdf:'):
        raise ValueError('Office 文档不支持 PDF 字段 %s' % key)
    if not (key in dmc.CORE_FIELDS or key in dmc.APP_FIELDS
            or key.startswith('custom:') or key in props):
        raise ValueError('未知字段 %s（可用字段见 --list-fields）' % key)
    return key


def _build_set_ops(entry, path, assigns):
    """把 字段=值 列表转为该文件的清除/修改操作。值留空 = 清除该字段。"""
    props = entry['props']
    checked, edits = [], {}
    for key, raw in assigns:
        key = _canon_key(key, entry)
        if key in (dmc.XMP_KEY, dmc.PAGEMETA_KEY):
            raise ValueError('%s 不支持逐项编辑，只能整体清除（clean --fields %s）'
                             % (key, key))
        if raw == '':
            if key in props:
                checked.append(key)
            else:
                raise ValueError('字段 %s 在该文件中不存在，无法清除' % key)
            continue
        err, value = dmc.normalize_edit_value(entry['kind'], key, raw)
        if err:
            raise ValueError('%s 的值无效：%s' % (key, err))
        edits[key] = {'value': value, 'all': False, 'path': path}
    return checked, edits


# ---------------------------------------------------------------------------
# 子命令
# ---------------------------------------------------------------------------

def cmd_scan(args):
    files, missing, unsupported = iter_files(args.paths)
    failed = _warn_paths(missing, unsupported)
    ok = 0
    for fp in files:
        name = os.path.basename(fp)
        try:
            entry = dmc.scan_file(fp)
        except Exception as exc:
            print('⚠ 无法读取 %s：%s' % (name, exc), file=sys.stderr)
            failed = True
            continue
        props = entry['props']
        suspects = set(entry['suspects'])
        print('%s （%s，%d 项属性，%d 项疑似 AI 痕迹）'
              % (name, _kind_name(entry['kind']), len(props), len(suspects)))
        if props:
            width = max(_width(dmc.field_display_name(k)) for k in props)
            for k in _sorted_keys(props):
                mark = '   ← 疑似 AI 痕迹' if k in suspects else ''
                print('  %s  %s%s'
                      % (_pad(dmc.field_display_name(k), width), _short(props[k]), mark))
        else:
            print('  （无属性）')
        print()
        ok += 1
    return _summary(ok, len(files), failed)


def cmd_clean(args):
    files, missing, unsupported = iter_files(args.paths)
    failed = _warn_paths(missing, unsupported)
    ok = 0
    for fp in files:
        name = os.path.basename(fp)
        try:
            entry = dmc.scan_file(fp)
            checked = resolve_fields(args.fields, entry)
        except Exception as exc:
            print('⚠ 跳过 %s：%s' % (name, exc), file=sys.stderr)
            failed = True
            continue
        if _run_clean(name, entry, checked, {}, args.mode):
            ok += 1
        else:
            failed = True
    return _summary(ok, len(files), failed)


def cmd_set(args):
    try:
        assigns = _parse_assigns(args.assigns)
    except ValueError as exc:
        print('⚠ %s' % exc, file=sys.stderr)
        return 2
    files, missing, unsupported = iter_files(args.paths)
    failed = _warn_paths(missing, unsupported)
    ok = 0
    for fp in files:
        name = os.path.basename(fp)
        try:
            entry = dmc.scan_file(fp)
            checked, edits = _build_set_ops(entry, fp, assigns)
        except Exception as exc:
            print('⚠ 跳过 %s：%s' % (name, exc), file=sys.stderr)
            failed = True
            continue
        if _run_clean(name, entry, checked, edits, args.mode):
            ok += 1
        else:
            failed = True
    return _summary(ok, len(files), failed)


def _parse_assigns(assign_pairs):
    assigns = []
    for pair in assign_pairs:
        key, sep, raw = pair.partition('=')
        key = key.strip()
        if not sep or not key:
            raise ValueError('赋值必须写成 字段=值 的形式：%r' % pair)
        assigns.append((key, raw.strip()))
    return assigns


def print_fields():
    rows = list(dmc.CORE_FIELDS.items()) + list(dmc.APP_FIELDS.items())
    rows += [('pdf:' + k, n) for k, n in dmc.PDF_FIELDS.items()]
    rows += [(dmc.XMP_KEY, dmc.XMP_NAME), (dmc.PAGEMETA_KEY, dmc.PAGEMETA_NAME)]
    width = max(_width(k) for k, _n in rows)
    print('字段 key                     名称')
    for k, n in rows:
        print('  %s  %s' % (_pad(k, width), n))
    print()
    print('Office 自定义属性为 custom: 开头的 key，只能修改已有的值，不能新建。')


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------

def build_parser():
    parser = argparse.ArgumentParser(
        prog='doc_meta_cli',
        description='文档AI属性清理工具（命令行版）：扫描、清理、修改 Word/Excel/PPT/PDF 的属性（元数据）',
        epilog='示例：python doc_meta_cli.py clean D:\\资料\\   '
               '# 递归清理目录内文档的疑似 AI 痕迹，另存副本')
    parser.add_argument('--version', action='version',
                        version='ai-metadata-cleaner ' + __version__ + '（MIT License）')
    parser.add_argument('--list-fields', action='store_true',
                        help='列出全部可操作的字段后退出')
    sub = parser.add_subparsers(dest='command', metavar='{scan,clean,set}')

    p_scan = sub.add_parser('scan', help='列出文档属性，标出疑似 AI 痕迹')
    p_scan.add_argument('paths', nargs='+', metavar='文件或目录')

    p_clean = sub.add_parser('clean', help='清除属性（默认清除疑似 AI 痕迹）')
    p_clean.add_argument('paths', nargs='+', metavar='文件或目录')
    p_clean.add_argument('--mode', choices=('copy', 'overwrite'), default='copy',
                         help='copy=另存副本（默认），overwrite=覆盖原文件')
    p_clean.add_argument('--fields', default='auto',
                         help='auto=仅疑似 AI 痕迹（默认）；all=全部属性；'
                              '或逗号分隔的字段 key/名称')

    p_set = sub.add_parser('set', help='修改属性：字段=值（值留空表示清除）')
    p_set.add_argument('paths', nargs='+', metavar='文件或目录')
    p_set.add_argument('assigns', nargs='+', metavar='字段=值')
    p_set.add_argument('--mode', choices=('copy', 'overwrite'), default='copy',
                       help='copy=另存副本（默认），overwrite=覆盖原文件')
    return parser


def main(argv=None):
    _setup_stdio()
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.list_fields:
        print_fields()
        return 0
    if not args.command:
        parser.print_help()
        return 2
    return {'scan': cmd_scan, 'clean': cmd_clean, 'set': cmd_set}[args.command](args)


if __name__ == '__main__':
    sys.exit(main())
