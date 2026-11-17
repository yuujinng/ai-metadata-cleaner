# -*- coding: utf-8 -*-
"""ai-metadata-cleaner 命令行入口：scan / clean / set。

与 doc_meta_cleaner.py（核心逻辑 + 图形界面）分离：
本文件只做参数解析与结果展示，所有能力都复用主模块的函数，
保证命令行与图形界面行为一致。

退出码：0=全部成功；1=有文件处理失败（被占用/加密/字段不存在/值非法，
明细见文本输出的 ⚠ 行或 JSON 的 errors/files[].error）；2=命令用法错误。
三个子命令都支持 --json：stdout 输出单个 JSON 文档（ensure_ascii，
任何控制台编码下都是合法 JSON），供脚本与 AI Agent 程序化解析。
"""
import argparse
import json
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


def _warn_paths(missing, unsupported, json_mode=False):
    """路径警告。JSON 模式下不打印（警告会进 JSON 文档），仍返回是否有问题。"""
    if not json_mode:
        for p in missing:
            print('⚠ 路径不存在：%s' % p, file=sys.stderr)
        for p in unsupported:
            print('⚠ 不支持的格式：%s' % p, file=sys.stderr)
    return bool(missing or unsupported)


def _emit_json(doc):
    """JSON 输出统一入口：ensure_ascii 保证任何控制台编码下都是合法 JSON。"""
    print(json.dumps(doc, ensure_ascii=True, indent=2))


def _summary(ok, total, failed, json_mode=False):
    if not json_mode:
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


def _run_clean(name, entry, checked, edits, mode, json_mode=False):
    """执行清理；单文件失败时给出警告而不中断整批。

    返回结果 dict：{name, ok, dest, changed, removed, edited, skipped, error}；
    json_mode=True 时只收集不打印。
    """
    result = {'name': name, 'ok': False, 'dest': None, 'changed': False,
              'removed': [], 'edited': [], 'skipped': [], 'error': None}

    def report_skipped(skipped):
        result['skipped'] = [{'key': k, 'reason': why} for k, why in skipped]
        result['error'] = '；'.join('%s：%s' % (k, why) for k, why in skipped)
        if not json_mode:
            print('⚠ %s：%s' % (name, result['error']), file=sys.stderr)

    try:
        deletes, sets, skipped = dmc.collect_ops(entry, checked, edits)
        if skipped:
            report_skipped(skipped)
            return result
        dest, removed, edited = dmc.clean_file(entry, checked, edits, mode)
    except PermissionError:
        msg = '文件被占用，请先关闭正在编辑它的程序'
        if not json_mode:
            print('⚠ %s：%s' % (name, msg), file=sys.stderr)
        result['error'] = msg
        return result
    except Exception as exc:
        if not json_mode:
            print('⚠ %s：%s' % (name, exc), file=sys.stderr)
        result['error'] = str(exc)
        return result
    result.update(dest=dest if dest != entry['path'] or mode == 'overwrite' else None,
                  changed=bool(removed or edited),
                  removed=list(removed),
                  edited=[{'key': k, 'value': v} for k, v in edited])
    unapplied = (set(deletes) - set(removed)) | (sets.keys() - dict(edited).keys())
    if unapplied:
        report_skipped([(k, '未能写入所请求的变更') for k in sorted(unapplied)])
        return result
    result['ok'] = True
    if not json_mode:
        _print_result(name, dest, removed, edited, mode)
    return result


def resolve_fields(spec, entry):
    """把 --fields 规格解析为该文件实际存在的 key 列表。

    auto=与界面相同的默认清理范围；all=全部属性；其余按逗号分隔。
    """
    props = entry['props']
    if spec == 'auto':
        return sorted(dmc.auto_clean_keys(entry))
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
        if key in (dmc.XMP_KEY, dmc.PAGEMETA_KEY, dmc.THUMB_KEY):
            raise ValueError('%s 不支持逐项编辑，只能整体清除（clean --fields %s）'
                             % (key, key))
        if not raw.strip():
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
    failed = _warn_paths(missing, unsupported, json_mode=args.json)
    ok = 0
    out, errors = [], []
    for fp in files:
        name = os.path.basename(fp)
        try:
            entry = dmc.scan_file(fp)
        except Exception as exc:
            if not args.json:
                print('⚠ 无法读取 %s：%s' % (name, exc), file=sys.stderr)
            errors.append({'path': fp, 'error': str(exc)})
            failed = True
            continue
        props = entry['props']
        suspects = sorted(entry['suspects'])
        out.append({'path': fp, 'name': name, 'kind': entry['kind'],
                    'props': props, 'suspects': suspects})
        if not args.json:
            print('%s （%s，%d 项属性，%d 项疑似 AI 痕迹）'
                  % (name, _kind_name(entry['kind']), len(props), len(suspects)))
            if props:
                width = max(_width(dmc.field_display_name(k)) for k in props)
                for k in _sorted_keys(props):
                    mark = '   ← 疑似 AI 痕迹' if k in entry['suspects'] else ''
                    print('  %s  %s%s'
                          % (_pad(dmc.field_display_name(k), width),
                             _short(props[k]), mark))
            else:
                print('  （无属性）')
            print()
        ok += 1
    if args.json:
        _emit_json({'command': 'scan', 'files': out, 'errors': errors,
                    'missing': missing, 'unsupported': unsupported,
                    'summary': {'total': len(files), 'ok': ok,
                                'failed': len(errors) + len(missing) + len(unsupported)}})
    return _summary(ok, len(files), failed, json_mode=args.json)


def cmd_clean(args):
    files, missing, unsupported = iter_files(args.paths)
    failed = _warn_paths(missing, unsupported, json_mode=args.json)
    ok = 0
    out, errors = [], []
    for fp in files:
        name = os.path.basename(fp)
        try:
            entry = dmc.scan_file(fp)
            checked = resolve_fields(args.fields, entry)
        except Exception as exc:
            if not args.json:
                print('⚠ 跳过 %s：%s' % (name, exc), file=sys.stderr)
            errors.append({'path': fp, 'error': str(exc)})
            failed = True
            continue
        result = _run_clean(name, entry, checked, {}, args.mode,
                            json_mode=args.json)
        if result['ok']:
            ok += 1
        else:
            failed = True
        result['path'] = fp
        out.append(result)
    if args.json:
        _emit_json({'command': 'clean', 'mode': args.mode, 'files': out,
                    'errors': errors, 'missing': missing,
                    'unsupported': unsupported,
                    'summary': {'total': len(files), 'ok': ok,
                                'failed': len(errors) + len(missing)
                                + len(unsupported) + sum(1 for r in out if not r['ok'])}})
    return _summary(ok, len(files), failed, json_mode=args.json)


def cmd_set(args):
    try:
        # 两组 nargs='+' 会把除最后一项外的所有赋值都误当成路径。
        # 第一个路径之后，从首个非现有路径的「字段=值」开始解析赋值。
        split = next((i for i, item in enumerate(args.items[1:], 1)
                      if '=' in item and not os.path.exists(item)), len(args.items))
        if split == len(args.items):
            raise ValueError('请在文件或目录之后提供至少一个 字段=值')
        assigns = _parse_assigns(args.items[split:])
    except ValueError as exc:
        if args.json:
            _emit_json({'command': 'set', 'error': str(exc),
                        'summary': {'total': 0, 'ok': 0, 'failed': 0}})
        else:
            print('⚠ %s' % exc, file=sys.stderr)
        return 2
    files, missing, unsupported = iter_files(args.items[:split])
    failed = _warn_paths(missing, unsupported, json_mode=args.json)
    ok = 0
    out, errors = [], []
    for fp in files:
        name = os.path.basename(fp)
        try:
            entry = dmc.scan_file(fp)
            checked, edits = _build_set_ops(entry, fp, assigns)
        except Exception as exc:
            if not args.json:
                print('⚠ 跳过 %s：%s' % (name, exc), file=sys.stderr)
            errors.append({'path': fp, 'error': str(exc)})
            failed = True
            continue
        result = _run_clean(name, entry, checked, edits, args.mode,
                            json_mode=args.json)
        if result['ok']:
            ok += 1
        else:
            failed = True
        result['path'] = fp
        out.append(result)
    if args.json:
        _emit_json({'command': 'set', 'mode': args.mode, 'files': out,
                    'errors': errors, 'missing': missing,
                    'unsupported': unsupported,
                    'summary': {'total': len(files), 'ok': ok,
                                'failed': len(errors) + len(missing)
                                + len(unsupported) + sum(1 for r in out if not r['ok'])}})
    return _summary(ok, len(files), failed, json_mode=args.json)


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
    rows += [(dmc.THUMB_KEY, dmc.THUMB_NAME),
             (dmc.XMP_KEY, dmc.XMP_NAME), (dmc.PAGEMETA_KEY, dmc.PAGEMETA_NAME)]
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

    def _add_json(p):
        p.add_argument('--json', action='store_true',
                       help='结果以 JSON 输出到 stdout，供脚本/Agent 程序化解析'
                            '（退出码 0=全部成功，1=有文件失败，2=用法错误）')

    p_scan = sub.add_parser('scan', help='列出文档属性，标出疑似 AI 痕迹')
    p_scan.add_argument('paths', nargs='+', metavar='文件或目录')
    _add_json(p_scan)

    p_clean = sub.add_parser('clean', help='清除属性（默认保留标题、关键词等内容字段）')
    p_clean.add_argument('paths', nargs='+', metavar='文件或目录')
    p_clean.add_argument('--mode', choices=('copy', 'overwrite'), default='copy',
                         help='copy=另存副本（默认），overwrite=覆盖原文件')
    p_clean.add_argument('--fields', default='auto',
                         help='auto=仅清除可自动勾选的疑似属性（默认）；all=全部属性；'
                              '或逗号分隔的字段 key/名称')
    _add_json(p_clean)

    p_set = sub.add_parser('set', help='修改属性：字段=值（值留空表示清除）')
    p_set.add_argument('items', nargs='+', metavar='文件或目录 … 字段=值')
    p_set.add_argument('--mode', choices=('copy', 'overwrite'), default='copy',
                       help='copy=另存副本（默认），overwrite=覆盖原文件')
    _add_json(p_set)
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
