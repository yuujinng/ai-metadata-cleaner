# -*- coding: utf-8 -*-
"""CLI 冒烟测试：用子进程真实调用 doc_meta_cli.py 的 scan/clean/set。

依赖 run_tests.py 先生成 out/sample.docx（含 AI 痕迹）。
"""
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
CLI = os.path.join(ROOT, 'doc_meta_cli.py')
OUT = os.path.join(HERE, 'out')

env = dict(os.environ, PYTHONIOENCODING='utf-8')

FAILURES = []


def check(name, cond, extra=''):
    print('[OK]  ' if cond else '[FAIL]', name,
          ('-- ' + str(extra)) if (extra and not cond) else '')
    if not cond:
        FAILURES.append(name)


def run(*args):
    return subprocess.run([sys.executable, CLI] + list(args),
                          capture_output=True, env=env,
                          encoding='utf-8', errors='replace')


sample = os.path.join(OUT, 'sample.docx')
assert os.path.exists(sample), '先运行 run_tests.py 生成样例'
work = os.path.join(OUT, 'cli_work')
shutil.rmtree(work, ignore_errors=True)
os.makedirs(work)
doc = shutil.copy(sample, work)

r = run('--version')
sys.path.insert(0, ROOT)
import version as _ver  # noqa: E402
check('版本号', r.returncode == 0 and 'ai-metadata-cleaner' in r.stdout
      and _ver.__version__ in r.stdout,
      (r.returncode, r.stdout, r.stderr))

r = run('scan', doc)
check('scan 发现 AI 痕迹', r.returncode == 0 and 'ChatGPT' in r.stdout
      and '疑似 AI 痕迹' in r.stdout, (r.returncode, r.stdout[:300], r.stderr))

r = run('scan', os.path.join(ROOT, '示例文件'))
check('scan 目录递归', r.returncode == 0 and 'sample.pdf' in r.stdout,
      (r.returncode, r.stdout[:300], r.stderr))

r = run('clean', doc, '--mode', 'copy')
check('clean 默认清 AI 痕迹并另存副本', r.returncode == 0 and '已清理' in r.stdout,
      (r.returncode, r.stdout[:300], r.stderr))
cleaned = os.path.join(work, 'sample_已清理.docx')
check('副本存在', os.path.exists(cleaned))

r = run('scan', cleaned)
check('清理后零痕迹', r.returncode == 0 and 'ChatGPT' not in r.stdout
      and '，0 项疑似 AI 痕迹' in r.stdout, (r.returncode, r.stdout[:300]))

r = run('set', cleaned, 'creator=张三', '--mode', 'overwrite')
check('set 修改作者（覆盖模式）', r.returncode == 0 and '已覆盖' in r.stdout,
      (r.returncode, r.stdout[:300], r.stderr))
r = run('scan', cleaned)
check('修改已生效', '张三' in r.stdout, r.stdout[:300])

r = run('set', cleaned, 'created=2026-01-02 03:04:05', '--mode', 'overwrite')
check('set 日期规范化', r.returncode == 0, (r.returncode, r.stdout[:200], r.stderr))
r = run('scan', cleaned)
# 本地墙钟输入按当前时区转成 UTC 存储，期望值动态计算
from datetime import datetime, timezone
expect_dt = datetime(2026, 1, 2, 3, 4, 5).astimezone(timezone.utc)
check('日期已生效', expect_dt.strftime('%Y-%m-%dT%H:%M:%SZ') in r.stdout, r.stdout[:300])

r = run('set', cleaned, 'created=明天下午', '--mode', 'overwrite')
check('非法日期被拒（exit 1）', r.returncode == 1
      and '无效' in (r.stdout + r.stderr), (r.returncode, r.stdout, r.stderr))

r = run('clean', doc, '--fields', '不存在的字段')
check('未知字段报错（exit 1）', r.returncode == 1
      and '可用' in (r.stdout + r.stderr), (r.returncode, r.stdout, r.stderr))

r = run('clean', doc, '--fields', 'title,创建者（作者）', '--mode', 'overwrite')
check('显式字段（key+中文名混用）', r.returncode == 0, (r.returncode, r.stdout[:200], r.stderr))
r = run('scan', doc)
check('指定字段已清除', 'AI 起草的测试报告' not in r.stdout
      and 'ChatGPT' not in r.stdout, r.stdout[:300])

r = run('set', cleaned, 'invalid-syntax')
check('非法赋值格式（exit 2）', r.returncode == 2, (r.returncode, r.stdout, r.stderr))

# 目录里混入 Word/WPS 的 ~$ 锁文件：应被跳过而不是当成坏文件报错
work2 = os.path.join(OUT, 'cli_locks')
shutil.rmtree(work2, ignore_errors=True)
os.makedirs(work2)
shutil.copy(sample, os.path.join(work2, 'a.docx'))
open(os.path.join(work2, '~$a.docx'), 'wb').close()
r = run('scan', work2)
check('scan 跳过 ~$ 锁文件', r.returncode == 0 and '~$' not in r.stdout
      and '~$' not in r.stderr and 'a.docx' in r.stdout,
      (r.returncode, r.stdout[:300], r.stderr[:300]))

# 文件被占用（父进程保持句柄打开）：单文件失败应有友好提示并继续，而不是 traceback 崩溃
handle = open(doc, 'rb')
try:
    r = run('clean', doc, '--fields', 'all', '--mode', 'overwrite')
finally:
    handle.close()
check('文件被占用时友好报错（exit 1，无 traceback）',
      r.returncode == 1 and '被占用' in (r.stdout + r.stderr)
      and 'Traceback' not in r.stderr,
      (r.returncode, r.stdout[:300], r.stderr[:500]))
r = run('clean', doc, '--fields', 'all', '--mode', 'overwrite')
check('占用解除后可正常清理', r.returncode == 0, (r.returncode, r.stdout[:200], r.stderr[:300]))

# ---- JSON 模式（Agent/脚本程序化解析）：用全新副本，不影响前面断言 ----
import json as _json

work3 = os.path.join(OUT, 'cli_json')
shutil.rmtree(work3, ignore_errors=True)
os.makedirs(work3)
jdoc = shutil.copy(sample, os.path.join(work3, 'json.docx'))

r = run('scan', jdoc, '--json')
check('scan --json 可解析且 exit 0', r.returncode == 0, (r.returncode, r.stderr[:200]))
data = _json.loads(r.stdout)
check('scan --json 结构与内容',
      data['command'] == 'scan'
      and data['files'][0]['props'].get('creator') == 'ChatGPT (OpenAI)'
      and 'creator' in data['files'][0]['suspects']
      and data['summary'] == {'total': 1, 'ok': 1, 'failed': 0},
      data['summary'])

r = run('clean', jdoc, '--json', '--mode', 'overwrite')
check('clean --json 成功且 exit 0', r.returncode == 0, (r.returncode, r.stderr[:200]))
data = _json.loads(r.stdout)
f0 = data['files'][0]
check('clean --json 结果字段',
      f0['ok'] is True and f0['changed'] is True and f0['dest'] == jdoc
      and 'creator' in f0['removed'] and f0['error'] is None
      and data['summary']['ok'] == 1, f0)

r = run('set', jdoc, 'title=ChatGPT 使用指南', 'creator=ChatGPT', '--mode', 'overwrite')
assert r.returncode == 0, r.stderr
r = run('clean', jdoc, '--json', '--mode', 'overwrite')
check('默认清理移除可疑作者，保留同样命中规则的标题',
      r.returncode == 0 and _json.loads(r.stdout)['files'][0]['removed'] == ['creator'])
r = run('scan', jdoc, '--json')
check('标题仍可扫描到并标记疑似',
      r.returncode == 0 and 'title' in _json.loads(r.stdout)['files'][0]['suspects'])

r = run('set', jdoc, 'title=应保留原值', 'custom:Missing=value',
        '--json', '--mode', 'overwrite')
f0 = _json.loads(r.stdout)['files'][0]
check('不支持的写入以退出码 1 和字段明细报告，且不做部分修改',
      r.returncode == 1 and not f0['ok'] and not f0['changed']
      and f0['skipped'][0]['key'] == 'custom:Missing', f0)
r = run('scan', jdoc, '--json')
check('失败后原文件未被部分修改',
      _json.loads(r.stdout)['files'][0]['props']['title'] == 'ChatGPT 使用指南')

r = run('clean', jdoc, '--json')
f0 = _json.loads(r.stdout)['files'][0]
check('无适用字段是成功的无改动，不生成副本',
      r.returncode == 0 and f0['ok'] and not f0['changed'] and f0['dest'] is None, f0)

r = run('set', jdoc, 'created=明天下午', '--json', '--mode', 'overwrite')
check('set --json 文件失败（exit 1）', r.returncode == 1, (r.returncode,))
data = _json.loads(r.stdout)
check('set --json 失败明细在 errors',
      data['summary']['failed'] == 1 and '无效' in data['errors'][0]['error'],
      data['errors'])

r = run('set', jdoc, 'invalid-syntax', '--json')
check('set --json 用法错误（exit 2）', r.returncode == 2, (r.returncode, r.stdout[:200]))
data = _json.loads(r.stdout)
check('set --json 用法错误结构', data['command'] == 'set' and 'error' in data, data)

r = run('scan', os.path.join(work3, '不存在.docx'), '--json')
check('scan --json 缺失路径（exit 1）', r.returncode == 1, (r.returncode,))
data = _json.loads(r.stdout)
check('scan --json missing 列表',
      data['missing'] == [os.path.join(work3, '不存在.docx')]
      and data['files'] == [] and data['summary']['failed'] == 1, data.get('missing'))

r = run('clean', jdoc, '--json', '--fields', '不存在的字段')
check('clean --json 字段解析失败进 errors（exit 1）', r.returncode == 1, (r.returncode,))
data = _json.loads(r.stdout)
check('clean --json errors 结构', data['files'] == []
      and '可用' in data['errors'][0]['error'], data['errors'])

print()
if FAILURES:
    print('FAILED:', len(FAILURES), FAILURES)
    sys.exit(1)
print('CLI SMOKE OK')
