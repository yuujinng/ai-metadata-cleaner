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
check('版本号', r.returncode == 0 and 'doc-meta-cleaner' in r.stdout,
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
check('日期已生效', '2026-01-02' in r.stdout, r.stdout[:300])

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

print()
if FAILURES:
    print('FAILED:', len(FAILURES), FAILURES)
    sys.exit(1)
print('CLI SMOKE OK')
