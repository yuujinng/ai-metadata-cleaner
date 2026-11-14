<div align="center">

[简体中文](README.md) | English

# 🧹 AI Metadata Cleaner

**One click to clear or edit the properties (metadata) of Word / Excel / PPT / PDF files — wipe AI-generated traces**

🔒 **100% local processing** — your files never leave your computer

[![Release](https://img.shields.io/github/v/release/yuujinng/ai-metadata-cleaner?style=flat-square)](../../releases)
[![下载量](https://img.shields.io/github/downloads/yuujinng/ai-metadata-cleaner/total?style=flat-square&label=%E4%B8%8B%E8%BD%BD)](../../releases)
[![CI](https://github.com/yuujinng/ai-metadata-cleaner/actions/workflows/ci.yml/badge.svg)](https://github.com/yuujinng/ai-metadata-cleaner/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/github/license/yuujinng/ai-metadata-cleaner?style=flat-square)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.11%2B-3776AB?style=flat-square&logo=python&logoColor=white)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/平台-Windows%2010%2F11-0078D6?style=flat-square&logo=windows11&logoColor=white)](https://github.com/yuujinng/ai-metadata-cleaner#-quick-start)

<img src="docs/screenshot.png?v=1.3" alt="App preview: suspected AI traces highlighted in red and pre-checked" width="720"/>

*Drop files in → suspected AI traces get highlighted and pre-checked → clean with one click*

<img src="docs/demo.gif?v=1.4" alt="Demo GIF: drop files → traces highlighted and checked → one-click cleanup with progress bar, files marked saved in the list" width="720"/>

</div>

---

## ✨ Features

- 🤖 **Automatic AI-trace detection** — sources such as `ChatGPT`, `Claude`, `python-docx`, `openpyxl`, `Skia/PDF` are highlighted automatically; identity fields (author / generator / time statistics) are pre-checked, while content-related fields such as Title are only flagged (to prevent accidental deletion)
- 🧬 **Structural fingerprinting** — spots one-shot machine generation without relying on keywords: word-count statistics of 0 while the body has text, document-internal times more than a year off the file's modification time (frozen generator-template dates), Mac Word leftovers, script-library version strings, WPS cloud fingerprints (ICV), thumbnails that Word/Excel would never write, and PDF's `anonymous`/`untitled` placeholder defaults
- 🧽 **Thorough cleanup** — author, last-modified-by, application name, created/modified time, editing statistics, company, custom properties, the document thumbnail, ZIP entry-timestamp fingerprints, plus PDF's Info dictionary and XMP metadata packet…
- ✏️ **Free editing** — set any property to any value; "apply to all files" batch-unifies a whole batch; dates are converted to the standard document format automatically
- 🖱️ **Drag & drop** — drop files or entire folders (recursively scanned) straight into the window
- ⚡ **Background processing** — scanning and cleaning run on a background thread with a progress bar, so even folder-sized batches keep the UI responsive; results are marked in the file list (hover for the full path) and your save-mode / auto-check / window preferences are remembered
- 🖥️ **Command-line mode** — `scan` / `clean` / `set` subcommands for batch processing and scripting, sharing the same core logic as the GUI; `--json` output + semantic exit codes make it easy to drive from scripts and AI agents
- 🔒 **Zero damage to content** — only property parts are rewritten; the rest of the document is preserved byte-for-byte; overwrite mode writes to a temp file first and replaces atomically, so failures never harm the original
- 🕐 **Leaves no processing traces** — cleaning does not touch the file's system timestamps (overwrite mode even restores the Windows creation time; copies keep the original's modification time)
- 📦 **Zero setup** — download a single exe from [Releases](../../releases) and run it: no Python, no network needed; running from source is also supported

## 🕵️ What traces can it spot?

| Trace type | Typical giveaway |
| --- | --- |
| Author / last modified by | `ChatGPT`, `Claude`, `python-docx`, etc. |
| Application name | Generator shows as a script library (openpyxl, pandoc, Skia/PDF, LibreOffice, pdfTeX, etc.) |
| Time pattern | Created and modified time identical down to the second; document-internal times more than a year older than the file's modification time (python-docx freezes its template at 2013) or in the future |
| Editing statistics | Total editing time of 0 minutes; word/character/line counts of 0 while the body has text (Office recalculates them on save, scripts don't) |
| Application mismatch | A "Microsoft Macintosh Word" document handed in on Windows; openpyxl's non-Office `3.1` version string; revision 1 with zero editing time |
| WPS cloud fingerprint | Custom properties `ICV` (a cloud-document GUID) and `KSOProductBuildVer`, whose values contain no keywords at all |
| Thumbnail leftover | Word/Excel never write `docProps/thumbnail.jpeg` by default — its presence points to a generator template |
| Embedded PDF info | Producer/Creator in the Info dictionary, XMP metadata packet, page-level embedded metadata, `anonymous`/`untitled` placeholder defaults |
| Others | Company, manager, custom properties, local paths in hyperlink bases, etc. |

> "Suspected AI traces" are detected by keywords and rules — there may be false positives/negatives, so please review the checklist before cleaning.

## 🚀 Quick start

> 💡 Just want to try it out? The [`示例文件/`](示例文件/) folder (sample files) contains 4 demo documents with AI traces — one for Word / Excel / PPT / PDF each. Drop them in and play.

**Option 1: download the exe (recommended, no Python needed)**

1. Open the [Releases](../../releases) page and download **AIMetadataCleaner.exe** (~15 MB) from the latest release;
2. Double-click to run. **No Python installation required** — everything is bundled, and it works offline. The exe file name is the ASCII `AIMetadataCleaner.exe` (for Release-asset compatibility); the file description in its properties stays Chinese.

> Your browser or Windows SmartScreen may warn about an "unknown publisher": the exe is not code-signed. Click "More info → Run anyway". Some antivirus tools may flag it (a common false positive for PyInstaller binaries) — add it to your trust list. The SHA256 checksum of the exe is listed on the Release page if you want to verify integrity.

**Option 2: run from source (needs Python, for developers)**

1. Install [Python 3.11+](https://www.python.org/downloads/) (check *"Add Python to PATH"*)
2. Download this repository (*Code → Download ZIP*, or `git clone`)
3. Double-click **文档AI属性清理工具.bat** (the double-click launcher)

On first launch it installs `pypdf` and `tkinterdnd2` automatically (once only). When the status shows 「PDF 支持：已就绪」 (PDF support: ready), everything is fine. A black console window stays next to the GUI while it runs and exits when you close the window.

**Option 3: command line**

```bash
git clone https://github.com/yuujinng/ai-metadata-cleaner.git
cd ai-metadata-cleaner
pip install -r requirements.txt
python doc_meta_cleaner.py
```

## 📖 Usage

**Clean properties, in four steps:**

1. **Drag files or an entire folder into the window** (folders are scanned recursively for supported documents), or use **Add files** / **Add folder**;
2. The right pane lists all properties; suspected AI traces are highlighted in red: identity fields such as **author, generator, time statistics are pre-checked**, while content-related fields like title/keywords are only flagged (check them manually if needed); click the ☑/☐ in the first column or double-click a row to toggle — the check state applies to all files in the list;
3. Choose how to save: **save as a copy** (`原名_已清理` / originalname_cleaned, recommended) or **overwrite the original**;
4. Click **Start cleaning** — the log below reports which fields were cleaned, file by file.

**Edit properties:** select a file on the left, **double-click the "current value" column** (or select a row and click **Edit selected…**), then type a new value in the dialog:

- Check "apply to all files in the list" to unify authors, etc. across a whole batch in one go;
- Date fields accept input like `2026-10-20 10:00:05`, interpreted in your local timezone and saved as standard document UTC format (Word/PDF will display what you typed); revision number and total editing time need integers;
- **Saving an empty value clears the field**;
- For the same field, "edit" and "clear" are mutually exclusive — the last one set wins;
- Fields the file doesn't have (e.g. category, PDF subject) are created automatically on save; custom properties can only be edited if they already exist;
- Click **Undo all edits** to cancel pending edit items;
- The PDF XMP packet cannot be edited item by item — it can only be cleared as a whole; but editing/clearing Info fields syncs XMP automatically, so no stale values remain.
- The document thumbnail can likewise only be cleared as a whole (checked removal also strips its package registration), not edited item by item.

## ⌨️ Command line

The GUI isn't required — the CLI shares the same core logic, handy for batch jobs and scripting:

```bash
# Scan: list properties; suspected AI traces are flagged (directories are expanded recursively)
python doc_meta_cli.py scan document.docx
python doc_meta_cli.py scan D:\materials\

# Clean: by default only suspected AI traces, saved as a copy (originalname_cleaned)
python doc_meta_cli.py clean D:\materials\
python doc_meta_cli.py clean document.pdf --mode overwrite            # overwrite originals
python doc_meta_cli.py clean document.docx --fields all               # clear every property
python doc_meta_cli.py clean document.docx --fields creator,创建时间   # pick fields (key or Chinese name)

# Edit: field=value; an empty value clears the field; multiple files per call
python doc_meta_cli.py set report.docx creator=ZhangSan
python doc_meta_cli.py set report.docx created="2026-10-20 10:00:05" --mode overwrite

# JSON output: a single JSON document on stdout, for scripts and AI agents
python doc_meta_cli.py scan paper.docx --json

python doc_meta_cli.py --list-fields    # list all operable fields
```

Exit codes: **0** all succeeded; **1** at least one file failed (in use / encrypted / unknown field / invalid value — details in the ⚠ lines or the JSON `errors`); **2** the command itself is wrong.

## 🤖 Working with AI agents

The tool's signature scenario is "students have an agent generate the homework" — so the smoothest usage is not dragging files into the GUI yourself, but **making cleanup a finishing rule the agent runs right after generating**. Paste the block below together with your generation request into any terminal-capable agent (Claude Code / Codex CLI / ZCode, etc.):

```text
[Cleanup document properties before delivery — REQUIRED]
Before handing me any generated Word/Excel/PPT/PDF document, you must finish these steps, no skipping:
1. If the cleaner isn't on this machine yet, clone it to a temp directory (on Windows you may use %TEMP% instead of /tmp):
   git clone --depth 1 https://github.com/yuujinng/ai-metadata-cleaner.git /tmp/ai-metadata-cleaner
2. For every generated .docx/.xlsx/.pptx/.pdf file, run (overwrite mode, cleans in place):
   python /tmp/ai-metadata-cleaner/doc_meta_cli.py clean "document path" --mode overwrite
3. Re-scan every file afterwards to confirm no "suspected AI traces" remain:
   python /tmp/ai-metadata-cleaner/doc_meta_cli.py scan "document path"
4. Both commands must exit 0 to count as done; if not, report the actual error to me honestly. Never fake success.
5. Clean document properties only (author / application / time statistics); never touch the body text.
```

**Programmatic interface** (for agents/scripts that parse results themselves): all three subcommands accept `--json` and print a single JSON document to stdout; Chinese values are `\uXXXX`-escaped so it stays valid JSON under any console encoding:

```bash
python doc_meta_cli.py scan paper.docx --json
```

```json
{
  "command": "scan",
  "files": [
    { "path": "paper.docx", "kind": "ooxml",
      "props": { "creator": "python-docx", "TotalTime": "0" },
      "suspects": ["creator", "created", "modified", "TotalTime"] }
  ],
  "errors": [],
  "missing": [],
  "unsupported": [],
  "summary": { "total": 1, "ok": 1, "failed": 0 }
}
```

For `clean` / `set`, each file in the JSON carries `ok`, `changed`, `dest` (output path), `removed` (cleared fields), `edited` (edited fields with new values) and `error` (per-file failure reason).

| Exit code | Meaning | What an agent should do |
| :---: | --- | --- |
| 0 | All succeeded | Proceed with delivery |
| 1 | At least one file failed (in use / encrypted / unknown field / invalid value) — see `errors` and `files[].error` | Read the details, fix, then retry |
| 2 | The command itself is wrong (missing args, bad assignment format) | Fix the command; don't retry as-is |

> 💡 For batch jobs pass a directory (expanded recursively): `clean D:\homework\`. To unify the author across a batch, use `set`: `python doc_meta_cli.py set D:\homework\ creator=ZhangSan --mode overwrite`.

## 📂 Supported formats

| Type | Extensions | Clean | Edit |
| --- | --- | :-: | :-: |
| Word | `.docx` `.docm` | ✅ | ✅ |
| Excel | `.xlsx` `.xlsm` | ✅ | ✅ |
| PPT | `.pptx` `.pptm` `.ppsx` `.ppsm` | ✅ | ✅ |
| PDF | `.pdf` | ✅ | ✅ (XMP cleared as a whole only) |

The legacy binary `.doc/.xls/.ppt` formats are not supported yet.

## 🔒 How cleanup works (safety boundaries)

- Office documents are essentially ZIP+XML packages. The tool only rewrites the three property parts `docProps/core.xml`, `docProps/app.xml`, `docProps/custom.xml` (clearing the thumbnail additionally removes `docProps/thumbnail.jpeg` together with its package relationship and [Content_Types] registration, leaving no dangling references) — **body text, styles, images and everything else is preserved byte-for-byte**, so documents display exactly as before. New fields are inserted at the correct position per the document spec; missing property parts are created and registered automatically so Word/Excel/PPT open the file normally.
- When ZIP entry timestamps look machine-made (all identical, or with abnormal years — Word saves progressively so its timestamps spread out, while one-shot script writes are all identical), the rebuilt package normalizes them to the file's modification time. Packages whose timestamps already look normal are left untouched.
- PDFs are cloned via pypdf, then the Info dictionary and /Metadata (XMP), including metadata embedded on page objects, are rewritten; page content is untouched. Editing/clearing Info fields syncs the XMP packet so no stale values remain.
- "Overwrite original" writes to a temp file first and replaces atomically; a failure never corrupts the original.
- Cleaning does not touch the file's system timestamps: overwrite mode fully restores created/modified/accessed times (including the Windows creation time), and copies keep the original's modification time — so nothing says "just processed".

## ⚠️ Notes

- Close documents that are open in Word/WPS/preview apps before cleaning — files in use cannot be written.
- This tool cleans document **properties**; it does not detect or modify the body content itself.
- AI detectors may also judge writing style — this tool only makes the metadata layer clean.

## ❓ FAQ

- **SmartScreen blocks the exe / antivirus alerts**: the exe is not code-signed and is packaged with PyInstaller — a common false positive. Click "More info → Run anyway", or use Option 2 (run from source) instead.
- **"Python not found"**: when installing Python 3, check "Add Python to PATH", or start via `python doc_meta_cleaner.py`; if you'd rather not install Python, just use the exe from Releases.
- **PDF files fail to load**: check whether the log at the top says 「PDF 支持：未安装」 (PDF support: not installed), then run `pip install pypdf` manually and restart the tool.
- **Dragging files into the window does nothing**: the optional component tkinterdnd2 is missing — run `pip install tkinterdnd2` and restart (other features still work).

## 🗂️ Project layout

```
ai-metadata-cleaner/
├── doc_meta_cleaner.py       Core logic + GUI
├── doc_meta_cli.py           CLI entry (scan / clean / set)
├── version.py                Single source of truth for the version (CLI / GUI title / exe version resource)
├── doc_meta_cleaner.spec     PyInstaller spec (builds the Windows exe)
├── README.en.md              English documentation
├── 文档AI属性清理工具.bat        Double-click launcher for running from source
├── 示例文件/                  4 demo documents with AI traces
├── docs/
│   ├── screenshot.png        README screenshot
│   └── demo.gif              README demo GIF
├── tests/
│   ├── run_tests.py          Core logic tests
│   ├── gui_smoke.py          GUI smoke test
│   ├── cli_smoke.py          CLI smoke test
│   ├── _make_bat.py          Regenerate the launcher
│   ├── _make_screenshot.py   Regenerate the README screenshot
│   ├── _make_gif.py          Regenerate the README demo GIF
│   └── _win_capture.py       Window-capture helper shared by screenshot/GIF
│                             scripts (falls back to PrintWindow when screen
│                             grabbing is unavailable, e.g. non-desktop sessions)
└── .github/
    ├── ISSUE_TEMPLATE/       Issue templates
    └── workflows/            CI & Release (a v* tag auto-publishes)
```

## 🛠️ Development & testing

```bash
python tests/run_tests.py         # Core logic tests (generate samples and verify cleanup)
python tests/gui_smoke.py         # GUI smoke test
python tests/cli_smoke.py         # CLI smoke test
python tests/_make_screenshot.py  # Regenerate the README screenshot
python tests/_make_gif.py         # Regenerate the README demo GIF
python tests/_make_bat.py         # Regenerate the launcher (keep ASCII + CRLF)
pyinstaller --noconfirm doc_meta_cleaner.spec   # Build a single-file exe locally (output in dist/)
```

Releases: set `__version__` in `version.py` to the new version, commit, then push a matching `vX.Y.Z` tag — the Release workflow first verifies the tag matches `version.py` (and fails otherwise), then runs the full test suite, builds the Windows single-file exe, and publishes it with the source zip as a GitHub Release.

> Note: the double-click launcher `文档AI属性清理工具.bat` must be saved as ASCII + CRLF — do not edit it and re-save as UTF-8 in an editor. If it breaks, regenerate it with `python tests/_make_bat.py`.

## 📄 License

[MIT](LICENSE) © 2026 yuujinng
