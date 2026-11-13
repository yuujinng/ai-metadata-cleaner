# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置（仅 Windows）：pyinstaller --noconfirm doc_meta_cleaner.spec

版本号通过环境变量 PYI_VERSION 注入（CI 打标签时由 release 工作流传入），
本地未设置时用占位版本 1.0.0.0。文件名保持 ASCII，避免 CI shell 的编码问题。
"""
import os

from PyInstaller.utils.hooks import collect_all
from PyInstaller.utils.win32.versioninfo import (
    FixedFileInfo,
    StringFileInfo,
    StringStruct,
    StringTable,
    VarFileInfo,
    VarStruct,
    VSVersionInfo,
)

# exe 文件名必须保持 ASCII：GitHub Release 资产上传 API 不接受非 ASCII 名称，
# 中文名会被 softprops/action-gh-release 降级成 default.exe。
# 中文产品名保留在下方版本资源（右键 → 详细信息）里。
APP_NAME = 'DocMetaCleaner'
APP_VERSION = os.environ.get('PYI_VERSION', '1.0.0.0').lstrip('v')


def _make_version_file(version):
    """生成 exe 属性（右键 → 详细信息）里的版本资源，返回文件路径。"""
    vers = (tuple(int(p) for p in version.split('.')) + (0, 0, 0, 0))[:4]
    info = VSVersionInfo(
        ffi=FixedFileInfo(
            filevers=vers,
            prodvers=vers,
            mask=0x3F,
            flags=0x0,
            OS=0x40004,
            fileType=0x1,
            subtype=0x0,
            date=(0, 0),
        ),
        kids=[
            StringFileInfo(
                [
                    StringTable(
                        '080404B0',
                        [
                            StringStruct('CompanyName', 'doc-meta-cleaner'),
                            StringStruct('FileDescription', '文档属性清理工具 - 清除 Word/Excel/PPT/PDF 的属性（元数据）'),
                            StringStruct('FileVersion', version),
                            StringStruct('InternalName', APP_NAME),
                            StringStruct('LegalCopyright', 'MIT License (c) 2026 yuujinng'),
                            StringStruct('OriginalFilename', APP_NAME + '.exe'),
                            StringStruct('ProductName', APP_NAME),
                            StringStruct('ProductVersion', version),
                        ],
                    )
                ]
            ),
            VarFileInfo([VarStruct('Translation', [2052, 1200])]),
        ],
    )
    path = os.path.join(SPECPATH, 'build', 'version_info.txt')
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(str(info))
    return path


# tkinterdnd2 自带 Tcl/Tk 扩展（TkDND 库与 pkgIndex.tcl），必须整体收集，否则拖拽不可用
tkdnd_datas, tkdnd_binaries, tkdnd_hiddenimports = collect_all('tkinterdnd2')

a = Analysis(
    ['doc_meta_cleaner.py'],
    pathex=[],
    binaries=tkdnd_binaries,
    datas=tkdnd_datas,
    hiddenimports=tkdnd_hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    # numpy 与本工具无关，但开发机全局环境若装了它会被 PyInstaller 顺着可选导入链
    # 打进包里（带 6MB OpenBLAS DLL），本地构建体积暴涨；CI 干净环境无此问题
    excludes=['numpy'],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version=_make_version_file(APP_VERSION),
)
