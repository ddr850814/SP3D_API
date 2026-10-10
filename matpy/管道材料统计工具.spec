# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

# cryptography 的 x509 等子模块仅被 _rust 扩展运行期动态导入（无静态
# import 链），hiddenimports 单列会漏，须全量收集；
# zipfile_deflate64：ISO 图纸页数要解 Deflate64 zip（SP3D 图纸 blob）
_hidden = (['pyodbc', 'oracledb', 'zipfile_deflate64']
           + collect_submodules('cryptography'))


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('64.ico', '.')],
    hiddenimports=_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # UI 已换 ttkbootstrap（tkinter 自带），Qt 全家不再打包
        'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'shiboken6', 'qdarktheme',
        # openpyxl 可选依赖（numbers.py/image.py 均 try/except 兜底；
        # utils/dataframe.py 虽硬 import numpy 但不在 openpyxl 导入链上）
        'numpy', 'pandas', 'psutil',
        # 注意 PIL 不能排除：ttkbootstrap 顶层经 widgets→dialogs→colorchooser
        # 硬 import PIL
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='SP3D管道报表工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['64.ico'],
)
