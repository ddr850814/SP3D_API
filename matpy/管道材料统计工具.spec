# -*- mode: python ; coding: utf-8 -*-
import os
import sys

from PyInstaller.utils.hooks import collect_submodules

# cryptography 的 x509 等子模块仅被 _rust 扩展运行期动态导入（无静态
# import 链），hiddenimports 单列会漏，须全量收集；
# zipfile_deflate64：ISO 图纸页数要解 Deflate64 zip（SP3D 图纸 blob）
_hidden = (['pyodbc', 'oracledb', 'zipfile_deflate64']
           + collect_submodules('cryptography'))

# cryptography 的 _rust.pyd（abi3 轮子）链接稳定 ABI 转发 DLL
# python3t.dll，PyInstaller 6.22 未自动捆扎（Python 3.15 里该 DLL 在
# <prefix>\abi3t-compat\ 子目录，超出扫描范围）——干净机器上
# ImportError: DLL load failed while importing _rust；须显式带入
_p3t_cands = [os.path.join(sys.base_prefix, 'python3t.dll'),
              os.path.join(sys.base_prefix, 'abi3t-compat', 'python3t.dll')]
a_binaries = [(next(p for p in _p3t_cands if os.path.exists(p)), '.')]


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=a_binaries,
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
