# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller配置文件 - ETF API调用工具
支持Mac Intel和Apple Silicon架构
"""

import sys
import os
from pathlib import Path

# 添加源码路径
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

block_cipher = None

a = Analysis(
    ['../src/etf_api_caller.py'],
    pathex=[str(Path(__file__).parent.parent)],
    binaries=[],
    datas=[
        ('../config/api_config.json', 'config'),
        ('../config/settings.json', 'config'),
    ],
    hiddenimports=[
        'requests',
        'json',
        'datetime',
        'pathlib',
        'argparse',
        'sys',
        'os',
        'time',
        'platform',
        're',
        'logging',
        'threading'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'PIL',
        'PyQt5',
        'PyQt6',
        'PySide2',
        'PySide6'
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='etf_api_caller',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
