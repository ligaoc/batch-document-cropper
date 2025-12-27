# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置"""

import sys
import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# 获取项目根目录
ROOT_DIR = os.path.dirname(os.path.abspath(SPEC))

# 收集 PyMuPDF 数据文件
datas = collect_data_files('fitz')

# 收集所有子模块
hiddenimports = collect_submodules('fitz') + [
    'PyQt5.sip',
    'PyQt5.QtCore',
    'PyQt5.QtGui',
    'PyQt5.QtWidgets',
    'src',
    'src.gui',
    'src.gui.main_window',
    'src.gui.file_list_widget',
    'src.gui.margin_panel',
    'src.gui.preview_widget',
    'src.gui.progress_widget',
    'src.core',
    'src.core.batch_processor',
    'src.core.document_converter',
    'src.core.docx_cropper',
    'src.core.file_validator',
    'src.core.output_manager',
    'src.core.pdf_cropper',
    'src.core.resolution_keeper',
    'src.models',
    'src.models.margin_settings',
    'src.models.task',
]

a = Analysis(
    ['run.py'],
    pathex=[ROOT_DIR],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'tkinter',
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
    name='BatchDocumentCropper',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 不显示控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 可以添加图标: icon='icon.ico'
)
