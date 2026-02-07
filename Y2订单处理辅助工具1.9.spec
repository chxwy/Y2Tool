# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['image_organizer.py'],
    pathex=[],
    binaries=[],
    datas=[('logo.ico', '.'), ('update_module.py', '.'), ('updater.py', '.')],
    hiddenimports=['requests', 'psutil'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['torch', 'torchvision', 'PySide6', 'shiboken6', 'matplotlib', 'cv2', 'pyarrow', 'scipy', 'seaborn', 'transformers', 'gradio', 'fastapi', 'streamlit'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Y2订单处理辅助工具1.9',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['logo.ico'],
)

# 更新助手程序
updater_a = Analysis(
    ['updater.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['psutil', 'zipfile'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
updater_pyz = PYZ(updater_a.pure)

updater_exe = EXE(
    updater_pyz,
    updater_a.scripts,
    [],
    exclude_binaries=True,
    name='updater',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,  # 更新助手需要控制台显示进度
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    updater_exe,
    a.binaries,
    a.datas,
    updater_a.binaries,
    updater_a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Y2订单处理辅助工具1.9',
)
