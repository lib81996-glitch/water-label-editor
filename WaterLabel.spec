# -*- mode: python ; coding: utf-8 -*-

analysis = Analysis(
    ["label_editor.pyw"],
    pathex=[],
    binaries=[],
    datas=[("assets", "assets"), ("waterlabel-app.ico", ".")],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(analysis.pure)
exe = EXE(
    pyz,
    analysis.scripts,
    analysis.binaries,
    analysis.datas,
    [],
    name="水洗标签编辑器",
    version="version_info.txt",
    icon="waterlabel-app.ico",
    debug=False,
    strip=False,
    upx=False,
    console=False,
)
