# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['lector.py'],
    pathex=['C:\\Users\\Cufa\\Documents\\Proyectos\\LibreriaMagic'],
    binaries=[],
    datas=[('C:\\Users\\Cufa\\Documents\\Proyectos\\LibreriaMagic\\precios.xlsx', 'precios.xlsx')],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Calculadora PDF V2.1',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon='C:\\Users\\Cufa\\Documents\\Proyectos\\LibreriaMagic\\calculator-icon_34473.ico'
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Calculadora PDF V2.1',
)
