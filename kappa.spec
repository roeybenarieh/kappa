# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

# collect_all grabs data files, native binaries, and hidden submodules for
# packages that PyInstaller can't fully analyse statically.
datas_ort,   bins_ort,   hidden_ort   = collect_all('ortools')
datas_pyd,   bins_pyd,   hidden_pyd   = collect_all('pydantic')
datas_pydcore, bins_pydcore, hidden_pydcore = collect_all('pydantic_core')
datas_lxml,  bins_lxml,  hidden_lxml  = collect_all('lxml')

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=bins_ort + bins_pyd + bins_pydcore + bins_lxml,
    datas=datas_ort + datas_pyd + datas_pydcore + datas_lxml,
    hiddenimports=[
        'create_excel',          # imported dynamically inside btn_get_template
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
    ] + hidden_ort + hidden_pyd + hidden_pydcore + hidden_lxml,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Kappa',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,   # no terminal window — GUI-only app
    icon=None,       # replace with 'icon.ico' to add a custom icon
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Kappa',
)
