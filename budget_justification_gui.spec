# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['budget_justification_gui.py'],
    pathex=[],
    binaries=[
        ('bundled_bin/pandoc-arm64', '.'),
        ('bundled_bin/pandoc-x86_64', '.'),
    ],
    datas=[
        ('OSP-3-Year-Budget-Template-August-2025.xlsx', '.'),
        ('OSP-5-Year-Budget-Template-August-2025.xlsx', '.'),
        ('OSP-10-Year-Budget-Template-August-2025.xlsx', '.'),
    ],
    hiddenimports=[
        'openpyxl',
        'openpyxl.cell',
        'openpyxl.cell._writer',
        'openpyxl.styles',
        'docx',
        'docx.shared',
        'docx.enum.text',
        'tkinter',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.ttk',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# Use onedir mode for proper code signing
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Budget Justification Generator',
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
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Budget Justification Generator',
)

app = BUNDLE(
    coll,
    name='Budget Justification Generator.app',
    icon=None,
    bundle_identifier='edu.nau.budgetjustification',
    info_plist={
        'NSPrincipalClass': 'NSApplication',
        'NSHighResolutionCapable': 'True',
        'CFBundleName': 'Budget Justification Generator',
        'CFBundleDisplayName': 'Budget Justification Generator',
        'CFBundleGetInfoString': 'NAU Budget Justification Generator',
        'CFBundleVersion': '1.0.0',
        'CFBundleShortVersionString': '1.0.0',
    },
)
