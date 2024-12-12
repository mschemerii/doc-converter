# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['doc_converter_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('*.py', '.'),
        ('README.md', '.'),
        ('README.pdf', '.')
    ],
    hiddenimports=[
        'docx',
        'tkinter',
        'Foundation',
        'AppKit',
        'ScriptingBridge'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'win32com',
        'pythoncom',
        'pypandoc'
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher
)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='doc-converter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    target_arch='x86_64',
    codesign_identity=None,
    entitlements_file='macos_entitlements.xml'
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='doc-converter'
)

app = BUNDLE(
    coll,
    name='doc-converter.app',
    icon=None,
    bundle_identifier='com.oracle.doc-converter',
    info_plist={
        'NSHighResolutionCapable': 'True',
        'LSMinimumSystemVersion': '10.15',
        'NSAppleEventsUsageDescription': 'This app needs to control Microsoft Word to convert documents.',
        'CFBundleShortVersionString': '1.0.0',
        'CFBundleVersion': '1.0.0'
    }
) 