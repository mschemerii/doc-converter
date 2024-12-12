# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['doc_converter_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('process_document.py', '.'),
        ('doc_to_docx_converter.py', '.'),
        ('python_version_check.py', '.'),
        ('modify_docx_tables.py', '.'),
        ('requirements.txt', '.'),
    ],
    hiddenimports=[
        'tkinter',
        'process_document',
        'subprocess',
        'threading',
        'logging',
        'pypandoc',
        'docx',
        'lxml',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)

def remove_duplicates(list_of_tuples):
    seen = set()
    return [x for x in list_of_tuples if not (x[0] in seen or seen.add(x[0]))]

a.datas = remove_duplicates(a.datas)
a.binaries = remove_duplicates(a.binaries)

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
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
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
    bundle_identifier='com.yourcompany.docconverter',
    info_plist={
        'LSEnvironment': {
            'PYTHONPATH': '@executable_path/../Resources:@executable_path/../Frameworks',
            'PATH': '/usr/local/bin:/usr/bin:/bin:@executable_path/../Resources'
        },
        'CFBundleDisplayName': 'Doc Converter',
        'CFBundleName': 'Doc Converter',
        'CFBundleExecutable': 'doc-converter',
        'CFBundleIdentifier': 'com.yourcompany.docconverter',
        'CFBundleShortVersionString': '1.0.0',
        'NSHighResolutionCapable': True,
        'NSPrincipalClass': 'NSApplication',
    }
) 