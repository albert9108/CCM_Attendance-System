# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['attendance system.py'],
    pathex=[r'C:\Users\alber\OneDrive - mmu.edu.my\Church\CCM_Attendance System\installer\2_install_app\attendance system final 1.7'],
    binaries=[
        ('libiconv.dll', '.'),
        ('libzbar-64.dll', '.')
    ],
    datas=[],
    hiddenimports=[],
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
    a.binaries,
    a.datas,
    [],
    name='CCM Attend System',
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
