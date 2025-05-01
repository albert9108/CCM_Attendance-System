# -*- mode: python ; coding: utf-8 -*-

# Create version info using a dictionary instead of the Version class
version_info = {
    'version': '1.1',
    'file_description': 'CCM Attendance System',
    'internal_name': 'CCMAttendanceSystem',
    'legal_copyright': '© 2025',
    'original_filename': 'CCM Attendance System.exe',
    'product_name': 'CCM Attendance System',
    'product_version': '1.1'
}
a = Analysis(
    ['improved_attendance_system.py'],
    pathex=[r'C:\Users\alber\OneDrive\attendance system final 1.7'],
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
    name='CCM Attedance System',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='ccmlogo_PaE_icon.ico',
    version_info = version_info
)
