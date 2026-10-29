from pathlib import Path
project_dir = Path(SPEC).resolve().parent
vendor_dir = project_dir / "vendor"

pywin32_binaries = [
    (str(vendor_dir / "pywin32_system32" / "pythoncom312.dll"), "vendor/pywin32_system32"),
    (str(vendor_dir / "pywin32_system32" / "pywintypes312.dll"), "vendor/pywin32_system32"),
    (str(vendor_dir / "win32" / "_win32sysloader.pyd"), "vendor/win32"),
]

a = Analysis(
    [str(project_dir / "rm_tracker_app.py")],
    pathex=[str(project_dir), str(vendor_dir), str(vendor_dir / "win32"), str(vendor_dir / "win32com"), str(vendor_dir / "win32comext")],
    binaries=pywin32_binaries,
    datas=[
        (str(project_dir / "app_logo.png"), "."),
        (str(project_dir / "app_icon.ico"), "."),
        (str(vendor_dir / "pythoncom.py"), "vendor"),
        (str(vendor_dir / "win32"), "vendor/win32"),
        (str(vendor_dir / "win32" / "lib"), "vendor/win32/lib"),
        (str(vendor_dir / "win32com"), "vendor/win32com"),
        (str(vendor_dir / "win32comext"), "vendor/win32comext"),
    ],
    hiddenimports=[
        "pythoncom", "pywintypes", "win32com", "win32com.client",
    ],    hookspath=[],
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
    name="RM_FollowUp_Automation_Pivot",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=str(project_dir / "app_icon.ico"),
)

