# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

# NOTE: PyInstaller executes .spec without defining __file__ in some contexts.
# Keep this path aligned with the project that is actually being built.
ROOT = Path(r"C:\Users\mokoh\document_generator_modern")

a = Analysis(
    [str(ROOT / 'app.py')],
    pathex=[str(ROOT)],
    binaries=[],
    datas=[
        (str(ROOT / 'ui'), 'ui'),
        (str(ROOT / 'coreclr.runtimeconfig.json'), '.'),
        (str(ROOT / 'assets' / 'fonts'), str(Path('assets') / 'fonts')),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Not used by InputStudio; excluding keeps build fast/small on dev machines
        "torch",
        "torchvision",
        "torchaudio",
        "tensorflow",
        "jax",
        "scipy",
        "numpy.f2py",
        "pandas",
        "matplotlib",
        "sympy",
        "pytest",
        "_pytest",
        "networkx",
        "sklearn",
        "cv2",
        "opencv",
        "IPython",
        "notebook",
        "jupyter",
        "PyQt5",
        "PySide6",
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='InputStudio',
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
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='InputStudio',
)
