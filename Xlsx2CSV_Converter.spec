# -*- mode: python ; coding: utf-8 -*-
"""Spec PyInstaller tối giản cho GUI SoilData2AI.

Chỉ đóng gói: PySide6 (QtCore/QtGui/QtWidgets), openpyxl, xlrd, xlsx2csv,
normalize và assets. Loại trừ các thư viện lớn không liên quan trong máy
(torch, pandas, scipy, cv2, numba, ...) để EXE nhẹ và build nhanh.
"""
hiddenimports = [
    "xlsx2csv",
    "normalize",
    "openpyxl",
    "xlrd",
]

excludes = [
    # ML / data science không dùng trong bài
    "torch", "torchvision", "torchaudio", "pandas", "numpy", "scipy",
    "sklearn", "cv2", "numba", "llvmlite", "av", "huggingface_hub",
    "transformers", "datasets", "tokenizers", "sentencepiece", "accelerate",
    "matplotlib", "seaborn", "plotly", "pyarrow", "polars", "h5py", "tables",
    # Notebook / dev tools
    "IPython", "jupyter", "notebook", "nbformat", "nbconvert", "pytest",
    "sphinx", "docutils", "setuptools", "pip", "wheel", "pkg_resources",
    "distutils", "tkinter", "unittest", "doctest",
    # GUI framework khác
    "PyQt5", "PyQt6", "PySide2",
    # Web / DB / cloud không dùng
    "requests", "urllib3", "certifi", "charset_normalizer", "idna",
    "boto3", "botocore", "s3transfer", "sqlalchemy", "pymysql",
    "psycopg2", "mysql", "redis", "flask", "django", "fastapi",
]

a = Analysis(
    ["src/gui.py"],
    pathex=["src"],
    binaries=[],
    datas=[("assets", "assets")],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Xlsx2CSV_Converter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="assets/app.ico",
)
