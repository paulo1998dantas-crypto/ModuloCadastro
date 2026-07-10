# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from PyInstaller.utils.hooks import collect_dynamic_libs


ROOT = Path.cwd()
TEMPLATE_DIR = ROOT / "_internal" / "templates"
TEMPLATE_FILES = [
    TEMPLATE_DIR / "cadastro_bancos.html",
    TEMPLATE_DIR / "opcoes.html",
    TEMPLATE_DIR / "planilha.html",
    TEMPLATE_DIR / "ponte.html",
]
DOWNLOAD_FILES = [
    ROOT / "local_bridge.py",
    ROOT / "iniciar_ponte_local.bat",
    ROOT / "ponte_config.example.json",
]
a = Analysis(
    ['desktop_launcher.py'],
    pathex=[],
    binaries=collect_dynamic_libs("pydantic_core"),
    datas=[(str(path), "_internal/templates") for path in TEMPLATE_FILES if path.exists()]
    + [(str(path), ".") for path in DOWNLOAD_FILES if path.exists()],
    hiddenimports=["pydantic_core._pydantic_core"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='ModuloCadastro',
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
)
