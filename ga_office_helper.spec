# ga_office_helper.spec  –  bundle whole _internal folder
# -------------------------------------------------------
from pathlib import Path

from PyInstaller.utils.hooks        import collect_submodules
from PyInstaller.building.build_main  import Analysis, PYZ, EXE, COLLECT
from PyInstaller.building.datastruct import Tree      # Tree lives here
import sys

APP_VERSION = "1.1.0"
DIST_NAME   = f"GA_office_helper_{APP_VERSION}"

SRC_DIR = Path(sys.argv[0]).resolve().parent          # ← fix here
INT_DIR = SRC_DIR / "_internal"                       # logo, templates …

# ── ANALYSIS ────────────────────────────────────────────────
a = Analysis(
    ["main_gui.py"],
    pathex=[str(SRC_DIR)],
    binaries=[],
    datas=[],
    hiddenimports=collect_submodules("mini_updater"),
)

pyz = PYZ(a.pure, a.zipped_data)

# ── EXE (launcher) ──────────────────────────────────────────
exe = EXE(
    pyz, a.scripts,
    exclude_binaries=True,
    name="GA Office Helper",
    icon=str(INT_DIR / "GA_Logo.ico"),
    console=False,                 # flip to True while debugging
)

# ── include entire _internal folder ------------------------
payload = Tree(str(INT_DIR), prefix="_internal")

# ── COLLECT  (final dist folder) ───────────────────────────
COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    payload,                       # full _internal copied over
    strip=False,
    upx=True,
    name=DIST_NAME,
)
