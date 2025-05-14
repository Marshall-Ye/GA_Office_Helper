# ga_office_helper.spec  –  bundle whole _internal folder
from pathlib import Path
import sys
from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.building.build_main import Analysis, PYZ, EXE, COLLECT, Tree

APP_VERSION = "1.0.3"
DIST_NAME   = f"GA_office_helper_{APP_VERSION}"
HERE        = Path(sys.argv[0]).resolve().parent

a = Analysis(
    ["main_gui.py"],
    pathex=[],
    binaries=[],
    datas=[],                              # we’ll add Tree() later
    hiddenimports=collect_submodules("mini_updater"),
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz, a.scripts,
    exclude_binaries=True,
    name="GA Office Helper",
    icon=str(HERE / "_internal" / "GA_Logo.ico"),
    console=True,          # flip to False for release
)

#  ⬇️  Copy entire _internal folder into the dist folder
internal = Tree(str(HERE / "_internal"), prefix="_internal")

COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    internal,          #  ←  THIS must be present
    strip=False,
    upx=True,
    name=DIST_NAME,
)

