# dicom_receiver_pusher.spec
# =============================================================================
# PyInstaller build spec for dicom_receiver_pusher-7-4.py
#
# WHY THIS FILE EXISTS (read this before changing anything)
# -----------------------------------------------------------------------------
# Building with a bare `pyinstaller dicom_receiver_pusher-7-4.py` (or an
# equivalent .spec without the extra collection calls below) produces an EXE
# that silently loses DICOM decompression for some files, even though
# pydicom/pynetdicom/python-gdcm/pylibjpeg* are all listed in requirements.txt
# and installed. This is not a "missing package" problem -- the compiled code
# for gdcm and pylibjpeg IS inside the frozen EXE either way. It's a *plugin
# discovery* problem:
#
#   * pylibjpeg-libjpeg, pylibjpeg-openjpeg and pylibjpeg-rle each register
#     themselves as pixel-data decoder/encoder plugins via Python package
#     *entry points* (see each package's .dist-info/entry_points.txt). At
#     runtime, pylibjpeg (and pydicom's own pylibjpeg bridge) discovers them
#     by calling importlib.metadata.entry_points(...). PyInstaller does NOT
#     bundle a package's .dist-info metadata by default, so in a frozen build
#     that lookup silently returns zero plugins -- pylibjpeg *imports* fine,
#     it just can't find its own codecs. This was directly reproduced and
#     confirmed while diagnosing this issue: entry_points(group=
#     "pylibjpeg.pixel_data_decoders") returns 12 entries when run normally
#     and 0 in a naive frozen build.
#   * The concrete plugin modules (the packages install top-level modules
#     literally named `libjpeg`, `openjpeg`, `rle` -- not sub-modules of
#     `pylibjpeg`) are loaded dynamically via those entry points rather than
#     a static `import libjpeg` anywhere in the source PyInstaller's scanner
#     can see, so they need to be collected explicitly too.
#
# Because pydicom can fall back from pylibjpeg to GDCM (or vice versa)
# depending on the transfer syntax, this bug does NOT break every file --
# only the specific compressed formats/edge cases where only one backend can
# actually decode/encode it. That's exactly the "some images work, some
# don't" pattern reported for both sending and Query/Retrieve.
#
# The `--collect-all` / `--copy-metadata` calls below fix this by bundling
# the entry-point metadata AND the concrete plugin modules for gdcm and the
# whole pylibjpeg family. copy_metadata() alone is not enough (it fixes
# discovery but not the ModuleNotFoundError when the plugin is actually
# invoked); collect_all() alone is not enough either (it bundles the module
# but not the metadata entry_points() needs to find it) -- both are required
# together, which is why they're both listed for every one of these packages.
#
# HOW TO BUILD
# -----------------------------------------------------------------------------
#   pip install -r requirements.txt
#   pip install pyinstaller
#   pyinstaller dicom_receiver_pusher.spec
#
# Output lands in dist/DicomReceiverPusher/ (onedir build -- see note below
# about why onedir is recommended over --onefile for this app).
# =============================================================================

from PyInstaller.utils.hooks import collect_all, copy_metadata

block_cipher = None

# ----------------------------------------------------------------------------
# Packages that need BOTH their metadata (for importlib.metadata / entry_point
# discovery) AND a full collect (submodules + binaries + data) explicitly.
# These are exactly the packages NOT already covered by an automatic hook
# from pyinstaller-hooks-contrib (pydicom itself, customtkinter, pystray,
# tkinterdnd2, docx, cryptography, reportlab, etc. already have community
# hooks that ship with PyInstaller and need no manual entry here).
# ----------------------------------------------------------------------------
datas = []
binaries = []
hiddenimports = []

for pkg in [
    "gdcm",              # python-gdcm's importable module name
    "libjpeg",            # pylibjpeg-libjpeg's importable module name
    "openjpeg",           # pylibjpeg-openjpeg's importable module name
    "rle",                # pylibjpeg-rle's importable module name
]:
    try:
        d, b, h = collect_all(pkg)
        datas += d
        binaries += b
        hiddenimports += h
    except Exception:
        # Package not installed in this build environment -- that's a valid
        # configuration (all codec packages are optional per requirements.txt),
        # just skip collecting it. The app already degrades gracefully and
        # logs a startup warning when a codec plugin is missing.
        pass

for dist_name in [
    "python-gdcm",
    "pylibjpeg",
    "pylibjpeg-libjpeg",
    "pylibjpeg-openjpeg",
    "pylibjpeg-rle",
    "pydicom",            # defensive: pydicom itself also exposes entry points
]:
    try:
        datas += copy_metadata(dist_name)
    except Exception:
        pass

# tkinterdnd2 ships a native tkdnd library as package data; the community
# hook for it (hook-tkinterdnd2.py, part of pyinstaller-hooks-contrib) should
# already handle this automatically, but it's optional in this app (drag-
# and-drop degrades to the file-picker button if unavailable) so we collect
# it defensively too in case the installed hooks-contrib version is older.
try:
    d, b, h = collect_all("tkinterdnd2")
    datas += d
    binaries += b
    hiddenimports += h
except Exception:
    pass

# Optional non-DICOM feature dependencies. Most of these already have
# community PyInstaller hooks (customtkinter, pystray, docx, cryptography,
# reportlab, ldap3), but listing the pure-Python ones as hidden imports here
# is cheap insurance against a hook-less older PyInstaller install silently
# dropping a lazily-imported optional module (see each try/except ImportError
# block near the top of the .py for how the app treats each as optional).
hiddenimports += [
    "psutil",
    "pyzipper",
    "ldap3",
    "matplotlib",
    "reportlab",
]

# The app's own local module (alongside the main script) providing the
# embedded toolbar/menu icon set. A plain top-level import, so PyInstaller's
# static analysis bundles it automatically -- listed here only as insurance;
# remove if you don't ship lucide_icons_data.py next to the main script.
hiddenimports += ["lucide_icons_data"]

# App's own visual assets. logo.png is optional -- _load_logo_pil() in the
# app returns None silently if it's missing, so this is a "nice to have",
# not a correctness requirement. Only include the add_data entry if the
# file actually exists next to this spec.
import os
extra_datas = []
if os.path.exists("logo.png"):
    extra_datas.append(("logo.png", "."))

a = Analysis(
    ["dicom_receiver_pusher-7-4.py"],
    pathex=[],
    binaries=binaries,
    datas=datas + extra_datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="DicomReceiverPusher",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,             # UPX-compressing native codec DLLs (gdcm/openjpeg)
                            # has a history of both false-positive AV flags
                            # and occasional runtime breakage -- left off.
    console=False,          # GUI app -- no console window.
    icon="icon.ico" if os.path.exists("icon.ico") else None,
)

# onedir (COLLECT), not --onefile: --onefile extracts the whole bundle to a
# temp directory on every launch, which measurably slows startup for an app
# this size and has caused real-world issues with antivirus software
# quarantining/scanning the temp extraction on every run. onedir also makes
# it far easier to verify what actually got bundled (see the "verifying the
# build" note below) since you can just browse dist/DicomReceiverPusher/.
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name="DicomReceiverPusher",
)

# -----------------------------------------------------------------------------
# VERIFYING THE BUILD ACTUALLY BUNDLED THE CODECS
# -----------------------------------------------------------------------------
# After building, confirm the fix worked before relying on it:
#
#   dist\DicomReceiverPusher\DicomReceiverPusher.exe
#
# then, in the app, check Admin > Settings (or wherever CODECS_FULLY_AVAILABLE
# / MISSING_CODEC_PLUGINS is surfaced) for codec warnings, and try opening/
# pushing a study that previously failed. You can also sanity-check the
# bundle contents directly without launching the GUI:
#
#   dir /s dist\DicomReceiverPusher\_internal\*pylibjpeg*
#   dir /s dist\DicomReceiverPusher\_internal\*dist-info*openjpeg*
#
# You should see both the compiled module directories (libjpeg/, openjpeg/,
# rle/) AND their *.dist-info folders. If the .dist-info folders are missing,
# entry-point discovery will still fail even though the modules are present.
# =============================================================================
