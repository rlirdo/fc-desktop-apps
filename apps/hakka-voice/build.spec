# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec — 客語語音小幫手 HakkaVoiceHelper 1.0.0

打包：
    pyinstaller --noconfirm build.spec
產物：
    Windows : dist/HakkaVoiceHelper.exe     （onefile、windowed）
    macOS   : dist/HakkaVoiceHelper.app     （onedir bundle）

語音辨識模型（約 0.8 GB）不包進 exe：放在 exe 旁 models\hakka-whisper-ct2-int8\。
不使用 VAD，排除 onnxruntime；也排除 torch 等大型套件，避免 exe 暴增。
"""
import os
import sys

from PyInstaller.utils.hooks import collect_dynamic_libs, collect_submodules

HERE = os.path.abspath(os.getcwd())
IS_MAC = sys.platform == "darwin"
NAME = "HakkaVoiceHelper"

icon_ico = os.path.join(HERE, "icon.ico")
icon_png = os.path.join(HERE, "icon.png")
if not IS_MAC and os.path.exists(icon_ico):
    icon_file = icon_ico
elif os.path.exists(icon_png):
    icon_file = icon_png
else:
    icon_file = None

datas = []
for f in ("icon.ico", "icon.png"):
    p = os.path.join(HERE, f)
    if os.path.exists(p):
        datas.append((p, "."))

binaries = collect_dynamic_libs("ctranslate2")
hidden = (["core_tts", "core_asr", "core_audio", "requests", "numpy", "tokenizers",
           "PIL.ImageGrab"]
          + collect_submodules("faster_whisper")
          + collect_submodules("ctranslate2", filter=lambda n: "converters" not in n)
          + collect_submodules("av"))

EXCLUDES = ["torch", "torchaudio", "torchvision", "onnxruntime", "tensorflow", "transformers",
            "scipy", "pandas", "matplotlib", "IPython", "jupyter", "notebook", "pytest",
            "sympy", "numba", "llvmlite", "cv2", "sklearn"]

a = Analysis(
    ["app.py"],
    pathex=[HERE],
    binaries=binaries,
    datas=datas,
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

if IS_MAC:
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name=NAME,
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,
        console=False,
        disable_windowed_traceback=False,
        argv_emulation=True,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon=icon_file,
    )
    coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name=NAME)
    app = BUNDLE(
        coll,
        name=NAME + ".app",
        icon=icon_file,
        bundle_identifier="tw.edu.ndhu.fcworkshop.hakkavoicehelper",
        info_plist={
            "CFBundleName": NAME,
            "CFBundleDisplayName": "客語語音小幫手",
            "CFBundleShortVersionString": "1.0.0",
            "NSHighResolutionCapable": True,
        },
    )
else:
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.datas,
        [],
        name=NAME,
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,
        runtime_tmpdir=None,
        console=False,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon=icon_file,
    )
