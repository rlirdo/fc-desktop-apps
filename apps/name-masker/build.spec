# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec — 姓名遮罩 NameMasker

打包：
    pyinstaller build.spec
產物：
    Windows : dist/NameMasker.exe        （onefile、windowed）
    macOS   : dist/NameMasker.app        （onedir bundle；macOS 的 .app 不支援 onefile）

--selftest 在 console=False 下仍會把結果寫到暫存資料夾的
NameMasker_selftest.log，並以回傳碼（0/1）表示成功與否；
若執行環境有主控台（例如從 cmd 重導向）也會直接印出。
"""
import os
import sys

HERE = os.path.abspath(os.getcwd())
IS_MAC = sys.platform == "darwin"

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
sample = os.path.join(HERE, "selftest", "樣本_合成資料.xlsx")
if os.path.exists(sample):
    datas.append((sample, "selftest"))

a = Analysis(
    ["app.py"],
    pathex=[HERE],
    binaries=[],
    datas=datas,
    hiddenimports=["openpyxl", "sample_data", "mask_names"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["PIL", "numpy", "pandas", "matplotlib", "pytest"],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

if IS_MAC:
    # macOS：onedir + .app bundle（PyInstaller 不建議 onefile 搭配 .app）
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name="NameMasker",
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,
        console=False,
        disable_windowed_traceback=False,
        argv_emulation=True,        # 讓 Finder 的「開啟檔案」事件變成 argv
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon=icon_file,
    )
    coll = COLLECT(
        exe,
        a.binaries,
        a.datas,
        strip=False,
        upx=False,
        name="NameMasker",
    )
    app = BUNDLE(
        coll,
        name="NameMasker.app",
        icon=icon_file,
        bundle_identifier="tw.edu.ndhu.fcworkshop.namemasker",
        info_plist={
            "CFBundleName": "NameMasker",
            "CFBundleDisplayName": "姓名遮罩",
            "CFBundleShortVersionString": "1.0.0",
            "NSHighResolutionCapable": True,
        },
    )
else:
    # Windows：單一 exe，方便把 Excel 檔拖到圖示上執行
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.datas,
        [],
        name="NameMasker",
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
