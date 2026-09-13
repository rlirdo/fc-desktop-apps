# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec — 仰望社團管理系統 ClubManager

打包：
    pyinstaller build.spec
產物：
    Windows : dist/ClubManager.exe        （onefile、windowed）
    macOS   : dist/ClubManager.app        （onedir bundle；macOS 的 .app 不支援 onefile）

web/（index.html、schools/ 九校版與清單、docs/）整包進 datas，執行期以
sys._MEIPASS 定位；HTML 一字未改。

--selftest 在 console=False 下仍會把結果寫到暫存資料夾的
ClubManager_selftest.log，並以回傳碼（0/1）表示成功與否；
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

# web/ 整個資料夾（含 schools/ 與 docs/）原封不動帶進封裝
datas = [(os.path.join(HERE, "web"), "web")]
for f in ("icon.ico", "icon.png"):
    p = os.path.join(HERE, f)
    if os.path.exists(p):
        datas.append((p, "."))

a = Analysis(
    ["app.py"],
    pathex=[HERE],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["PIL", "numpy", "pandas", "matplotlib", "openpyxl", "pytest",
              "pptx", "jieba", "wordcloud"],
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
        name="ClubManager",
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
        icon=icon_file,
    )
    coll = COLLECT(
        exe,
        a.binaries,
        a.datas,
        strip=False,
        upx=False,
        name="ClubManager",
    )
    app = BUNDLE(
        coll,
        name="ClubManager.app",
        icon=icon_file,
        bundle_identifier="tw.edu.ndhu.fcworkshop.clubmanager",
        info_plist={
            "CFBundleName": "ClubManager",
            "CFBundleDisplayName": "仰望社團管理系統",
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
        name="ClubManager",
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
