# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec — FC文獻缺口搜尋工具 FC LitGap Finder 1.0.1

打包：
    pyinstaller --noconfirm build.spec
產物：
    Windows : dist/FCLitGapFinder.exe   （onefile、windowed；交付時複製改名為 FC文獻缺口搜尋工具.exe）
    macOS   : dist/FCLitGapFinder.app   （onedir bundle；macOS 的 .app 不支援 onefile）

--selftest 在 console=False 下把結果寫到暫存資料夾的 FCLitGapFinder_selftest.log，
並以回傳碼（0/1）表示成功與否。
"""
import os
import sys

from PyInstaller.utils.hooks import collect_data_files

HERE = os.path.abspath(os.getcwd())
IS_MAC = sys.platform == "darwin"
NAME = "FCLitGapFinder"

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
_samples = os.path.join(HERE, "selftest", "samples.json")
if os.path.exists(_samples):
    datas.append((_samples, "selftest"))

# python-docx 範本（default.docx、頁首頁尾 xml）
datas += collect_data_files("docx")
# macOS：docx/parts/ 必須真實存在（見 docx_parts.keep 說明）
_keep = os.path.join(HERE, "docx_parts.keep")
if os.path.exists(_keep):
    datas.append((_keep, os.path.join("docx", "parts")))
# reportlab 內建字型與 rl_config 資料（無 Word 時的 PDF 後備）
datas += collect_data_files("reportlab")

hidden = ["openpyxl", "docx", "lxml", "lxml.etree", "lxml._elementpath", "requests", "sqlite3",
          "PIL.Image", "PIL.ImageDraw", "PIL.ImageFont", "reportlab.pdfbase.ttfonts",
          "reportlab.pdfbase.cidfonts", "reportlab.platypus"]
# core／ui 子模組明列（collect_submodules 需在隔離子程序 import，路徑不穩）
for _pkg in ("core", "ui"):
    for _f in sorted(os.listdir(os.path.join(HERE, _pkg))):
        if _f.endswith(".py"):
            hidden.append(_pkg if _f == "__init__.py" else f"{_pkg}.{_f[:-3]}")
if not IS_MAC:
    hidden += ["win32com", "win32com.client", "pythoncom", "pywintypes"]

a = Analysis(
    ["app.py"],
    pathex=[HERE],
    binaries=[],
    datas=datas,
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["numpy", "pandas", "matplotlib", "pytest", "IPython", "scipy"],
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
        argv_emulation=False,
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
        bundle_identifier="tw.edu.ndhu.fcworkshop.litgapfinder",
        info_plist={
            "CFBundleName": NAME,
            "CFBundleDisplayName": "FC文獻缺口搜尋工具",
            "CFBundleShortVersionString": "1.0.1",
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
