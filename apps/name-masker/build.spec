# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec — 姓名遮罩與學生編號 NameMasker 2.2

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

from PyInstaller.utils.hooks import collect_data_files

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
for f in ("樣本_Zuvio合成資料.xlsx", "樣本_名冊合成資料.xlsx", "樣本_通用格式.csv",
          "樣本_原始名單_含序號.xlsx", "樣本_原始名單_無序號.xlsx", "樣本_原始名單.csv",
          "樣本_選課名單_合成.pdf", "樣本_選課名單_合成_黏連格式.pdf"):
    sample = os.path.join(HERE, "selftest", f)
    if os.path.exists(sample):
        datas.append((sample, "selftest"))

# python-docx（2.2）：templates/ 內的 default.docx、default-header.xml… 要一起帶走
datas += collect_data_files("docx")
# 關鍵（macOS）：python-docx 用 docx/parts/../templates/x.xml 讀範本；純 Python 模組在 PYZ 內，
# 磁碟上沒有 docx/parts/ 目錄，POSIX 逐層解析 ".." 會 ENOENT（Windows 會先字串正規化所以沒事）。
# 放一個佔位檔讓 docx/parts/ 真實存在（只有 parts/ 用到 ".."，opc/oxml 沒有）。
_keep = os.path.join(HERE, "docx_parts.keep")
if os.path.exists(_keep):
    datas.append((_keep, os.path.join("docx", "parts")))

a = Analysis(
    ["app.py"],
    pathex=[HERE],
    binaries=[],
    datas=datas,
    hiddenimports=["openpyxl", "sample_data", "core_mask", "roster", "doc_mask",
                   "pypdf", "docx", "lxml", "lxml.etree", "lxml._elementpath"],
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
            "CFBundleDisplayName": "姓名遮罩與學生編號",
            "CFBundleShortVersionString": "2.2.0",
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
