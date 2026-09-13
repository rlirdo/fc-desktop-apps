# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec：週報產生器（WeeklyReportMaker）
  Windows :  pyinstaller build.spec   →  dist\WeeklyReportMaker.exe
  macOS   :  pyinstaller build.spec   →  dist/WeeklyReportMaker.app

onefile＋GUI（console=False）。--selftest 在 GUI 模式下仍然可用：
程式會接上呼叫端的主控台印出結果，同時把記錄寫到暫存資料夾的
WeeklyReportMaker_selftest.log，並用回傳碼（0 成功 / 1 失敗）表示結果。

注意：不打包任何字型檔（微軟正黑體有授權限制）；字型只寫名稱進 pptx。
"""
import os
import sys

APP = "WeeklyReportMaker"
HERE = os.path.abspath(SPECPATH)

datas = [
    (os.path.join(HERE, "selftest", "sample_data.json"), "selftest"),
    (os.path.join(HERE, "icon.png"), "."),
]
if os.path.exists(os.path.join(HERE, "icon.ico")):
    datas.append((os.path.join(HERE, "icon.ico"), "."))
# python-pptx 的 templates/（default.pptx、notesMaster.xml…）在 macOS 端不會被 hook 自動帶入，
# 無條件收集，否則 notes_slide 會 FileNotFoundError（CI macos-latest 實測）。
from PyInstaller.utils.hooks import collect_data_files
datas += collect_data_files("pptx")
# 再以實際檔案系統逐檔加入（雙保險），並印出數量供 CI 記錄比對
import pptx as _pptx
_tpl = os.path.join(os.path.dirname(_pptx.__file__), "templates")
_tpl_files = [(os.path.join(_tpl, f), os.path.join("pptx", "templates")) for f in os.listdir(_tpl)]
datas += _tpl_files
print("[spec] pptx templates 逐檔加入 %d 個：%s" % (len(_tpl_files), sorted(os.listdir(_tpl))))
# 關鍵（macOS）：python-pptx 用 pptx/oxml/../templates/x.xml 讀範本；純 Python 模組在 PYZ 內，
# 磁碟上沒有 pptx/oxml/ 目錄，POSIX 逐層解析 ".." 會 ENOENT（Windows 會先字串正規化所以沒事）。
# 放一個佔位檔讓 pptx/oxml/ 真實存在。
datas.append((os.path.join(HERE, "pptx_oxml.keep"), os.path.join("pptx", "oxml")))

a = Analysis(
    [os.path.join(HERE, "app.py")],
    pathex=[HERE],
    binaries=[],
    datas=datas,
    hiddenimports=["core", "report_builder", "validator"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["numpy", "pandas", "matplotlib", "scipy", "pytest", "setuptools",
              "IPython", "PyQt5", "PySide6", "tkinter.test", "test"],
    noarchive=False,
)
pyz = PYZ(a.pure)

icon = os.path.join(HERE, "icon.ico" if sys.platform.startswith("win") else "icon.png")

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=APP,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,              # GUI 模式；--selftest 仍會接上父行程的主控台
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon,
)

if sys.platform == "darwin":
    app = BUNDLE(
        exe,
        name=APP + ".app",
        icon=os.path.join(HERE, "icon.png"),
        bundle_identifier="tw.edu.ndhu.biomimicry.weeklyreportmaker",
        info_plist={
            "CFBundleName": APP,
            "CFBundleDisplayName": "週報產生器",
            "CFBundleShortVersionString": "1.0.0",
            "NSHighResolutionCapable": True,
        },
    )
