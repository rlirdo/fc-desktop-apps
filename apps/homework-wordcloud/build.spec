# -*- mode: python ; coding: utf-8 -*-
"""
build.spec — PyInstaller 打包設定（單一執行檔）

    pyinstaller build.spec

產出：dist\\HomeworkWordCloud.exe（Windows）／dist\\HomeworkWordCloud.app（macOS）

要點：
  * onefile + console=False（雙擊不跳黑視窗）；--selftest 仍會回傳 exit code，
    並把完整紀錄寫到系統暫存資料夾的 HomeworkWordCloud_selftest.log。
  * jieba 的 dict.txt 在套件內，用 collect_data_files('jieba') 帶進來
    （lac_small 是 paddle 模型，用不到又佔 13MB，剔除）。
    程式端會用 sys._MEIPASS 明確 jieba.set_dictionary()，不依賴 pkg_resources。
  * wordcloud 的 stopwords 與 DroidSansMono.ttf 同樣用 collect_data_files 帶進來
    （找不到系統中文字型時的最後保險字型）。
  * 中文字型一律用「使用者電腦上的系統字型」，不打包微軟正黑體（授權限制）。
"""
import os
import sys

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

HERE = os.path.abspath(SPECPATH)          # noqa: F821  (SPECPATH 由 PyInstaller 提供)

# ---------------------------------------------------------------- 資料檔
datas = [
    (os.path.join(HERE, 'config.json'), '.'),
    (os.path.join(HERE, 'assets', 'logo_placeholder.png'), 'assets'),
    (os.path.join(HERE, 'icon.ico'), '.'),
]
for fn in sorted(os.listdir(os.path.join(HERE, 'sample_input'))):
    if fn.lower().endswith(('.xlsx', '.csv')):
        datas.append((os.path.join(HERE, 'sample_input', fn), 'sample_input'))

# selftest 用的合成名單 PDF（虛構姓名／學號），--selftest 的端對端 PDF 解析要用
_st = os.path.join(HERE, 'selftest', '樣本_選課名單_合成.pdf')
if os.path.exists(_st):
    datas.append((_st, 'selftest'))

# jieba：只要 dict.txt 等字典檔，跳過 lac_small 的 paddle 模型
datas += [(src, dst) for src, dst in collect_data_files('jieba')
          if 'lac_small' not in src.replace('/', os.sep).split(os.sep)]

# wordcloud：stopwords ＋ 內建 DroidSansMono.ttf
datas += collect_data_files('wordcloud')

# python-pptx：templates/（default.pptx、notesMaster.xml…）在 macOS 端不會被 hook 自動帶入，
# 無條件收集，否則 notes_slide 會 FileNotFoundError（CI macos-latest 實測）。
datas += collect_data_files('pptx')
# 再以實際檔案系統逐檔加入（雙保險），並印出數量供 CI 記錄比對
import pptx as _pptx
_tpl = os.path.join(os.path.dirname(_pptx.__file__), 'templates')
_tpl_files = [(os.path.join(_tpl, f), os.path.join('pptx', 'templates')) for f in os.listdir(_tpl)]
datas += _tpl_files
print('[spec] pptx templates 逐檔加入 %d 個：%s' % (len(_tpl_files), sorted(os.listdir(_tpl))))
# 關鍵（macOS）：python-pptx 用 pptx/oxml/../templates/x.xml 讀範本；純 Python 模組在 PYZ 內，
# 磁碟上沒有 pptx/oxml/ 目錄，POSIX 逐層解析 ".." 會 ENOENT（Windows 會先字串正規化所以沒事）。
# 放一個佔位檔讓 pptx/oxml/ 真實存在。
datas.append((os.path.join(HERE, 'pptx_oxml.keep'), os.path.join('pptx', 'oxml')))

# ---------------------------------------------------------------- 模組
hiddenimports = [
    'jieba', 'jieba.finalseg', 'jieba.finalseg.prob_start',
    'jieba.finalseg.prob_trans', 'jieba.finalseg.prob_emit',
    'wordcloud', 'wordcloud.query_integral_image',
    'PIL.Image', 'PIL.ImageDraw', 'PIL.ImageFont', 'PIL.ImageColor',
    'numpy', 'openpyxl', 'pptx',
    # 2.1：東華教務系統選課名單 PDF 的名單解析（純 Python，不用 PyMuPDF）
    'pypdf', 'pypdf.generic', 'pypdf._page', 'pypdf._reader',
]
hiddenimports += collect_submodules('core')
hiddenimports += collect_submodules('pypdf')

excludes = [
    # 注意：不要排除 unittest／test，pyparsing（matplotlib 相依）會 import 它們
    'tkinter.test', 'pydoc_data',
    'pandas', 'scipy', 'IPython', 'jupyter', 'notebook',
    'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'wx',
    'matplotlib.backends.backend_qt5agg', 'matplotlib.backends.backend_qtagg',
    'matplotlib.backends.backend_webagg', 'matplotlib.backends.backend_gtk3agg',
]

a = Analysis(                                                     # noqa: F821
    [os.path.join(HERE, 'app.py')],
    pathex=[HERE],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)                                                 # noqa: F821

icon = os.path.join(HERE, 'icon.png' if sys.platform == 'darwin' else 'icon.ico')

exe = EXE(                                                        # noqa: F821
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='HomeworkWordCloud',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon,
)

if sys.platform == 'darwin':
    app = BUNDLE(                                                 # noqa: F821
        exe,
        name='HomeworkWordCloud.app',
        icon=icon,
        bundle_identifier='tw.edu.ndhu.fc.homeworkwordcloud',
        info_plist={
            'CFBundleName': 'HomeworkWordCloud',
            'CFBundleDisplayName': '學生作業文字雲',
            'CFBundleShortVersionString': '2.2.0',
            'NSHighResolutionCapable': True,
        },
    )
