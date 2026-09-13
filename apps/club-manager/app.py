#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""仰望社團管理系統（ClubManager）——桌面版啟動器。

把既有的純前端網頁系統（index.html + schools/ 九校版 + docs/）包成一支
執行檔：雙擊後在本機起一個只綁 127.0.0.1 的小型 HTTP 伺服器，再用系統
預設瀏覽器打開，並保留一個 tkinter 控制視窗負責開啟／關閉。

HTML 一字未改，全部原封不動放在 web/ 之下（打包時進 datas，執行期用
sys._MEIPASS 定位）。資料仍然存在使用者瀏覽器的 localStorage。

CLI：
    app.py              啟動 GUI（預設行為）
    app.py --selftest   無 GUI 自我測試（不開瀏覽器），成功印 SELFTEST OK 並 exit 0
    app.py --version    印出版本

環境變數：
    CLUBMANAGER_NO_BROWSER=1   測試模式：啟動 GUI 時不自動開瀏覽器，
                               偵測到系統已在執行時也不跳對話框（直接結束）
"""
from __future__ import annotations

import functools
import http.server
import os
import socket
import subprocess
import sys
import tempfile
import threading
import urllib.error
import urllib.request

APP_NAME = "ClubManager"
APP_TITLE = "仰望社團管理系統"
VERSION = "1.0.0"

HOST = "127.0.0.1"
DEFAULT_PORT = 8765
PORT_TRIES = 10                      # 8765 ~ 8774
PING_PATH = "/__clubmanager_ping"
PING_BODY = "ClubManager OK"

BACKUP_DIR_NAME = "ClubManager_備份"

SCHOOL_PAGE_COUNT = 9


# --------------------------------------------------------------------------
# 基礎工具
# --------------------------------------------------------------------------
def safe_print(*args) -> None:
    """在 console=False 的封裝版本裡 sys.stdout 可能是 None，印字不得讓程式炸掉。"""
    try:
        if sys.stdout is None:
            return
        print(*args)
        sys.stdout.flush()
    except Exception:
        pass


def resource_root() -> str:
    """回傳 web/ 的絕對路徑（原始碼執行與 PyInstaller onefile 皆適用）。"""
    base = getattr(sys, "_MEIPASS", None)
    if not base:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "web")


def backup_dir() -> str:
    """~/Desktop/ClubManager_備份（桌面不存在時退回家目錄；家目錄也查不到才用暫存區）。"""
    home = os.path.expanduser("~")
    if home == "~" or not os.path.isdir(home):
        # 極簡環境（例如清掉環境變數的測試）查不到家目錄，不要在工作目錄亂建資料夾
        home = tempfile.gettempdir()
    desktop = os.path.join(home, "Desktop")
    parent = desktop if os.path.isdir(desktop) else home
    return os.path.join(parent, BACKUP_DIR_NAME)


def ensure_backup_dir() -> str:
    path = backup_dir()
    try:
        os.makedirs(path, exist_ok=True)
    except Exception:
        pass
    return path


def open_folder(path: str) -> None:
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa: S606  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


def school_pages(web_root: str) -> list[str]:
    """回傳 schools/ 下 9 校頁面的檔名（不含 index.html），依檔名排序。"""
    folder = os.path.join(web_root, "schools")
    names = [
        f for f in os.listdir(folder)
        if f.lower().endswith(".html") and f.lower() != "index.html"
    ]
    return sorted(names)


# --------------------------------------------------------------------------
# HTTP 伺服器
# --------------------------------------------------------------------------
class ClubHandler(http.server.SimpleHTTPRequestHandler):
    """只服務 web/ 底下的靜態檔，只回應本機，HTML 一律 utf-8。"""

    server_version = "ClubManager/" + VERSION
    sys_version = ""

    # --- 不要把每個請求都印到 stdout（封裝版沒有 console） ---
    def log_message(self, fmt, *args):  # noqa: A003
        return

    # --- Content-Type：中文網頁一定要帶 charset=utf-8 ---
    def guess_type(self, path):
        ext = os.path.splitext(str(path))[1].lower()
        table = {
            ".html": "text/html; charset=utf-8",
            ".htm": "text/html; charset=utf-8",
            ".md": "text/plain; charset=utf-8",
            ".txt": "text/plain; charset=utf-8",
            ".css": "text/css; charset=utf-8",
            ".js": "text/javascript; charset=utf-8",
            ".json": "application/json; charset=utf-8",
            ".svg": "image/svg+xml",
        }
        if ext in table:
            return table[ext]
        return super().guess_type(path)

    # --- 禁止跳出 web/ 之外 ---
    def translate_path(self, path):
        resolved = super().translate_path(path)
        root = os.path.realpath(self.directory)
        target = os.path.realpath(resolved)
        if target != root and not target.startswith(root + os.sep):
            # 指向一個必定不存在的名字 → 交給標準流程回 404
            return os.path.join(root, "__forbidden__")
        return resolved

    # --- 僅限本機 ---
    def _is_local(self) -> bool:
        host = self.client_address[0] if self.client_address else ""
        return host in ("127.0.0.1", "::1", "::ffff:127.0.0.1")

    def _handle_ping(self) -> bool:
        if self.path.split("?", 1)[0] != PING_PATH:
            return False
        body = PING_BODY.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)
        return True

    def do_GET(self):  # noqa: N802
        if not self._is_local():
            self.send_error(403, "Forbidden")
            return
        if self._handle_ping():
            return
        super().do_GET()

    def do_HEAD(self):  # noqa: N802
        if not self._is_local():
            self.send_error(403, "Forbidden")
            return
        if self._handle_ping():
            return
        super().do_HEAD()


class LocalServer:
    """包住 ThreadingHTTPServer，負責選埠、背景執行與關閉。"""

    def __init__(self, web_root: str):
        self.web_root = web_root
        self.httpd: http.server.ThreadingHTTPServer | None = None
        self.thread: threading.Thread | None = None
        self.port: int = 0
        self.port_changed: bool = False

    def start(self, port: int | None = None, tries: int = PORT_TRIES) -> int:
        handler = functools.partial(ClubHandler, directory=self.web_root)
        first = DEFAULT_PORT if port is None else port
        candidates = [0] if first == 0 else [first + i for i in range(tries)]
        last_err: Exception | None = None
        for p in candidates:
            try:
                httpd = http.server.ThreadingHTTPServer((HOST, p), handler)
            except OSError as exc:
                last_err = exc
                continue
            httpd.daemon_threads = True
            self.httpd = httpd
            self.port = httpd.server_address[1]
            self.port_changed = (first != 0 and self.port != first)
            self.thread = threading.Thread(
                target=httpd.serve_forever, kwargs={"poll_interval": 0.2}, daemon=True
            )
            self.thread.start()
            return self.port
        raise OSError(f"找不到可用的埠（{candidates[0]}～{candidates[-1]}）：{last_err}")

    def stop(self) -> None:
        if self.httpd is not None:
            try:
                self.httpd.shutdown()
            except Exception:
                pass
            try:
                self.httpd.server_close()
            except Exception:
                pass
        self.httpd = None
        if self.thread is not None:
            self.thread.join(timeout=3)
        self.thread = None

    def url(self, path: str = "/index.html") -> str:
        if not path.startswith("/"):
            path = "/" + path
        return f"http://{HOST}:{self.port}{path}"


def ping(port: int, timeout: float = 0.8) -> bool:
    """檢查該埠是不是「本程式」已經在服務。"""
    try:
        with urllib.request.urlopen(
            f"http://{HOST}:{port}{PING_PATH}", timeout=timeout
        ) as resp:
            return resp.status == 200 and PING_BODY in resp.read().decode("utf-8", "replace")
    except Exception:
        return False


def port_busy(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.5)
        return s.connect_ex((HOST, port)) == 0


# --------------------------------------------------------------------------
# 自我測試（無 GUI、不開瀏覽器）
# --------------------------------------------------------------------------
def selftest() -> int:
    log_path = os.path.join(tempfile.gettempdir(), f"{APP_NAME}_selftest.log")
    lines: list[str] = []
    failures: list[str] = []

    def note(msg: str) -> None:
        lines.append(msg)
        safe_print(msg)

    server = LocalServer(resource_root())
    note(f"{APP_TITLE} {APP_NAME} v{VERSION} 自我測試開始")
    note(f"web 根目錄：{server.web_root}")
    try:
        if not os.path.isdir(server.web_root):
            raise FileNotFoundError(f"找不到 web 資料夾：{server.web_root}")

        pages = school_pages(server.web_root)
        note(f"偵測到各校頁面 {len(pages)} 個")
        if len(pages) != SCHOOL_PAGE_COUNT:
            failures.append(f"各校頁面數量應為 {SCHOOL_PAGE_COUNT}，實得 {len(pages)}")

        port = server.start(port=0)          # 隨機埠，避開使用者正在用的 8765
        note(f"伺服器已啟動：{server.url('/')}（隨機埠 {port}）")

        # (path, 必含字串)
        checks: list[tuple[str, tuple[str, ...]]] = [
            ("/index.html", ("社團", "localStorage")),
            ("/", ("社團", "localStorage")),
            # schools/index.html 是純連結清單頁，本身不含 localStorage 程式碼
            ("/schools/index.html", ("社團",)),
        ]
        for name in pages:
            checks.append(("/schools/" + urllib.request.quote(name), ("社團", "localStorage")))

        for path, needles in checks:
            try:
                with urllib.request.urlopen(server.url(path), timeout=10) as resp:
                    status = resp.status
                    ctype = resp.headers.get("Content-Type", "")
                    body = resp.read().decode("utf-8")
            except Exception as exc:
                failures.append(f"{path} 取得失敗：{exc}")
                note(f"  [FAIL] {path} → {exc}")
                continue
            problems = []
            if status != 200:
                problems.append(f"狀態碼 {status}")
            if ctype.replace(" ", "").lower() != "text/html;charset=utf-8":
                problems.append(f"Content-Type={ctype!r}")
            for needle in needles:
                if needle not in body:
                    problems.append(f"缺少「{needle}」")
            if problems:
                failures.append(f"{path}：" + "、".join(problems))
                note(f"  [FAIL] {path} → " + "、".join(problems))
            else:
                note(f"  [OK] {path}（{status}、{len(body):,} 字元、{ctype}）")

        # ping 端點
        if ping(port):
            note(f"  [OK] {PING_PATH}（重複啟動偵測可用）")
        else:
            failures.append(f"{PING_PATH} 沒有正確回應")
            note(f"  [FAIL] {PING_PATH}")

        # 目錄外存取必須被擋
        for evil in ("/../app.py", "/%2e%2e/app.py", "/..%2fapp.py"):
            code = 0
            try:
                with urllib.request.urlopen(server.url(evil), timeout=5) as resp:
                    code = resp.status
                    leaked = b"ClubHandler" in resp.read()
            except urllib.error.HTTPError as exc:
                code = exc.code
                leaked = False
            except Exception:
                code = -1
                leaked = False
            if code == 200 and leaked:
                failures.append(f"目錄外存取沒被擋：{evil}")
                note(f"  [FAIL] 目錄外存取 {evil} → 200 且讀到原始碼")
            else:
                note(f"  [OK] 目錄外存取被擋：{evil}（{code}）")

        # 備份資料夾可建立
        bdir = ensure_backup_dir()
        if os.path.isdir(bdir):
            note(f"  [OK] 備份資料夾：{bdir}")
        else:
            failures.append(f"備份資料夾無法建立：{bdir}")
            note(f"  [FAIL] 備份資料夾：{bdir}")

    except Exception as exc:  # noqa: BLE001
        failures.append(f"未預期的錯誤：{exc}")
        note(f"  [FAIL] {exc}")
    finally:
        server.stop()
        note("伺服器已關閉")

    if failures:
        note("SELFTEST FAILED")
        for f in failures:
            note("  - " + f)
    else:
        note("SELFTEST OK")

    try:
        with open(log_path, "w", encoding="utf-8") as fh:
            fh.write("\n".join(lines) + "\n")
        safe_print(f"log：{log_path}")
    except Exception:
        pass
    return 1 if failures else 0


# --------------------------------------------------------------------------
# GUI
# --------------------------------------------------------------------------
def pick_font() -> str:
    import tkinter.font as tkfont
    try:
        families = set(tkfont.families())
    except Exception:
        return "TkDefaultFont"
    for name in ("Microsoft JhengHei UI", "微軟正黑體", "Microsoft JhengHei",
                 "PingFang TC", "PingFang SC", "Heiti TC", "Noto Sans CJK TC"):
        if name in families:
            return name
    return "TkDefaultFont"


def run_gui() -> int:
    import tkinter as tk
    from tkinter import messagebox, ttk

    no_browser = os.environ.get("CLUBMANAGER_NO_BROWSER", "") == "1"

    def browse(url: str) -> None:
        if no_browser:
            return
        import webbrowser
        try:
            webbrowser.open(url, new=2)
        except Exception:
            pass

    # 重複啟動：8765 已經是本程式在服務 → 只開瀏覽器，不再起第二個伺服器
    if ping(DEFAULT_PORT):
        browse(f"http://{HOST}:{DEFAULT_PORT}/index.html")
        msg = ("系統已經在執行中了。\n\n已經幫您把瀏覽器打開，"
               "請切換到瀏覽器視窗使用。\n（不會重複開啟第二套系統，資料才不會分家。）")
        if no_browser:
            # 測試模式（CLUBMANAGER_NO_BROWSER=1）：不開瀏覽器也不跳對話框，直接結束
            safe_print("ALREADY RUNNING")
            return 0
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(APP_TITLE, msg)
        root.destroy()
        return 0

    web_root = resource_root()
    if not os.path.isdir(web_root):
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(APP_TITLE, f"找不到網頁檔案資料夾：\n{web_root}")
        root.destroy()
        return 1

    server = LocalServer(web_root)
    try:
        server.start()
    except OSError as exc:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(APP_TITLE, f"無法啟動本機伺服器：\n{exc}")
        root.destroy()
        return 1

    ensure_backup_dir()

    NAVY, TEAL, INK, GREY, BG = "#0B1F3A", "#1C7293", "#1E293B", "#64748B", "#F5F9FA"
    fam = pick_font()

    root = tk.Tk()
    root.title(f"{APP_TITLE}　{APP_NAME} v{VERSION}")
    root.configure(bg=BG)
    root.geometry("640x520")
    root.minsize(560, 480)

    header = tk.Frame(root, bg=NAVY)
    header.pack(fill="x")
    tk.Label(header, text=APP_TITLE, bg=NAVY, fg="white",
             font=(fam, 18, "bold")).pack(anchor="w", padx=20, pady=(16, 2))
    tk.Label(header, text="系統已經啟動，請在瀏覽器裡使用；這個小視窗關掉，系統就會停止。",
             bg=NAVY, fg="#B9D6E4", font=(fam, 10)).pack(anchor="w", padx=20, pady=(0, 14))

    body = tk.Frame(root, bg=BG)
    body.pack(fill="both", expand=True, padx=20, pady=16)

    status = tk.Label(body, text=f"網址：{server.url('/index.html')}",
                      bg=BG, fg=TEAL, font=(fam, 11, "bold"), anchor="w", justify="left")
    status.pack(fill="x")

    if server.port_changed:
        tk.Label(
            body,
            text=(f"注意：原本的 {DEFAULT_PORT} 埠被別的程式占用了，這次改用 {server.port}。\n"
                  f"瀏覽器會把不同埠當成不同網站，所以您在 {DEFAULT_PORT} 存過的資料在這裡看不到。\n"
                  "請先關掉占用的程式，再重新開啟本系統。"),
            bg="#FDF3E0", fg="#8A5A00", font=(fam, 10), justify="left",
            anchor="w", padx=10, pady=8, relief="flat",
        ).pack(fill="x", pady=(10, 0))

    btns = tk.Frame(body, bg=BG)
    btns.pack(fill="x", pady=(18, 6))

    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("CM.TButton", font=(fam, 11), padding=(10, 10))

    def open_main():
        browse(server.url("/index.html"))

    def open_schools():
        browse(server.url("/schools/index.html"))

    def open_backup():
        open_folder(ensure_backup_dir())

    ttk.Button(btns, text="開啟系統（總版）", style="CM.TButton",
               command=open_main).grid(row=0, column=0, sticky="ew", padx=(0, 8), pady=4)
    ttk.Button(btns, text="開啟各校版清單", style="CM.TButton",
               command=open_schools).grid(row=0, column=1, sticky="ew", padx=(0, 0), pady=4)
    ttk.Button(btns, text="開啟備份資料夾", style="CM.TButton",
               command=open_backup).grid(row=1, column=0, sticky="ew", padx=(0, 8), pady=4)
    btns.columnconfigure(0, weight=1)
    btns.columnconfigure(1, weight=1)

    note = tk.Text(body, height=9, wrap="word", bg="white", fg=INK,
                   font=(fam, 10), relief="flat", padx=12, pady=10,
                   highlightthickness=1, highlightbackground="#D7E3E8")
    note.pack(fill="both", expand=True, pady=(12, 10))
    note.insert("1.0",
                "資料存在哪裡？\n"
                "　所有社團、學生、出席、經費資料都存在「這台電腦的這個瀏覽器」裡"
                "（瀏覽器的 localStorage），不會上傳到網路，也不會存在這支程式裡面。\n\n"
                "所以請養成習慣：\n"
                "　1. 每次用完，在系統裡按一次「匯出備份」，把 .json 檔存到下面這個資料夾。\n"
                f"　2. 備份資料夾：{backup_dir()}\n"
                "　3. 換電腦或重灌前，一定要先匯出備份；只複製這支程式是搬不走資料的。\n\n"
                "提醒：清除瀏覽器資料、換瀏覽器、用無痕視窗，都會看不到原本的資料。")
    note.configure(state="disabled")

    bottom = tk.Frame(body, bg=BG)
    bottom.pack(fill="x")
    tk.Label(bottom, text=f"本機伺服器：{HOST}:{server.port}（只有這台電腦連得到）",
             bg=BG, fg=GREY, font=(fam, 9)).pack(side="left")

    closing = {"done": False}

    def do_close():
        if closing["done"]:
            return
        closing["done"] = True
        try:
            server.stop()
        finally:
            try:
                root.destroy()
            except Exception:
                pass

    ttk.Button(bottom, text="關閉（停止系統）", style="CM.TButton",
               command=do_close).pack(side="right")

    root.protocol("WM_DELETE_WINDOW", do_close)

    if not no_browser:
        root.after(300, open_main)

    try:
        root.mainloop()
    finally:
        do_close()
    return 0


# --------------------------------------------------------------------------
def main(argv: list[str]) -> int:
    args = [a.lower() for a in argv[1:]]
    if "--version" in args or "-v" in args:
        safe_print(f"{APP_TITLE}（{APP_NAME}）v{VERSION}")
        return 0
    if "--selftest" in args:
        return selftest()
    return run_gui()


if __name__ == "__main__":
    sys.exit(main(sys.argv))
