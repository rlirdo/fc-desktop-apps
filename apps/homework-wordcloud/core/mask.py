# -*- coding: utf-8 -*-
"""
mask.py — 姓名遮罩＋學生編號（HomeworkWordCloud 2.1）

這個檔把「姓名遮罩與學生編號」（NameMasker 2.0）的核心規則**複製**一份進來，
讓本程式在拿到「原始 Zuvio 匯出檔」時也能自己完成去識別化，
不需要 import 另一支 App 的路徑（兩支程式各自獨立、各自可打包）。

規則（與 NameMasker 2.0 相同）：
  1. 姓名：中文姓名的第 2 個字改成 O
         王小明 → 王O明　　李明 → 李O　　歐陽小花 → 歐O小花
     已經遮罩過的姓名（王O明）再遮一次不會變形（冪等）。
  2. 學號／電子郵件：每個字元都換成 O，長度不變
         990000043 → OOOOOOOOO　　abc@x.tw → OOOOOOOO
  3. 學生編號：`{學期}_{課程縮寫}_{n}`，例如 `115-1_EC_3`
     **2.1 起的預設**：n 來自 TA 提供的原始名單所產生的「固定對照表」
     （`core/roster.py`），同一位同學整學期、跨題、跨週都是同一個編號（`FixedCoder`）。
     名單外的作答者（退選等）自 101 起持久登記在對照表裡。
     只有明確選擇「沒有名單」時才退回 2.0 的行為（`StudentCoder`：依檔案內
     首次出現順序編號 —— 同一位同學在不同檔案會拿到不同編號，不建議）。
     編號鍵優先用 9 碼學號，沒有學號才用姓名。
  4. 自由文字：出現在名冊上的真實姓名 → 換成該生的學生編號
         「和王小明討論」→「和 115-1_EC_3 討論」

任何寫到硬碟、簡報、網路的資料都必須是去識別化後的版本；
原始 xlsx 只留在使用者自己的 input\\，永遠不上傳。
"""
import re

# 連續 2 個以上中文字（含擴充 A 區），用來辨識「這一段是姓名」
NAME_RUN = re.compile("[一-鿿㐀-䶿]{2,}")

# 系統佔位字串，不是真人姓名，不遮罩、不編號
NOT_A_NAME = {"匿名作答者", "未作答學生", "匿名", "無", "測試帳號", "示範帳號",
              "姓名", "學號", "學生編號", "-", "—", ""}

# 9 碼學號
SID9 = re.compile(r"^\d{9}$")
# 掃描用：前後都不可以是英數底線，避免把 PyInstaller 的 _MEI000105442
# 這類暫存路徑誤判成學號（真正的學號在 CSV/JSON 欄位裡一定是獨立的）
SID9_IN_TEXT = re.compile(r"(?<![0-9A-Za-z_])\d{9}(?![0-9A-Za-z_])")
DIGIT_RUN = re.compile(r"\d{9,}")   # 遮罩用：文字內任何 ≥9 碼數字串

# 電子郵件
EMAIL = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

# 學生編號：115-1_EC_12（隱私掃描時這是「允許出現」的字串）
STUDENT_CODE = re.compile(r"^\d{3}-[12]_[A-Z]{2}_\d+$")
STUDENT_CODE_IN_TEXT = re.compile(r"\d{3}-[12]_[A-Z]{2}_\d+")

COURSE_CODE_RE = re.compile(r"^[A-Z]{2}$")
SEMESTER_RE = re.compile(r"^\d{3}-[12]$")

DEFAULT_COURSE = "EC"
DEFAULT_SEMESTER = "115-1"


# ------------------------------------------------------------------ 欄位判斷
def is_student_id(v):
    """是不是 9 碼學號。"""
    return bool(SID9.match(str(v or "").strip()))


def is_email(v):
    return bool(EMAIL.fullmatch(str(v or "").strip()))


def is_student_code(v):
    return bool(STUDENT_CODE.match(str(v or "").strip()))


def looks_like_name(v):
    """兩個字以上的純中文 → 視為真人姓名。"""
    v = str(v or "").strip()
    return len(v) >= 2 and v not in NOT_A_NAME and bool(NAME_RUN.fullmatch(v))


# ------------------------------------------------------------------ 遮罩
def mask_name(text):
    """把字串中第一段連續中文字的第 2 個字改為 O；無中文則原樣返回（冪等）。"""
    if not isinstance(text, str) or not text:
        return text
    if text.strip() in NOT_A_NAME:
        return text
    m = NAME_RUN.search(text)
    if not m:
        return text
    run = m.group(0)
    masked = run[0] + "O" + run[2:]
    return text[: m.start()] + masked + text[m.start() + len(run):]


def mask_id(v):
    """學號／電子郵件：每個字元換成 O，長度不變（冪等）。"""
    s = str(v or "").strip()
    if not s or s in NOT_A_NAME:
        return s
    return "O" * len(s)


def mask_ids_in_text(t):
    """把自由文字裡夾帶的 9 碼學號與電子郵件也換成等長的 O。"""
    if not isinstance(t, str) or not t:
        return t
    # 學號可能與其他字黏在一起（「4113xxxxx115-1」），遮罩時用無邊界的 ≥9 碼；掃描仍用 SID9_IN_TEXT
    t = DIGIT_RUN.sub(lambda m: "O" * len(m.group(0)), t)
    t = EMAIL.sub(lambda m: "O" * len(m.group(0)), t)
    return t


# ------------------------------------------------------------------ 學生編號
def normalize_course_code(s, default=DEFAULT_COURSE):
    s = str(s or "").strip().upper()
    return s if COURSE_CODE_RE.match(s) else default


def normalize_semester(s, default=DEFAULT_SEMESTER):
    s = str(s or "").strip().replace("／", "/").replace("/", "-")
    return s if SEMESTER_RE.match(s) else default


def make_code(semester, course, n):
    return f"{normalize_semester(semester)}_{normalize_course_code(course)}_{int(n)}"


class StudentCoder:
    """把（學號, 姓名）對應到學生編號；**依首次出現順序**編號。

    ⚠ 2.0 的舊行為：同一位同學在不同檔案會拿到不同編號，沒辦法跨題跨週追蹤。
    2.1 只有在使用者明確選擇「沒有名單」時才會用到它；正常請用 `FixedCoder`。
    """

    def __init__(self, semester=DEFAULT_SEMESTER, course=DEFAULT_COURSE):
        self.semester = normalize_semester(semester)
        self.course = normalize_course_code(course)
        self._map = {}          # key -> code
        self._name = {}         # code -> 真實姓名（只留在記憶體，不寫檔）
        self._n = 0

    @staticmethod
    def key_of(sid, name):
        sid = str(sid or "").strip()
        name = str(name or "").strip()
        if is_student_id(sid):
            return "S:" + sid
        if name and name not in NOT_A_NAME:
            return "N:" + name
        return ""

    def code_for(self, sid, name):
        """回傳學生編號；匿名／非真人回傳空字串（不編號）。"""
        name_s = str(name or "").strip()
        if name_s in NOT_A_NAME and not is_student_id(str(sid or "").strip()):
            return ""
        k = self.key_of(sid, name)
        if not k:
            return ""
        if k not in self._map:
            self._n += 1
            self._map[k] = make_code(self.semester, self.course, self._n)
            if looks_like_name(name_s):
                self._name[self._map[k]] = name_s
        elif looks_like_name(name_s):
            self._name.setdefault(self._map[k], name_s)
        return self._map[k]

    def name_to_code(self):
        """{真實姓名: 學生編號}，供自由文字替換用（只在記憶體）。"""
        return {v: k for k, v in self._name.items()}

    @property
    def count(self):
        return self._n


class FixedCoder:
    """2.1 預設：依 TA 提供的**固定對照表**給學生編號。

    * 名單內的同學：編號來自對照表「學生編號對照」工作表，整學期固定。
    * 名單外的作答者（退選、旁聽…）：由對照表持久登記，自 101 起遞增，
      跨檔跨週一致（處理完會回寫對照表）。
    * 自由文字替換用的 `name_to_code()` 涵蓋**名單內所有姓名**，
      即使這個檔案的姓名欄沒有出現過（SPEC §1.4）。

    參數 codebook 是 `core.roster.Codebook`（只用 duck typing，避免循環匯入）。
    """

    def __init__(self, codebook, source_file=""):
        self.cb = codebook
        self.source_file = source_file
        self.semester = getattr(codebook, "semester", DEFAULT_SEMESTER)
        self.course = getattr(codebook, "course", DEFAULT_COURSE)

    def set_source(self, source_file):
        """換檔案時記下檔名，名單外作答者的「首次出現檔案」才寫得對。"""
        self.source_file = source_file or ""

    def code_for(self, sid, name):
        name_s = str(name or "").strip()
        sid_s = str(sid or "").strip()
        if name_s in NOT_A_NAME and not is_student_id(sid_s):
            return ""                       # 匿名作答者：不編號
        if not sid_s and not name_s:
            return ""
        return self.cb.code_for(sid_s, name_s, self.source_file)

    def name_to_code(self):
        return dict(self.cb.name_to_code())

    def roster_names(self):
        return list(self.cb.names())

    @property
    def count(self):
        return self.cb.count


# ------------------------------------------------------------------ 自由文字
def build_roster(*name_lists):
    """回傳 [(真實姓名, 遮罩姓名)]，依長度由長到短排序（給 --check-privacy 用）。"""
    names = set()
    for lst in name_lists:
        for n in lst:
            n = str(n or "").strip()
            if looks_like_name(n):
                names.add(n)
    return sorted(((n, mask_name(n)) for n in names), key=lambda t: -len(t[0]))


def build_replacer(roster, name_codes=None):
    """回傳 [(真實姓名, 取代字串)]；有學生編號就換成編號，沒有就換成遮罩姓名。"""
    name_codes = name_codes or {}
    out = []
    for real, fake in roster:
        out.append((real, name_codes.get(real) or fake))
    return sorted(out, key=lambda t: -len(t[0]))


def mask_in_text(ans, replacer):
    """把名冊中的真實姓名在自由文字裡換掉（長名優先，避免部分覆蓋），
    順便把夾帶的 9 碼學號與電子郵件換成 O。"""
    if not isinstance(ans, str) or not ans:
        return ans
    for real, sub in (replacer or []):
        if real and real in ans:
            ans = ans.replace(real, sub)
    return mask_ids_in_text(ans)


def roster_names(roster):
    """只取真實姓名清單（供 --check-privacy 掃描用）。"""
    return [real for real, _ in roster]
