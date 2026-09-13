# -*- coding: utf-8 -*-
"""
parse_zuvio.py — 把 Zuvio 老師端「下載數據」匯出的 xlsx 轉成每題一份遮罩版 CSV。

通用版：不綁任何課程、不需要 map.json。
    * input 資料夾裡有幾個 xlsx，就依檔名排序編成 Q01、Q02、Q03…
    * 題目名稱直接讀檔案內的「題幹 / 資料夾名稱」，讀不到就用檔名。
    * 也吃通用 CSV（欄位含「姓名」與「作答」），方便其他平台匯出轉檔後使用。

Zuvio xlsx 的兩種資料區塊（已於真實匯出檔確認）：
    A. 上方彙總表：學號 | 姓名 | 作答時間 | 回答 | 回答 | …（一欄一子題）
       子題標籤在「問題類型」列、子題題目在「問題敘述」列，欄位對齊。
       ※ 已知在部分「題組問答」匯出檔會出現整體位移一列的對位錯誤。
    B. 下方逐子題明細區塊：每個子題一段，開頭是「第N題:題型」，
       接著「問題敘述」，再來表頭可能是
           學號 | 姓名 | 回答
           學號 | 姓名 | 正解 | 回答            （單選／是非題）
           學號 | 姓名 | 答對選項數量 | 答錯選項數量 | 回答   （多選題）
       這一區塊的學號／姓名／回答對位是正確的。
    -> 因此只要 B 區塊存在，一律以 B 為準；沒有 B 才退回 A。

輸出 CSV 欄位：題號,子題號,子題題目,學號,姓名,作答時間,作答內容
（姓名已遮罩、作答內容中的名冊姓名也已遮罩）
"""
import os
import re
import csv
import glob

import openpyxl

from .mask import mask_name, mask_in_text, build_roster

CSV_COLS = ["題號", "子題號", "子題題目", "學號", "姓名", "作答時間", "作答內容"]

META_KEYS = ("資料夾名稱", "問題題型", "是否分組", "是否匿名", "題幹", "問題敘述",
             "正解", "總分")


def norm(v):
    if v is None:
        return ""
    return str(v).replace("　", " ").strip()


def safe_name(s, n=12):
    s = re.sub(r"https?://\S+", "", s or "")
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f\s]+', "", s)
    return s[:n] or "untitled"


def read_grid(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    grid = [[norm(c) for c in row] for row in ws.iter_rows(values_only=True)]
    wb.close()
    return grid


# ------------------------------------------------------------------ B 區塊
def parse_details(grid):
    """逐子題明細區塊（第N題:題型）。回傳 [{label, text, answers:[(學號,姓名,回答)]}]。"""
    marks = [i for i, r in enumerate(grid) if r and re.match(r"^第\d+題", r[0])]
    out = []
    for k, i in enumerate(marks):
        end = marks[k + 1] if k + 1 < len(marks) else len(grid)
        label = grid[i][0]
        text = ""
        hdr_i = None
        for j in range(i + 1, end):
            r = grid[j]
            if not r:
                continue
            if r[0] == "問題敘述" and len(r) > 1 and not text:
                text = r[1]
            if r[0] == "學號" and len(r) > 1 and r[1] == "姓名":
                hdr_i = j
                break
        if hdr_i is None:
            continue
        hdr = grid[hdr_i]
        try:
            acol = hdr.index("回答")
        except ValueError:
            acol = 2
        answers = []
        for j in range(hdr_i + 1, end):
            r = grid[j]
            if not r or not r[0] or r[0] == "學號":
                break
            ans = r[acol] if acol < len(r) else ""
            answers.append((r[0], r[1] if len(r) > 1 else "", ans))
        out.append({"label": label, "text": text, "answers": answers})
    return out


# ------------------------------------------------------------------ A 區塊
def parse_summary(grid):
    """上方彙總表。回傳 dict(meta, subs, rows, unanswered)。"""
    meta = {}
    for r in grid[:12]:
        if len(r) >= 2 and r[0] in META_KEYS:
            meta.setdefault(r[0], r[1])

    types_row = next((r for r in grid[:14] if r and r[0] == "問題類型"), None)
    desc_row = next((r for r in grid[:14] if r and r[0] == "問題敘述" and len(r) > 3), None)

    hdr_i = None
    for i, r in enumerate(grid):
        if len(r) >= 3 and r[0] == "學號" and r[1] == "姓名" and r[2] == "作答時間":
            hdr_i = i
            break
    if hdr_i is None:
        return {"meta": meta, "subs": [], "rows": [], "unanswered": []}
    hdr = grid[hdr_i]

    subs = []
    for ci in range(3, len(hdr)):
        if hdr[ci] not in ("回答", "總分") and not hdr[ci].startswith("第"):
            continue                      # 跳過 答對/答錯選項數量 等統計欄
        if hdr[ci] == "總分":
            continue
        label = types_row[ci] if (types_row and ci < len(types_row)) else ""
        text = desc_row[ci] if (desc_row and ci < len(desc_row)) else ""
        if not label and hdr[ci].startswith("第"):
            label = hdr[ci]
        if not label and hdr[ci] == "回答":
            label = "第1題"
        subs.append({"col": ci, "label": label or f"第{len(subs) + 1}題", "text": text})

    rows = []
    i = hdr_i + 1
    while i < len(grid):
        r = grid[i]
        if not any(r) or (r and r[0] in ("未作答學生", "學號")) or (r and r[0].startswith("第")):
            break
        if r and r[0]:
            rows.append(r)
        i += 1

    unanswered = []
    for j in range(i, len(grid)):
        if grid[j] and grid[j][0] == "未作答學生":
            k = j + 2                      # 跳過 學號|姓名 表頭
            while k < len(grid) and grid[k] and grid[k][0] and grid[k][0] != "學號":
                unanswered.append({"學號": grid[k][0],
                                   "姓名": grid[k][1] if len(grid[k]) > 1 else ""})
                k += 1
            break
    return {"meta": meta, "subs": subs, "rows": rows, "unanswered": unanswered}


# ------------------------------------------------------------------ 通用 CSV
def parse_generic_csv(path):
    """其他平台匯出的 CSV：只要欄位名含「姓名」與「作答」就能吃。"""
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    if not rows:
        return None
    cols = list(rows[0].keys())

    def pick(*keys):
        for k in keys:
            for c in cols:
                if k in c:
                    return c
        return None

    c_name = pick("姓名", "名字", "name")
    c_ans = pick("作答", "回答", "答案", "answer")
    if not c_name or not c_ans:
        return None
    c_sid = pick("學號", "編號", "id")
    c_time = pick("時間", "time")
    c_sub = pick("子題", "題號", "題目")
    out = []
    for r in rows:
        out.append({
            "學號": (r.get(c_sid) or "").strip(),
            "姓名": (r.get(c_name) or "").strip(),
            "作答時間": (r.get(c_time) or "").strip() if c_time else "",
            "子題號": (r.get(c_sub) or "第1題").strip() if c_sub else "第1題",
            "作答內容": (r.get(c_ans) or "").strip(),
        })
    return out


# ------------------------------------------------------------------ 主流程
def list_inputs(input_dir):
    """回傳要處理的檔案清單（xlsx 優先，其次 csv），依檔名排序。"""
    files = sorted(glob.glob(os.path.join(input_dir, "*.xlsx")))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    files += sorted(glob.glob(os.path.join(input_dir, "*.csv")))
    return files


def parse_file(path, no):
    """單一檔案 → (rows, rec)；rows 是已遮罩的 CSV 列，rec 是統計資訊。"""
    qno = f"Q{no:02d}"
    base = os.path.splitext(os.path.basename(path))[0]

    if path.lower().endswith(".csv"):
        generic = parse_generic_csv(path)
        if generic is None:
            return [], {"題號": qno, "來源檔": os.path.basename(path),
                        "狀態": "失敗", "原因": "CSV 缺少「姓名」或「作答」欄"}
        roster = build_roster([g["姓名"] for g in generic])
        out_rows, n_masked = [], 0
        for g in generic:
            if not g["作答內容"]:
                continue
            m = mask_name(g["姓名"])
            n_masked += (m != g["姓名"])
            out_rows.append([qno, g["子題號"], "", g["學號"], m, g["作答時間"],
                             mask_in_text(g["作答內容"], roster)])
        rec = {"題號": qno, "來源檔": os.path.basename(path), "狀態": "成功",
               "題目": base, "題型": "通用CSV", "作答人數": len({g["學號"] or g["姓名"] for g in generic}),
               "未作答人數": 0, "全班人數": len({g["學號"] or g["姓名"] for g in generic}),
               "子題數": len({r[1] for r in out_rows}), "有效作答列數": len(out_rows),
               "取用區塊": "通用CSV", "姓名遮罩": "已套用(第2字改O)", "遮罩列數": n_masked,
               "子題": [{"label": s, "text": "", "作答數": sum(1 for r in out_rows if r[1] == s)}
                        for s in sorted({r[1] for r in out_rows})]}
        return out_rows, rec

    grid = read_grid(path)
    p = parse_summary(grid)
    details = parse_details(grid)
    title = p["meta"].get("題幹") or p["meta"].get("資料夾名稱") or base

    # 作答時間只在彙總表有 -> 用學號回查
    tmap = {r[0]: (r[2] if len(r) > 2 else "") for r in p["rows"] if r and r[0]}

    roster = build_roster(
        [r[1] for r in p["rows"] if len(r) > 1],
        [u["姓名"] for u in p["unanswered"]],
        [x[1] for d in details for x in d["answers"]],
    )

    out_rows, n_masked = [], 0
    if details:
        source = "逐子題明細區塊(第N題)"
        for d in details:
            for sid, name, ans in d["answers"]:
                if not ans:
                    continue
                m = mask_name(name)
                n_masked += (m != name)
                out_rows.append([qno, d["label"], d["text"], sid, m,
                                 tmap.get(sid, ""), mask_in_text(ans, roster)])
        subs_meta = [{"label": d["label"], "text": d["text"],
                      "作答數": sum(1 for x in d["answers"] if x[2])} for d in details]
    else:
        source = "彙總表"
        for r in p["rows"]:
            sid = r[0]
            name = r[1] if len(r) > 1 else ""
            ts = r[2] if len(r) > 2 else ""
            m = mask_name(name)
            n_masked += (m != name)
            for s in p["subs"]:
                ans = r[s["col"]] if s["col"] < len(r) else ""
                if not ans:
                    continue
                out_rows.append([qno, s["label"], s["text"], sid, m, ts,
                                 mask_in_text(ans, roster)])
        subs_meta = [{"label": s["label"], "text": s["text"],
                      "作答數": sum(1 for r in p["rows"]
                                 if s["col"] < len(r) and r[s["col"]])}
                     for s in p["subs"]]

    anon = sum(1 for r in p["rows"] if r and r[0] == "匿名作答者")
    named = len(p["rows"]) - anon
    rec = {
        "題號": qno,
        "來源檔": os.path.basename(path),
        "狀態": "成功" if out_rows else "無有效作答",
        "題目": title,
        "題型": p["meta"].get("問題題型", ""),
        "資料夾": p["meta"].get("資料夾名稱", ""),
        "作答人數": len(p["rows"]),
        "實名作答": named,
        "匿名作答": anon,
        "未作答人數": len(p["unanswered"]),
        "全班人數": named + len(p["unanswered"]),
        "子題數": len(subs_meta),
        "子題": subs_meta,
        "有效作答列數": len(out_rows),
        "取用區塊": source,
        "姓名遮罩": "已套用(第2字改O)",
        "遮罩列數": n_masked,
    }
    return out_rows, rec


def parse_dir(input_dir, out_dir, log=print):
    """把 input_dir 內所有檔案轉成 out_dir 內的 Q0N_*.csv，回傳 index 清單。"""
    os.makedirs(out_dir, exist_ok=True)
    files = list_inputs(input_dir)
    if not files:
        raise SystemExit(f"[錯誤] {input_dir} 裡找不到任何 .xlsx 或 .csv。\n"
                         f"       請先到 Zuvio 老師端「下載數據」，把 xlsx 放進這個資料夾。")
    index = []
    for i, path in enumerate(files, 1):
        rows, rec = parse_file(path, i)
        fn = f"{rec['題號']}_{safe_name(rec.get('題目', ''))}.csv"
        rec["檔名"] = fn
        with open(os.path.join(out_dir, fn), "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(CSV_COLS)
            w.writerows(rows)
        index.append(rec)
        log(f"  {rec['題號']}  {rec.get('題目', '')[:26]} -> {fn}"
            f"（作答 {rec.get('作答人數', 0)} 人／子題 {rec.get('子題數', 0)}／"
            f"有效 {rec.get('有效作答列數', 0)} 列／姓名遮罩 {rec.get('遮罩列數', 0)} 列）")
    return index


def roster_from_dir(input_dir):
    """從原始 input 檔案重建名冊（僅供 --check-privacy 在本機比對，不寫出任何檔案）。"""
    names = []
    for path in list_inputs(input_dir):
        if path.lower().endswith(".csv"):
            g = parse_generic_csv(path)
            if g:
                names += [x["姓名"] for x in g]
            continue
        grid = read_grid(path)
        p = parse_summary(grid)
        names += [r[1] for r in p["rows"] if len(r) > 1]
        names += [u["姓名"] for u in p["unanswered"]]
        for d in parse_details(grid):
            names += [x[1] for x in d["answers"]]
    return build_roster(names)
