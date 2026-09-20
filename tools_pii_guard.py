# -*- coding: utf-8 -*-
"""CI 防線：repo 內不得出現「看起來像真實學號」的 9 碼數字（41/61/81 開頭）。
樣本學號一律使用 9900xxxxx 號段。發現即以非零碼結束，阻止封裝。"""
import os, re, sys, zipfile
PAT = re.compile(r'(?<!\d)(?:41|61|81)\d{7}(?!\d)')
bad = []
for root, dirs, files in os.walk('apps'):
    dirs[:] = [d for d in dirs if d not in ('dist', 'build', '__pycache__')]
    for f in files:
        p = os.path.join(root, f); ext = f.lower().rsplit('.', 1)[-1]
        try:
            if ext in ('py', 'md', 'json', 'txt', 'csv', 'spec', 'html'):
                t = open(p, encoding='utf-8', errors='ignore').read()
            elif ext in ('xlsx', 'docx', 'pptx'):
                z = zipfile.ZipFile(p)
                t = ''.join(z.read(n).decode('utf-8', 'ignore') for n in z.namelist() if n.endswith('.xml'))
            else:
                continue
        except Exception:
            continue
        n = len(PAT.findall(t))
        if n:
            bad.append((p, n))
if bad:
    for p, n in bad:
        print('PII GUARD: %s 有 %d 個疑似真實學號樣式' % (p, n))
    sys.exit(1)
print('PII GUARD OK')
