# -*- coding: utf-8 -*-
"""Ghép ocr<N>_out_b*.json vào data/kttt<N>_01.json (KTTT) + data/batch_<NN>.json (GD02/GD03) và đăng ký key vào build_docx.py.
Dùng: python tools/merge_ocr.py 4 13"""
import json, glob, os, re, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__)); DATA = os.path.join(HERE, "..", "data")
n, bn = sys.argv[1], sys.argv[2]
kt, gd = [], []
for p in sorted(glob.glob(os.path.join(DATA, f"ocr{n}_out_b*.json"))):
    for it in json.load(open(p, encoding="utf-8")):
        for k, v in list(it.items()):
            if isinstance(v, str): it[k] = re.sub(r"\btấn\b", "thùng", v.replace("(?)", "")).strip()
        (kt if "m1" in it else gd).append(it)
if kt: json.dump(kt, open(os.path.join(DATA, f"kttt{n}_01.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
if gd: json.dump(gd, open(os.path.join(DATA, f"batch_{bn}.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
bp = os.path.join(HERE, "build_docx.py"); s = open(bp, encoding="utf-8").read()
def add(name, keys):
    global s
    m = re.search(name + r"\s*=\s*\{(.*?)\n\}", s, re.S); have = set(re.findall(r'"([^"]+)"', m.group(1)))
    new = [k for k in keys if k not in have]
    if new:
        ins = "".join(f'\n    "{k}",' for k in new)
        s = s[:m.end(1)] + ins + s[m.end(1):]
    return new
a = add("KTTT_KEYS", [it["key"] for it in kt]); b = add("FREEFORM_KEYS", [it["key"] for it in gd if "siêu thị" in it["key"]])
open(bp, "w", encoding="utf-8").write(s)
print(f"KTTT: {len(kt)} (mới đăng ký {len(a)}) | GD02/GD03: {len(gd)} (freeform mới {len(b)})")
