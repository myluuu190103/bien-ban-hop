# -*- coding: utf-8 -*-
"""Áp kết quả rà số liệu (fix2_b*.json) + soi ô trống (blanks_check2.json) vào data/*.json."""
import json, glob, os, re, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
DATA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data")

GD_MAP = {"ton_kho": ("muc1", "ton_kho"), "bao_cao": ("muc2", "bao_cao"), "muc1_khac": ("muc1", "khac"),
          "ho_tro": ("muc2", "ho_tro"), "muc5": ("muc5",), "nvbh": ("nvbh",), "tai_npp": ("tai_npp",)}
KT_MAP = {"m1_tinh_hinh": "m1", "m2_doanh_so": "m2", "m3_khuyen_mai": "m3", "m1": "m1", "lam_viec_voi": "lam_viec_voi", "npp": "npp"}

fixes = []
for p in sorted(glob.glob(os.path.join(DATA, "fix2_b*.json"))):
    for f in json.load(open(p, encoding="utf-8")):
        fixes.append((f["key"], f["field"], f["gia_tri_moi"], os.path.basename(p)))
bp = os.path.join(DATA, "blanks_check2.json")
if os.path.exists(bp):
    for f in json.load(open(bp, encoding="utf-8")):
        fixes.append((f["key"], f["field"], f["value"], "blanks"))

def clean(v):
    v = (v or "").replace(" (?)", "").replace("(?)", "")
    return re.sub(r"\s{2,}", " ", v).strip()

applied, skipped = [], []
for p in sorted(glob.glob(os.path.join(DATA, "batch_*.json"))):
    items = json.load(open(p, encoding="utf-8")); ch = False
    for it in items:
        for key, field, val, srcf in fixes:
            if it["key"] != key or field not in GD_MAP: continue
            path = GD_MAP[field]; tgt = it
            for part in path[:-1]: tgt = tgt.setdefault(part, {})
            old = tgt.get(path[-1]); new = clean(val)
            if old != new:
                tgt[path[-1]] = new; ch = True; applied.append((key, field, old, new, srcf))
    if ch: json.dump(items, open(p, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
for p in sorted(glob.glob(os.path.join(DATA, "kttt*.json"))):
    items = json.load(open(p, encoding="utf-8")); ch = False
    for it in items:
        for key, field, val, srcf in fixes:
            if it["key"] != key or field not in KT_MAP: continue
            f2 = KT_MAP[field]; old = it.get(f2); new = clean(val)
            if old != new:
                it[f2] = new; ch = True; applied.append((key, field, old, new, srcf))
    if ch: json.dump(items, open(p, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

print(f"fixes nhận: {len(fixes)} | áp dụng thật (giá trị đổi): {len(applied)}")
for key, field, old, new, srcf in applied:
    print(f"[{srcf}] {key} | {field}: {old!r} -> {new!r}")
