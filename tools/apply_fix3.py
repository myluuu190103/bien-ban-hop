# -*- coding: utf-8 -*-
"""Áp fix3_<prefix>_b*.json (field dạng 'muc1.ton_kho' hoặc 'm2'/'nvbh'...) vào data/batch_*.json và kttt*.json."""
import json, glob, os, re, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
DATA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data")
prefix = sys.argv[1] if len(sys.argv) > 1 else "full"

src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "build_docx.py"), encoding="utf-8").read()
m = re.search(r"KTTT_KEYS\s*=\s*\{(.*?)\n\}", src, re.S)
KTTT = set(re.findall(r'"([^"]+)"', m.group(1)))
m2 = re.search(r"ASM_CHECK_KEYS\s*=\s*\{(.*?)\n\}", src, re.S)
KTTT |= set(re.findall(r'"([^"]+)"', m2.group(1))) if m2 else set()

fixes = []
for p in sorted(glob.glob(os.path.join(DATA, f"fix3_{prefix}*_b*.json"))):
    for f in json.load(open(p, encoding="utf-8")):
        fixes.append((f["key"], f["field"], f.get("gia_tri_moi", ""), os.path.basename(p)))

def clean(v):
    v = (v or "").replace(" (?)", "").replace("(?)", "")
    v = re.sub(r"\btấn\b", "thùng", v)
    return re.sub(r"\s{2,}", " ", v).strip()

def set_path(obj, path, val):
    parts = path.split("."); tgt = obj
    for part in parts[:-1]:
        if not isinstance(tgt.get(part), dict): tgt[part] = {}
        tgt = tgt[part]
    old = tgt.get(parts[-1]); tgt[parts[-1]] = val
    return old

applied, skipped = [], []
for p in sorted(glob.glob(os.path.join(DATA, "batch_*.json"))) + sorted(glob.glob(os.path.join(DATA, "kttt*.json"))):
    is_kt = os.path.basename(p).startswith("kttt")
    items = json.load(open(p, encoding="utf-8")); ch = False
    for it in items:
        for key, field, val, srcf in fixes:
            if it["key"] != key: continue
            if is_kt != (key in KTTT): continue          # đúng loại file cho loại biên bản
            if field == "key": continue
            if field in ("ngay", "thang", "nam"): continue           # ngày lấy theo tên file / bảng tuyến
            if (val or "").strip() == "Liêm Xuân": continue         # tên NPP theo bảng tuyến: Xuân Liêm
            # chấp nhận cả tên trường KTTT không dấu chấm và GD02 có dấu chấm
            if is_kt and "." in field: continue
            if not is_kt and field not in it and "." not in field and field not in ("muc3", "muc4", "muc5", "nvbh", "asm", "gsbh", "tai_npp", "dia_ban", "ngay", "thang", "nam"):
                skipped.append((key, field, srcf)); continue
            new = clean(val); old = set_path(it, field, new)
            if old != new: ch = True; applied.append((key, field, old, new, srcf))
    if ch: json.dump(items, open(p, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

print(f"fixes nhận: {len(fixes)} | áp dụng (giá trị đổi): {len(applied)} | bỏ qua: {len(skipped)}")
for key, field, old, new, srcf in applied:
    o = (old or ""); n = new
    print(f"[{srcf}] {key} | {field}: {o[:70]!r} -> {n[:70]!r}")
for s in skipped: print("SKIP", s)
