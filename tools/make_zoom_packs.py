# -*- coding: utf-8 -*-
"""Tạo ảnh zoom (3 băng ×1.8) cho mọi biên bản trong khoảng ngày + gói review đầy đủ trường cho agent."""
import json, glob, os, re, sys, io, datetime, unicodedata, argparse
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import cv2, numpy as np

HERE = os.path.dirname(os.path.abspath(__file__)); DATA = os.path.join(HERE, "..", "data")
CROP_DIR = r"C:\Users\desig\Downloads\Anh_bien_ban_da_cat"
ZOOM_DIR = os.path.join(DATA, "zoom"); os.makedirs(ZOOM_DIR, exist_ok=True)
ap = argparse.ArgumentParser(); ap.add_argument("--start", required=True); ap.add_argument("--end", required=True)
ap.add_argument("--prefix", required=True); ap.add_argument("--per", type=int, default=8)
a = ap.parse_args()

def sd(s): return "".join(c for c in unicodedata.normalize("NFD", s or "") if unicodedata.category(c) != "Mn").lower()
src = open(os.path.join(HERE, "build_docx.py"), encoding="utf-8").read()
def ks(n):
    m = re.search(n + r"\s*=\s*\{(.*?)\n\}", src, re.S); return set(re.findall(r'"([^"]+)"', m.group(1)))
KTTT, FREE, DUP, ASMCHK = ks("KTTT_KEYS"), ks("FREEFORM_KEYS"), ks("DUP_DROP_KEYS"), ks("ASM_CHECK_KEYS")
d0 = datetime.datetime.strptime(a.start, "%d-%m-%Y").date(); d1 = datetime.datetime.strptime(a.end, "%d-%m-%Y").date()
def dk(m): return datetime.date(int(m["yyyy"]), int(m["mm"]), int(m["dd"]))
meetings = [m for m in json.load(open(os.path.join(DATA, "meetings.json"), encoding="utf-8"))
            if m["key"] not in DUP and d0 <= dk(m) <= d1]
gd = {}; kt = {}
for p in glob.glob(os.path.join(DATA, "batch_*.json")):
    for it in json.load(open(p, encoding="utf-8")): gd[it["key"]] = it
for p in glob.glob(os.path.join(DATA, "kttt*.json")):
    for it in json.load(open(p, encoding="utf-8")): kt[it["key"]] = it
ns = json.load(open(os.path.join(DATA, "nhansu2_raw.json"), encoding="utf-8"))
nvbh_rows = [(r[1], r[2]) for r in ns["NVBH"][3:] if r[1]]
gsbh_route = {r[1]: r[2] for r in ns["GSBH"][4:] if r[1] and r[2]}
asm_names = [r[1] for r in ns["ASM"][4:] if r[1]]
def route_cands(person):
    words = [sd(w.strip()) for w in re.split("[,;]", gsbh_route.get(person, "")) if w.strip()]
    return sorted({n for n, r in nvbh_rows if r and any(w in sd(r) for w in words)})

def make_zoom(img_path, stem):
    img = cv2.imdecode(np.fromfile(img_path, dtype=np.uint8), cv2.IMREAD_COLOR)
    h, w = img.shape[:2]; outs = []
    for i, (y0, y1) in enumerate([(0.0, 0.40), (0.30, 0.72), (0.62, 1.0)], start=1):
        band = img[int(h * y0):int(h * y1), :]
        f = min(1.8, 2400 / band.shape[1])
        band = cv2.resize(band, None, fx=f, fy=f, interpolation=cv2.INTER_CUBIC)
        out = os.path.join(ZOOM_DIR, f"{stem}_{i}.jpg")
        cv2.imencode(".jpg", band, [cv2.IMWRITE_JPEG_QUALITY, 90])[1].tofile(out); outs.append(out)
    return outs

pack = []
for m in meetings:
    key = m["key"]; stem = re.sub(r"[^\w]+", "_", sd(key))[:80]
    zooms = make_zoom(os.path.join(CROP_DIR, m["image"]), stem)
    loai = "ASM_CHECK" if key in ASMCHK else ("KTTT" if key in KTTT else ("FREEFORM" if key in FREE else "GD02"))
    cur = kt[key] if (key in KTTT or key in ASMCHK) else gd[key]
    pack.append({"key": key, "loai": loai, "anh_goc": os.path.join(CROP_DIR, m["image"]), "anh_zoom": zooms,
                 "nguoi_gui": m["nguoi_gui"], "nvbh_trong_tuyen": route_cands(m["nguoi_gui"]),
                 "gsbh_tuyen": gsbh_route.get(m["nguoi_gui"], ""), "asm_danh_sach": asm_names,
                 "du_lieu_hien_co": {k: v for k, v in cur.items() if k != "key"}})
n = (len(pack) + a.per - 1) // a.per
for i in range(n):
    json.dump(pack[i * a.per:(i + 1) * a.per], open(os.path.join(DATA, f"{a.prefix}_b{i+1}.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(f"{len(pack)} biên bản → {n} gói ({a.prefix}_b1..b{n}), zoom tại {ZOOM_DIR}")
