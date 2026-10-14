# -*- coding: utf-8 -*-
"""Gói OCR có zoom cho các key vừa ingest: python tools/make_ocr_packs.py <crop_report.json> <prefix> [--per 8] [--drop key ...]"""
import json, os, re, sys, io, unicodedata, argparse
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import cv2, numpy as np
HERE = os.path.dirname(os.path.abspath(__file__)); DATA = os.path.join(HERE, "..", "data")
CROP = r"C:\Users\desig\Downloads\Anh_bien_ban_da_cat"; ZOOM = os.path.join(DATA, "zoom")
ap = argparse.ArgumentParser(); ap.add_argument("report"); ap.add_argument("prefix"); ap.add_argument("--per", type=int, default=8)
ap.add_argument("--drop", nargs="*", default=[]); a = ap.parse_args()
def sd(s): return "".join(c for c in unicodedata.normalize("NFD", s or "") if unicodedata.category(c) != "Mn").lower()
files = {r["file"] for r in json.load(open(a.report, encoding="utf-8"))}
meetings = [m for m in json.load(open(os.path.join(DATA, "meetings.json"), encoding="utf-8")) if m["image"] in files and m["key"] not in a.drop]
ns = json.load(open(os.path.join(DATA, "nhansu2_raw.json"), encoding="utf-8"))
nvbh_rows = [(r[1], r[2]) for r in ns["NVBH"][3:] if r[1]]
gsbh_route = {r[1]: r[2] for r in ns["GSBH"][4:] if r[1] and r[2]}
asm = [r[1] for r in ns["ASM"][4:] if r[1]]
def cands(p):
    ws = [sd(w.strip()) for w in re.split("[,;]", gsbh_route.get(p, "")) if w.strip()]
    return sorted({n for n, r in nvbh_rows if r and any(w in sd(r) for w in ws)})
def zoom(path, stem):
    img = cv2.imdecode(np.fromfile(path, dtype=np.uint8), cv2.IMREAD_COLOR); h, w = img.shape[:2]; outs = []
    for i, (y0, y1) in enumerate([(0, .40), (.30, .72), (.62, 1)], 1):
        b = img[int(h * y0):int(h * y1)]; f = min(1.8, 2400 / b.shape[1]); b = cv2.resize(b, None, fx=f, fy=f, interpolation=cv2.INTER_CUBIC)
        o = os.path.join(ZOOM, f"{stem}_{i}.jpg"); cv2.imencode(".jpg", b, [cv2.IMWRITE_JPEG_QUALITY, 90])[1].tofile(o); outs.append(os.path.abspath(o))
    return outs
pack = []
for x in meetings:
    stem = re.sub(r"[^\w]+", "_", sd(x["key"]))[:80]
    pack.append({"key": x["key"], "loai": "GD03" if "siêu thị" in x["key"] else "KTTT hoặc GD02 (xem tiêu đề tờ giấy)",
                 "anh_goc": os.path.join(CROP, x["image"]), "anh_zoom": zoom(os.path.join(CROP, x["image"]), stem),
                 "nguoi_gui": x["nguoi_gui"], "npp_theo_ten_file": x["npp_ten"], "nvbh_trong_tuyen": cands(x["nguoi_gui"]),
                 "gsbh_tuyen": gsbh_route.get(x["nguoi_gui"], ""), "asm_danh_sach": asm})
n = (len(pack) + a.per - 1) // a.per
for i in range(n):
    json.dump(pack[i * a.per:(i + 1) * a.per], open(os.path.join(DATA, f"{a.prefix}_b{i+1}.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(f"{len(pack)} biên bản → {n} gói {a.prefix}_b1..b{n}")
