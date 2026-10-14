# -*- coding: utf-8 -*-
"""Bước 1+5: gom ảnh thư mục mới vào data/meetings.json (bỏ key đã có), cắt ảnh, contact sheet để soi selfie/cắt hụt.
Dùng: python tools/ingest_folder.py "<thư mục ảnh>" [--pick "key=Ảnh 02" ...]"""
import os, re, sys, io, json, argparse, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import cv2, numpy as np
from autocrop import process_folder
HERE = os.path.dirname(os.path.abspath(__file__)); DATA = os.path.join(HERE, "..", "data")
CROP_DIR = r"C:\Users\desig\Downloads\Anh_bien_ban_da_cat"
ap = argparse.ArgumentParser(); ap.add_argument("src"); ap.add_argument("--pick", nargs="*", default=[])
a = ap.parse_args()
pick = dict(x.split("=", 1) for x in a.pick)
mp = os.path.join(DATA, "meetings.json"); meetings = json.load(open(mp, encoding="utf-8"))
have = {m["key"] for m in meetings}
pat = re.compile(r"^(?P<ng>.+?) - (?P<npp>.+?) - (?P<dd>\d\d)-(?P<mm>\d\d)-(?P<yyyy>\d{4}) - (?P<anh>Ảnh \d\d)(?: - \d)?\.jpe?g$", re.I)
groups = {}
for fn in sorted(os.listdir(a.src)):
    m = pat.match(fn)
    if not m: print("BỎ QUA (tên lạ):", fn); continue
    key = f"{m['ng']} - {m['npp']} - {m['dd']}-{m['mm']}-{m['yyyy']}"
    groups.setdefault(key, {"m": m, "files": []})["files"].append(fn)
new, files = [], []
for key, g in groups.items():
    if key in have: print("ĐÃ CÓ:", key); continue
    m = g["m"]; fs = g["files"]
    want = pick.get(key, "Ảnh 01")
    cand = [f for f in fs if f" - {want}.jp" in f or f" - {want} - " in f] or fs
    img = sorted(cand, key=len)[0]
    new.append({"key": key, "image": img, "nguoi_gui": m["ng"], "npp_ten": m["npp"], "dd": m["dd"], "mm": m["mm"],
                "yyyy": m["yyyy"], "n_images": len(fs)})
    files.append(img)
meetings += new
json.dump(meetings, open(mp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
rep = os.path.join(DATA, f"crop_report_{stamp}.json")
report = process_folder(a.src, CROP_DIR, files, rep)
print(f"Thêm {len(new)} biên bản mới (tổng {len(meetings)}), cắt {len(files)} ảnh, report: {rep}")
for r in report:
    if r.get("status") == "cropped":
        bf = r["box"][3] / r["orig"][1]
        print(f"  {r['file']}: giảm {r.get('area_reduction_pct', 0):.0f}%{'  ⚠ đáy ' + format(bf, '.0%') if bf < 0.72 else ''}")
    else: print("  ", r["file"], r["status"])
# contact sheet: mọi ảnh trong thư mục (gốc) + ảnh đã cắt của key mới
def sheet(paths, out, w=420):
    tiles = []
    for p in paths:
        im = cv2.imdecode(np.fromfile(p, dtype=np.uint8), cv2.IMREAD_COLOR)
        if im is None: continue
        h = int(im.shape[0] * w / im.shape[1]); im = cv2.resize(im, (w, h)); im = im[:560] if h > 560 else np.vstack([im, np.full((560 - h, w, 3), 255, np.uint8)])
        cv2.putText(im, os.path.basename(p)[:60], (4, 18), cv2.FONT_HERSHEY_SIMPLEX, 0.45, (0, 0, 255), 1); tiles.append(im)
    cols = 5; rows = [np.hstack(tiles[i:i + cols] + [np.full((560, w, 3), 255, np.uint8)] * (cols - len(tiles[i:i + cols]))) for i in range(0, len(tiles), cols)]
    cv2.imencode(".jpg", np.vstack(rows), [cv2.IMWRITE_JPEG_QUALITY, 80])[1].tofile(out)
sheet([os.path.join(a.src, f) for f in sorted(os.listdir(a.src)) if f.lower().endswith(".jpg")], os.path.join(DATA, f"sheet_goc_{stamp}.jpg"))
sheet([os.path.join(CROP_DIR, f) for f in files], os.path.join(DATA, f"sheet_cat_{stamp}.jpg"))
print("contact sheets:", os.path.join(DATA, f"sheet_goc_{stamp}.jpg"), os.path.join(DATA, f"sheet_cat_{stamp}.jpg"))
