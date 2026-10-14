# -*- coding: utf-8 -*-
"""QA khó tính cho 1 tuần: ngày lệch, ô trống đáng ngờ, tên không khớp, ảnh cắt hụt, đối chiếu bảng tuyến."""
import json, re, glob, os, sys, io, datetime, argparse, unicodedata
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "..", "data")
ap = argparse.ArgumentParser()
ap.add_argument("--start", required=True); ap.add_argument("--end", required=True)
ap.add_argument("--xlsx", default=None)
ap.add_argument("--cropreport", nargs="*", default=[])
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

issues = []
def flag(sev, key, msg): issues.append((sev, key, msg))

# 1. ngày trên giấy vs ngày tên file
for m in meetings:
    r = kt[m["key"]] if (m["key"] in KTTT or m["key"] in ASMCHK) else gd[m["key"]]
    ng, th, nm = (r.get("ngay") or ""), (r.get("thang") or ""), (r.get("nam") or "")
    if ng and th and (ng.zfill(2), th.zfill(2)) != (m["dd"], m["mm"]):
        flag("NGÀY", m["key"], f"giấy ghi {ng}/{th}/{nm} ≠ tên file {m['dd']}/{m['mm']}")

# 2. ô trống đáng ngờ
for m in meetings:
    k = m["key"]
    if k in ASMCHK:
        r = kt[k]
        if not (r.get("ho_ten") or "").strip(): flag("TRỐNG", k, "ASM_CHECK trống: ho_ten")
    elif k in KTTT:
        r = kt[k]
        empty = [f for f in ("ho_ten", "npp", "m1", "m2") if not (r.get(f) or "").strip()]
        if empty: flag("TRỐNG", k, "KTTT trống: " + ", ".join(empty))
        if not (r.get("lam_viec_voi") or "").strip(): flag("TRỐNG-nhẹ", k, "KTTT: 'Làm việc với' trống")
    elif k in FREE:
        r = gd[k]
        if not any((r.get("muc1", {}) or {}).values()) and not r.get("muc4") and not r.get("muc5"):
            flag("TRỐNG", k, "freeform không có nội dung")
    else:
        r = gd[k]; m1 = r.get("muc1", {}) or {}; m2 = r.get("muc2", {}) or {}
        empty = []
        if not (r.get("tai_npp") or "").strip(): empty.append("tai_npp")
        if not (r.get("gsbh") or "").strip(): empty.append("gsbh")
        if not (m1.get("ton_kho") or "").strip(): empty.append("ton_kho")
        if not (r.get("muc5") or "").strip(): empty.append("muc5")
        if not (m2.get("ho_tro") or "").strip(): empty.append("ho_tro")
        if not (r.get("nvbh") or "").strip(): empty.append("nvbh")
        if empty: flag("TRỐNG", k, "GD02 trống: " + ", ".join(empty))

# 3. tai_npp / gsbh vs tên file
for m in meetings:
    k = m["key"]
    if k in ASMCHK:
        continue  # mẫu kiểm tra nội bộ ASM/GS, không gắn 1 NPP cụ thể
    r = kt[k] if k in KTTT else gd[k]
    npp_paper = r.get("npp") if k in KTTT else r.get("tai_npp")
    if npp_paper and sd(m["npp_ten"]) not in sd(npp_paper) and sd(npp_paper) not in sd(m["npp_ten"]):
        flag("NPP", k, f"giấy ghi NPP '{npp_paper}' ≠ tên file '{m['npp_ten']}'")
    who = r.get("ho_ten") if k in KTTT else r.get("gsbh")
    if who and sd(m["nguoi_gui"]) != sd(who) and sd(m["nguoi_gui"]) not in sd(who):
        flag("NGƯỜI", k, f"giấy ghi {'Họ tên' if k in KTTT else 'GSBH'} '{who}' ≠ người gửi '{m['nguoi_gui']}'")

# 4. đối chiếu bảng tuyến
master = []; cur = None
ws = openpyxl.load_workbook(a.xlsx, data_only=True)["Tổng hợp"] if a.xlsx else None
for r_ in (range(3, ws.max_row + 1) if ws else []):
    ht, npp, ng = ws.cell(row=r_, column=2).value, ws.cell(row=r_, column=3).value, ws.cell(row=r_, column=5).value
    if ht: cur = ht
    if npp:
        for d in ([x.strip() for x in str(ng).split(",")] if ng else []):
            master.append((cur, npp, d))
mine = {}
for m in meetings: mine.setdefault((sd(m["npp_ten"]), f"{m['dd']}/{m['mm']}/{m['yyyy']}"), []).append(m["nguoi_gui"])
for ht, npp, d in master:
    if (sd(npp), d) not in mine: flag("TUYẾN", f"{ht} - {npp}", f"bảng tuyến có họp {d} nhưng không có biên bản")
    elif ht not in mine[(sd(npp), d)]: flag("TUYẾN-nhẹ", f"{ht} - {npp}", f"{d}: người gửi {mine[(sd(npp), d)]} ≠ GSBH tuyến {ht}")
mset = {(sd(npp), d) for _, npp, d in master}
for m in (meetings if a.xlsx else []):
    if (sd(m["npp_ten"]), f"{m['dd']}/{m['mm']}/{m['yyyy']}") not in mset:
        flag("TUYẾN", m["key"], "có biên bản nhưng bảng tuyến không ghi")

# 5. ảnh cắt hụt
for cr in a.cropreport:
    for r in json.load(open(cr, encoding="utf-8")):
        if r.get("status") == "cropped":
            bf = r["box"][3] / r["orig"][1]
            if bf < 0.72: flag("ẢNH", r["file"], f"đáy cắt ở {bf:.0%} chiều cao ảnh gốc → nghi cắt hụt chữ ký")

# 6. còn dấu (?) / đơn vị tấn
for m in meetings:
    r = json.dumps(kt[m["key"]] if (m["key"] in KTTT or m["key"] in ASMCHK) else gd[m["key"]], ensure_ascii=False)
    if "(?)" in r: flag("KÝ HIỆU", m["key"], "còn (?)")
    if re.search(r"\btấn\b", r): flag("ĐƠN VỊ", m["key"], "còn 'tấn'")

sev_order = {"NGÀY": 0, "TUYẾN": 1, "NPP": 2, "NGƯỜI": 3, "ẢNH": 4, "TRỐNG": 5, "ĐƠN VỊ": 6, "KÝ HIỆU": 7, "TUYẾN-nhẹ": 8, "TRỐNG-nhẹ": 9}
issues.sort(key=lambda x: (sev_order.get(x[0], 99), x[1]))
print(f"Tuần {a.start} → {a.end}: {len(meetings)} biên bản, {len(issues)} điểm cần xem\n")
for sev, key, msg in issues: print(f"[{sev}] {key}: {msg}")
