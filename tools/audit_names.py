# -*- coding: utf-8 -*-
import json, re, difflib, unicodedata, glob, os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
SCRATCH = os.path.join(HERE, "..", "data")

def strip_diac(s):
    s = unicodedata.normalize("NFD", s or "")
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower()

def load_key_set(name):
    src = open(os.path.join(HERE, "build_docx.py"), encoding="utf-8").read()
    m = re.search(name + r"\s*=\s*\{(.*?)\n\}", src, re.S)
    return set(re.findall(r'"([^"]+)"', m.group(1)))

KTTT_KEYS = load_key_set("KTTT_KEYS")
FREEFORM_KEYS = load_key_set("FREEFORM_KEYS")
DUP_DROP_KEYS = load_key_set("DUP_DROP_KEYS")

with open(os.path.join(SCRATCH, "nhansu2_raw.json"), encoding="utf-8") as f:
    ns = json.load(f)
nvbh_rows = [(r[1], r[2]) for r in ns["NVBH"][3:] if r[1]]
gsbh_route = {r[1]: r[2] for r in ns["GSBH"][4:] if r[1] and r[2]}
nvbh_names = [n for n, _ in nvbh_rows]
gsbh_names = list(gsbh_route.keys())
asm_names = [r[1] for r in ns["ASM"][4:] if r[1]]
all_staff = nvbh_names + gsbh_names + asm_names

with open(os.path.join(SCRATCH, "meetings.json"), encoding="utf-8") as f:
    meetings = [m for m in json.load(f) if m["key"] not in DUP_DROP_KEYS]

data_by_key = {}
for p in sorted(glob.glob(os.path.join(SCRATCH, "batch_*.json"))):
    for it in json.load(open(p, encoding="utf-8")):
        data_by_key[it["key"]] = (it, os.path.basename(p))
kttt_by_key = {}
for p in sorted(glob.glob(os.path.join(SCRATCH, "kttt*.json"))):
    for it in json.load(open(p, encoding="utf-8")):
        kttt_by_key[it["key"]] = (it, os.path.basename(p))

def route_cands(person):
    route = gsbh_route.get(person, "")
    words = [strip_diac(w.strip()) for w in re.split("[,;]", route) if w.strip()]
    return [n for n, r in nvbh_rows if r and any(w in strip_diac(r) for w in words)]

def split_names(s):
    s = re.sub(r"\(\?\)", "", s or "")
    parts = re.split(r"[,&;/]|\bvà\b|\+", s)
    return [re.sub(r"\s+", " ", p).strip(" .") for p in parts if p.strip(" .")]

print("=== NVBH (GD02/freeform) can ra soat ===")
for m in meetings:
    key = m["key"]
    if key in KTTT_KEYS:
        continue
    if key not in data_by_key:
        continue
    d, fn = data_by_key[key]
    nvbh = d.get("nvbh") or ""
    if not nvbh.strip():
        continue
    names = split_names(nvbh)
    bad = []
    for n in names:
        if n in nvbh_names:
            continue
        cands = route_cands(m["nguoi_gui"])
        close = difflib.get_close_matches(n, cands, n=2, cutoff=0.45) or difflib.get_close_matches(n, nvbh_names, n=2, cutoff=0.75)
        bad.append((n, close))
    if bad:
        print(f"{key} [{fn}] | raw={nvbh!r} | route={gsbh_route.get(m['nguoi_gui'])}")
        for n, close in bad:
            print(f"     ? {n!r} -> {close}")

print()
print("=== KTTT: ho_ten / lam_viec_voi ===")
for m in meetings:
    key = m["key"]
    if key not in KTTT_KEYS:
        continue
    if key not in kttt_by_key:
        continue
    k, fn = kttt_by_key[key]
    print(f"{key} [{fn}] | ho_ten={k.get('ho_ten')!r} chuc_vu={k.get('chuc_vu')!r} | lvv={k.get('lam_viec_voi')!r} | route={gsbh_route.get(m['nguoi_gui'])}")

print()
print("=== so luong dau (?) theo file ===")
for p in sorted(glob.glob(os.path.join(SCRATCH, "batch_*.json")) + glob.glob(os.path.join(SCRATCH, "kttt*.json"))):
    txt = open(p, encoding="utf-8").read()
    print(os.path.basename(p), txt.count("(?)"))
