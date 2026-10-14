# -*- coding: utf-8 -*-
import json, re, glob, os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
SCRATCH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data")

# (key, field-path, new value). field-path dùng dấu chấm cho nested.
FIXES = [
    # ---- NVBH (GD02 / freeform) ----
    ("Cầm Thế Khái - Vĩnh Hoàn - 03-09-2026", "nvbh", "Dương Văn Ngọ"),
    ("Lê Quốc Việt - Vy Đức Anh - 31-08-2026", "nvbh", "Vũ Thị Phương Lê"),
    ("Nguyễn Quốc Thắm - Bảo Cường - 04-09-2026", "nvbh", "Tạ Thị Trung Kiên"),
    ("Nguyễn Văn Tuấn - Hữu Hiền - 04-09-2026", "nvbh", "Phạm Thu Hữu"),
    ("Nguyễn Văn Tuấn - Định Nhuận - 05-09-2026", "nvbh", "Đào Xuân Hoá"),
    ("Nguyễn Xuân Thắng - Hoàng Phát - 04-09-2026", "nvbh", "Hoàng Trường Sinh"),
    ("Nguyễn Xuân Thắng - Minh Khuê - 05-09-2026", "nvbh", "Đào Văn Tắc"),
    ("Nguyễn Xuân Trường - AI - 04-09-2026", "nvbh", "Phạm Thị Lưu"),
    ("Phan Ngọc Bích - Hoa Hồng - 04-09-2026", "nvbh", "Đặng Bá Hải"),
    ("Đinh Văn Huấn - Tuấn Anh - 04-09-2026", "nvbh", "Bùi Thái Toàn, Lê Văn Thỏa"),
    ("Bùi Thị Thái Hà - Anh Thư - 07-09-2026", "nvbh", "Nguyễn Thọ Nam"),
    ("Chu Công Kiên - Huy Đạt - 07-09-2026", "nvbh", "Đỗ Ngọc Hậu, Phùng Văn Đại"),
    ("Chu Công Kiên - Mạnh Long - 08-09-2026", "nvbh", "Vũ Xuân Toàn, Kim Đức Tiệp"),
    ("Mai Thị Hạnh - Hiền Tuệ - 10-09-2026", "nvbh", "Đào Thị Diệu, Bùi Thị Thủy Chung"),
    ("Nguyễn Hữu Dực - Ánh Thu - 07-09-2026", "nvbh", "Linh Thị Tình, Đinh Thị Hồng, Nguyễn Văn Mạnh"),
    ("Nguyễn Quốc Thắm - Bảo Cường - 07-09-2026", "nvbh", "Tạ Thị Trung Kiên"),
    ("Nguyễn Tuệ Tuấn - Đức Nam - 07-09-2026", "nvbh", "Đỗ Thị Mến"),
    ("Nguyễn Xuân Trường - Hoa Hùng - 07-09-2026", "nvbh", "Nguyễn Hữu Bằng"),
    ("Nguyễn Đức Tôn - Liêm Xuân việt trì - 08-09-2026", "nvbh", "Nguyễn Xuân Thể"),
    ("Nguyễn Đức Tôn - Liêm Xuân việt trì - 08-09-2026", "tai_npp", "Xuân Liêm"),
    ("Ngô Huy Thanh - Phương Đông - 07-09-2026", "nvbh", "Nguyễn Thị Thu Hương, Phạm Vân Anh, Ngô Thị Ngần, Hà Thị Vinh"),
    ("Trần Đức Phong - Long Giang - 08-09-2026", "nvbh", "Trần Văn Luyến"),
    ("Đàm Ngọc Sơn - Trung Hiếu - 08-09-2026", "nvbh", "Nguyễn Quang Chinh"),
    ("Đào Văn Sơn - Thịnh Long - 07-09-2026", "nvbh", "Đinh Thị Thuỳ Linh"),
    ("Nguyễn Văn Đồng - Minh Lâm - 12-09-2026", "nvbh", "Nguyễn Thị Thu Quyên"),
    ("Trần Đức Phong - Minh Lâm - 11-09-2026", "nvbh", "Nguyễn Thị Thu Quyên"),
    # ---- ASM / GSBH ----
    ("Nguyễn Thị Phương Thúy - Huệ Phát - 04-09-2026", "asm", "Đinh Thị Huệ"),
    ("Nguyễn Thị Phương Thúy - Đặng Tiến Trung - 04-09-2026", "asm", "Đinh Thị Huệ"),
    ("Bùi Thị Thái Hà - Anh Thư - 07-09-2026", "gsbh", "Bùi Thị Thái Hà"),
    ("Nguyễn Phú Lộc - Họp Kinh doanh siêu thị - 07-09-2026", "asm", "Nguyễn Phú Lộc"),
    ("Nguyễn Xuân Trường - Hoa Hùng - 07-09-2026", "asm", "Nguyễn Văn Đồng"),
    ("Trần Đức Phong - Minh Lâm - 11-09-2026", "asm", "Nguyễn Văn Đồng"),
    # ---- các ô không đọc được -> để trống / bỏ ghi chú ----
    ("Lê Quốc Việt - Vy Đức Anh - 31-08-2026", "muc1.ton_kho", ""),
    ("Trần Đức Phong - Long Giang - 08-09-2026", "muc1.ton_kho", ""),
    ("Trần Đức Phong - Long Giang - 08-09-2026", "muc5", "Nhập trước ngày 20 và bán"),
    ("Trần Đức Phong - Minh Lâm - 11-09-2026", "muc1.ton_kho", "15,39 thùng"),
    # ---- KTTT ----
    ("Bùi Thị Thái Hà - Vạn An Lộc - 03-09-2026", "ho_ten", "Nguyễn Văn Đồng"),
    ("Bùi Thị Thái Hà - Vạn An Lộc - 03-09-2026", "lam_viec_voi", "Bùi Thị Thái Hà - GSBH; NVBH Tạ Thị Phương Hoa"),
    ("Chu Công Kiên - Mạnh Long - 04-09-2026", "ho_ten", "Nguyễn Đình Trọng"),
    ("Chu Công Kiên - Mạnh Long - 04-09-2026", "lam_viec_voi", "Chu Công Kiên - GSBH; NVBH Kim Đức Tiệp, Vũ Xuân Toàn"),
    ("Lê Thị Trâm - Minh Anh - 31-08-2026", "chuc_vu", "GSBH"),
    ("Lê Thị Trâm - Tâm An - 03-09-2026", "chuc_vu", "GSBH"),
    ("Lê Thị Trâm - Tâm An - 03-09-2026", "lam_viec_voi", "Mai Thị Tuyết - NVBH - Quảng Xương"),
    ("Lê Thị Trâm - Tâm An - 03-09-2026", "m1", "Hàng ổn định, đáp ứng đủ nhu cầu"),
    ("Nguyễn Đình Trọng - Huy Đạt - 03-09-2026", "lam_viec_voi", "Chu Công Kiên - GSBH; NVBH Phùng Văn Đại, Đỗ Ngọc Hậu"),
    ("Nhữ Hồng Huy - Minh Phúc HD - 03-09-2026", "lam_viec_voi", "Đỗ Thị Tâm - Đại diện NPP; NVBH Lê Quý Quỳnh, Lê Văn Tuyến"),
    ("Bùi Thị Thái Hà - Hùng Phát lẻ - 08-09-2026", "ho_ten", "Bùi Thị Thái Hà"),
    ("Cầm Thế Khái - Vĩnh Hoàn - 10-09-2026", "lam_viec_voi", "NVBH Nông Thị Nhung, Dương Văn Ngọ"),
    ("Lê Thị Trâm - Tâm An - 11-09-2026", "lam_viec_voi", "Mai Thị Tuyết - NVBH - Quảng Xương"),
    ("Nguyễn Tuệ Tuấn - Mạnh Lực - 11-09-2026", "lam_viec_voi", "ASM Nguyễn Đình Trọng; chị Trang - chủ NPP; NVBH Phạm Thị Hảo"),
    ("Nguyễn Xuân Thắng - Hoa Thắng - 08-09-2026", "lam_viec_voi", "Hồ Văn Sang - NVBH khu vực Hưng Yên"),
    ("Nguyễn Xuân Thắng - Hoàng Phát - 07-09-2026", "lam_viec_voi", "Hoàng Trường Sinh - NVBH khu vực Hưng Yên"),
    ("Nhữ Hồng Huy - Cường Long - 10-09-2026", "lam_viec_voi", "NPP Cường Long - Ngô Thị Chi; NVBH Nguyễn Đình Phương"),
    ("Trần Hữu Đoàn - Mạnh Nam - 10-09-2026", "lam_viec_voi", "NVBH Nguyễn Thị Thạo"),
    ("Bùi Thị Thái Hà - Vạn An Lộc - 11-09-2026", "ho_ten", "Bùi Thị Thái Hà"),
    ("Bùi Thị Thái Hà - Vạn An Lộc - 11-09-2026", "lam_viec_voi", "NVBH Tạ Thị Phương Hoa - Long Biên"),
    ("Lê Văn Thế - Phương Nam - 11-09-2026", "lam_viec_voi", "NVBH Hà Thị Dung - Khu vực Thọ Xuân"),
    ("Lại Văn Thịnh - Ngọc Bích - 12-09-2026", "lam_viec_voi", "Vũ Tiến Công (NPP); NVBH Trần Đức Quang, Nguyễn Văn Thắng - Hà Nam"),
    ("Nguyễn Thị Phương Thúy - Đặng Tiến Trung - 12-09-2026", "lam_viec_voi", "NPP Đặng Tiến Trung; NVBH Trần Thị Thu Hiền"),
    ("Nguyễn Văn Tuấn - Định Nhuận - 12-09-2026", "lam_viec_voi", "NVBH Đào Xuân Hoá"),
    ("Nguyễn Xuân Thắng - Minh Khuê - 12-09-2026", "lam_viec_voi", "NVBH Đào Văn Tắc"),
    ("Nhữ Hồng Huy - Minh Phúc HD - 11-09-2026", "lam_viec_voi", "Đỗ Thị Tâm - Đại diện NPP; NVBH Lê Quý Quỳnh, Lê Văn Tuyến"),
]

NOTE_PREFIX = re.compile(r"^\s*\[[^\]]*\]\s*")          # "[Ảnh dùng mẫu ...] " ở đầu
NOTE_TRAILING = re.compile(r"\s*[-–—]\s*(phần còn lại )?(số liệu )?không đọc rõ.*$", re.I)
UNREADABLE_ONLY = re.compile(r"^\s*(\(?không đọc được\)?|KHÔNG ĐỌC ĐƯỢC.*|Không đọc rõ số liệu|\(\?\))\s*$", re.I)

def clean_str(s):
    if not isinstance(s, str):
        return s
    if UNREADABLE_ONLY.match(s):
        return ""
    s = s.replace(" (?)", "").replace("(?)", "")
    s = NOTE_PREFIX.sub("", s)
    s = NOTE_TRAILING.sub("", s)
    s = re.sub(r"\s{2,}", " ", s).strip()
    s = re.sub(r"\s+([,.;:])", r"\1", s)
    return s

def clean(obj):
    if isinstance(obj, dict):
        return {k: clean(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [clean(v) for v in obj]
    return clean_str(obj)

files = sorted(glob.glob(os.path.join(SCRATCH, "batch_*.json")) + glob.glob(os.path.join(SCRATCH, "kttt*.json")))
applied = set()
for p in files:
    items = json.load(open(p, encoding="utf-8"))
    for it in items:
        for key, path, val in FIXES:
            if it["key"] != key:
                continue
            parts = path.split(".")
            tgt = it
            for part in parts[:-1]:
                tgt = tgt.setdefault(part, {})
            # kttt files chỉ nhận field kttt, batch files chỉ nhận field gd02
            if parts[0] in it or parts[0] in ("muc1", "muc2"):
                tgt[parts[-1]] = val
                applied.add((key, path))
    items = clean(items)
    json.dump(items, open(p, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

missed = [(k, f) for k, f, _ in FIXES if (k, f) not in applied]
print("applied:", len(applied), "missed:", missed)
left = 0
for p in files:
    left += open(p, encoding="utf-8").read().count("(?)")
print("remaining (?):", left)
