# -*- coding: utf-8 -*-
import json, os, sys, datetime
from collections import OrderedDict
import openpyxl
from PIL import Image
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

import argparse, re, unicodedata
def _sd(x): return "".join(c for c in unicodedata.normalize("NFD", x or "") if unicodedata.category(c) != "Mn").lower()
try:
    _ns = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data", "nhansu2_raw.json"), encoding="utf-8"))
    NVBH_ROSTER = sorted({r[1].strip() for r in _ns["NVBH"][3:] if r[1]}, key=len, reverse=True)
except Exception:
    NVBH_ROSTER = []
_ap = argparse.ArgumentParser(description="Ghép file Word biên bản họp theo tuần")
_ap.add_argument("--start", required=True, help="dd-mm-yyyy (ngày đầu tuần)")
_ap.add_argument("--end", required=True, help="dd-mm-yyyy (ngày cuối tuần)")
_ap.add_argument("--xlsx", default=None, help="file tong-hop-bien-ban-hop tuần đó (sheet Tổng hợp + Thống kê)")
_ap.add_argument("--out", required=True, help="đường dẫn file .docx xuất ra")
_ap.add_argument("--only-keys", default=None, help="chỉ in chi tiết (gõ lại + ảnh) cho các key này, cách nhau bằng ';' — bảng thống kê ngày vẫn đủ toàn bộ")
_args = _ap.parse_args()
SOURCE_XLSX = _args.xlsx
ONLY_KEYS = set(_args.only_keys.split(";")) if _args.only_keys else None

WEEKDAY_VN = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy", "Chủ Nhật"]

def date_key(dd, mm, yyyy):
    return datetime.date(int(yyyy), int(mm), int(dd))

def vn_weekday(dd, mm, yyyy):
    return WEEKDAY_VN[date_key(dd, mm, yyyy).weekday()]

SCRATCH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data")
IMG_DIR = r"C:\Users\desig\Downloads\Anh_bien_ban_da_cat"
MEETINGS_JSON = os.path.join(SCRATCH, "meetings.json")
OUT_PATH = _args.out

KTTT_KEYS = {
    "Bùi Thị Thái Hà - Anh Thư - 04-09-2026",
    "Bùi Thị Thái Hà - Hùng Phát lẻ - 31-08-2026",
    "Bùi Thị Thái Hà - Vạn An Lộc - 03-09-2026",
    "Chu Công Kiên - Mạnh Long - 04-09-2026",
    "Lê Quốc Việt - An Vy - 03-09-2026",
    "Lê Thị Trâm - Minh Anh - 31-08-2026",
    "Lê Thị Trâm - Tâm An - 03-09-2026",
    "Nguyễn Văn Đồng - Vạn An Lộc - 03-09-2026",
    "Nguyễn Đình Trọng - Huy Đạt - 03-09-2026",
    "Nguyễn Đình Trọng - Mạnh Long - 04-09-2026",
    "Nhữ Hồng Huy - Minh Phúc HD - 03-09-2026",
    "Trần Ngọc Lâm - Phương Anh - 04-09-2026",
    "Đinh Thị Huệ - Huệ Phát - 04-09-2026",
    "Đinh Thị Huệ - Đặng Tiến Trung - 04-09-2026",
    # đợt 2 (07-11/09/2026)
    "Bùi Thị Thái Hà - Hùng Phát lẻ - 08-09-2026",
    "Cầm Thế Khái - Vĩnh Hoàn - 10-09-2026",
    "Lê Quốc Việt - Vy Đức Anh - 07-09-2026",
    "Lê Thị Trâm - Minh Anh - 07-09-2026",
    "Lê Thị Trâm - Tâm An - 11-09-2026",
    "Lê Văn Thế - Hồng Minh - 10-09-2026",
    "Nguyễn Hữu Dực - Thi Hương - 11-09-2026",
    "Nguyễn Quốc Thắm - Linh Anh - 11-09-2026",
    "Nguyễn Thị Phương Thúy - Huệ Phát - 11-09-2026",
    "Nguyễn Tuệ Tuấn - Mạnh Lực - 11-09-2026",
    "Nguyễn Văn Toàn - An Viên Phát - 10-09-2026",
    "Nguyễn Văn Tuấn - Hữu Hiền - 11-09-2026",
    "Nguyễn Xuân Thắng - Hoa Thắng - 08-09-2026",
    "Nguyễn Xuân Thắng - Hoàng Phát - 07-09-2026",
    "Nguyễn Đình Trọng - Mạnh Lực - 11-09-2026",
    "Nguyễn Đình Trọng - Phương Anh - 07-09-2026",
    "Nguyễn Đình Trọng - Đức Mạnh - 10-09-2026",
    "Nhữ Hồng Huy - Cường Long - 10-09-2026",
    "Trần Hữu Đoàn - Mạnh Nam - 10-09-2026",
    "Trần Ngọc Lâm - Phương Anh - 07-09-2026",
    "Trần Ngọc Lâm - Đức Mạnh - 10-09-2026",
    "Đinh Thị Huệ - Hà Sơn - 10-09-2026",
    "Đàm Ngọc Sơn - Long Minh Trang - 11-09-2026",
    # đợt 3 (12/09/2026)
    "Bùi Thị Thái Hà - Vạn An Lộc - 11-09-2026",
    "Lê Văn Thế - Phương Nam - 11-09-2026",
    "Lại Văn Thịnh - Ngọc Bích - 12-09-2026",
    "Nguyễn Thị Liên - Hùng Mạnh Lai Châu - 12-09-2026",
    "Nguyễn Thị Phương Thúy - Đặng Tiến Trung - 12-09-2026",
    "Nguyễn Văn Tuấn - Định Nhuận - 12-09-2026",
    "Nguyễn Xuân Thắng - Minh Khuê - 12-09-2026",
    "Nhữ Hồng Huy - Minh Phúc HD - 11-09-2026",
    "Phan Ngọc Bích - Hoa Hồng - 11-09-2026",
    "Phan Ngọc Bích - Minh Hưng - 12-09-2026",
    "Phạm Thanh Sơn - Bảo Thắm - 12-09-2026",
    "Phạm Thanh Sơn - Duyên Hùng - 11-09-2026",
    "Toàn Văn Thuận - Thành Trung LS - 12-09-2026",
    "Trần Hữu Đoàn - Đại Dương - 12-09-2026",
    "Chu Công Kiên - Huy Đạt - 14-09-2026",
    "Chu Công Kiên - Mạnh Long - 15-09-2026",
    "Cầm Thế Khái - Vĩnh Hoàn - 15-09-2026",
    "Lê Thị Trâm - Minh Anh - 15-09-2026",
    "Mai Thị Hạnh - Hiền Tuệ - 14-09-2026",
    "Nguyễn Quốc Thắm - Bảo Cường - 14-09-2026",
    "Nguyễn Quốc Thắm - Hồng Hạnh - 15-09-2026",
    "Nguyễn Thị Phương Thúy - Hương Nhài - 14-09-2026",
    "Nguyễn Văn Toàn - An Viên Phát - 14-09-2026",
    "Nguyễn Đức Tôn - Liêm Xuân - 14-09-2026",
    "Ngô Huy Thanh - Phương Đông - 14-09-2026",
    "Phan Thị Thanh Hải - Phượng Hoàn - 15-09-2026",
    "Đàm Ngọc Sơn - Trung Hiếu - 15-09-2026",
    "Đào Văn Sơn - Thịnh Long - 14-09-2026",
    "Đỗ Quang Tý - Hiếu Thảo - 14-09-2026",
    "Lê Văn Thế - Hồng Minh - 16-09-2026",
    "Nguyễn Quốc Thắm - Linh Anh - 16-09-2026",
    "Nguyễn Thị Liên - Hùng Mạnh - 16-09-2026",
    "Trần Ngọc Lâm - Đức Mạnh - 16-09-2026",
    "Trịnh Văn Dũng - Trình Oanh - 15-09-2026",
    "Bùi Thị Thái Hà - Anh Thư - 16-09-2026",
    "Lê Quốc Việt - An Vy - 16-09-2026",
    "Nguyễn Đình Trọng - Thảo Nam - 16-09-2026",
    "Đinh Thị Huệ - An Vy - 16-09-2026",
    "Lê Quốc Việt - Vy Đức Anh - 17-09-2026",
    "Nguyễn Văn Tuấn - Hữu Hiền - 17-09-2026",
    "Nhữ Hồng Huy - Minh Phúc HD - 17-09-2026",
    "Đàm Ngọc Sơn - Long Minh Trang - 17-09-2026",
    "Đinh Thị Huệ - Phương Đông - 17-09-2026",
    "Phan Ngọc Bích - Hoa Hồng - 17-09-2026",
    "Bùi Thị Thái Hà - Hùng Phát lẻ - 17-09-2026",
    "Bùi Thị Thái Hà - Vạn An Lộc - 17-09-2026",
    "Lê Văn Thế - Phương Nam - 18-09-2026",
    "Nguyễn Thị Phương Thúy - Huệ Phát - 18-09-2026",
    "Nguyễn Tuệ Tuấn - Mạnh Lực - 18-09-2026",
    "Trần Ngọc Lâm - Phương Anh - 18-09-2026",
    # đợt 19-09
    "Phạm Thanh Sơn - Duyên Hùng - 18-09-2026",
    "Trần Hữu Đoàn - Đại Dương - 18-09-2026",
    "Đào Văn Sơn - Đức Long - 18-09-2026",
    "Lê Thị Trâm - Tâm An - 19-09-2026",
    "Lại Văn Thịnh - Ngọc Bích - 19-09-2026",
    "Nguyễn Thị Phương Thúy - Đặng Tiến Trung - 19-09-2026",
    "Nguyễn Văn Toàn - Hùng Phát sỉ - 19-09-2026",
    "Nguyễn Văn Tuấn - Định Nhuận - 19-09-2026",
    "Nguyễn Xuân Thắng - Hoàng Phát - 19-09-2026",
    "Nhữ Hồng Huy - Cường Long - 19-09-2026",
    "Phan Ngọc Bích - Minh Hưng - 19-09-2026",
    "Phạm Thanh Sơn - Bảo Thắm - 19-09-2026",
    "Toàn Văn Thuận - Thành Trung LS - 19-09-2026",
    "Trần Hữu Đoàn - Minh Ý - 19-09-2026",
    "Hoàng Thu Thủy - Đức Thủy - 19-09-2026",
    "Nguyễn Tuệ Tuấn - Đức Nam - 19-09-2026",
    "Trần Hữu Đoàn - Mạnh Nam - 19-09-2026",
    # đợt 21/09-22/09
    "Chu Công Kiên - Huy Đạt - 21-09-2026",
    "Nguyễn Quốc Thắm - Bảo Cường - 21-09-2026",
    "Nguyễn Thị Phương Thúy - Hương Nhài - 21-09-2026",
    "Nguyễn Tuệ Tuấn - Đức Nam - 21-09-2026",
    "Nguyễn Văn Tuấn - Hữu Hiền - 21-09-2026",
    "Nguyễn Đức Tôn - liêm Xuân việt trì - 21-09-2026",
    "Ngô Huy Thanh - Phương Đông - 21-09-2026",
    "Phạm Thanh Sơn - Duyên Hùng - 21-09-2026",
    "Trần Hữu Đoàn - Mạnh Nam - 21-09-2026",
    "Trịnh Văn Dũng - Trình Oanh - 21-09-2026",
    "Đàm Ngọc Sơn - Trung Hiếu - 21-09-2026",
    "Đào Văn Sơn - Thịnh Long - 21-09-2026",
    "Đào Xuân Chung - Tùng Anh - 21-09-2026",
    "Chu Công Kiên - Mạnh Long - 22-09-2026",
    "Cầm Thế Khái - Vĩnh Hoàn - 22-09-2026",
    "Lê Quốc Việt - Vy Đức Anh - 22-09-2026",
    "Nguyễn Quốc Thắm - Hồng Hạnh - 22-09-2026",
    "Nguyễn Đức Tôn - Thảo Nam Phú Thọ - 22-09-2026",
    "Đỗ Quang Tý - Hiếu Thảo - 22-09-2026",
    "Nguyễn Văn Toàn - An Viên Phát - 22-09-2026",
    "Lê Thị Trâm - Minh Anh - 23-09-2026",
    "Lê Văn Thế - Hồng Minh - 23-09-2026",
    "Nguyễn Quốc Thắm - Linh Anh - 23-09-2026",
    "Nguyễn Thị Phương Thúy - Đặng Tiến Trung - 23-09-2026",
    "Nhữ Hồng Huy - Minh Phúc HD - 23-09-2026",
    "Trần Ngọc Lâm - Đức Mạnh - 23-09-2026",
    "Hoàng Thu Thủy - Đức Thủy - 26-09-2026",
    "Lê Quốc Việt - An Vy - 24-09-2026",
    "Lê Thị Trâm - Tâm An - 26-09-2026",
    "Lê Văn Thế - Phương Nam - 26-09-2026",
    "Lại Văn Thịnh - Ngọc Bích - 26-09-2026",
    "Nguyễn Thị Liên - Hùng Mạnh - 25-09-2026",
    "Nguyễn Thị Phương Thúy - Huệ Phát - 25-09-2026",
    "Nguyễn Tuệ Tuấn - Mạnh Lực - 25-09-2026",
    "Nguyễn Văn Toàn - Hùng Phát sỉ - 25-09-2026",
    "Nguyễn Văn Tuấn - Định Nhuận - 24-09-2026",
    "Nguyễn Văn Đồng - Minh Ý - 26-09-2026",
    "Nhữ Hồng Huy - Cường Long - 25-09-2026",
    "Phan Ngọc Bích - Minh Hưng - 25-09-2026",
    "Phạm Thanh Sơn - Bảo Thắm - 26-09-2026",
    "Trần Hữu Đoàn - Minh Ý - 26-09-2026",
    "Trần Hữu Đoàn - Đại Dương - 25-09-2026",
    "Trần Ngọc Lâm - Phương Anh - 24-09-2026",
    "Đàm Ngọc Sơn - Long Minh Trang - 24-09-2026",
    "Chu Công Kiên - Huy Đạt - 28-09-2026",
    "Nguyễn Văn Toàn - An Viên Phát - 28-09-2026",
    "Nguyễn Đức Tôn - Liêm Xuân việt trì - 28-09-2026",
    "Ngô Huy Thanh - Phương Đông - 28-09-2026",
    "Đàm Ngọc Sơn - Trung Hiếu - 28-09-2026",
    "Đào Văn Sơn - Thịnh Long - 28-09-2026",
    "Đào Xuân Chung - Tùng Anh - 28-09-2026",
    "Đỗ Quang Tý - Hiếu Thảo - 28-09-2026",
    "Chu Công Kiên - Mạnh Long - 29-09-2026",
    "Cầm Thế Khái - Vĩnh Hoàn - 29-09-2026",
    "Lê Văn Thế - Hồng Minh - 29-09-2026",
    "Nguyễn Văn Tuấn - Hữu Hiền - 28-09-2026",
    "Nguyễn Văn Đồng - Long Minh Trang - 29-09-2026",
    "Phan Ngọc Bích - Minh Hưng - 29-09-2026",
    "Lê Thị Trâm - Minh Anh - 30-09-2026",
    "Nguyễn Văn Toàn - Hùng Phát sỉ - 30-09-2026",
    "Nguyễn Xuân Thắng - Hoàng Phát - 30-09-2026",
    "Phan Ngọc Bích - Hoa Hồng - 30-09-2026",
    "Phạm Thanh Sơn - Duyên Hùng - 30-09-2026",
    "Trần Ngọc Lâm - Đức Mạnh - 30-09-2026",
    "Bùi Thị Thái Hà - Vạn An Lộc - 01-10-2026",
    "Lê Quốc Việt - An Vy - 01-10-2026",
    "Lại Văn Thịnh - Ngọc Bích - 01-10-2026",
    "Nguyễn Thị Phương Thúy - Đặng Tiến Trung - 01-10-2026",
    "Nguyễn Văn Tuấn - Định Nhuận - 01-10-2026",
    "Nguyễn Xuân Thắng - Hoa Thắng - 01-10-2026",
    "Nguyễn Đình Trọng - Trình Oanh - 01-10-2026",
    "Nhữ Hồng Huy - Minh Phúc HD - 01-10-2026",
    "Trịnh Văn Dũng - Trình Oanh - 01-10-2026",
    "Đinh Thị Huệ - An Vy - 01-10-2026",
    "Đàm Ngọc Sơn - Long Minh Trang - 01-10-2026",
    "Hoàng Thu Thủy - Đức Thủy - 02-10-2026",
    "Lê Thị Trâm - Tâm An - 02-10-2026",
    "Lê Văn Thế - Phương Nam - 02-10-2026",
    "Nguyễn Thị Liên - Hùng Mạnh - 02-10-2026",
    "Nguyễn Thị Phương Thúy - Huệ Phát - 02-10-2026",
    "Nguyễn Đình Trọng - Đức Thủy - 02-10-2026",
    "Nhữ Hồng Huy - Cường Long - 02-10-2026",
    "Trần Ngọc Lâm - Phương Anh - 02-10-2026",
    "Nguyễn Tuệ Tuấn - Đức Nam - 02-10-2026",
    "Nguyễn Xuân Trường - Hoa Hùng - 02-10-2026",
    "Phạm Thanh Sơn - Bảo Thắm - 02-10-2026",
    "Toàn Văn Thuận - Thành Trung LS - 02-10-2026",
    "Trần Hữu Đoàn - Minh Ý - 02-10-2026",
    "Đinh Văn Huấn - Tuấn Anh - 02-10-2026",
    "Nguyễn Tuệ Tuấn - Mạnh Lực - 03-10-2026",
    "Nguyễn Xuân Thắng - Minh Khuê - 03-10-2026",
    "Nguyễn Xuân Vinh - Thành Thanh - 03-10-2026",
    "Trần Hữu Đoàn - Đại Dương - 03-10-2026",
    "Chu Công Kiên - Huy Đạt - 05-10-2026",
    "Nguyễn Quốc Thắm - Bảo Cường - 05-10-2026",
    "Nguyễn Thị Phương Thúy - Hương Nhài - 05-10-2026",
    "Nguyễn Thị Phương Thúy - Đặng Tiến Trung - 05-10-2026",
    "Nguyễn Văn Toàn - An Viên Phát - 05-10-2026",
    "Nguyễn Đức Tôn - Liêm Xuân Việt Trì - 05-10-2026",
    "Ngô Huy Thanh - Phương Đông - 05-10-2026",
    "Đào Văn Sơn - Thịnh Long - 05-10-2026",
    "Đào Xuân Chung - Nam Huệ - 05-10-2026",
    "Phạm Thanh Sơn - Duyên Hùng - 05-10-2026",
    "Trần Hữu Đoàn - Mạnh Nam - 05-10-2026",
    "Trịnh Văn Dũng - Trình Oanh - 05-10-2026",
    "Đàm Ngọc Sơn - Trung Hiếu - 05-10-2026",
    "Cầm Thế Khái - Vĩnh Hoàn - 06-10-2026",
    "Hoàng Thu Thủy - Đức Thủy - 06-10-2026",
    "Lê Thị Trâm - Minh Anh - 06-10-2026",
    "Nguyễn Quốc Thắm - Hồng Hạnh - 06-10-2026",
    "Nguyễn Đức Tôn - Thảo Nam Phú Thọ - 06-10-2026",
    "Bùi Thị Thái Hà - Anh Thư - 07-10-2026",
    "Chu Công Kiên - Mạnh Long - 07-10-2026",
    "Nguyễn Quốc Thắm - Linh Anh - 07-10-2026",
    "Trần Đức Phong - Minh Lâm - 07-10-2026",
    "Đỗ Quang Tý - Hiếu Thảo - 07-10-2026",
    "Nguyễn Đình Trọng - Cường Long - 07-10-2026",
    "Nhữ Hồng Huy - Cường Long - 07-10-2026",
    "Lê Quốc Việt - Vy Đức Anh - 07-10-2026",
    "Đào Văn Sơn - Đức Long - 07-10-2026",
    "Lê Thị Trâm - Tâm An - 08-10-2026",
    "Nguyễn Hữu Dực - Thi Hương - 08-10-2026",
    "Nguyễn Hữu Dực - Ánh Thu - 08-10-2026",
    "Nguyễn Thị Phương Thúy - Huệ Phát - 08-10-2026",
    "Nguyễn Xuân Thắng - Hoàng Phát - 08-10-2026",
    "Nguyễn Xuân Trường - AI - 08-10-2026",
    "Trần Ngọc Lâm - Phương Anh - 08-10-2026",
    "Đinh Thị Huệ - Thi Hương - 08-10-2026",
    "Đinh Thị Huệ - Ánh Thu - 08-10-2026",
    "Đinh Văn Huấn - Tuấn Anh - 08-10-2026",
    "Đàm Ngọc Sơn - Long Minh Trang - 08-10-2026",
    "Nhữ Hồng Huy - Minh Phúc HD - 08-10-2026",
    "Bùi Thị Thái Hà - Vạn An Lộc - 09-10-2026",
    "Lê Quốc Việt - An Vy - 09-10-2026",
    "Nguyễn Thị Liên - Hùng Mạnh - 09-10-2026",
    "Nguyễn Văn Tuấn - Định Nhuận - 09-10-2026",
    "Nguyễn Xuân Trường - Hoa Hùng - 09-10-2026",
    "Nguyễn Xuân Vinh - Thành Thanh - 09-10-2026",
    "Nguyễn Đình Trọng - Minh Phúc HD - 09-10-2026",
    "Toàn Văn Thuận - Thành Trung LS - 09-10-2026",
    "Đinh Thị Huệ - Thành Trung LS - 09-10-2026",
    "Nguyễn Xuân Thắng - Hoa Thắng - 09-10-2026",
    "Nguyễn Xuân Trường - Kiên Cường - 09-10-2026",
    "Phan Ngọc Bích - Hoa Hồng - 09-10-2026",
    "Trần Hữu Đoàn - Đại Dương - 09-10-2026",
    "Lại Văn Thịnh - Ngọc Bích - 10-10-2026",
    "Nguyễn Tuệ Tuấn - Mạnh Lực - 10-10-2026",
    "Nguyễn Tuệ Tuấn - Đức Nam - 10-10-2026",
    "Nguyễn Văn Toàn - Hùng Phát sỉ - 10-10-2026",
    "Nguyễn Văn Đồng - Hải Long Giang - 10-10-2026",
    "Nguyễn Xuân Thắng - Minh Khuê - 10-10-2026",
    "Nguyễn Đình Trọng - Đức Mạnh - 10-10-2026",
    "Phan Ngọc Bích - Minh Hưng - 10-10-2026",
    "Phạm Thanh Sơn - Bảo Thắm - 10-10-2026",
    "Trần Hữu Đoàn - Minh Ý - 10-10-2026",
    "Trần Ngọc Lâm - Đức Mạnh - 10-10-2026",
    "Đinh Thị Huệ - Phương Đông - 10-10-2026",
    "Chu Công Kiên - Huy Đạt - 12-10-2026",
    "Nguyễn Hữu Dực - Ánh Thu - 12-10-2026",
    "Nguyễn Quốc Thắm - Bảo Cường - 12-10-2026",
    "Nguyễn Thị Phương Thúy - Hương Nhài - 12-10-2026",
    "Nguyễn Văn Tuấn - Hữu Hiền - 12-10-2026",
    "Nguyễn Đức Tôn - liêm Xuân - 12-10-2026",
    "Ngô Huy Thanh - Phương Đông - 12-10-2026",
    "Đào Văn Sơn - Thịnh Long - 12-10-2026",
    "Đào Xuân Chung - Tùng Anh - 12-10-2026",
    "Đỗ Quang Tý - Hiếu Thảo - 12-10-2026",
    "Nguyễn Thị Phương Thúy - Đặng Tiến Trung - 12-10-2026",
    "Phạm Thanh Sơn - Duyên Hùng - 12-10-2026",
    "Trịnh Văn Dũng - Trình Oanh - 12-10-2026",
    "Đàm Ngọc Sơn - Trung Hiếu - 12-10-2026",
    "Chu Công Kiên - Mạnh Long - 13-10-2026",
    "Cầm Thế Khái - Vĩnh Hoàn - 13-10-2026",
    "Lê Quốc Việt - An Vy - 13-10-2026",
    "Lê Thị Trâm - Minh Anh - 13-10-2026",
    "Nguyễn Quốc Thắm - Hồng Hạnh - 13-10-2026",
    "Nguyễn Văn Toàn - An Viên Phát - 13-10-2026",
    "Nguyễn Xuân Trường - AI - 13-10-2026",
    "Nguyễn Đình Trọng - An Viên Phát - 13-10-2026",
    "Nguyễn Đức Tôn - Thảo Nam Phú Thọ - 13-10-2026",
    "Trần Đức Phong - Minh Lâm - 13-10-2026",
    "Đinh Thị Huệ - An Vy - 13-10-2026",
}

ASM_CHECK_KEYS = {
    # mẫu "BIÊN BẢN KIỂM TRA THỊ TRƯỜNG (ASM, GS, NVBH)" — kiểm tra nội bộ ASM/GS, không gắn NPP cụ thể
    "Mai Thị Hạnh - Phạm Bảo Khánh - 26-09-2026",
    "Đinh Thị Huệ - Hà Sơn - 24-09-2026",
    "Đinh Thị Huệ - Hà Sơn - 26-09-2026",
    "Mai Thị Hạnh - Hiền Tuệ - 21-09-2026",
    "Nguyễn Quốc Thắm - Bảo Cường - 28-09-2026",
    "Nguyễn Quốc Thắm - Hồng Hạnh - 29-09-2026",
    "Đinh Thị Huệ - Hiền Tuệ - 29-09-2026",
    "Nguyễn Quốc Thắm - Linh Anh - 30-09-2026",
    "Đinh Thị Huệ - Phúc Thành Đạt - 30-09-2026",
    "Đinh Thị Huệ - Phượng Hoàn - 30-09-2026",
    "Mai Thị Hạnh - Hiền Tuệ - 02-10-2026",
    "Mai Thị Hạnh - Phạm Bảo Khánh - 03-10-2026",
    "Nguyễn Văn Tuấn - Hữu Hiền - 05-10-2026",
    "Phan Thị Thanh Hải - Phượng Hoàn - 05-10-2026",
    "Mai Thị Hạnh - Hiền Tuệ - 07-10-2026",
    "Phan Thị Thanh Hải - Phúc Thành Đạt - 08-10-2026",
    "Mai Thị Hạnh - Phạm Bảo Khánh - 09-10-2026",
    "Phan Thị Thanh Hải - Phượng Hoàn - 13-10-2026",
}

FREEFORM_KEYS = {
    "Nguyễn Đình Trọng - Đức Mạnh - 31-08-2026",
    "Trần Ngọc Lâm - Đức Mạnh - 31-08-2026",
    "Đinh Thị Huệ - Vy Đức Anh - 31-08-2026",
    "Nguyễn Phú Lộc - Họp Kinh doanh siêu thị - 31-08-2026",
    "Nguyễn Phú Lộc - Họp Kinh doanh siêu thị - 07-09-2026",
    "Đinh Thị Huệ - Hà Sơn - 11-09-2026",
    "Nguyễn Phú Lộc - Họp Kinh doanh siêu thị - 14-09-2026",
    "Đinh Thị Huệ - Ánh Thu - 18-09-2026",
    "Mai Thị Hạnh - Phạm Bảo Khánh - 19-09-2026",
    "Nguyễn Phú Lộc - Họp Kinh doanh siêu thị - 21-09-2026",
    "Đinh Thị Huệ - Vy Đức Anh - 22-09-2026",
    "Nguyễn Đình Trọng - Vĩnh Hoàn - 23-09-2026",
    "Nguyễn Phú Lộc - Họp Kinh doanh siêu thị - 28-09-2026",
}

DUP_DROP_KEYS = {
    "Nguyễn Văn Đồng - Kiên Cường - 31-08-2026",       # trùng Nguyễn Xuân Trường - Kiên Cường (ASM chụp lại)
    "Đinh Thị Huệ - Huệ Phát - 04-09-2026",             # trùng Nguyễn Thị Phương Thúy - Huệ Phát
    "Đinh Thị Huệ - Đặng Tiến Trung - 04-09-2026",      # trùng Nguyễn Thị Phương Thúy - Đặng Tiến Trung
    "Nguyễn Văn Đồng - Bảo Cường - 04-09-2026",         # trùng Nguyễn Quốc Thắm - Bảo Cường
    "Nguyễn Văn Đồng - Vạn An Lộc - 03-09-2026",        # trùng Bùi Thị Thái Hà - Vạn An Lộc
    "Nguyễn Đình Trọng - Mạnh Long - 04-09-2026",       # trùng Chu Công Kiên - Mạnh Long
    "Nguyễn Đình Trọng - Đức Mạnh - 31-08-2026",        # trùng Trần Ngọc Lâm - Đức Mạnh
    "Đinh Thị Huệ - Vy Đức Anh - 31-08-2026",           # trùng Lê Quốc Việt - Vy Đức Anh
    # đợt 2 (07-11/09/2026) — cả 3 đều là Nguyễn Đình Trọng (ASM) chụp lại đúng biên bản GSBH đã gửi
    "Nguyễn Đình Trọng - Mạnh Lực - 11-09-2026",        # trùng Nguyễn Tuệ Tuấn - Mạnh Lực
    "Nguyễn Đình Trọng - Phương Anh - 07-09-2026",      # trùng Trần Ngọc Lâm - Phương Anh
    "Nguyễn Đình Trọng - Đức Mạnh - 10-09-2026",        # trùng Trần Ngọc Lâm - Đức Mạnh
    "Nguyễn Văn Đồng - Hoa Thắng - 15-09-2026",
    "Nguyễn Văn Đồng - Hoa Thắng - 16-09-2026",
    "Nguyễn Văn Đồng - Minh Khuê - 17-09-2026",       # trùng Nguyễn Xuân Thắng - Minh Khuê (ASM chụp lại)
    "Nguyễn Văn Đồng - Kiên Cường - 19-09-2026",      # trùng Nguyễn Xuân Trường - Kiên Cường (ASM chụp lại)
    "Nguyễn Văn Đồng - Hùng Phát - 22-09-2026",       # trùng Bùi Thị Thái Hà - Hùng Phát lẻ (ASM chụp lại)
    "Nguyễn Văn Đồng - Minh Ý - 26-09-2026",          # trùng Trần Hữu Đoàn - Minh Ý (ASM chụp lại đúng ảnh của GSBH)
    "Nguyễn Văn Đồng - Anh Thư - 07-10-2026",         # ASM chụp lại tờ của Bùi Thị Thái Hà
}

with open(MEETINGS_JSON, "r", encoding="utf-8") as f:
    all_meetings = json.load(f)
meetings = [m for m in all_meetings if m["key"] not in DUP_DROP_KEYS]
_d0 = datetime.datetime.strptime(_args.start, "%d-%m-%Y").date()
_d1 = datetime.datetime.strptime(_args.end, "%d-%m-%Y").date()
meetings = [m for m in meetings if _d0 <= date_key(m["dd"], m["mm"], m["yyyy"]) <= _d1]
NOTES = {
    "Nhữ Hồng Huy - Minh Phúc HD - 08-10-2026": "Giấy ghi ngày 07/10/2026 (tên NPP trên giấy ghi Cường Long) - chấp nhận tính họp NPP Minh Phúc HD",
}
FIRST_KEYS = {
    "Lê Văn Thế - Hồng Minh - 06-10-2026",
    "Bùi Thị Thái Hà - Hùng Phát lẻ - 05-10-2026",
    "Phạm Thanh Sơn - Duyên Hùng - 05-10-2026",
    "Trần Hữu Đoàn - Mạnh Nam - 05-10-2026",
    "Trịnh Văn Dũng - Trình Oanh - 05-10-2026",
    "Đàm Ngọc Sơn - Trung Hiếu - 05-10-2026",
}
meetings.sort(key=lambda m: (int(m["yyyy"]), int(m["mm"]), int(m["dd"]), 0 if m["key"] in FIRST_KEYS else 1, m["nguoi_gui"]))
print(f"Meetings: {len(all_meetings)} total, {len(meetings)} after removing {len(DUP_DROP_KEYS)} duplicates, sorted by date")

import glob
data_by_key = {}
for p in sorted(glob.glob(os.path.join(SCRATCH, "batch_*.json"))):
    with open(p, "r", encoding="utf-8") as f:
        items = json.load(f)
    for it in items:
        data_by_key[it["key"]] = it

kttt_by_key = {}
for p in sorted(glob.glob(os.path.join(SCRATCH, "kttt*.json"))):
    with open(p, "r", encoding="utf-8") as f:
        items = json.load(f)
    for it in items:
        kttt_by_key[it["key"]] = it

missing = [m["key"] for m in meetings if m["key"] not in data_by_key and m["key"] not in kttt_by_key]
if missing:
    print("MISSING TRANSCRIPTS:", missing)
    sys.exit(1)

missing_kttt = [m["key"] for m in meetings if m["key"] in KTTT_KEYS and m["key"] not in kttt_by_key]
if missing_kttt:
    print("MISSING KTTT TRANSCRIPTS:", missing_kttt)
    sys.exit(1)

def set_font(run, size=13, bold=False, name="Times New Roman"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = name
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'), {})
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), name)

def add_p(doc, text="", size=13, bold=False, align=None, space_after=6, underline=False):
    p = doc.add_paragraph()
    if align:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    r = p.add_run(text)
    set_font(r, size=size, bold=bold)
    r.underline = underline
    return p

def label_line(doc, label, value, label_bold=True, indent=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    r1 = p.add_run(label)
    set_font(r1, size=13, bold=label_bold)
    r2 = p.add_run(value if value else "")
    set_font(r2, size=13, bold=False)
    return p

def sub_field(doc, label, value):
    label_line(doc, "+ " + label, value, label_bold=False, indent=0.8)

def render_gd02(doc, m, d):
    ngay = d.get("ngay") or m["dd"]
    thang = d.get("thang") or m["mm"]
    nam = d.get("nam") or m["yyyy"]

    add_p(doc, "BIÊN BẢN HỌP GĐKD VÀ NHÀ PHÂN PHỐI", size=15, bold=True,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    add_p(doc, f"Hôm nay, ngày {ngay} tháng {thang} năm {nam}", size=13)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("Tại NPP: ")
    set_font(r, bold=True)
    r = p.add_run(d.get("tai_npp") or "")
    set_font(r)
    r = p.add_run("          địa bàn: ")
    set_font(r, bold=True)
    r = p.add_run(d.get("dia_ban") or "")
    set_font(r)

    label_line(doc, "ASM: ", d.get("asm") or "")
    label_line(doc, "GSBH: ", d.get("gsbh") or "")
    label_line(doc, "NVBH: ", d.get("nvbh") or "")

    add_p(doc, "NỘI DUNG HỌP", size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
    add_p(doc, "Ý KIẾN TRAO ĐỔI:", size=13, bold=True, underline=True, space_after=6)

    add_p(doc, "1. Trao đổi với nhà phân phối", size=13, bold=True, space_after=4)
    m1 = d.get("muc1", {}) or {}
    sub_field(doc, "Tồn kho: ", m1.get("ton_kho") or "")
    sub_field(doc, "Thị trường: ", m1.get("thi_truong") or "")
    sub_field(doc, "Nhân sự: ", m1.get("nhan_su") or "")
    sub_field(doc, "Khác: ", m1.get("khac") or "")

    add_p(doc, "2. NVBH:", size=13, bold=True, space_after=4)
    m2 = d.get("muc2", {}) or {}
    sub_field(doc, "Hàng hóa: ", m2.get("hang_hoa") or "")
    sub_field(doc, "Thị trường: ", m2.get("thi_truong") or "")
    sub_field(doc, "Chính sách: ", m2.get("chinh_sach") or "")
    sub_field(doc, "Báo cáo: ", m2.get("bao_cao") or "")
    sub_field(doc, "Hỗ trợ của ASM-GSBH: ", m2.get("ho_tro") or "")

    add_p(doc, "3. ASM-GSBH", size=13, bold=True, space_after=4)
    add_p(doc, d.get("muc3") or "", size=13, space_after=10)

    add_p(doc, "4. Ý kiến chung", size=13, bold=True, space_after=4)
    add_p(doc, d.get("muc4") or "", size=13, space_after=10)

    add_p(doc, "5. Kế hoạch nhập - bán hàng và định hướng:", size=13, bold=True, space_after=4)
    add_p(doc, d.get("muc5") or "", size=13, space_after=10)

def render_kttt(doc, m, k):
    ngay = k.get("ngay") or m["dd"]
    thang = k.get("thang") or m["mm"]
    nam = k.get("nam") or m["yyyy"]

    add_p(doc, "BIÊN BẢN KIỂM TRA THỊ TRƯỜNG NPP", size=15, bold=True,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_p(doc, f"Ngày {ngay} tháng {thang} năm {nam}", size=13,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)

    label_line(doc, "Họ và tên: ", (k.get("ho_ten") or "") + (
        "   - Chức vụ: " + k["chuc_vu"] if k.get("chuc_vu") else ""))
    add_p(doc, "Làm việc với", size=13, bold=True, space_after=4)
    label_line(doc, "Họ và tên - Chức vụ - Địa bàn quản lý: ", k.get("lam_viec_voi") or "")

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run("NPP: ")
    set_font(r, bold=True)
    r = p.add_run(k.get("npp") or "")
    set_font(r)
    r = p.add_run("      Địa chỉ: ")
    set_font(r, bold=True)
    r = p.add_run(k.get("dia_chi") or "")
    set_font(r)

    add_p(doc, "*Các bên tham gia đã thống nhất các nội dung làm việc như sau:", size=13, space_after=8)

    items = [
        ("1. Về tình hình thị trường chung: ", k.get("m1")),
        ("2. Về doanh số nhập - bán: ", k.get("m2")),
        ("3. Về chương trình khuyến mại: ", k.get("m3")),
        ("4. Về hàng trôi: ", k.get("m4")),
        ("5. Các mặt hàng/sản phẩm chính NPP đang bán: ", k.get("m5")),
        ("6. Đánh giá của NPP về nhân sự (NV-GSBH): ", k.get("m6")),
        ("7. Các ý kiến đánh giá, trao đổi khác: ", k.get("m7")),
        ("8. Các đề xuất của NPP: ", k.get("m8")),
        ("9. Các nội dung khác: ", k.get("m9")),
    ]
    for label, val in items:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(8)
        r1 = p.add_run(label)
        set_font(r1, bold=True)
        r2 = p.add_run(val or "")
        set_font(r2)

def render_asm_check(doc, m, k):
    ngay = k.get("ngay") or m["dd"]
    thang = k.get("thang") or m["mm"]
    nam = k.get("nam") or m["yyyy"]

    add_p(doc, "BIÊN BẢN KIỂM TRA THỊ TRƯỜNG (ASM, GS, NVBH)", size=15, bold=True,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_p(doc, f"Ngày {ngay} tháng {thang} năm {nam}", size=13,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)

    label_line(doc, "Họ và tên: ", (k.get("ho_ten") or "") + (
        "   - Chức vụ: " + k["chuc_vu"] if k.get("chuc_vu") else ""))
    add_p(doc, "Làm việc với", size=13, bold=True, space_after=4)
    label_line(doc, "Họ và tên - Chức vụ - Địa bàn quản lý: ", k.get("lam_viec_voi") or "")

    add_p(doc, "*Các bên tham gia đã thống nhất các nội dung làm việc như sau:", size=13, space_after=8)

    items = [
        ("1. Về nhạc chờ Minh Trung: ", k.get("m1")),
        ("2. Về đồng phục Minh Trung: ", k.get("m2")),
        ("3. Về hình đại diện zalo: ", k.get("m3")),
        ("4. Về đăng ký các kênh youtube Tập đoàn: ", k.get("m4")),
        ("5. Về vị trí làm việc đúng theo báo cáo tuần: ", k.get("m5")),
        ("6. Về thực hiện các quy định khác của Tập đoàn: ", k.get("m6")),
        ("7. Về tình hình thị trường: ", k.get("m7")),
        ("8. Về nhân sự (khi làm việc với GS, ASM): ", k.get("m8")),
        ("9. Về NPP: ", k.get("m9")),
        ("10. Các đề xuất: ", k.get("m10")),
        ("11. Các nội dung khác: ", k.get("m11")),
    ]
    for label, val in items:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(8)
        r1 = p.add_run(label)
        set_font(r1, bold=True)
        r2 = p.add_run(val or "")
        set_font(r2)

def render_freeform(doc, m, d):
    ngay = d.get("ngay") or m["dd"]
    thang = d.get("thang") or m["mm"]
    nam = d.get("nam") or m["yyyy"]

    add_p(doc, "BIÊN BẢN HỌP GĐKD VÀ NHÀ PHÂN PHỐI", size=15, bold=True,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    add_p(doc, f"Hôm nay, ngày {ngay} tháng {thang} năm {nam}", size=13)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("Tại NPP: ")
    set_font(r, bold=True)
    r = p.add_run(d.get("tai_npp") or "")
    set_font(r)
    r = p.add_run("          địa bàn: ")
    set_font(r, bold=True)
    r = p.add_run(d.get("dia_ban") or "")
    set_font(r)

    label_line(doc, "ASM: ", d.get("asm") or "")
    label_line(doc, "GSBH: ", d.get("gsbh") or "")
    label_line(doc, "NVBH: ", d.get("nvbh") or "")

    add_p(doc, "NỘI DUNG HỌP", size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)

    m1 = d.get("muc1", {}) or {}
    m2 = d.get("muc2", {}) or {}
    pieces = []
    for v in (m1.get("khac"), m1.get("ton_kho"), m1.get("thi_truong"), m1.get("nhan_su")):
        if v:
            pieces.append(v)
    for v in (m2.get("hang_hoa"), m2.get("thi_truong"), m2.get("chinh_sach"), m2.get("bao_cao")):
        if v:
            pieces.append(v)
    for v in (d.get("muc3"), d.get("muc4"), d.get("muc5")):
        if v:
            pieces.append(v)
    for text in pieces:
        add_p(doc, text, size=13, space_after=8)

doc = Document()
sec = doc.sections[0]
sec.left_margin = Cm(2.5)
sec.right_margin = Cm(2)
sec.top_margin = Cm(2)
sec.bottom_margin = Cm(2)

def set_cell_text(cell, text, bold=False, size=12, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(2)
    if align:
        p.alignment = align
    r = p.add_run(text if text is not None else "")
    set_font(r, size=size, bold=bold)

def add_table(doc, headers, rows, col_widths_cm):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        set_cell_text(hdr_cells[i], h, bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER)
    for row_vals in rows:
        cells = table.add_row().cells
        for i, v in enumerate(row_vals):
            align = WD_ALIGN_PARAGRAPH.CENTER if i == 0 else None
            set_cell_text(cells[i], str(v) if v is not None else "", size=12, align=align)
    for i, w_cm in enumerate(col_widths_cm):
        for row in table.rows:
            row.cells[i].width = Cm(w_cm)
    return table

def add_grouped_table(doc, headers, group_rows, col_widths_cm):
    """group_rows: list of (stt, ho_ten, [ (npp, dia_chi), ... ]) — STT/HỌ TÊN
    merged vertically across the person's NPP rows, matching mẫu Tổng hợp."""
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        set_cell_text(hdr_cells[i], h, bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER)
    for stt, ho_ten, npp_list in group_rows:
        start_row_idx = len(table.rows)
        for j, (npp, dia_chi) in enumerate(npp_list):
            cells = table.add_row().cells
            set_cell_text(cells[0], str(stt) if j == 0 else "", size=12, align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_text(cells[1], ho_ten if j == 0 else "", size=12)
            set_cell_text(cells[2], npp, size=12)
            set_cell_text(cells[3], dia_chi, size=12)
        if len(npp_list) > 1:
            first_row = table.rows[start_row_idx]
            last_row = table.rows[start_row_idx + len(npp_list) - 1]
            first_row.cells[0].merge(last_row.cells[0])
            first_row.cells[1].merge(last_row.cells[1])
    for i, w_cm in enumerate(col_widths_cm):
        for row in table.rows:
            row.cells[i].width = Cm(w_cm)
    return table

def load_full_week_detail(path, nvbh_lookup):
    """Đọc trọn sheet 'Tổng hợp' (STT/HỌ TÊN/NPP/ĐỊA CHỈ/Ngày họp/NVBH/GHI CHÚ) —
    bảng chi tiết đầy đủ từng NPP từng GSBH trong tuần, giữ nguyên nhóm gộp
    STT/HỌ TÊN như file nguồn, kể cả các NPP chưa có ngày họp."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Tổng hợp"]
    title = ws.cell(row=1, column=1).value or ""
    group_rows = []
    cur_ho_ten = None
    for r in range(3, ws.max_row + 1):
        stt = ws.cell(row=r, column=1).value
        ho_ten = ws.cell(row=r, column=2).value
        npp = ws.cell(row=r, column=3).value
        dia_chi = ws.cell(row=r, column=4).value
        ngay = ws.cell(row=r, column=5).value or ""
        ghi_chu = ws.cell(row=r, column=6).value or ""
        if npp is None and stt is None and ho_ten is None:
            continue
        if ho_ten is not None:
            cur_ho_ten = ho_ten
        if stt is not None:
            group_rows.append([stt, ho_ten, []])
        nvbh = nvbh_lookup.get((cur_ho_ten, npp), "")
        group_rows[-1][2].append((npp, dia_chi, ngay, nvbh, ghi_chu))
    return group_rows, title

def add_grouped_table_6col(doc, headers, group_rows, col_widths_cm):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        set_cell_text(hdr_cells[i], h, bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    for stt, ho_ten, sub_rows in group_rows:
        start_row_idx = len(table.rows)
        for j, (npp, dia_chi, ngay, nvbh, ghi_chu) in enumerate(sub_rows):
            cells = table.add_row().cells
            set_cell_text(cells[0], str(stt) if j == 0 else "", size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_text(cells[1], ho_ten if j == 0 else "", size=10)
            set_cell_text(cells[2], npp, size=10)
            set_cell_text(cells[3], dia_chi, size=10)
            set_cell_text(cells[4], ngay, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_text(cells[5], nvbh, size=10)
            set_cell_text(cells[6], ghi_chu, size=10)
        if len(sub_rows) > 1:
            first_row = table.rows[start_row_idx]
            last_row = table.rows[start_row_idx + len(sub_rows) - 1]
            first_row.cells[0].merge(last_row.cells[0])
            first_row.cells[1].merge(last_row.cells[1])
    for i, w_cm in enumerate(col_widths_cm):
        for row in table.rows:
            row.cells[i].width = Cm(w_cm)
    return table

def load_npp_chua_hop(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Thống kê"]
    header_idx = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=1).value == "NPP CHƯA HỌP":
            header_idx = r
            break
    if header_idx is None:
        return [], None, {}
    date_range = ws.cell(row=2, column=1).value or ""
    stats = {}
    for r in range(4, 10):
        label = ws.cell(row=r, column=1).value
        val = ws.cell(row=r, column=2).value
        if label:
            stats[label] = val
    rows = []
    r = header_idx + 2
    while r <= ws.max_row and ws.cell(row=r, column=1).value:
        rows.append([ws.cell(row=r, column=c).value for c in range(1, 4)])
        r += 1
    group_rows = []
    stt = 0
    for ho_ten, npp, dia_chi in rows:
        if group_rows and group_rows[-1][1] == ho_ten:
            group_rows[-1][2].append((npp, dia_chi))
        else:
            stt += 1
            group_rows.append([stt, ho_ten, [(npp, dia_chi)]])
    return group_rows, date_range, stats

def get_nvbh(m):
    key = m["key"]
    if key in KTTT_KEYS:
        # mẫu Kiểm tra thị trường: NV đi kèm nằm ở dòng "Làm việc với" (bỏ tiền tố NVBH/NV)
        lv = (kttt_by_key.get(key) or {}).get("lam_viec_voi") or ""
        found = [n for n in NVBH_ROSTER if _sd(n) in _sd(lv)]
        return ", ".join(found)
    d = data_by_key.get(key) or {}
    return d.get("nvbh") or ""

by_date = OrderedDict()
for m in meetings:
    by_date.setdefault((m["dd"], m["mm"], m["yyyy"]), []).append(m)
dates_sorted = list(by_date.keys())

# ---------- trang tổng kết cả tuần (đầu tài liệu) ----------
add_p(doc, "TỔNG KẾT TUẦN BIÊN BẢN HỌP KD", size=17, bold=True,
      align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
first_d, last_d = dates_sorted[0], dates_sorted[-1]
add_p(doc, f"Từ ngày {first_d[0]}/{first_d[1]}/{first_d[2]} đến ngày {last_d[0]}/{last_d[1]}/{last_d[2]}",
      size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=16)

week_rows = []
total_all = 0
for k in dates_sorted:
    cnt = len(by_date[k])
    total_all += cnt
    week_rows.append((f"{k[0]}/{k[1]}/{k[2]}", vn_weekday(*k), cnt))
week_rows.append(("Tổng cộng", "", total_all))
add_table(doc, ["Ngày", "Thứ", "Số biên bản"], week_rows, [4, 5, 4])
add_p(doc, "")
add_p(doc, f"Tổng số biên bản trong tuần: {total_all}", size=13, bold=True, space_after=4)
add_p(doc, f"Số ngày có biên bản: {len(dates_sorted)}", size=13, bold=True, space_after=4)
doc.add_page_break()

# ---------- trang chi tiết đầy đủ cả tuần (đúng mẫu "Tổng hợp", không viết tắt) ----------
nvbh_lookup = {}
for m in all_meetings:
    nvbh_lookup[(m["nguoi_gui"], m["npp_ten"])] = get_nvbh(m)

if SOURCE_XLSX:
  week_detail_groups, week_detail_title = load_full_week_detail(SOURCE_XLSX, nvbh_lookup)
  add_p(doc, week_detail_title.split("\n")[0] if week_detail_title else "TỔNG HỢP BIÊN BẢN HỌP KD",
        size=15, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
  if week_detail_title and "\n" in week_detail_title:
      add_p(doc, week_detail_title.split("\n", 1)[1], size=12,
            align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
  else:
      add_p(doc, "", size=12, space_after=12)
  add_grouped_table_6col(doc, ["STT", "HỌ TÊN", "NPP", "ĐỊA CHỈ", "Ngày họp", "NVBH đi kèm", "GHI CHÚ"],
                          week_detail_groups, [0.9, 2.6, 2.4, 2.4, 2.7, 2.9, 2.6])
  doc.add_page_break()

  # ---------- trang "TỔNG HỢP DANH SÁCH CHƯA GỬI BIÊN BẢN HỌP TẠI NPP" (đúng mẫu) ----------
  chua_hop_groups, chua_hop_date_range, week_stats = load_npp_chua_hop(SOURCE_XLSX)
  add_p(doc, "TỔNG HỢP DANH SÁCH CHƯA GỬI BIÊN BẢN HỌP TẠI NPP", size=16, bold=True,
        align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
  add_p(doc, f"- {chua_hop_date_range}" if chua_hop_date_range else "", size=12,
        align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14)
  if chua_hop_groups:
      add_grouped_table(doc, ["STT", "HỌ TÊN", "NPP", "ĐỊA CHỈ"], chua_hop_groups, [1.8, 5, 5, 5])
  else:
      add_p(doc, "(Không có NPP nào chưa gửi biên bản trong tuần)", size=12)
  add_p(doc, "")
  if week_stats:
      for label in ("Tổng NPP trong tuyến", "NPP đã họp", "NPP chưa họp", "Người đã có biên bản",
                    "Số biên bản trong khoảng", "Biên bản chưa khớp tuyến"):
          if label in week_stats:
              add_p(doc, f"{label}: {week_stats[label]}", size=12, space_after=2)
  doc.add_page_break()

# ---------- từng ngày: trang tổng kết ngày + các biên bản trong ngày ----------
rendered_meetings = [m for m in meetings if ONLY_KEYS is None or m["key"] in ONLY_KEYS]
last_meeting_overall = rendered_meetings[-1] if rendered_meetings else None
for k in dates_sorted:
    group = by_date[k]
    add_p(doc, f"BIÊN BẢN HỌP KD NGÀY {k[0]}/{k[1]}/{k[2]} ({vn_weekday(*k)})", size=16, bold=True,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=16)
    day_rows = [(i, m["nguoi_gui"], m["npp_ten"], get_nvbh(m), NOTES.get(m["key"], "")) for i, m in enumerate(group, start=1)]
    add_table(doc, ["STT", "NGƯỜI GỬI (GSBH/ASM)", "NPP", "NVBH ĐI KÈM", "GHI CHÚ"], day_rows, [1.3, 4, 3.7, 4.3, 3])
    add_p(doc, "")
    add_p(doc, f"Tổng số biên bản ngày {k[0]}/{k[1]}/{k[2]}: {len(group)}", size=13, bold=True, space_after=0)
    doc.add_page_break()

    render_group = [m for m in group if ONLY_KEYS is None or m["key"] in ONLY_KEYS]
    for m in render_group:
        key = m["key"]
        d = data_by_key.get(key) or kttt_by_key.get(key) or {}

        add_p(doc, f"NGÀY {m['dd']}/{m['mm']}/{m['yyyy']}", size=11, bold=True,
              align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=4)

        if key in ASM_CHECK_KEYS:
            render_asm_check(doc, m, kttt_by_key[key])
        elif key in KTTT_KEYS:
            render_kttt(doc, m, kttt_by_key[key])
        elif key in FREEFORM_KEYS:
            render_freeform(doc, m, d)
        else:
            render_gd02(doc, m, d)

        add_p(doc, f"(Người gửi: {m['nguoi_gui']} — NPP: {m['npp_ten']})", size=10)
        if m["key"] in NOTES: add_p(doc, "Ghi chú: " + NOTES[m["key"]], size=10, bold=True)

        doc.add_page_break()

        img_path = os.path.join(IMG_DIR, m["image"])
        with Image.open(img_path) as im:
            w, h = im.size
        avail_w = sec.page_width - sec.left_margin - sec.right_margin
        avail_h = sec.page_height - sec.top_margin - sec.bottom_margin - Cm(2)
        ratio = min(avail_w / w, avail_h / h)
        pic_w = w * ratio
        pic_h = h * ratio

        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(img_path, width=pic_w, height=pic_h)

        if m is not last_meeting_overall:
            doc.add_page_break()

doc.save(OUT_PATH)
print("SAVED:", OUT_PATH, "meetings:", len(meetings), "days:", len(dates_sorted))
