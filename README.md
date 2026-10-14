# Biên bản họp GĐKD – NPP: quy trình hằng tuần

Mỗi tuần nhận 1 thư mục ảnh (`Anh bien ban hop - dd-mm-yyyy_hhmm`) + file `tong-hop-bien-ban-hop-<từ>-den-<đến>.xlsx`
(sheet `Tổng hợp` + `Thống kê`) + (nếu có) `danh-sach-nhan-su-<ngày>.xlsx`. Đầu ra: 1 file Word cho tuần đó.

Python: `C:\Users\desig\AppData\Local\Programs\Python\Python311\python.exe` (python-docx, openpyxl, opencv-headless, PIL có sẵn).
Dữ liệu tích lũy nằm trong `data/` (KHÔNG xóa – tuần sau vẫn cần để lọc trùng).

## Các bước

1. **Gom ảnh mới, bỏ trùng**: nhóm theo `key = "Người gửi - NPP - dd-mm-yyyy"` (bỏ hậu tố ` - Ảnh NN` / ` - 2`);
   chỉ giữ key chưa có trong `data/meetings.json`. Ảnh 02/03 hay `- 2` thường là chụp lại cùng tờ → chỉ dùng Ảnh 01,
   NHƯNG kiểm tra: đã gặp 3 lần Ảnh 01 là selfie/ảnh tập thể, phải đổi sang Ảnh 02 (Ngô Huy Thanh, Lại Văn Thịnh).
2. **OCR** bằng agent (8 ảnh/agent, chạy song song), schema GD02: `key, ngay, thang, nam, tai_npp, dia_ban, asm, gsbh, nvbh,
   muc1{ton_kho,thi_truong,nhan_su,khac}, muc2{hang_hoa,thi_truong,chinh_sach,bao_cao,ho_tro}, muc3, muc4, muc5`
   → lưu `data/batch_NN.json`. Agent phải báo mẫu khác: **KTTT** ("BIÊN BẢN KIỂM TRA THỊ TRƯỜNG NPP", 9 mục) và
   **GD03** (họp siêu thị). Đơn vị tồn kho luôn là **thùng** (không có "tấn").
3. **KTTT**: đọc lại các key đó theo schema 9 mục
   `key, ngay, thang, nam, ho_ten, chuc_vu, lam_viec_voi, npp, dia_chi, m1..m9` → `data/ktttN_NN.json`;
   thêm key vào `KTTT_KEYS` trong `tools/build_docx.py`. GD03 / GD02 viết tự do → thêm vào `FREEFORM_KEYS`.
4. **Trùng ASM/GSBH**: cùng NPP + cùng ngày mà 2 người gửi ⇒ ASM chụp lại tờ của GSBH → giữ bản GSBH,
   thêm key ASM vào `DUP_DROP_KEYS`. Tên NPP trong file ảnh đôi khi thừa địa danh ("Thảo Nam Phú Thọ", "Liêm Xuân việt trì",
   "Hùng Mạnh Lai Châu") → sửa `npp_ten` trong `meetings.json` theo sheet Tổng hợp (Thảo Nam, Xuân Liêm, Hùng Mạnh).
5. **Cắt ảnh**: `autocrop.process_folder(src_dir, "C:\Users\desig\Downloads\Anh_bien_ban_da_cat", files, report)`.
   Sau đó xem contact sheet + check `box[3]/orig_h < 0.7` → ảnh nào cắt hụt chữ ký thì copy ảnh gốc đè lên (đã gặp 2 lần).
6. **Rà tên**: `python tools/audit_names.py` (đối chiếu ASM/GSBH/NVBH với `data/nhansu2_raw.json` theo tuyến) →
   ghi sửa vào `FIXES` trong `tools/apply_fixes.py` → chạy nó (bỏ luôn mọi "(?)" và ghi chú "không đọc được").
   Quy tắc của anh Mỹ: không dùng "(?)" – đọc được thì ghi, không thì để trống; tên NVBH mờ thì gắn đúng người trong tuyến.
7. **Đối chiếu bảng tuyến**: so (npp, ngày) của tuần với sheet Tổng hợp → thiếu/thừa/lệch tên. Lệch tên vì ASM đi thay GSBH
   thì giữ (không phải lỗi). "Cầm Thế Khái" đúng theo danh sách nhân sự (sheet Tổng hợp gõ "Khải").
8. **Xuất Word**:
   ```
   python tools/build_docx.py --start 07-09-2026 --end 13-09-2026 ^
     --xlsx "C:\Users\desig\Downloads\tong-hop-bien-ban-hop-2026-09-07-den-2026-09-13.xlsx" ^
     --out "C:\Users\desig\Downloads\BIEN_BAN_HOP_TUAN_07-09_den_13-09-2026.docx"
   ```
   Cấu trúc: trang tổng kết số BB theo ngày → bảng chi tiết cả tuần (đúng sheet Tổng hợp) → danh sách NPP chưa gửi BB
   (đúng "Mẫu tổng hợp Họp và chưa họp") → mỗi ngày: trang tổng kết ngày + [trang gõ lại | trang ảnh] cho từng biên bản.

## Bước 9 – QA khó tính (bắt buộc trước khi gửi)
`python tools/qa_week.py --start … --end … --xlsx … --cropreport <crop_report*.json>` → liệt kê: ngày giấy≠tên file, ô trống đáng ngờ,
NPP/GSBH lệch tên file, ảnh cắt hụt (<72% chiều cao – xem contact sheet, thường là dương tính giả), lệch bảng tuyến, còn (?)/tấn.
Rồi giao 8 agent đọc lại ĐỘC LẬP các con số (numbers_review2_b*.json → fix2_b*.json) + 1 agent soi ô trống (blanks_check2.json),
chạy `python tools/apply_qa2.py`, xuất lại. Tuần 07–13/09: vòng này bắt được 26 lỗi (10,1%→101%, 106→1061 thùng, 1589→1539…).

## Bước 10 – Rà phóng to toàn trường (từ 15/09/2026, thay cho vòng số liệu ở bước 9)
`python tools/make_zoom_packs.py --start … --end … --prefix fullN --per 8` → mỗi biên bản 3 băng zoom ×1,8 + gói `data/fullN_bK.json`
(kèm nvbh_trong_tuyen để agent gắn tên). Giao mỗi gói 1 agent đối chiếu TỪNG trường → `data/fix3_fullN_bK.json`
`[{key, field ("muc1.ton_kho" | "m2" | "nvbh"…), gia_tri_moi, ly_do}]`. Chạy `python tools/apply_fix3.py fullN` (tự bỏ "(?)", tấn→thùng,
KHÔNG áp ngay/thang/nam và "Liêm Xuân"). Kết quả: tuần 31/08–06/09 sửa 359 chỗ, tuần 07–13/09 sửa 279 chỗ (nhiều số sai dấu phẩy:
6,87→687 thùng; nội dung bịa; tên đối thủ "cháo Hữu Nghị"). Chạy 8 agent/đợt – 16 agent cùng lúc dễ dính session limit.

## Thư mục mới nhanh gọn (đã dùng cho 14–15/09/2026)
1. `python tools/ingest_folder.py "<thư mục ảnh>" [--pick "key=Ảnh 02"]` → thêm meetings.json, cắt ảnh, in report + 2 contact sheet
   (`data/sheet_goc_*.jpg` để soi selfie, `data/sheet_cat_*.jpg` để soi cắt hụt). Sửa `npp_ten` lệch (Liêm Xuân→Xuân Liêm) trong meetings.json.
2. `python tools/make_ocr_packs.py data/crop_report_<stamp>.json ocrN --drop "<key ASM trùng>"` (soi contact sheet trước: cùng NPP cùng tờ mà 2 người gửi → thêm key ASM vào DUP_DROP_KEYS; ảnh nào rõ nhất thì trỏ `image` của key GSBH sang ảnh đó) → agent gõ theo schema
   KTTT/GD02/GD03 → `data/ocrN_out_b*.json` → `python tools/merge_ocr.py N <số batch>` (ghi kttt N_01.json + batch_NN.json, tự đăng ký KTTT_KEYS/FREEFORM_KEYS).
3. Bước 10 với `--prefix fullM`, rồi `build_docx.py` (từ 15/09 `--xlsx` là tùy chọn: chưa có bảng tuyến thì bỏ 2 trang tổng hợp/chưa họp)
   và `qa_week.py` (cũng không bắt buộc `--xlsx`).
Cột "NVBH đi kèm" với mẫu KTTT lấy từ dòng "Làm việc với", chỉ giữ tên có trong danh sách nhân sự.
File Word đang mở trong Word thì không ghi đè được → xuất tên `_v2`.

## Prompt agent OCR và kiểm định ngữ nghĩa

Prompt dùng cho agent OCR nằm ở [`docs/agent-ocr-prompt.md`](docs/agent-ocr-prompt.md). Agent phải đọc theo từng vùng phóng to, đối chiếu roster và chỉ ghi nội dung khi vừa nhìn rõ vừa có ý nghĩa nghiệp vụ. Số liệu không có đơn vị/ngữ cảnh, tên không khớp tuyến, câu cụt hoặc chuỗi ký tự vô nghĩa phải để trống.

Quy trình tối thiểu trước khi gửi Word:

1. Ingest ảnh và kiểm tra bản trùng ASM/GSBH.
2. OCR độc lập từng mẫu GD02/KTTT/GD03.
3. Đối chiếu tên với roster mới nhất; không dùng roster để tự bịa nội dung bị mờ.
4. QA số liệu, đơn vị, ngày, tên, ô trống và crop.
5. Chỉ sau QA mới chạy `build_docx.py` và gửi file. Nếu Word đang mở, dùng hậu tố `_v2`.
