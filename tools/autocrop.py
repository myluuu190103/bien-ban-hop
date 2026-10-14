# -*- coding: utf-8 -*-
import cv2
import numpy as np
import os, sys, json

def detect_crop_box(img):
    h, w = img.shape[:2]
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # paper mask: bright + low saturation (white/off-white/gray paper, incl.
    # faint blue ink areas) AND locally smooth — this last bit is what tells
    # solid paper apart from lacy white tablecloths/doilies, which are also
    # bright and low-saturation but have a fine perforated texture (high
    # local intensity variance) that plain paper (blank, or sparse text)
    # doesn't have.
    S = hsv[:, :, 1]
    V = hsv[:, :, 2]
    gray_f = gray.astype(np.float32)
    mean = cv2.blur(gray_f, (9, 9))
    sq_mean = cv2.blur(gray_f * gray_f, (9, 9))
    local_std = np.sqrt(np.maximum(sq_mean - mean * mean, 0))
    paper_mask = ((S < 70) & (V > 110) & (local_std < 11)).astype(np.uint8) * 255

    # clean up
    kernel = np.ones((15, 15), np.uint8)
    paper_mask = cv2.morphologyEx(paper_mask, cv2.MORPH_CLOSE, kernel)
    paper_mask = cv2.morphologyEx(paper_mask, cv2.MORPH_OPEN, kernel)

    contours, _ = cv2.findContours(paper_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    paper_box = None
    if contours:
        # The paper can be split into several disconnected blobs (folds,
        # shadows, watermark bands breaking up the mask) — union the
        # bounding boxes of every sizeable blob rather than just the
        # single largest one, or a multi-page/split photo gets truncated.
        big = [c for c in contours if cv2.contourArea(c) > 0.03 * w * h]
        total_area = sum(cv2.contourArea(c) for c in big)
        if big and total_area > 0.15 * w * h:
            xs0, ys0, xs1, ys1 = [], [], [], []
            for c in big:
                x, y, bw, bh = cv2.boundingRect(c)
                xs0.append(x); ys0.append(y); xs1.append(x + bw); ys1.append(y + bh)
            paper_box = (min(xs0), min(ys0), max(xs1), max(ys1))

    # black stamp band: rows near bottom where most pixels are very dark
    row_dark_frac = (gray < 55).mean(axis=1)
    band_rows = np.where(row_dark_frac > 0.55)[0]
    band_box = None
    if len(band_rows) > 0:
        # keep only rows in bottom half (the stamp band is always near the bottom)
        band_rows = band_rows[band_rows > h * 0.5]
    if len(band_rows) > 0:
        band_top = int(band_rows.min())
        band_bottom = h
        # horizontal extent of the dark band
        band_region = gray[band_top:h, :]
        col_dark_frac = (band_region < 55).mean(axis=0)
        dark_cols = np.where(col_dark_frac > 0.5)[0]
        if len(dark_cols) > 0:
            band_left = int(dark_cols.min())
            band_right = int(dark_cols.max()) + 1
        else:
            band_left, band_right = 0, w
        band_box = (band_left, band_top, band_right, band_bottom)

    if paper_box is None:
        # Fallback for tricky lighting (glass tables, blue tint reflections)
        # where the HSV paper mask fails: Otsu-threshold the grayscale,
        # excluding the bottom stamp-band area from the search.
        band_est = int(h * 0.85)
        roi = gray[:band_est, :]
        _, mask = cv2.threshold(roi, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        k2 = np.ones((21, 21), np.uint8)
        mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, k2)
        mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, k2)
        contours2, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        big2 = [c for c in contours2 if cv2.contourArea(c) > 0.2 * w * h]
        if big2:
            xs0, ys0, xs1, ys1 = [], [], [], []
            for c in big2:
                x, y, bw, bh = cv2.boundingRect(c)
                xs0.append(x); ys0.append(y); xs1.append(x + bw); ys1.append(y + bh)
            paper_box = (min(xs0), min(ys0), max(xs1), max(ys1))

    # Crop tightly to the paper itself only — this also drops the black
    # location/time stamp band below it (not part of the bien-ban document),
    # which is what actually saves ink on print.
    if paper_box is None:
        if band_box is None:
            return None
        left, top, right, bottom = band_box
    else:
        left, top, right, bottom = paper_box

    margin = int(0.012 * min(w, h))
    left = max(0, left - margin)
    top = max(0, top - margin)
    right = min(w, right + margin)
    bottom = min(h, bottom + margin)

    box_w, box_h = right - left, bottom - top
    if box_w < 0.3 * w or box_h < 0.3 * h:
        return None

    return (left, top, right, bottom)


def process_folder(src_dir, dst_dir, filenames, report_path):
    os.makedirs(dst_dir, exist_ok=True)
    report = []
    for fn in filenames:
        src = os.path.join(src_dir, fn)
        dst = os.path.join(dst_dir, fn)
        img = cv2.imdecode(np.fromfile(src, dtype=np.uint8), cv2.IMREAD_COLOR)
        if img is None:
            report.append({"file": fn, "status": "read_error"})
            continue
        h, w = img.shape[:2]
        box = detect_crop_box(img)
        if box is None:
            cv2.imencode(os.path.splitext(fn)[1], img)[1].tofile(dst)
            report.append({"file": fn, "status": "fallback_no_crop", "orig": [w, h]})
            continue
        left, top, right, bottom = box
        cropped = img[top:bottom, left:right]
        ok, buf = cv2.imencode(os.path.splitext(fn)[1], cropped, [cv2.IMWRITE_JPEG_QUALITY, 92])
        buf.tofile(dst)
        reduction = 1 - (cropped.shape[0] * cropped.shape[1]) / (h * w)
        report.append({
            "file": fn, "status": "cropped",
            "orig": [w, h], "crop": [right - left, bottom - top],
            "box": [left, top, right, bottom],
            "area_reduction_pct": round(reduction * 100, 1)
        })
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=1)
    return report


if __name__ == "__main__":
    src_dir = sys.argv[1]
    dst_dir = sys.argv[2]
    report_path = sys.argv[3]
    filenames = sys.argv[4:]
    process_folder(src_dir, dst_dir, filenames, report_path)
    print("DONE", len(filenames))
