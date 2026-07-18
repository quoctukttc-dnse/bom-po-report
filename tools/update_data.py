#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cập nhật data.js cho BOM/PO Dashboard từ file "BOM PO Status.xlsx" xuất từ hệ thống.

Cách dùng:
    python3 tools/update_data.py "đường/dẫn/BOM PO Status.xlsx"

Script sẽ ghi đè file data.js ở thư mục gốc repo. Sau đó commit + push để
GitHub Pages cập nhật dữ liệu mặc định. (Dashboard cũng hỗ trợ kéo-thả file
Excel trực tiếp trên trình duyệt, không bắt buộc chạy script này.)

Yêu cầu: pip install pandas openpyxl
"""
import sys, os, json, datetime, re
import pandas as pd

def fix_date(v):
    """Sửa lỗi ngày bị đảo dd/mm <-> mm/dd.

    File xuất ra chứa ngày dạng chuỗi 'dd/mm/yyyy' lẫn ô kiểu datetime.
    Các ô datetime là do ngày dd<=12 bị Excel hiểu nhầm theo kiểu Mỹ (mm/dd),
    nên phải hoán đổi ngày/tháng lại. Chuỗi thì parse day-first bình thường.
    Trả về (date_iso, time_str hoặc None).
    """
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None, None
    if isinstance(v, datetime.datetime):
        if v.year <= 1900:
            return None, None
        # hoán đổi: stored(month, day) -> thực tế (day=month, month=day)
        try:
            real = datetime.datetime(v.year, v.day, v.month, v.hour, v.minute, v.second)
        except ValueError:
            real = v
        t = real.strftime('%H:%M:%S') if (real.hour or real.minute or real.second) else None
        return real.strftime('%Y-%m-%d'), t
    s = str(v).strip()
    if not s:
        return None, None
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$', s)
    if not m:
        return None, None
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if y <= 1900:
        return None, None
    try:
        datetime.date(y, mo, d)
    except ValueError:
        return None, None
    t = None
    if m.group(4):
        t = '%02d:%02d:%02d' % (int(m.group(4)), int(m.group(5)), int(m.group(6) or 0))
    return '%04d-%02d-%02d' % (y, mo, d), t

def dt_iso(v):
    d, t = fix_date(v)
    if d is None:
        return None
    return d + (' ' + t if t else '')

def after_dash(v):
    """'3 - Confirmed' -> 'Confirmed'; '100100 - HBI' -> 'HBI'"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v)
    return s.split(' - ', 1)[1].strip() if ' - ' in s else s.strip()

def clean(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    return s if s else None

def main(xlsx_path, out_path):
    print('Đang đọc', xlsx_path, '...')
    bom_raw = pd.read_excel(xlsx_path, sheet_name='BOM')
    po_raw = pd.read_excel(xlsx_path, sheet_name='PO', header=None, skiprows=1)

    VALID_BOM_TYPES = {'Purchasing', 'Development', 'Manufacturing', 'Costing'}
    bom_rows = []
    for r in bom_raw.itertuples(index=False):
        d, _ = fix_date(getattr(r, '_17'))  # Created Date (cột 18)
        creator = clean(r[18])              # Created By
        btype = clean(r[9])                 # BOM Type
        # loại dòng rác (remark tràn dòng): không có ngày tạo/người tạo,
        # hoặc BOM Type không hợp lệ
        if not d or not creator or btype not in VALID_BOM_TYPES:
            continue
        bom_rows.append([
            d,                       # 0 created date ISO
            creator,                 # 1 created by
            clean(r[6]),             # 2 customer
            btype,                   # 3 bom type
            clean(r[12]),            # 4 stage
            clean(r[10]),            # 5 order type
            clean(r[2]),             # 6 season
            clean(r[4]),             # 7 year
            clean(r[1]),             # 8 product code
            clean(r[7]),             # 9 BOM no
            clean(r[23]),            # 10 import (Checked/Unchecked/Indeterminate)
            dt_iso(r[24]),           # 11 upload SAP datetime
            dt_iso(r[25]),           # 12 upload Centric datetime
        ])

    po_rows = []
    for r in po_raw.itertuples(index=False):
        pono = clean(r[0])
        d, _ = fix_date(r[15])       # Created Date
        creator = clean(r[16])       # Created By
        if not pono or not d or not creator:
            continue
        amt = r[8]
        try:
            amt = round(float(amt), 2) if amt is not None and not pd.isna(amt) else None
        except (TypeError, ValueError):
            amt = None
        po_rows.append([
            d,                       # 0 created date ISO
            creator,                 # 1 created by
            after_dash(r[4]),        # 2 customer
            after_dash(r[1]),        # 3 PO type
            after_dash(r[3]),        # 4 status
            pono,                    # 5 PO no
            after_dash(r[7]),        # 6 vendor
            amt,                     # 7 amount
            clean(r[9]),             # 8 currency
            fix_date(r[2])[0],       # 9 document date
            fix_date(r[19])[0],      # 10 approved date
            clean(r[20]),            # 11 approved by
        ])

    payload = {
        'generatedAt': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
        'sourceFile': os.path.basename(xlsx_path),
        'bom': bom_rows,
        'po': po_rows,
    }
    js = 'window.DEFAULT_DATA = ' + json.dumps(payload, ensure_ascii=False, separators=(',', ':')) + ';\n'
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(js)
    size_mb = os.path.getsize(out_path) / 1e6
    print('Đã ghi %s: %d dòng BOM, %d dòng PO (%.1f MB)' % (out_path, len(bom_rows), len(po_rows), size_mb))

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    main(sys.argv[1], os.path.join(root, 'data.js'))
