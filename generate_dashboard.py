#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_dashboard.py
======================
Tu dong cap nhat "index.html" (bao cao lanh dao - kho than NMND Song Hau 1)
tu file Excel nguon "So do luu kho than.xlsx" nam CUNG THU MUC voi script nay.

Cach dung:
    python generate_dashboard.py

Yeu cau:
    pip install openpyxl

Script nay trich xuat du lieu tho tu cac sheet Excel (Sheet1, Bao cao san
luong, Tau OG tai VTau- Cang SH1, va 2 sheet tuy chon: Tau OG Indonesia-VTau,
Theo doi KL HD giao nhan) va ghi de vao khoi "var RAW_DATA = {...};" trong
index.html. Toan bo tinh toan/dien giai (KPI, bang, bieu do, cau cau canh
bao...) nam trong JS (computeReport/renderText) o index.html va se tu dong
chay lai moi khi trang duoc mo - script Python nay khong can biet gi ve
cach hien thi, chi can trich xuat dung so lieu.

An toan: neu co loi (thieu sheet, thieu file, cau truc index.html bi doi),
script se dung lai va KHONG ghi de index.html, de tranh lam hong dashboard
dang chay.
"""
import json
import os
import re
import sys

# QUAN TRONG (loi thuc te da gap trong log tu dong chay tren may Windows): khi script chay
# qua Task Scheduler va bi redirect output vao file log (">> log.txt"), Python/Windows mac
# dinh dung code page cua may (thuong la cp1252), khong phai UTF-8 — bat ky ky tu tieng Viet
# co dau nao (a, e, o, u, d...) trong thong bao loi se lam print() nem UnicodeEncodeError,
# khien chinh buoc BAO LOI lai bi crash va che mat thong bao loi that su trong log. Ep stdout/
# stderr sang UTF-8 ngay tu dau de tranh hoan toan loi nay.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(HERE, "So do luu kho than.xlsx")
HTML_PATH = os.path.join(HERE, "index.html")

SHEET_STOCK = "Sheet1"
SHEET_PRODUCTION = "Báo cáo sản lượng"
SHEET_VESSELS = "Tàu OG tại VTau- Cảng SH1"
# 2 sheet duoi day la TUY CHON (khong bat buoc) — neu chua co / bi doi ten, script
# van chay binh thuong va chi bo qua phan du lieu tuong ung (mang rong []), KHONG
# lam dung toan bo script, vi day la du lieu bo sung cho slide 7/8, khong phai du
# lieu loi cot cua dashboard.
SHEET_OG_INDO = "Tau OG Indonesia-VTau"
SHEET_HD_TRACKING = "Theo dõi KL HĐ giao nhận"


def log(msg):
    # phong ngua kep: du reconfigure() o tren co that bai vi ly do gi, log() van khong duoc
    # phep tu crash — neu in UTF-8 loi thi thu lai voi errors='replace' thay vi nem exception.
    try:
        print(msg, flush=True)
    except UnicodeEncodeError:
        try:
            enc = sys.stdout.encoding or "ascii"
            print(msg.encode(enc, errors="replace").decode(enc), flush=True)
        except Exception:
            print(msg.encode("ascii", errors="replace").decode("ascii"), flush=True)


def die(msg):
    log("LOI: " + msg)
    sys.exit(1)


def extract_stock_rows(ws):
    """Cac lo than hien dang o bai: dang luu kho binh thuong, dang nhap, hoac dang cap.
    Trang thai (cot 'status') xac dinh theo cot NgayCapHet (cot L/12, UU TIEN cao nhat)
    va cot TrangThaiKho (cot O/15):
      - NgayCapHet CO gia tri (khac rong) -> da cap het / da xu ly xong -> KHONG con o bai
        nua -> bo qua hoan toan, bat ke TrangThaiKho la gi (tranh hien thi nham cac lo
        da xong tu lau nhung con sot co TrangThaiKho='Nhap'/'Cap' cu chua don dep).
      - NgayCapHet rong VA TrangThaiKho == 'Nhap' -> status 'nhap' (dang nhap -> do)
      - NgayCapHet rong VA TrangThaiKho == 'Cap'  -> status 'cap'  (dang cap  -> vang)
      - NgayCapHet rong VA TrangThaiKho rong       -> status 'luukho' (dang luu kho -> xanh la)
    """
    rows = []
    r = 2
    empty_streak = 0
    while empty_streak < 30 and r < 5000:
        ten_tau = ws.cell(r, 1).value
        if ten_tau is None:
            empty_streak += 1
            r += 1
            continue
        empty_streak = 0

        trang_thai_raw = ws.cell(r, 15).value
        trang_thai = str(trang_thai_raw).strip() if trang_thai_raw is not None else ""
        ngay_cap_het = ws.cell(r, 12).value

        if ngay_cap_het is not None:
            status = None
        elif trang_thai == "Nhap":
            status = "nhap"
        elif trang_thai == "Cap":
            status = "cap"
        else:
            status = "luukho"

        if status:
            yard = ws.cell(r, 4).value
            zone = ws.cell(r, 5).value
            layer = ws.cell(r, 6).value
            segment = ws.cell(r, 7).value
            ngay_nhap = ws.cell(r, 8).value
            kl_nhap = ws.cell(r, 9).value
            tau_sb = ws.cell(r, 2).value  # cot TauSB: ten tau/sa lan con da trung chuyen dung lo nay vao bai
            if yard and zone and layer is not None and segment and ngay_nhap and kl_nhap is not None:
                rows.append({
                    "yard": str(yard).strip(),
                    "zone": str(zone).strip(),
                    "layer": int(layer),
                    "segment": str(segment).strip(),
                    "vessel": str(ten_tau).strip(),
                    "sbVessel": str(tau_sb).strip() if tau_sb else None,
                    "date": ngay_nhap.strftime("%d/%m/%Y"),
                    "qty": round(float(kl_nhap), 2),
                    "status": status,
                })
        r += 1
    return rows


def extract_month_days(ws):
    """Doc toan bo cac ngay co du lieu ton kho (cot F/6), roi chi lay cac
    ngay thuoc THANG cua ngay cuoi cung co du lieu (thang bao cao hien tai)."""
    all_rows = []
    r = 3
    empty_streak = 0
    while empty_streak < 60 and r < 3000:
        d = ws.cell(r, 1).value
        if d is None:
            empty_streak += 1
            r += 1
            continue
        tonkho = ws.cell(r, 6).value
        if tonkho is not None:
            nhap = ws.cell(r, 4).value or 0
            tieuthu = ws.cell(r, 5).value or 0
            all_rows.append((d, float(nhap), float(tieuthu), float(tonkho)))
            empty_streak = 0
        else:
            empty_streak += 1
        r += 1
    if not all_rows:
        return None, None, None, []
    last_date = all_rows[-1][0]
    month, year = last_date.month, last_date.year
    month_days = [
        {"day": d.day, "nhap": round(n, 2), "tieuthu": round(t, 2), "tonkho": round(tk, 2)}
        for (d, n, t, tk) in all_rows if d.month == month and d.year == year
    ]
    month_days.sort(key=lambda x: x["day"])
    return last_date.strftime("%d/%m/%Y"), month, year, month_days


def extract_vessels(ws):
    """'Loai' == 'SB' -> tau con/sa lan trung chuyen; con lai (OGxx) -> tau me."""
    og_vessels, sb_vessels = [], []
    r = 2
    empty_streak = 0
    while empty_streak < 15 and r < 500:
        name = ws.cell(r, 2).value
        if name is None:
            empty_streak += 1
            r += 1
            continue
        empty_streak = 0
        loai = ws.cell(r, 4).value
        kl = ws.cell(r, 5).value
        done = ws.cell(r, 6).value
        remain = ws.cell(r, 7).value
        contract = ws.cell(r, 3).value
        if loai and kl is not None:
            name_clean = str(name).strip()
            loai_s = str(loai).strip().upper()
            done_v = float(done) if done is not None else 0.0
            remain_v = float(remain) if remain is not None else max(float(kl) - done_v, 0.0)
            if loai_s == "SB":
                m = re.search(r"\(OG\s*0*(\d+)\)", name_clean, re.IGNORECASE)
                parent_code = ("OG%02d" % int(m.group(1))) if m else None
                sb_vessels.append({
                    "name": re.sub(r"\s*\(OG\s*\d+\)", "", name_clean, flags=re.IGNORECASE).strip(),
                    "parentCode": parent_code,
                    "klNor": round(float(kl), 2),
                    "done": round(done_v, 2),
                    "remain": round(remain_v, 2),
                })
            else:
                og_vessels.append({
                    "name": name_clean,
                    "contract": str(contract).strip() if contract else "",
                    "code": loai_s,
                    "klNor": round(float(kl), 2),
                    "done": round(done_v, 2),
                    "remain": round(remain_v, 2),
                })
        r += 1
    return og_vessels, sb_vessels


def extract_og_indo_vtau(ws):
    """Sheet 'Tau OG Indonesia-VTau': vi tri/ETA cua tau me OG tren hanh trinh
    Indonesia -> Vung Tau -> Song Hau 1. Cot: B=ten, C=vi tri/ETA, D=ngay roi
    cang Indonesia, E=ngay du kien den Vtau, F=hop dong, G=loai(ma tau),
    H=khoi luong hang, I=da xep hang tai Indonesia, J=con lai."""
    rows = []
    r = 2
    empty_streak = 0
    while empty_streak < 15 and r < 500:
        name = ws.cell(r, 2).value
        if name is None:
            empty_streak += 1
            r += 1
            continue
        empty_streak = 0
        position = ws.cell(r, 3).value
        depart = ws.cell(r, 4).value
        eta = ws.cell(r, 5).value
        contract = ws.cell(r, 6).value
        code = ws.cell(r, 7).value
        kl_hang = ws.cell(r, 8).value
        da_xep = ws.cell(r, 9).value
        con_lai = ws.cell(r, 10).value
        if kl_hang is not None:
            da_xep_v = float(da_xep) if da_xep is not None else 0.0
            con_lai_v = float(con_lai) if con_lai is not None else max(float(kl_hang) - da_xep_v, 0.0)
            rows.append({
                "name": str(name).strip(),
                "position": str(position).strip() if position else "",
                "departIndo": depart.strftime("%d-%m-%y") if hasattr(depart, "strftime") else (str(depart) if depart else ""),
                "etaVtau": eta.strftime("%d-%m-%y") if hasattr(eta, "strftime") else (str(eta) if eta else ""),
                "contract": str(contract).strip() if contract else "",
                "code": str(code).strip() if code else "",
                "klHang": round(float(kl_hang), 2),
                "daXepIndo": round(da_xep_v, 2),
                "conLai": round(con_lai_v, 2),
            })
        r += 1
    return rows


def extract_hd_tracking(ws):
    """Sheet 'Theo dõi KL HĐ giao nhận': theo doi khoi luong giao nhan theo
    tung hop dong. Cot: B=ten HD, C=bat dau giao, D=ket thuc, E=KL giao,
    F=KL da xep, G=KL da do vao kho, H=con lai."""
    rows = []
    r = 2
    empty_streak = 0
    while empty_streak < 15 and r < 500:
        name = ws.cell(r, 2).value
        if name is None:
            empty_streak += 1
            r += 1
            continue
        empty_streak = 0
        start = ws.cell(r, 3).value
        end = ws.cell(r, 4).value
        kl_giao = ws.cell(r, 5).value
        kl_da_xep = ws.cell(r, 6).value
        kl_da_do_kho = ws.cell(r, 7).value
        con_lai = ws.cell(r, 8).value
        if kl_giao is not None:
            kl_da_xep_v = float(kl_da_xep) if kl_da_xep is not None else 0.0
            kl_da_do_kho_v = float(kl_da_do_kho) if kl_da_do_kho is not None else 0.0
            con_lai_v = float(con_lai) if con_lai is not None else max(float(kl_giao) - kl_da_do_kho_v, 0.0)
            rows.append({
                "name": str(name).strip(),
                "start": start.strftime("%d-%m-%y") if hasattr(start, "strftime") else (str(start) if start else ""),
                "end": end.strftime("%d-%m-%y") if hasattr(end, "strftime") else (str(end) if end else ""),
                "klGiao": round(float(kl_giao), 2),
                "klDaXep": round(kl_da_xep_v, 2),
                "klDaDoKho": round(kl_da_do_kho_v, 2),
                "conLai": round(con_lai_v, 2),
            })
        r += 1
    return rows


def main():
    try:
        import openpyxl
    except ImportError:
        die("chua cai thu vien openpyxl. Chay lenh sau roi thu lai:\n       pip install openpyxl")

    if not os.path.exists(EXCEL_PATH):
        die("khong tim thay file Excel:\n       " + EXCEL_PATH)
    if not os.path.exists(HTML_PATH):
        die("khong tim thay file:\n       " + HTML_PATH)

    try:
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    except Exception as e:
        die("khong mo duoc file Excel (co the dang mo bang Excel tren may, hay dong file lai): " + str(e))

    for sheet_name in (SHEET_STOCK, SHEET_PRODUCTION, SHEET_VESSELS):
        if sheet_name not in wb.sheetnames:
            die("khong tim thay sheet '%s' trong file Excel. Cac sheet hien co: %s" % (sheet_name, ", ".join(wb.sheetnames)))

    stock_rows = extract_stock_rows(wb[SHEET_STOCK])
    if not stock_rows:
        die("khong trich xuat duoc lo than nao dang ton tai bai tu sheet '%s'." % SHEET_STOCK)

    report_date, month, year, month_days = extract_month_days(wb[SHEET_PRODUCTION])
    if not month_days:
        die("khong tim thay du lieu ton kho theo ngay trong sheet '%s'." % SHEET_PRODUCTION)

    og_vessels, sb_vessels = extract_vessels(wb[SHEET_VESSELS])
    if not og_vessels:
        die("khong trich xuat duoc tau me (OG) nao tu sheet '%s'." % SHEET_VESSELS)

    if SHEET_OG_INDO in wb.sheetnames:
        og_indo_vtau = extract_og_indo_vtau(wb[SHEET_OG_INDO])
    else:
        og_indo_vtau = []
        log("CANH BAO: khong tim thay sheet '%s' - bo qua phan vi tri tau OG Indonesia-VTau "
            "(the/bang lien quan o slide 07/08 se hien trong)." % SHEET_OG_INDO)

    if SHEET_HD_TRACKING in wb.sheetnames:
        hd_tracking = extract_hd_tracking(wb[SHEET_HD_TRACKING])
    else:
        hd_tracking = []
        log("CANH BAO: khong tim thay sheet '%s' - bo qua phan theo doi hop dong giao nhan "
            "(bang lien quan o slide 07 se hien trong)." % SHEET_HD_TRACKING)

    raw_data = {
        "reportDate": report_date,
        "month": month,
        "year": year,
        "monthDays": month_days,
        "ogVessels": og_vessels,
        "sbVessels": sb_vessels,
        "stockRows": stock_rows,
        "ogIndonesiaVtau": og_indo_vtau,
        "hdTracking": hd_tracking,
    }

    with open(HTML_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    # QUAN TRONG: khong gia dinh so khoang trang truoc "};" (ban goc viet tay thut 2 dau
    # cach, nhung json.dumps() lai xuat dau "}" dong ngoai cung KHONG thut dau) — neu quy
    # dinh cung so khoang trang, lan chay thu 2 tro di se khop nham qua ca cac khoi
    # "var yardOrder / segMeters / yardSegments" phia sau va xoa mat chung.
    # Noi dung JSON ben trong khong bao gio chua chuoi "};" (JSON chi dung dau phay/dong
    # ngoac don thuan, khong co dau cham phay) nen "};" dau tien gap duoc chinh la diem
    # ket thuc that su cua "var RAW_DATA = {...};", bat ke thut dau the nao.
    pattern = re.compile(r"var RAW_DATA = \{.*?\};", re.DOTALL)
    if not pattern.search(html):
        die("khong tim thay khoi 'var RAW_DATA' trong index.html — cau truc file mau co the da bi thay doi. "
            "Khong ghi de de tranh lam hong dashboard, vui long kiem tra lai thu cong.")

    new_block = "var RAW_DATA = " + json.dumps(raw_data, ensure_ascii=False, indent=2) + ";"
    # dung ham (khong phai chuoi) lam repl de tranh re.sub dien giai \1, \g<...> trong noi dung JSON
    new_html = pattern.sub(lambda m: new_block, html, count=1)

    # kiem tra nhanh: so luong the <script> khong doi (khong lam hong HTML)
    if html.count("<script>") != new_html.count("<script>"):
        die("phat hien bat thuong sau khi ghi du lieu moi — dung lai, khong luu file.")

    # kiem tra an toan bo sung: cac khai bao quan trong ngay sau RAW_DATA phai con nguyen
    # (phong truong hop pattern lo khop qua tay trong tuong lai neu cau truc file doi khac)
    for must_have in ("var yardOrder", "var segMeters", "var yardSegments", "function computeReport"):
        if must_have in html and must_have not in new_html:
            die("phat hien mat noi dung '%s' sau khi thay RAW_DATA — dung lai, khong luu file "
                "de tranh lam hong dashboard. Vui long bao lai de kiem tra generate_dashboard.py." % must_have)

    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(new_html)

    log("OK: da cap nhat index.html tu du lieu ngay %s — %d lo than dang ton tai bai, "
        "%d ngay du lieu trong thang %d/%d, %d tau me (OG), %d tau con/sa lan (SB), "
        "%d tau OG Indonesia-VTau, %d hop dong theo doi giao nhan." % (
            report_date, len(stock_rows), len(month_days), month, year, len(og_vessels), len(sb_vessels),
            len(og_indo_vtau), len(hd_tracking)))


if __name__ == "__main__":
    main()
