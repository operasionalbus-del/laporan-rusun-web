import datetime
import re
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from mapping import mapping

print("APP_LOGIC VERSION 2026-02-16 STABLE ENGINE")

# =========================
# Helper
# =========================
def normalize_text(s):
    if not s:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())

def normalize_key(s):
    return re.sub(r"[^a-z]", "", s.lower())

def safe_int(val):
    if not val:
        return 0
    val = re.sub(r"[^0-9]", "", str(val))
    return int(val) if val else 0

def safe_clear_cell(ws, cell):
    if not isinstance(ws[cell], MergedCell):
        ws[cell] = None

from openpyxl.styles import Font

def analisis_rekap(ws):
    START_ROW = 6
    END_ROW = 69

    anomali = 0
    duplicate = False

    body_list = []

    for row in range(START_ROW, END_ROW + 1):

        body = ws[f"C{row}"].value

        if not body:
            continue

        body = str(body).strip().upper()

        if body in body_list:
            duplicate = True
        else:
            body_list.append(body)

        # ==========================
        # SHIFT 1
        # ==========================
        fp1 = ws[f"D{row}"].value
        ep1 = ws[f"E{row}"].value
        lg1 = ws[f"F{row}"].value

        if fp1 is not None or ep1 is not None or lg1 is not None:

            fp1 = fp1 or 0
            ep1 = ep1 or 0
            lg1 = lg1 or 0

            # Rule 1
            if fp1 == 0 and ep1 == 0 and lg1 == 0:
                anomali += 1

            # Rule 2
            if fp1 > 500 or ep1 > 500 or lg1 > 500:
                anomali += 1

            # Rule 3
            if ws[f"D{row}"].value is None or ws[f"E{row}"].value is None or ws[f"F{row}"].value is None:
                anomali += 1

        # ==========================
        # SHIFT 2
        # ==========================
        fp2 = ws[f"M{row}"].value
        ep2 = ws[f"N{row}"].value
        lg2 = ws[f"O{row}"].value

        if fp2 is not None or ep2 is not None or lg2 is not None:

            fp2 = fp2 or 0
            ep2 = ep2 or 0
            lg2 = lg2 or 0

            if fp2 == 0 and ep2 == 0 and lg2 == 0:
                anomali += 1

            if fp2 > 500 or ep2 > 500 or lg2 > 500:
                anomali += 1

            if ws[f"M{row}"].value is None or ws[f"N{row}"].value is None or ws[f"O{row}"].value is None:
                anomali += 1

    status = "[OK] TKA, SIAP KIRIM"

    if anomali > 0 or duplicate:
        status = "[WARNING] PERLU VERIFIKASI SEBELUM DIKIRIM"

    return anomali, duplicate, status

# =========================
# STEP 1: Filter chat by DATE
# =========================
def filter_orderan_from_text(text, tanggal_target):
    lines = text.splitlines()
    reports = []
    buffer = []
    active = False

    date_pattern = re.compile(r"^(\d{2}/\d{2}/\d{2})\s+\d{2}\.\d{2}\s+-")

    for line in lines:
        line = line.strip()
        m = date_pattern.match(line)

        if m:
            tanggal_line = m.group(1)

            if tanggal_line != tanggal_target:
                active = False
                continue
            else:
                active = True

            if ":" in line:
                msg = line.split(":", 1)[1].strip()
            else:
                msg = ""
        else:
            msg = line

        if not active:
            continue

        if msg.lower().startswith("<media"):
            continue
        if "pesan ini dihapus" in msg.lower():
            continue

        # DETEKSI AWAL REPORT (fleksibel)
        if re.search(r"\bshift\b", msg.lower()):
            if buffer:
                reports.append("\n".join(buffer))
                buffer = []
            buffer.append(msg)
            continue

        if msg:
            buffer.append(msg)

    if buffer:
        reports.append("\n".join(buffer))

    print("TOTAL REPORT FOUND:", len(reports))
    return reports


# =========================
# STEP 2: Parsing laporan (FIX SHIFT TOTAL)
# =========================
def parse_report(text):
    data = {}

    for line in text.splitlines():
        line = line.strip()
        line = line.replace("<Pesan ini diedit>", "").strip()

        # Deteksi shift langsung (paling aman)
        m_shift = re.search(r"\bshift\s*:?\s*(\d)", line.lower())
        if m_shift:
            data["shift"] = m_shift.group(1)

        if ":" in line:
            key, val = line.split(":", 1)
            key = normalize_key(key)
            val = val.strip()
            data[key] = val

    return data



# =========================
# STEP 3: Isi template (STABLE VERSION)
# =========================
def isi_template(template_path, chat_text, tanggal_target, output_file):
    reports = filter_orderan_from_text(chat_text, tanggal_target)

    wb = load_workbook(template_path)
    ws = wb.active

    DATA_START_ROW = 6

    # CLEAR DATA
    for row in range(DATA_START_ROW, ws.max_row + 1):
        for col in ["C", "D", "E", "F", "L", "M", "N", "O"]:
            safe_clear_cell(ws, f"{col}{row}")

    print("TEMPLATE CLEARED")

    # HEADER tanggal
    tanggal = datetime.datetime.strptime(tanggal_target, "%d/%m/%y").date()
    hari_id = {
        "Monday": "Senin",
        "Tuesday": "Selasa",
        "Wednesday": "Rabu",
        "Thursday": "Kamis",
        "Friday": "Jumat",
        "Saturday": "Sabtu",
        "Sunday": "Minggu"
    }

    ws["A1"] = f"HARI/TANGGAL : {hari_id[tanggal.strftime('%A')]} {tanggal.strftime('%d %B %Y')}"

    # =========================
    # ENGINE ROW ALLOCATION (DETERMINISTIC)
    # =========================
    body_row_map = {}

    for rep in reports:
        data = parse_report(rep)

        shift = data.get("shift", "").strip()
        kode_rute = normalize_text(data.get("koderute", ""))

        no_body_raw = data.get("nobody", "")
        no_body_clean = normalize_text(no_body_raw)

        tob_fp = safe_int(data.get("tobfp"))
        tob_ep = safe_int(data.get("tobep"))
        tob_lg = safe_int(data.get("toblg"))
        tap_out = safe_int(data.get("tapout"))

        if not kode_rute or not no_body_clean:
            continue

        if kode_rute not in mapping:
            print("ROUTE NOT FOUND:", kode_rute)
            continue

        rows = mapping[kode_rute]
        key = (kode_rute, no_body_clean)

        # Tentukan row
        if key not in body_row_map:
            used_rows = set(body_row_map.values())
            target_row = None

            for r in rows:
                if r not in used_rows:
                    target_row = r
                    break

            if not target_row:
                print("NO SLOT:", kode_rute, no_body_clean)
                continue

            body_row_map[key] = target_row

        target_row = body_row_map[key]

        # =========================
        # TULIS DATA
        # =========================

        # BODY SELALU DITULIS
        ws[f"C{target_row}"] = no_body_raw.upper()

        if shift == "1":
            ws[f"D{target_row}"] = tob_fp
            ws[f"E{target_row}"] = tob_ep
            ws[f"F{target_row}"] = tob_lg

        elif shift == "2":
            ws[f"L{target_row}"] = tap_out
            ws[f"M{target_row}"] = tob_fp
            ws[f"N{target_row}"] = tob_ep
            ws[f"O{target_row}"] = tob_lg

        else:
            print("SHIFT TIDAK TERBACA:", no_body_raw)

    # ==========================
    # ANALISIS REKAP
    # ==========================

    anomali, duplicate, status = analisis_rekap(ws)

    start = 72

    ws[f"A{start}"] = "HASIL ANALISIS REKAP"
    ws[f"A{start}"].font = Font(bold=True)

    ws[f"A{start+1}"] = f"{'🟡' if anomali else '🟢'} Anomali Pelanggan : {'Tidak Ada' if anomali == 0 else str(anomali) + ' Temuan'}"

    ws[f"A{start+2}"] = f"{'🔴' if duplicate else '🟢'} Duplikasi Body : {'Ada' if duplicate else 'Tidak Ada'}"

    ws[f"A{start+4}"] = "Status Rekap"
    ws[f"A{start+5}"] = status
    ws[f"A{start+5}"].font = Font(bold=True)
    
    
    wb.save(output_file)
    print("FILE SAVED:", output_file)
    return output_file

