import datetime
import re
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from rapidfuzz import process
from mapping import mapping

print("APP_LOGIC VERSION 2026-02-16 DATE FILTER + MERGED SAFE")

# =========================
# Helper normalization
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


# =========================
# Safe clear cell (anti merged error)
# =========================
def safe_clear_cell(ws, cell):
    if not isinstance(ws[cell], MergedCell):
        ws[cell] = None


# =========================
# STEP 1: Filter chat by DATE
# =========================
def filter_orderan_from_text(text, tanggal_target):
    lines = text.splitlines()
    reports = []
    buffer = []
    capture = False

    for line in lines:
        # contoh: 15/02/26, 08:20 - Nama: Shift 1
        if re.match(rf"^{tanggal_target},", line):
            capture = True

        # jika ketemu tanggal lain → stop capture
        elif re.match(r"\d{2}/\d{2}/\d{2},", line):
            capture = False

        if not capture:
            continue

        # ambil pesan setelah "Nama:"
        if " - " in line and ":" in line:
            msg = line.split(":", 1)[1].strip()
        else:
            msg = line.strip()

        # deteksi shift baru
        if re.match(r"^shift", msg.lower()):
            if buffer:
                reports.append("\n".join(buffer))
                buffer = []

        if msg:
            buffer.append(msg)

    if buffer:
        reports.append("\n".join(buffer))

    print("TOTAL REPORT FOUND:", len(reports))
    return reports


# =========================
# STEP 2: Parsing fleksibel
# =========================
def parse_report(text):
    data = {}

    for line in text.splitlines():
        line = line.strip()

        if ":" in line:
            key, val = line.split(":", 1)
            key = normalize_key(key)
            val = val.strip()
            data[key] = val
        else:
            # handle "Shift 2"
            m = re.match(r"(shift)\s*(\d)", line.lower())
            if m:
                data["shift"] = m.group(2)

    return data


# =========================
# STEP 3: Isi template
# =========================
def isi_template(template_path, chat_text, tanggal_target, output_file):
    reports = filter_orderan_from_text(chat_text, tanggal_target)

    wb = load_workbook(template_path)
    ws = wb.active

    # =========================
    # CLEAR TEMPLATE FIRST
    # =========================
    for row in range(1, ws.max_row + 1):
        for col in ["C", "D", "E", "F", "L", "M", "N", "O"]:
            safe_clear_cell(ws, f"{col}{row}")

    print("TEMPLATE CLEARED")

    # =========================
    # Header tanggal
    # =========================
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
    # Proses laporan
    # =========================
    for rep in reports:
        data = parse_report(rep)

        shift = data.get("shift", "").strip()
        kode_rute_input = normalize_text(data.get("koderute", ""))

        no_body_raw = data.get("nobody", "")
        no_body_clean = normalize_text(no_body_raw)

        tob_fp = safe_int(data.get("tobfp"))
        tob_ep = safe_int(data.get("tobep"))
        tob_lg = safe_int(data.get("toblg"))
        tap_out = safe_int(data.get("tapout"))

        if not kode_rute_input or not no_body_clean:
            continue

        best_match, score, _ = process.extractOne(kode_rute_input, mapping.keys())

        if score < 70:
            print("SKIP ROUTE:", kode_rute_input)
            continue

        rows = mapping[best_match]
        target_row = None

        # cari body yang sama
        for r in rows:
            cell_val = ws[f"C{r}"].value
            if normalize_text(cell_val) == no_body_clean:
                target_row = r
                break

        # cari slot kosong
        if not target_row:
            for r in rows:
                if ws[f"C{r}"].value in (None, ""):
                    target_row = r
                    break

        if not target_row:
            print("NO SLOT:", best_match, no_body_clean)
            continue

        # =========================
        # Tulis ke Excel
        # =========================
        if shift == "1":
            ws[f"C{target_row}"] = no_body_raw.upper()
            ws[f"D{target_row}"] = tob_fp
            ws[f"E{target_row}"] = tob_ep
            ws[f"F{target_row}"] = tob_lg

        elif shift == "2":
            ws[f"L{target_row}"] = tap_out
            ws[f"M{target_row}"] = tob_fp
            ws[f"N{target_row}"] = tob_ep
            ws[f"O{target_row}"] = tob_lg

    wb.save(output_file)
    print("FILE SAVED:", output_file)
    return output_file
