import datetime
import re
from openpyxl import load_workbook
from rapidfuzz import process
from mapping import mapping

print("APP_LOGIC VERSION 2026-02-15 AUTO CLEAR TEMPLATE")

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
    val = re.sub(r"[^0-9]", "", val)
    return int(val) if val else 0


# =========================
# STEP 1: Filter laporan
# =========================
def filter_orderan_from_text(text):
    lines = text.splitlines()
    reports = []
    buffer = []

    for line in lines:
        if " - " in line and ":" in line:
            msg = line.split(":", 1)[1].strip()
        else:
            msg = line.strip()

        if re.match(r"^shift", msg.lower()):
            if buffer:
                reports.append("\n".join(buffer))
                buffer = []

        if msg:
            buffer.append(msg)

    if buffer:
        reports.append("\n".join(buffer))

    return reports


# =========================
# STEP 2: Parsing fleksibel
# =========================
def parse_report(text):
    data = {}

    for line in text.splitlines():
        if ":" in line:
            key, val = line.split(":", 1)
            key = normalize_key(key)
            val = val.strip()
            data[key] = val
        else:
            m = re.match(r"(shift)\s*(\d)", line.lower())
            if m:
                data["shift"] = m.group(2)

    return data


# =========================
# STEP 3: Isi template (AUTO CLEAR)
# =========================
def isi_template(template_path, chat_text, tanggal_target, output_file):
    reports = filter_orderan_from_text(chat_text)

    wb = load_workbook(template_path)
    ws = wb.active

    # =========================
    # CLEAR TEMPLATE (ONE TASK KILLER)
    # =========================
    for row in range(1, ws.max_row + 1):
        ws[f"C{row}"] = None
        ws[f"D{row}"] = None
        ws[f"E{row}"] = None
        ws[f"F{row}"] = None
        ws[f"L{row}"] = None
        ws[f"M{row}"] = None
        ws[f"N{row}"] = None
        ws[f"O{row}"] = None

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
    # Isi data dari chat
    # =========================
    for rep in reports:
        data = parse_report(rep)

        shift = data.get("shift", "").strip()
        kode_rute_input = normalize_text(data.get("koderute", ""))

        no_body_raw = data.get("nobody", "").upper()
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

        # cari body sama
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

        if shift == "1":
            ws[f"C{target_row}"] = no_body_raw
            ws[f"D{target_row}"] = tob_fp
            ws[f"E{target_row}"] = tob_ep
            ws[f"F{target_row}"] = tob_lg

        elif shift == "2":
            ws[f"L{target_row}"] = tap_out
            ws[f"M{target_row}"] = tob_fp
            ws[f"N{target_row}"] = tob_ep
            ws[f"O{target_row}"] = tob_lg

    wb.save(output_file)
    return output_file
