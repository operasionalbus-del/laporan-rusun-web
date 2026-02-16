import datetime
import re
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from rapidfuzz import process
from mapping import mapping

print("APP_LOGIC VERSION 2026-02-16 BODY_ROW_MAP FIX")

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
# STEP 2: Parsing laporan
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

            #tambahan khusus shift
            if key == "shift": 
                num = re.search(r"\d+", val) 
                if num: 
                    data["shift"] = num.group(0)
        
        else:
            m = re.match(r"(shift)\s*:?(\s*\d)", line.lower())
            if m:
                data["shift"] = m.group(2).strip()

    return data


# =========================
# STEP 3: Isi template (FIXED)
# =========================
def isi_template(template_path, chat_text, tanggal_target, output_file):
    reports = filter_orderan_from_text(chat_text, tanggal_target)

    wb = load_workbook(template_path)
    ws = wb.active

    DATA_START_ROW = 6

    # CLEAR DATA (tanpa rusak merged cell)
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
    # BODY ROW MAP (KUNCI FIX)
    # =========================
    body_row_map = {}

    # Proses laporan
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

        key = (best_match, no_body_clean)
        target_row = None

        # Jika sudah pernah dicatat
        if key in body_row_map:
            target_row = body_row_map[key]
        else:
            # cari di excel
            for r in rows:
                if normalize_text(ws[f"C{r}"].value) == no_body_clean:
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

            body_row_map[key] = target_row

        # =========================
        # TULIS DATA
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

