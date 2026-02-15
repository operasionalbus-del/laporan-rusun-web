import datetime
from openpyxl import load_workbook
from rapidfuzz import process
from mapping import mapping
import re

print("APP_LOGIC VERSION 2026-02-15 FINAL")

# =========================
# Helper: normalisasi teks
# =========================
def clean_text(s):
    if not s:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())


def safe_int(val):
    if not val:
        return 0
    val = val.strip()
    val = val.replace("O", "0")
    digits = re.findall(r"\d+", val)
    return int(digits[0]) if digits else 0


# =========================
# STEP 1: Filter laporan dari chat text
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
# STEP 2: Parsing laporan
# =========================
def parse_report(text):
    data = {}
    for line in text.splitlines():
        if ":" in line:
            key, val = line.split(":", 1)
            key = key.lower().replace(" ", "").replace(".", "")
            data[key] = val.strip()
    return data


# =========================
# STEP 3: Isi template Excel
# =========================
def isi_template(template_path, chat_text, tanggal_target, output_file):
    reports = filter_orderan_from_text(chat_text)

    wb = load_workbook(template_path)
    ws = wb.active

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

    for rep in reports:
        data = parse_report(rep)

        shift = data.get("shift", "").replace(" ", "")
        kode_rute_input = clean_text(data.get("koderute", ""))

        no_body_raw = data.get("nobody", "").upper()
        no_body_clean = clean_text(no_body_raw)

        tob_fp = safe_int(data.get("tobfp", "0"))
        tob_ep = safe_int(data.get("tobep", "0"))
        tob_lg = safe_int(data.get("toblg", "0"))
        tap_out = safe_int(data.get("tapout", "0"))

        if not kode_rute_input or not no_body_clean:
            continue

        best_match, score, _ = process.extractOne(kode_rute_input, mapping.keys())
        if score < 70:
            continue

        rows = mapping[best_match]
        target_row = None

        # cari baris dengan no body yang sama (untuk shift 2)
        for r in rows:
            cell_value = ws[f"C{r}"].value
            if clean_text(cell_value) == no_body_clean:
                target_row = r
                break

        # jika belum ada (shift 1)
        if not target_row:
            for r in rows:
                if ws[f"C{r}"].value in (None, ""):
                    target_row = r
                    break

        if not target_row:
            continue

        if shift == "1":
            ws[f"C{target_row}"] = no_body_raw
            ws[f"D{target_row}"] = tob_fp
            ws[f"E{target_row}"] = tob_ep
            ws[f"F{target_row}"] = tob_lg

        elif shift == "2":
            ws[f"M{target_row}"] = tob_fp
            ws[f"N{target_row}"] = tob_ep
            ws[f"O{target_row}"] = tob_lg
            ws[f"L{target_row}"] = tap_out

    wb.save(output_file)
    return output_file
