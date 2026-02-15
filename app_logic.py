import datetime
import re
from openpyxl import load_workbook
from rapidfuzz import process
from mapping import mapping


# =========================
# Helper
# =========================
def clean_text(s):
    if not s:
        return ""
    return str(s).replace(" ", "").replace("\n", "").upper()


def safe_int(value):
    try:
        return int(str(value).strip())
    except:
        return 0


# =========================
# STEP 1: Filter laporan dari chat text
# =========================
def filter_orderan_from_text(text):
    if not text:
        return []

    lines = text.splitlines()

    reports = []
    buffer = []

    for line in lines:
        try:
            if " - " in line and ":" in line:
                msg = line.split(":", 1)[1].strip()
            else:
                msg = line.strip()
        except:
            continue

        if msg.lower().startswith("shift"):
            if buffer:
                reports.append("\n".join(buffer))
                buffer = []

        if msg != "":
            buffer.append(msg)

    if buffer:
        reports.append("\n".join(buffer))

    return reports


# =========================
# STEP 2: Parsing laporan (pakai regex fleksibel)
# =========================
def parse_report(text):
    data = {}
    if not text:
        return data

    lines = text.splitlines()

    for line in lines:
        line = line.strip()

        # SHIFT (Shift 2 / Shift : 2)
        m = re.search(r"shift\s*:?\s*(\d)", line, re.IGNORECASE)
        if m:
            data["shift"] = m.group(1)

        # KODE RUTE
        m = re.search(r"kode\s*rute\s*:?\s*([A-Za-z0-9]+)", line, re.IGNORECASE)
        if m:
            data["koderute"] = m.group(1)

        # NO BODY
        m = re.search(r"no\s*body\s*:?\s*(.+)", line, re.IGNORECASE)
        if m:
            data["nobody"] = m.group(1).strip()

        # TOB FP
        m = re.search(r"tob\s*fp\s*:?\s*(\d+)", line, re.IGNORECASE)
        if m:
            data["tobfp"] = m.group(1)

        # TOB EP
        m = re.search(r"tob\s*ep\s*:?\s*(\d+)", line, re.IGNORECASE)
        if m:
            data["tobep"] = m.group(1)

        # TOB LG
        m = re.search(r"tob\s*lg\s*:?\s*(\d+)", line, re.IGNORECASE)
        if m:
            data["toblg"] = m.group(1)

        # TAP OUT
        m = re.search(r"tap\s*out\s*:?\s*(\d+)", line, re.IGNORECASE)
        if m:
            data["tapout"] = m.group(1)

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
        if not rep.strip():
            continue

        data = parse_report(rep)

        shift = data.get("shift", "")
        kode_rute_input = clean_text(data.get("koderute", ""))

        no_body_raw = data.get("nobody", "").upper()
        no_body_clean = clean_text(no_body_raw)

        tob_fp = safe_int(data.get("tobfp"))
        tob_ep = safe_int(data.get("tobep"))
        tob_lg = safe_int(data.get("toblg"))
        tap_out = safe_int(data.get("tapout"))

        if not kode_rute_input:
            continue

        best_match = process.extractOne(kode_rute_input, mapping.keys())
        if not best_match:
            continue

        best_match, score, _ = best_match
        if score < 80:
            continue

        rows = mapping[best_match]
        target_row = None

        # cari baris dengan no body sama (shift 2)
        for r in rows:
            cell_value = ws[f"C{r}"].value
            if clean_text(cell_value) == no_body_clean:
                target_row = r
                break

        # jika belum ada, cari baris kosong (shift 1)
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
