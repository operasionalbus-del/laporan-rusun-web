import streamlit as st
from app_logic import isi_template
import tempfile
import base64
import os
from datetime import datetime


# =========================
# Buat Logo / Image Helper
# =========================
def get_base64_image(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()


# =========================
# SET PAGE
# =========================
st.set_page_config(
    page_title="Laporan Rusun",
    page_icon="building.png",
    layout="centered"
)


# =========================
# BACKGROUND IMAGE + CSS
# =========================
def set_background(image_file):
    if not os.path.exists(image_file):
        st.warning("Background image tidak ditemukan.")
        return

    with open(image_file, "rb") as f:
        encoded = base64.b64encode(f.read()).decode()

    # PENTING: karena css di bawah adalah f-string, setiap kurung kurawal
    # murni CSS harus ditulis GANDA ({{ }}) supaya tidak dibaca Python
    # sebagai placeholder f-string. Ini penyebab error sebelumnya.
    css = f"""
    <style>
    .stApp {{
        background-image: url("data:image/jpg;base64,{encoded}");
        background-size: cover;
        background-position: center;
        background-attachment: fixed;
    }}

    /* =========================
       KONTEN UTAMA
    ========================= */
    .block-container {{
        background-color: rgba(255, 255, 255, 0.90);
        padding: 2.2rem 2.4rem 2.6rem 2.4rem;
        border-radius: 20px;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.12);
        margin-top: 1.5rem;
    }}

    /* HEADER CONTAINER */
    .header-box {{
        background: linear-gradient(135deg, #0E7490, #155E75);
        padding: 2rem;
        border-radius: 18px;
        text-align: center;
        margin-bottom: 25px;
        box-shadow: 0 6px 18px rgba(14, 116, 144, 0.35);
    }}

    .header-box img {{
        width: 80px;
        margin-bottom: 12px;
    }}

    .header-box h1 {{
        color: #ffffff;
        font-size: 1.5rem;
        margin: 0 0 6px 0;
    }}

    .header-box p {{
        color: rgba(255, 255, 255, 0.9);
        margin: 0;
        font-size: 0.95rem;
    }}

    /* LABEL */
    label {{
        color: #0F172A !important;
        font-weight: 600 !important;
    }}

    /* INPUT BOX BACKGROUND */
    div[data-baseweb="input"] > div,
    div[data-baseweb="file-uploader"] {{
        background-color: rgba(240, 240, 240, 0.95);
        border-radius: 12px;
        padding: 4px;
    }}

    section[data-testid="stFileUploaderDropzone"] {{
        border-radius: 14px;
        border: 1.5px dashed #0E7490;
        background-color: rgba(14, 116, 144, 0.05);
    }}

    /* =========================
       BUTTON GENERATE
    ========================= */
    .stButton > button {{
        width: 100%;
        background-color: #0E7490;
        color: white;
        border-radius: 12px;
        border: none;
        font-weight: bold;
        padding: 12px;
        transition: 0.25s ease-in-out;
    }}

    .stButton > button:hover {{
        background-color: #155E75;
        transform: translateY(-1px);
    }}

    /* =========================
       BUTTON DOWNLOAD
    ========================= */
    .stDownloadButton > button {{
        width: 100%;
        background-color: #16A34A;
        color: white;
        border-radius: 12px;
        border: none;
        font-weight: bold;
        padding: 12px;
        transition: 0.25s ease-in-out;
    }}

    .stDownloadButton > button:hover {{
        background-color: #15803D;
        transform: translateY(-1px);
    }}

    /* FOOTER */
    .footer-box {{
        margin-top: 40px;
        padding: 12px;
        background: rgba(255, 255, 255, 0.9);
        border-radius: 10px;
        text-align: center;
        font-weight: 600;
        color: #334155;
    }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)


# pakai background
set_background("rusun oke.jpg")


# =========================
# HEADER
# =========================
logo_base64 = get_base64_image("tap.png")

st.markdown(f"""
<div class="header-box">
    <img src="data:image/png;base64,{logo_base64}">
    <h1>Web App Laporan TOB Rute Integrasi Rusun</h1>
    <p>Upload chat WhatsApp &amp; generate Excel otomatis</p>
</div>
""", unsafe_allow_html=True)


# =========================
# INPUT
# =========================
tanggal_target = st.text_input("📅 Masukkan tanggal (dd/mm/yy)", "15/02/26")

uploaded_file = st.file_uploader(
    "📄 Upload file chat WhatsApp (.txt)",
    type=["txt"]
)


# =========================
# VALIDASI FORMAT TANGGAL
# =========================
def validasi_tanggal(tgl):
    try:
        datetime.strptime(tgl, "%d/%m/%y")
        return True
    except ValueError:
        return False


# =========================
# PROCESS
# =========================
if uploaded_file and tanggal_target:

    if not validasi_tanggal(tanggal_target):
        st.error("❌ Format tanggal harus dd/mm/yy (contoh: 15/02/26)")
        st.stop()

    chat_text = uploaded_file.read().decode("utf-8")

    if st.button("🚀 Generate Excel"):

        with st.spinner("⏳ Memproses laporan..."):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    output_file = tmp.name

                template_path = "template_bersih.xlsx"

                isi_template(template_path, chat_text, tanggal_target, output_file)

                with open(output_file, "rb") as f:
                    excel_bytes = f.read()

                st.success(
                    "✅ Rekap berhasil diproses.\n\n"
                    "Silakan klik tombol **Download laporan Excel** di bawah untuk mengunduh hasil."
                )

                st.download_button(
                    label="📥 Download laporan Excel",
                    data=excel_bytes,
                    file_name=f"laporan_{tanggal_target.replace('/', '-')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error("❌ Terjadi error saat memproses file")
                st.exception(e)


# =========================
# FOOTER
# =========================
st.markdown("""
<div class="footer-box">
    Developed for Laporan Rusun • 2026
</div>
""", unsafe_allow_html=True)
