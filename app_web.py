import streamlit as st
from app_logic import isi_template
import tempfile
import base64
import os
from datetime import datetime


# =========================
# SET PAGE
# =========================
st.set_page_config(
    page_title="Laporan Rusun",
    page_icon="📊",
    layout="centered"
)


# =========================
# BACKGROUND IMAGE
# =========================
def set_background(image_file):
    if not os.path.exists(image_file):
        st.warning("Background image tidak ditemukan.")
        return

    with open(image_file, "rb") as f:
        encoded = base64.b64encode(f.read()).decode()

    css = f"""
    <style>
    .stApp {{
        background-image: url("data:image/jpg;base64,{encoded}");
        background-size: cover;
        background-position: center;
        background-attachment: fixed;
    }}

    /* HEADER CONTAINER */
    .header-box {{
        background-color: rgba(200, 200, 200, 0.9);
        padding: 20px;
        border-radius: 15px;
        text-align: center;
        margin-bottom: 20px;
        box-shadow: 0px 2px 6px rgba(0,0,0,0.2);
    }}

    /* MAIN CONTENT */
    .main-box {{
        background-color: rgba(255,255,255,0.9);
        padding: 25px;
        border-radius: 15px;
        box-shadow: 0px 2px 6px rgba(0,0,0,0.2);
    }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)


set_background("rusun oke.jpg")


# =========================
# HEADER (WITH CONTAINER)
# =========================
st.markdown("""
<div class="header-box">
    <h1>📊 Web App Laporan TOB Rute Integrasi Rusun</h1>
    <p>Upload chat WhatsApp & generate Excel otomatis</p>
</div>
""", unsafe_allow_html=True)


# =========================
# MAIN CONTAINER
# =========================
st.markdown('<div class="main-box">', unsafe_allow_html=True)


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
    except:
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
                    st.download_button(
                        label="📥 Download laporan Excel",
                        data=f,
                        file_name=f"laporan_{tanggal_target.replace('/','-')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                st.success("✅ File berhasil dibuat!")

            except Exception as e:
                st.error("❌ Terjadi error saat memproses file")
                st.exception(e)


# =========================
# FOOTER
# =========================
st.markdown("---")
st.caption("Developed for Laporan Rusun • 2026")

st.markdown('</div>', unsafe_allow_html=True)
