import streamlit as st
from app_logic import isi_template
import tempfile

st.set_page_config(page_title="Laporan Rusun", layout="centered")

st.title("📊 Web App Laporan Rusun")
st.write("Upload chat WhatsApp & generate Excel otomatis")

tanggal_target = st.text_input("Masukkan tanggal (dd/mm/yy)", "12/02/26")

uploaded_file = st.file_uploader("Upload file chat WhatsApp (.txt)", type=["txt"])

if uploaded_file and tanggal_target:
    chat_text = uploaded_file.read().decode("utf-8")

    if st.button("🚀 Generate Excel"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            output_file = tmp.name

        template_path = "template_bersih.xlsx"

        isi_template(template_path, chat_text, tanggal_target, output_file)

        with open(output_file, "rb") as f:
            st.download_button(
                label="📥 Download laporan Excel",
                data=f,
                file_name="laporan_otomatis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.success("File berhasil dibuat!")

