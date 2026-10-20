import streamlit as st
from docx import Document
import os

st.title("Form Laporan Perjalanan Dinas")
st.write("Silakan isi form di bawah ini untuk mengunduh dokumen Anda.")

# Form Input Web
nama = st.text_input("NAMA PETUGAS")
NIP = st.text_input("NIP")

# File template harus diletakkan di folder yang sama di server
template_path = "Laporan Perjalanan Dinas.docx"

if st.button("Buat Dokumen Word"):
    if not nama or not NIP:
        st.warning("Semua kolom form harus diisi!")
    elif not os.path.exists(template_path):
        st.error("File 'Laporan Perjalanan Dinas.docx' tidak ditemukan di server.")
    else:
        try:
            doc = Document(template_path)

            #Ganti placeholder di tabel
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                if "{{nama}}" in run.text:
                                    run.text = run.text.replace("{{nama}}", nama)
                                if "{{nip}}" in run.text:
                                    run.text = run.text.replace("{{nip}}", NIP)
                                    
            # Ganti placeholder di paragraf
            for p in doc.paragraphs:
                for run in p.runs:
                    if "{{nama}}" in run.text:
                        run.text = run.text.replace("{{nama}}", nama)
                    if "{{nip}}" in run.text:
                        run.text = run.text.replace("{{nip}}", NIP)

            # Ganti placeholder di tabel
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if "{{nama}}" in cell.text:
                            cell.text = cell.text.replace("{{nama}}", nama)
                        if "{{nip}}" in cell.text:
                            cell.text = cell.text.replace("{{nip}}", NIP)

            output_path = f"LaporanJalDis_{nama.replace(' ', '_')}.docx"
            doc.save(output_path)

            # Sediakan tombol download di web untuk pengguna
            with open(output_path, "rb") as file:
                st.download_button(
                    label="📥 Unduh Dokumen Word Anda",
                    data=file,
                    file_name=output_path,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            st.success("Dokumen berhasil diproses!")
            
            # Hapus file temporary di server setelah selesai
            os.remove(output_path)

        except Exception as e:
            st.error(f"Terjadi kesalahan: {str(e)}")
