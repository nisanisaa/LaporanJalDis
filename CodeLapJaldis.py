import streamlit as st
from docx import Document
import os

st.title("Form Laporan Perjalanan Dinas")
st.write("Silakan isi form di bawah ini untuk mengunduh dokumen Anda.")

# DATABASE JENIS MONEV
DATABASE_MONEV = {
    "BPOPP PENCAIRAN TAHAP 2": {
        "no_spt": "000.1.2.3/6007/101.1/2026",
        "tgl_spt": "6 Oktober 2026",
        "tujuan": "Melaksanakan Tugas Verifikasi dan Persiapan Pencairan dana BPOPP Swasta Tahap II Tahun 2026 di Lingkungan Dinas Pendidikan Provinsi Jawa Timur"
    }
}

# DATABASE PETUGAS
DATABASE_PETUGAS = {
    "FITROTUN NISA, S.Kom.": "20010105 202504 2 006",
    "SULTHAN RAFI ANWAR, S.Tr.I.P.": "20010413 202510 1 001"
}

# Form Input Pilihan Monev dan Nama Petugas (Dropdown)
jenis_spt = list(DATABASE_MONEV.keys())
jenis = st.selectbox("KEGIATAN MONEV:", ["-- Silakan Pilih Kegiatan Monev --"] + jenis_spt)
nama_petugas = list(DATABASE_PETUGAS.keys())
nama = st.selectbox("NAMA PETUGAS:", ["-- Silakan Pilih Nama --"] + nama_petugas)

# Logika Isian Otomatis
tujuan = ""
no_spt = ""
tgl_spt = ""
NIP = ""
if jenis != "-- Silakan Pilih Kegiatan Monev --":
    # Cari no, tgl, dan tujuan berdasarkan jenis monev yang dipilih
    tujuan = DATABASE_MONEV[jenis]["tujuan"]
    no_spt = DATABASE_MONEV[jenis]["no_spt"]
    tgl_spt = DATABASE_MONEV[jenis]["tgl_spt"]
if nama != "-- Silakan Pilih Nama --":
    # Cari NIP berdasarkan nama yang dipilih
    NIP = DATABASE_PETUGAS[nama]

# Tampilkan kolom NIP (Disabled/ReadOnly agar tidak bisa diubah manual)
st.text_input("NIP (Terisi Otomatis):", NIP, disabled=True)

# Form Input Web
waktu = st.text_input("WAKTU PELAKSANAAN")
wilayah = st.text_input("DAERAH TUJUAN/PERTEMUAN")
tamu = st.text_input("HADIR DALAM PERTEMUAN")
arahan = st.text_input("PETUNJUK/ARAHAN YANG DIBERIKAN")
masalah = st.text_input("MASALAH/TEMUAN")
saran = st.text_input("SARAN/TINDAKAN")
tindakan = st.text_input("LAIN-LAIN")

# File template harus diletakkan di folder yang sama di server
template_path = "Laporan Perjalanan Dinas.docx"

if st.button("Buat Dokumen Word"):
    if nama == "-- Silakan Pilih Nama --":
        st.warning("Silakan pilih nama terlebih dahulu!")
    elif not os.path.exists(template_path):
        st.error("File 'Laporan Perjalanan Dinas.docx' tidak ditemukan di server.")
    else:
        try:
            doc = Document(template_path)

            # Fungsi pembantu untuk mengganti teks placeholder
            def ganti_placeholder(teks_sumber):
                if "{{no_spt}}" in teks_sumber:
                    teks_sumber = teks_sumber.replace("{{no_spt}}", no_spt)
                if "{{tgl_spt}}" in teks_sumber:
                    teks_sumber = teks_sumber.replace("{{tgl_spt}}", tgl_spt)
                if "{{tujuan}}" in teks_sumber:
                    teks_sumber = teks_sumber.replace("{{tujuan}}", tujuan)
                return teks_sumber

            #Ganti placeholder di tabel
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                if "{{tujuan}}" in run.text:
                                    run.text = run.text.replace("{{tujuan}}", tujuan)
                                if "{{nama}}" in run.text:
                                    run.text = run.text.replace("{{nama}}", nama)
                                if "{{no_spt}}" in run.text:
                                    run.text = run.text.replace("{{no_spt}}", no_spt)
                                if "{{tgl_spt}}" in run.text:
                                    run.text = run.text.replace("{{tgl_spt}}", tgl_spt)
                                if "{{waktu}}" in run.text:
                                    run.text = run.text.replace("{{waktu}}", waktu)
                                if "{{wilayah}}" in run.text:
                                    run.text = run.text.replace("{{wilayah}}", wilayah)
                                if "{{tamu}}" in run.text:
                                    run.text = run.text.replace("{{tamu}}", tamu)
                                if "{{arahan}}" in run.text:
                                    run.text = run.text.replace("{{arahan}}", arahan)
                                if "{{masalah}}" in run.text:
                                    run.text = run.text.replace("{{masalah}}", masalah)
                                if "{{saran}}" in run.text:
                                    run.text = run.text.replace("{{saran}}", saran)
                                if "{{tindakan}}" in run.text:
                                    run.text = run.text.replace("{{tindakan}}", tindakan)
                                if "{{tgl}}" in run.text:
                                    run.text = run.text.replace("{{tgl}}", tgl)
                                    
            # Ganti placeholder di paragraf
            for p in doc.paragraphs:
                for run in p.runs:
                    if "{{nama}}" in run.text:
                        run.text = run.text.replace("{{nama}}", nama)
                    if "{{nip}}" in run.text:
                        run.text = run.text.replace("{{nip}}", NIP)

            output_path = f"LaporanJalDis_{nama.replace(' ', '_')}.docx"
            doc.save(output_path)

            # Tombol download
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
