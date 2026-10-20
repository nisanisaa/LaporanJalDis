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
    tujuan = DATABASE_MONEV[jenis]["tujuan"]
    no_spt = DATABASE_MONEV[jenis]["no_spt"]
    tgl_spt = DATABASE_MONEV[jenis]["tgl_spt"]

if nama != "-- Silakan Pilih Nama --":
    NIP = DATABASE_PETUGAS[nama]

# Tampilkan kolom NIP (Disabled/ReadOnly)
st.text_input("NIP (Terisi Otomatis):", NIP, disabled=True)

# Form Input Web
waktu = st.text_input("WAKTU PELAKSANAAN")
wilayah = st.text_input("DAERAH TUJUAN/PERTEMUAN")
tamu = st.text_input("HADIR DALAM PERTEMUAN (Pisahkan dengan koma jika lebih dari satu)")
arahan = st.text_input("PETUNJUK/ARAHAN YANG DIBERIKAN")
masalah = st.text_input("MASALAH/TEMUAN")
saran = st.text_input("SARAN/TINDAKAN")
lain = st.text_input("LAIN-LAIN")
tgl = st.text_input("TANGGAL LAPORAN")

# File template
template_path = "Laporan Perjalanan Dinas.docx"

# Fungsi pembantu untuk mengganti teks secara konsisten di paragraf
def replace_text_in_paragraph(paragraph, replacements):
    for key, val in replacements.items():
        if key in paragraph.text:
            # Coba ganti langsung per run jika tidak terpisah
            for run in paragraph.runs:
                if key in run.text:
                    # Khusus untuk teks multi-baris (seperti tamu), gunakan .add_break() agar rapi di Word
                    if key == "{{tamu}}" and "\n" in val:
                        run.text = run.text.replace(key, "") # Bersihkan placeholder dulu
                        items = val.split("\n")
                        for idx, item in enumerate(items, 1):
                            p = paragraph
                            if idx > 1:
                                p = paragraph.insert_paragraph_before() # Atau tambahkan baris baru
                            # Atau cara aman dengan string replacement biasa yang mendukung \n di docx:
                    
                    run.text = run.text.replace(key, val)
            
            # Jika placeholder terpecah antar-run, ganti secara utuh di level paragraf
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, val)

if st.button("Buat Dokumen Word"):
    # 1. Validasi Semua Isian Wajib Diisi
    if jenis == "-- Silakan Pilih Kegiatan Monev --":
        st.warning("⚠️ Silakan pilih Kegiatan Monev terlebih dahulu!")
    elif nama == "-- Silakan Pilih Nama --":
        st.warning("⚠️ Silakan pilih Nama Petugas terlebih dahulu!")
    elif not waktu.strip():
        st.warning("⚠️ Waktu Pelaksanaan wajib diisi!")
    elif not wilayah.strip():
        st.warning("⚠️ Daerah Tujuan/Pertemuan wajib diisi!")
    elif not tamu.strip():
        st.warning("⚠️ Hadir Dalam Pertemuan wajib diisi!")
    elif not arahan.strip():
        st.warning("⚠️ Petunjuk/Arahan yang Diberikan wajib diisi!")
    elif not masalah.strip():
        st.warning("⚠️ Masalah/Temuan wajib diisi!")
    elif not saran.strip():
        st.warning("⚠️ Saran/Tindakan wajib diisi!")
    elif not lain.strip():
        st.warning("⚠️ Lain-lain wajib diisi!")
    elif not tgl.strip():
        st.warning("⚠️ Tanggal Laporan wajib diisi!")
    elif not os.path.exists(template_path):
        st.error(f"❌ File '{template_path}' tidak ditemukan di server.")
    else:
        try:
            doc = Document(template_path)

            # Dictionary Pemetaan Placeholder
            replacements = {
                "{{no_spt}}": no_spt,
                "{{tgl_spt}}": tgl_spt,
                "{{tujuan}}": tujuan,
                "{{nama}}": nama,
                "{{nip}}": NIP,
                "{{waktu}}": waktu,
                "{{wilayah}}": wilayah,
                "{{tamu}}": tamu,
                "{{arahan}}": arahan,
                "{{masalah}}": masalah,
                "{{saran}}": saran,
                "{{lain}}": lain,
                "{{tgl}}": tgl,
            }

            # Replace pada paragraf biasa
            for p in doc.paragraphs:
                replace_text_in_paragraph(p, replacements)

            # Replace pada tabel
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            replace_text_in_paragraph(p, replacements)

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
            st.success("✅ Dokumen berhasil diproses dan siap diunduh!")

            # Hapus file temporary di server setelah selesai
            os.remove(output_path)

        except Exception as e:
            st.error(f"Terjadi kesalahan: {str(e)}")
