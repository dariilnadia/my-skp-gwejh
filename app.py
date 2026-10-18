import streamlit as st
import pandas as pd
from datetime import datetime
import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# --- KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="SKP & Eviden - Serang Gawe FM",
    page_icon="🎙️",
    layout="wide"
)

# --- TAMPILAN KUSTOM (TEMA BIRU SERANG GAWE FM) ---
st.markdown("""
    <style>
    .stApp {
        background-color: #f4f7fc;
    }
    .main-header {
        background-color: #0d47a1;
        color: white;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 25px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .stButton>button {
        background-color: #1565c0;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 10px 24px;
        border: none;
        width: 100%;
    }
    .stButton>button:hover {
        background-color: #0d47a1;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

# --- HEADER APLIKASI ---
st.markdown("""
    <div class="main-header">
        <h1>🎙️ PORTAL EVIDEN SKP - SERANG GAWE FM</h1>
        <p>Sistem Rekapitulasi Dokumentasi & Laporan Kinerja Bulanan Karyawan</p>
    </div>
""", unsafe_allow_html=True)

# Folder penyimpanan foto
if not os.path.exists("uploads"):
    os.makedirs("uploads")

# File database lokal
DATA_FILE = "data_eviden.csv"

if os.path.exists(DATA_FILE):
    df_data = pd.read_csv(DATA_FILE)
else:
    df_data = pd.DataFrame(columns=["Tanggal", "Waktu", "Program", "Narasumber", "Tema", "Foto_Path"])

# --- MENU TAB ---
tab1, tab2, tab3 = st.tabs(["📝 Input Eviden Harian", "📊 Rekap Data Bulanan", "📄 Generate Laporan SKP"])

# --- TAB 1: FORM INPUT ---
with tab1:
    st.subheader("📌 Form Input Kegiatan Harian")
    
    with st.form("form_eviden", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            tanggal = st.date_input("Tanggal Kegiatan", datetime.now())
            waktu = st.selectbox("Waktu/Sesi Siaran", ["SIARAN PAGI", "OBRAS","MENJELANG SORE", "SERGAW MALAM", "NGECAS IMAN", "DOKUMENTASI TEKNIS / PROSES"])
            program = st.text_input("Nama Program / Jenis Kegiatan", placeholder="Contoh: Zona Atensi / OBRAS / Proses Pencarian Narasumber")
        
        with col2:
            narasumber = st.text_area("Nama Narasumber & Jabatan", placeholder="Contoh: Drs. Lutfi Mujahidin, M.Si. (Kepala BPBD)")
            tema = st.text_area("Tema / Judul Pembahasan", placeholder="Contoh: Status Siaga Darurat Bencana")
        
        uploaded_file = st.file_uploader("Upload Foto Dokumentasi/Evidence (PNG/JPG)", type=["jpg", "jpeg", "png"])
        
        submit = st.form_submit_button("💾 Simpan Eviden")
        
        if submit:
            if uploaded_file is not None:
                # Simpan foto
                img_name = f"{tanggal.strftime('%Y%m%d')}_{uploaded_file.name}"
                img_path = os.path.join("uploads", img_name)
                with open(img_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                # Simpan data ke CSV
                new_entry = {
                    "Tanggal": tanggal.strftime("%Y-%m-%d"),
                    "Waktu": waktu,
                    "Program": program,
                    "Narasumber": narasumber,
                    "Tema": tema,
                    "Foto_Path": img_path
                }
                df_data = pd.concat([df_data, pd.DataFrame([new_entry])], ignore_index=True)
                df_data.to_csv(DATA_FILE, index=False)
                st.success("✅ Data eviden berhasil disimpan!")
            else:
                st.error("⚠️ Harap upload foto bukti dokumentasi!")

# --- TAB 2: REKAP DATA ---
with tab2:
    st.subheader("📋 Rekapitulasi Data Eviden")
    if not df_data.empty:
        st.dataframe(df_data[["Tanggal", "Waktu", "Program", "Narasumber", "Tema"]], use_container_width=True)
    else:
        st.info("Belum ada data yang diinputkan bulan ini.")

# --- TAB 3: GENERATE LAPORAN DOCX ---
with tab3:
    st.subheader("🖨️ Cetak Laporan Bulanan (Format SKP)")
    
    col_a, col_b = st.columns(2)
    with col_a:
        nama_pegawai = st.text_input("Nama Pegawai", value="NADIYYATUL FIZRIYA")
        nip_pegawai = st.text_input("NIP / ID Pegawai", value="200112222025212026")
    with col_b:
        bulan_tahun = st.text_input("Bulan & Tahun Laporan", value="JANUARI 2026")
    
    if st.button("🚀 Buat Laporan Word (.docx) Otomatis"):
        if not df_data.empty:
            doc = Document()
            
            # --- HALAMAN 1: SAMPUL ---
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run("LAMPIRAN\nSKP BULAN " + bulan_tahun.upper() + "\n\n")
            run.font.size = Pt(16)
            run.font.bold = True
            
            run_nama = p.add_run(nama_pegawai.upper() + "\n")
            run_nama.font.size = Pt(18)
            run_nama.font.bold = True
            
            run_nip = p.add_run(nip_pegawai)
            run_nip.font.size = Pt(14)
            
            doc.add_page_break()
            
            # --- HALAMAN 2: TABEL SKP ---
            doc.add_heading(f"REKAPITULASI SIARAN & KEGIATAN - {bulan_tahun.upper()}", level=1)
            
            table = doc.add_table(rows=1, cols=5)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'NO'
            hdr_cells[1].text = 'HARI / TANGGAL'
            hdr_cells[2].text = 'NAMA PROGRAM'
            hdr_cells[3].text = 'NARASUMBER'
            hdr_cells[4].text = 'TEMA'
            
            for idx, row in df_data.iterrows():
                row_cells = table.add_row().cells
                row_cells[0].text = str(idx + 1)
                row_cells[1].text = f"{row['Tanggal']}"
                row_cells[2].text = f"{row['Program']} ({row['Waktu']})"
                row_cells[3].text = str(row['Narasumber'])
                row_cells[4].text = str(row['Tema'])
            
            doc.add_page_break()
            
            # --- HALAMAN 3: GRID DOKUMENTASI ---
            doc.add_heading(f"DOKUMENTASI {bulan_tahun.upper()}", level=1)
            
            doc_table = doc.add_table(rows=0, cols=3)
            
            row_cells = None
            for idx, row in df_data.iterrows():
                if idx % 3 == 0:
                    row_cells = doc_table.add_row().cells
                col_idx = idx % 3
                
                cell = row_cells[col_idx]
                p_cell = cell.paragraphs[0]
                p_cell.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_cell.add_run(f"**{row['Tanggal']}**\n({row['Waktu']})\n\n")
                
                if os.path.exists(row['Foto_Path']):
                    p_cell.add_run().add_picture(row['Foto_Path'], width=Inches(1.8))
            
            file_output = f"Laporan_SKP_{bulan_tahun.replace(' ', '_')}.docx"
            doc.save(file_output)
            
            with open(file_output, "rb") as fp:
                st.download_button(
                    label="📥 Unduh File Laporan (.docx)",
                    data=fp,
                    file_name=file_output,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.error("Tidak ada data untuk dicetak!")
