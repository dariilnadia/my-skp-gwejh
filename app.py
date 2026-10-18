import streamlit as st
import pandas as pd
from datetime import datetime
import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

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
        <p>Sistem Rekapitulasi Dokumentasi & Laporan Kinerja Mandiri (Spesial OBRAS & Talkshow)</p>
    </div>
""", unsafe_allow_html=True)

# Folder utama penyimpanan
if not os.path.exists("data_pegawai"):
    os.makedirs("data_pegawai")

# --- IDENTITAS PENGGUNA DI SIDEBAR ---
st.sidebar.markdown("### 👤 Identitas Pengguna")
nama_pegawai_input = st.sidebar.text_input("Masukkan Nama Anda:", value="NADIYYATUL FIZRIYA")
nip_pegawai_input = st.sidebar.text_input("Masukkan NIP / ID:", value="200112222025212026")

clean_name = nama_pegawai_input.strip().replace(" ", "_").lower()
if not clean_name:
    clean_name = "default_user"

DATA_FILE = f"data_pegawai/data_{clean_name}.csv"
UPLOAD_DIR = f"data_pegawai/uploads_{clean_name}"

if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

if os.path.exists(DATA_FILE):
    df_data = pd.read_csv(DATA_FILE)
else:
    df_data = pd.DataFrame(columns=["Kategori", "Tanggal", "Waktu", "Program", "Narasumber", "Tema", "Foto_Path"])

# --- HELPER FUNCTION: BORDER TABEL ---
def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key, attr in [('val', 'w:val'), ('color', 'w:color'), ('sz', 'w:sz'), ('space', 'w:space')]:
                if key in edge_data:
                    element.set(qn(attr), str(edge_data[key]))

def set_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        r'<w:tblBorders %s>'
        r'  <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'  <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'  <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'</w:tblBorders>' % nsdecls('w')
    )
    tblPr.append(borders)

# --- MENU TAB ---
tab1, tab2, tab3 = st.tabs(["📝 Input Eviden Harian", "📊 Rekap Data Bulanan", "📄 Generate Laporan SKP"])

# --- TAB 1: FORM INPUT ---
with tab1:
    st.subheader(f"📌 Form Input Kegiatan Harian: {nama_pegawai_input.upper()}")
    
    # Pilih jenis input: Talkshow Biasa atau OBRAS
    jenis_kegiatan = st.radio("Pilih Jenis Kegiatan Siaran:", ["Talkshow / Siaran Umum (Pakai Narasumber)", "Program OBRAS (Materi / Informasi Harian)"])
    
    with st.form("form_eviden", clear_on_submit=True):
        if jenis_kegiatan == "Talkshow / Siaran Umum (Pakai Narasumber)":
            cat = "TALKSHOW"
            col1, col2 = st.columns(2)
            with col1:
                tanggal = st.date_input("Tanggal Kegiatan", datetime.now())
                waktu = st.selectbox("Sesi Siaran", ["SIARAN PAGI", "MENJELANG SORE", "SERGAW MALAM", "NGECAS IMAN"])
                program = st.text_input("Nama Program", placeholder="Contoh: Zona Atensi")
            with col2:
                narasumber = st.text_area("Nama Narasumber & Jabatan", placeholder="Contoh: Drs. Lutfi (Kepala BPBD)")
                tema = st.text_area("Tema / Judul Pembahasan", placeholder="Contoh: Siaga Darurat Bencana")
        else:
            cat = "OBRAS"
            col1, col2 = st.columns(2)
            with col1:
                tanggal = st.date_input("Tanggal Kegiatan", datetime.now())
                waktu = "OBRAS (10.00 - 13.00 WIB)"
                program = "OBRAS (Obrolan Siang)"
            with col2:
                narasumber = "-"
                tema = st.text_area("Materi / Topik Berita Harian", placeholder="Contoh: Informasi Cuaca & Tips Kesehatan Kulit di Musim Panas")
        
        uploaded_file = st.file_uploader("Upload Foto Dokumentasi/Evidence (PNG/JPG)", type=["jpg", "jpeg", "png"])
        
        submit = st.form_submit_button("💾 Simpan Eviden")
        
        if submit:
            if uploaded_file is not None:
                img_name = f"{tanggal.strftime('%Y%m%d')}_{uploaded_file.name}"
                img_path = os.path.join(UPLOAD_DIR, img_name)
                with open(img_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                new_entry = {
                    "Kategori": cat,
                    "Tanggal": tanggal.strftime("%d %B %Y").upper(),
                    "Waktu": waktu,
                    "Program": program,
                    "Narasumber": narasumber,
                    "Tema": tema,
                    "Foto_Path": img_path
                }
                df_data = pd.concat([df_data, pd.DataFrame([new_entry])], ignore_index=True)
                df_data.to_csv(DATA_FILE, index=False)
                st.success(f"✅ Eviden {cat} berhasil disimpan!")
            else:
                st.error("⚠️ Harap upload foto bukti dokumentasi!")

# --- TAB 2: REKAP DATA ---
with tab2:
    st.subheader(f"📋 Rekapitulasi Data: {nama_pegawai_input.upper()}")
    if not df_data.empty:
        st.dataframe(df_data[["Kategori", "Tanggal", "Waktu", "Program", "Narasumber", "Tema"]], use_container_width=True)
    else:
        st.info("Belum ada data yang diinputkan.")

# --- TAB 3: GENERATE LAPORAN DOCX ---
with tab3:
    st.subheader("🖨️ Cetak Laporan Bulanan (Format Terpisah Talkshow & OBRAS)")
    
    bulan_tahun = st.text_input("Bulan & Tahun Laporan", value="JANUARI 2026")
    
    if st.button("🚀 Buat Laporan Word (.docx) Otomatis"):
        if not df_data.empty:
            doc = Document()
            
            # --- HALAMAN 1: SAMPUL ---
            p_cover = doc.add_paragraph()
            p_cover.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            run1 = p_cover.add_run("LAMPIRAN\nSKP BULAN " + bulan_tahun.upper() + "\n\n\n\n\n\n")
            run1.font.size = Pt(16)
            run1.font.bold = True
            
            run_nama = p_cover.add_run(nama_pegawai_input.upper() + "\n")
            run_nama.font.size = Pt(18)
            run_nama.font.bold = True
            
            run_nip = p_cover.add_run(nip_pegawai_input)
            run_nip.font.size = Pt(14)
            
            doc.add_page_break()
            
            # --- BAGIAN 1: DOKUMENTASI TALKSHOW ---
            df_talkshow = df_data[df_data["Kategori"] == "TALKSHOW"]
            if not df_talkshow.empty:
                p_head = doc.add_paragraph()
                p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_h = p_head.add_run(f"DOKUMENTASI KEGIATAN TALKSHOW {bulan_tahun.upper()}\n")
                run_h.font.bold = True
                run_h.font.size = Pt(14)
                
                grid_table = doc.add_table(rows=0, cols=3)
                grid_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                
                row_cells = None
                for idx, row in df_talkshow.reset_index().iterrows():
                    if idx % 3 == 0:
                        row_cells = grid_table.add_row().cells
                    col_idx = idx % 3
                    cell = row_cells[col_idx]
                    
                    inner_table = cell.add_table(rows=1, cols=1)
                    inner_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    inner_cell = inner_table.rows[0].cells[0]
                    
                    border_spec = {"val": "single", "sz": "8", "color": "000000"}
                    set_cell_border(inner_cell, top=border_spec, bottom=border_spec, left=border_spec, right=border_spec)
                    
                    p_frame = inner_cell.paragraphs[0]
                    p_frame.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    label_text = f"{row['Tanggal']}\n({row['Waktu']})"
                    r_title = p_frame.add_run(f"{label_text}\n\n")
                    r_title.font.bold = True
                    r_title.font.size = Pt(9.5)
                    
                    if os.path.exists(str(row['Foto_Path'])):
                        p_frame.add_run().add_picture(row['Foto_Path'], width=Inches(1.6))
                        
                    cell.add_paragraph().paragraph_format.space_after = Pt(8)
                
                doc.add_page_break()
            
            # --- BAGIAN 2: DOKUMENTASI KHUSUS OBRAS ---
            df_obras = df_data[df_data["Kategori"] == "OBRAS"]
            if not df_obras.empty:
                p_head_obras = doc.add_paragraph()
                p_head_obras.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_ho = p_head_obras.add_run(f"DOKUMENTASI PROGRAM ACARA OBRAS (OBROLAN SIANG) {bulan_tahun.upper()}\n")
                run_ho.font.bold = True
                run_ho.font.size = Pt(13)
                
                grid_obras = doc.add_table(rows=0, cols=3)
                grid_obras.alignment = WD_TABLE_ALIGNMENT.CENTER
                
                row_cells_o = None
                for idx, row in df_obras.reset_index().iterrows():
                    if idx % 3 == 0:
                        row_cells_o = grid_obras.add_row().cells
                    col_idx = idx % 3
                    cell = row_cells_o[col_idx]
                    
                    inner_table = cell.add_table(rows=1, cols=1)
                    inner_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    inner_cell = inner_table.rows[0].cells[0]
                    
                    border_spec = {"val": "single", "sz": "8", "color": "000000"}
                    set_cell_border(inner_cell, top=border_spec, bottom=border_spec, left=border_spec, right=border_spec)
                    
                    p_frame = inner_cell.paragraphs[0]
                    p_frame.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    label_text = f"OBRAS\n{row['Tanggal']}"
                    r_title = p_frame.add_run(f"{label_text}\n\n")
                    r_title.font.bold = True
                    r_title.font.size = Pt(9.5)
                    
                    if os.path.exists(str(row['Foto_Path'])):
                        p_frame.add_run().add_picture(row['Foto_Path'], width=Inches(1.6))
                        
                    cell.add_paragraph().paragraph_format.space_after = Pt(8)
                
                doc.add_page_break()
            
            # --- BAGIAN 3: TABEL REKAPITULASI GABUNGAN/TANGGAL ---
            doc.add_heading(f"REKAPITULASI KEGIATAN SIARAN", level=2)
            
            skp_table = doc.add_table(rows=1, cols=5)
            skp_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            set_table_borders(skp_table)
            
            hdr_cells = skp_table.rows[0].cells
            headers = ['NO', 'HARI/TANGGAL', 'NAMA PROGRAM', 'NARASUMBER', 'TEMA / MATERI']
            for i, h in enumerate(headers):
                hdr_cells[i].text = h
                p = hdr_cells[i].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    r.font.bold = True
                    r.font.size = Pt(10)
            
            for idx, row in df_data.reset_index().iterrows():
                row_cells = skp_table.add_row().cells
                row_cells[0].text = str(idx + 1)
                row_cells[1].text = str(row['Tanggal'])
                row_cells[2].text = f"{row['Program']}\n({row['Waktu']})"
                row_cells[3].text = str(row['Narasumber'])
                row_cells[4].text = str(row['Tema'])
                
                row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            file_output = f"Laporan_SKP_{clean_name}_{bulan_tahun.replace(' ', '_')}.docx"
            doc.save(file_output)
            
            with open(file_output, "rb") as fp:
                st.download_button(
                    label="📥 Unduh File Laporan Word (.docx)",
                    data=fp,
                    file_name=file_output,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.error("Belum ada data yang diinputkan!")
                
