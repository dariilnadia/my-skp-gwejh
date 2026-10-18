import streamlit as st
import pandas as pd
from datetime import datetime
import os
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

st.set_page_config(page_title="SKP & Eviden - Serang Gawe FM", page_icon="🎙️", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #f4f7fc; }
    .main-header {
        background-color: #0d47a1; color: white; padding: 20px;
        border-radius: 10px; text-align: center; margin-bottom: 25px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .stButton>button {
        background-color: #1565c0; color: white; font-weight: bold;
        border-radius: 8px; padding: 10px 24px; border: none; width: 100%;
    }
    .stButton>button:hover { background-color: #0d47a1; color: white; }
    h3 { color: #0d47a1 !important; }
    
    /* MEMBUAT TEKS MENU TAB JELAS & BERWARNA BIRU */
    .stTabs [data-baseweb="tab"] p {
        color: #0d47a1 !important;
        font-weight: bold !important;
        font-size: 16px !important;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="main-header">
        <h1>🎙️ PORTAL EVIDEN SKP - SERANG GAWE FM</h1>
        <p>Sistem Rekapitulasi Dokumentasi & Laporan Kinerja Mandiri</p>
    </div>
""", unsafe_allow_html=True)

if not os.path.exists("data_pegawai"):
    os.makedirs("data_pegawai")

st.sidebar.markdown("### 👤 Identitas Pengguna")
nama_pegawai_input = st.sidebar.text_input("Masukkan Nama Anda:", value="NADIYYATUL FIZRIYA")
nip_pegawai_input = st.sidebar.text_input("Masukkan NIP / ID:", value="200112222025212026")

clean_name = nama_pegawai_input.strip().replace(" ", "_").lower()
if not clean_name: clean_name = "default_user"

DATA_FILE = f"data_pegawai/data_{clean_name}.csv"
UPLOAD_DIR = f"data_pegawai/uploads_{clean_name}"

if not os.path.exists(UPLOAD_DIR): os.makedirs(UPLOAD_DIR)

if os.path.exists(DATA_FILE):
    df_data = pd.read_csv(DATA_FILE)
else:
    df_data = pd.DataFrame(columns=["Tanggal", "Waktu", "Program", "Narasumber", "Tema", "Foto_Path"])

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
                if key in edge_data: element.set(qn(attr), str(edge_data[key]))

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

tab1, tab2, tab3 = st.tabs(["📝 Input Eviden Harian", "📊 Rekap Data Bulanan", "📄 Generate Laporan SKP"])

with tab1:
    st.markdown(f"### 📌 Form Input Kegiatan Harian: {nama_pegawai_input.upper()}")
    
    with st.form("form_eviden", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            tanggal = st.date_input("Tanggal Kegiatan", datetime.now())
            waktu = st.selectbox("Sesi / Program Siaran", [
                "SIARAN PAGI", 
                "OBRAS (OBROLAN SIANG)", 
                "MENJELANG SORE", 
                "SERGAW MALAM", 
                "NGECAS IMAN",
                "PROSES PENCARIAN NARASUMBER",
                "PROSES MEMBUAT RANGKUMAN LIVE"
            ])
            program = st.text_input("Nama Program", placeholder="Contoh: Zona Atensi / Yuk Cari Tau")
        with col2:
            narasumber = st.text_area("Narasumber / Materi Harian", placeholder="Isi materi jika Obras/tanpa narasumber")
            tema = st.text_area("Tema / Berita", placeholder="Contoh: Pemprov Banten Gelar Mudik Gratis / Berita Harian")
        
        uploaded_file = st.file_uploader("Upload Foto Dokumentasi (PNG/JPG)", type=["jpg", "jpeg", "png"])
        submit = st.form_submit_button("💾 Simpan Eviden")
        
        if submit:
            if uploaded_file is not None:
                img_name = f"{tanggal.strftime('%Y%m%d')}_{uploaded_file.name}"
                img_path = os.path.join(UPLOAD_DIR, img_name)
                with open(img_path, "wb") as f: f.write(uploaded_file.getbuffer())
                
                new_entry = {
                    "Tanggal": tanggal.strftime("%d %B %Y").upper(),
                    "Waktu": waktu,
                    "Program": program,
                    "Narasumber": narasumber,
                    "Tema": tema,
                    "Foto_Path": img_path
                }
                df_data = pd.concat([df_data, pd.DataFrame([new_entry])], ignore_index=True)
                df_data.to_csv(DATA_FILE, index=False)
                st.success("✅ Eviden berhasil disimpan!")
            else:
                st.error("⚠️ Harap upload foto bukti dokumentasi!")

with tab2:
    st.markdown(f"### 📋 Rekapitulasi Data: {nama_pegawai_input.upper()}")
    if not df_data.empty:
        st.dataframe(df_data[["Tanggal", "Waktu", "Program", "Narasumber", "Tema"]], use_container_width=True)
    else:
        st.info("Belum ada data yang diinputkan.")

with tab3:
    st.markdown("### 🖨️ Cetak Laporan Bulanan (Format Presisi Kotak & Tabel)")
    bulan_tahun = st.text_input("Bulan & Tahun Laporan", value="JANUARI 2026")
    
    if st.button("🚀 Buat Laporan Word (.docx) Otomatis"):
        if not df_data.empty:
            doc = Document()
            
            p_cover = doc.add_paragraph()
            p_cover.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run1 = p_cover.add_run("LAMPIRAN\nSKP BULAN " + bulan_tahun.upper() + "\n\n\n\n\n\n")
            run1.font.size = Pt(16); run1.font.bold = True
            
            run_nama = p_cover.add_run(nama_pegawai_input.upper() + "\n")
            run_nama.font.size = Pt(18); run_nama.font.bold = True
            
            run_nip = p_cover.add_run(nip_pegawai_input)
            run_nip.font.size = Pt(14)
            doc.add_page_break()
            
            p_head = doc.add_paragraph()
            p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_h = p_head.add_run(f"DOKUMENTASI {bulan_tahun.upper()}\n")
            run_h.font.bold = True; run_h.font.size = Pt(14)
            
            p_narrative = doc.add_paragraph(
                "Keberhasilan pelaksanaan kegiatan sosialisasi kepada masyarakat di wilayah Kabupaten Serang tidak terlepas dari alur persiapan dan koordinasi intensif yang dilakukan di balik layar. Seluruh dokumen dan foto yang dilampirkan dalam bagian ini merupakan rekapitulasi dari peran operasional teknis yang telah dilaksanakan secara berkesinambungan.\n\n"
                "Evidence ini mencakup aktivitas riset dan komunikasi dalam menjaring narasumber yang relevan, pengelolaan administratif, hingga dukungan operasional teknis acara. Penulisan dan penyampaian eviden ini disusun untuk memberikan gambaran utuh mengenai kontribusi nyata serta realisasi sasaran kinerja pegawai pada periode ini.\n"
            )
            p_narrative.paragraph_format.line_spacing = 1.15
            p_narrative.paragraph_format.space_after = Pt(14)
            
            grid_table = doc.add_table(rows=0, cols=3)
            grid_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            row_cells = None
            for idx, row in df_data.reset_index().iterrows():
                if idx % 3 == 0: row_cells = grid_table.add_row().cells
                col_idx = idx % 3
                cell = row_cells[col_idx]
                
                inner_table = cell.add_table(rows=1, cols=1)
                inner_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                inner_cell = inner_table.rows[0].cells[0]
                
                border_spec = {"val": "single", "sz": "8", "color": "000000"}
                set_cell_border(inner_cell, top=border_spec, bottom=border_spec, left=border_spec, right=border_spec)
                
                p_frame = inner_cell.paragraphs[0]
                p_frame.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                if "PROSES" in str(row['Waktu']).upper():
                    label_text = f"{row['Waktu']}"
                else:
                    label_text = f"{row['Tanggal']}\n({row['Waktu']})"
                    
                r_title = p_frame.add_run(f"{label_text}\n\n")
                r_title.font.bold = True; r_title.font.size = Pt(9.5)
                
                if os.path.exists(str(row['Foto_Path'])):
                    p_frame.add_run().add_picture(row['Foto_Path'], width=Inches(1.6))
                    
                cell.add_paragraph().paragraph_format.space_after = Pt(8)
            
            doc.add_page_break()
            
            doc.add_heading(f"REKAPITULASI KEGIATAN SIARAN", level=2)
            
            skp_table = doc.add_table(rows=1, cols=5)
            skp_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            set_table_borders(skp_table)
            
            col_widths = [Inches(0.5), Inches(1.2), Inches(1.5), Inches(1.8), Inches(1.8)]
            
            hdr_cells = skp_table.rows[0].cells
            headers = ['NO', 'HARI/TANGGAL', 'NAMA PROGRAM', 'NARASUMBER / MATERI', 'TEMA / BERITA']
            for i, h in enumerate(headers):
                hdr_cells[i].text = h
                hdr_cells[i].width = col_widths[i]
                p = hdr_cells[i].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs: r.font.bold = True; r.font.size = Pt(10)
            
            for idx, row in df_data.reset_index().iterrows():
                row_cells = skp_table.add_row().cells
                
                row_cells[0].text = str(idx + 1)
                row_cells[1].text = str(row['Tanggal'])
                row_cells[2].text = f"{row['Program']}\n({row['Waktu']})" if pd.notna(row['Program']) else str(row['Waktu'])
                row_cells[3].text = str(row['Narasumber']) if pd.notna(row['Narasumber']) else "-"
                row_cells[4].text = str(row['Tema']) if pd.notna(row['Tema']) else "-"
                
                for i in range(5):
                    row_cells[i].width = col_widths[i]
                
                row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            file_output = f"Laporan_SKP_{clean_name}_{bulan_tahun.replace(' ', '_')}.docx"
            doc.save(file_output)
            
            with open(file_output, "rb") as fp:
                st.download_button(
                    label="📥 Unduh File Laporan Word (.docx)",
                    data=fp, file_name=file_output,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.error("Belum ada data yang diinputkan!")
            
