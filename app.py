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
                program =
                
