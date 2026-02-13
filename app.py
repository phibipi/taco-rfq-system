import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from docxtpl import DocxTemplate
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import nsdecls, qn  
from docx.oxml import parse_xml, OxmlElement 
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT


# --- MAIN APP ---
st.set_page_config(page_title="TACO Procurement", layout="wide", page_icon="🚛")

# --- SHEET CONNECTION ---
SPREADSHEET_ID = "1j9GCq8Wwm-MM8hOamsH26qlmjNwuDBuEMnbw6ORzTQk"


# --- UI PAGE FONT BUTTON ETC ---
def init_style():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=Montserrat:wght@500;600;700&display=swap');
        
        /* GLOBAL TEXT */
        html, body, p, label, span, div, .stMarkdown, .stText, h1, h2, h3, h4, h5, h6, th, td {
            font-family: 'Inter', sans-serif;
            color: #111827 !important;
        }
        .stApp { background-color: #F3F4F6; }
        
        /* HEADER */
        h1, h2, h3, h4 { font-family: 'Montserrat', sans-serif !important; font-weight: 600 !important; }
        
        /* CARD */
        div[data-testid="stVerticalBlockBorderWrapper"] {
            background-color: #FFFFFF; border-radius: 10px; border: 1px solid #E5E7EB; padding: 18px; box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        }
        
        /* INPUTS */
        .stTextInput input, .stSelectbox div[data-baseweb="select"] > div, .stNumberInput input, textarea {
            background-color: #F9FAFB !important; border: 1px solid #D1D5DB !important; border-radius: 6px; color: #111827 !important;
        }
        
        /* DROPDOWN & MULTISELECT */
        div[data-baseweb="popover"], div[data-baseweb="menu"], ul[data-baseweb="menu"] { background-color: #FFFFFF !important; }
        li[role="option"] { background-color: #FFFFFF !important; color: #111827 !important; }
        li[role="option"]:hover { background-color: #F3F4F6 !important; }
        span[data-baseweb="tag"] { background-color: #E5E7EB !important; color: #111827 !important; }
        
        /* --- STANDARD BUTTONS --- */
        div[data-testid="stButton"] button {
            background-color: #FCA568 !important; color: #FFFFFF !important; border: 2px solid #F58536 !important;
            border-radius: 8px !important; font-weight: 600 !important; padding: 0.5rem 1.2rem !important;
            box-shadow: 0 2px 4px rgba(249, 115, 22, 0.2);
        }
        div[data-testid="stButton"] button:hover {
            background-color: #EA580C !important; transform: translateY(-1px);
        }
        button[kind="secondary"] { background-color: #E5E7EB !important; color: #111827 !important; box-shadow: none !important; }
        button[kind="secondary"]:hover { background-color: #D1D5DB !important; }
        
        /* --- TOMBOL SIMPAN (SUBMIT) LEBIH BESAR --- */
        div[data-testid="stFormSubmitButton"] button {
            background-color: #FCA568 !important; 
            color: #FFFFFF !important; 
            border: 2px solid #F58536 !important;
            border-radius: 12px !important; 
            font-weight: 900 !important;      /* Lebih Tebal */
            font-size: 24px !important;       /* Huruf Lebih Besar */
            padding: 0.8rem 2rem !important;  /* Ukuran Tombol Lebih Besar */
            box-shadow: 0 4px 8px rgba(249, 115, 22, 0.3);
            width: 100% !important;           /* Tombol Memanjang Penuh */
            transition: all 0.3s ease;
        }
        div[data-testid="stFormSubmitButton"] button:hover {
            background-color: #EA580C !important; 
            transform: translateY(-2px);
            box-shadow: 0 6px 12px rgba(249, 115, 22, 0.4);
        }
        
        /* --- DISABLED BUTTON STYLE (ABU TUA) --- */
        div[data-testid="stButton"] button:disabled {
            background-color: #6B7280 !important; /* Cool Gray 500 (Abu Tua) */
            color: #FFFFFF !important;
            border: 1px solid #4B5563 !important;
            cursor: not-allowed !important;
            opacity: 1 !important; /* Memaksa warna keluar jelas */
            box-shadow: none !important;
        }
        
        /* TABLE FIX */
        div[data-testid="stDataFrame"] { background-color: #F8FAFC !important; border-radius: 8px; border: 1px solid #E5E7EB; }
        div[data-testid="stDataEditor"] input { color: #000000 !important; background-color: #FFFFFF !important; caret-color: #000000 !important; -webkit-text-fill-color: #000000 !important; }
        
        /* STATUS */
        .status-done { color: #065F46 !important; background-color: #D1FAE5; padding: 5px 12px; border-radius: 6px; font-size: 14px; font-weight: 700; }
        .status-pending { color: #B91C1C !important; background-color: #FEE2E2 !important; padding: 6px 16px; border-radius: 8px; font-size: 15px !important; font-weight: 800 !important; border: 1px solid #FECACA; display: inline-block; }
        .route-dest-list { font-size: 13px; color: #4B5563 !important; line-height: 1.4; }
        .streamlit-expanderHeader { background-color: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 8px; color: #111827 !important; font-weight: 600; }
        
        hr { margin-top: 0.5em; margin-bottom: 0.5em; border: none; height: 1px; background-color: #E5E7EB; }
        
        /* TABS */
        .stTabs [data-baseweb="tab-list"] { gap: 8px; }
        .stTabs [data-baseweb="tab"] { background-color: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 6px; padding: 4px 16px; color: #111827 !important; }
        .stTabs [data-baseweb="tab"][aria-selected="true"] { background-color: #B5B2B0 !important; color: #FFFFFF !important; border: none; }
       
        /* --- PERBAIKAN TAMPILAN EXPANDER (DROPDOWN) --- */
        
        /* 1. Target Kotak Utama Expander */
        div[data-testid="stExpander"] {
            background-color: #FFFFFF !important;
            border: 1px solid #E5E7EB !important;
            border-radius: 8px !important;
            color: #111827 !important; /* Warna teks isi gelap */
            box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        }

        /* 2. Target Bagian Detail (Isi Dalam) */
        div[data-testid="stExpander"] details {
            background-color: #FFFFFF !important;
            border-radius: 8px !important;
        }

        /* 3. Target Judul (Header) Expander */
        div[data-testid="stExpander"] summary {
            background-color: #FFFFFF !important; /* Paksa Background Putih */
            color: #111827 !important;            /* Paksa Teks Hitam */
            font-weight: 600 !important;
            border-radius: 8px !important;
        }
        
        /* 4. Target Ikon Panah Kecil (Agar tidak putih/hilang) */
        div[data-testid="stExpander"] summary svg {
            fill: #6B7280 !important;  /* Warna panah abu tua */
            color: #6B7280 !important;
        }

        /* 5. Efek Hover (Opsional: Garis jadi Orange) */
        div[data-testid="stExpander"]:hover {
            border-color: #FCA568 !important;
        }
        
        </style>
    """, unsafe_allow_html=True)

# --- KONEKSI & CACHE ---
# --- KONEKSI & CACHE (SMART DETECTION) ---
@st.cache_resource
def connect_to_gsheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    
    try:
        # 1. CEK APAKAH DI STREAMLIT CLOUD (SECRETS)
        if "gcp_service_account" in st.secrets:
            # st.write("Mendeteksi environment Cloud...") # Debugging (Boleh dihapus)
            creds_dict = st.secrets["gcp_service_account"]
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            client = gspread.authorize(creds)
            return client.open_by_key(SPREADSHEET_ID)

        # 2. CEK APAKAH DI LAPTOP (FILE JSON)
        elif os.path.exists("kunci_rahasia.json"):
            # st.write("Mendeteksi environment Lokal...") # Debugging (Boleh dihapus)
            creds = ServiceAccountCredentials.from_json_keyfile_name("kunci_rahasia.json", scope)
            client = gspread.authorize(creds)
            return client.open_by_key(SPREADSHEET_ID)

        # 3. JIKA KEDUANYA TIDAK ADA
        else:
            st.error("CRITICAL ERROR: Kredensial tidak ditemukan!")
            st.warning("""
            Penyebab:
            1. Jika di Laptop: File 'kunci_rahasia.json' tidak ada di folder ini.
            2. Jika di Cloud: Anda belum setting 'Secrets' di dashboard Streamlit.
            """)
            return None

    except Exception as e:
        st.error(f"GAGAL KONEKSI: {e}")
        return None

@st.cache_data(ttl=10, show_spinner=False)
def get_data(sheet_name):
    sh = connect_to_gsheet()
    if sh:
        try:
            ws = sh.worksheet(sheet_name)
            all_values = ws.get_all_values()
            if len(all_values) > 0:
                headers = all_values[0]
                data = all_values[1:]
                return pd.DataFrame(data, columns=headers)
            else:
                return pd.DataFrame()
        except Exception as e:
            st.error(f"ERROR GSPREAD: {e}")  # Tampilkan Error Asli di Layar
            return pd.DataFrame()
    else:
        st.error("ERROR: Gagal Login ke Google (Kredensial Salah/Tidak Ditemukan)")
        return pd.DataFrame()

# --- FUNGSI SAVE (UPSERT UNTUK ID UNIK) ---
def save_data(sheet_name, new_data_list):
    sh = connect_to_gsheet()
    if sh:
        try:
            ws = sh.worksheet(sheet_name)
            existing_data = ws.get_all_values()
            if not existing_data or len(existing_data) < 2:
                ws.append_rows(new_data_list)
                get_data.clear()
                return True
            headers = existing_data[0]
            old_rows = existing_data[1:]
            new_ids = {str(row[0]) for row in new_data_list}
            kept_rows = [row for row in old_rows if str(row[0]) not in new_ids]
            final_rows = kept_rows + new_data_list
            ws.clear()
            ws.append_rows([headers] + final_rows)
            get_data.clear()
            return True
        except Exception as e:
            st.error(f"Error: {str(e)}")
            return False
    return False

# --- FUNGSI UPDATE STATUS LOCK/UNLOCK (OPTIMIZED) ---
def update_status_locked(ids_to_lock, status_value="Locked"):
    sh = connect_to_gsheet()
    if sh:
        try:
            ws = sh.worksheet("Price_Data")
            vals = ws.get_all_values()
            
            if not vals: return False

            # Cari index kolom
            header = vals[0]
            try:
                id_idx = header.index("id_transaksi")
                status_idx = header.index("status")
            except ValueError:
                return False
            
            # Modifikasi di memori
            is_changed = False
            for i in range(1, len(vals)):
                if vals[i][id_idx] in ids_to_lock:
                    vals[i][status_idx] = status_value
                    is_changed = True
            
            # Upload ulang jika ada perubahan (Batch Update)
            if is_changed:
                ws.update(vals)
            
            get_data.clear()
            return True
        except Exception as e:
            st.error(f"Error update: {e}")
            return False
    return False

# --- ID GENERATORS ---
def generate_next_id(df, id_col, prefix, digits=3):
    if df.empty: return f"{prefix}{str(1).zfill(digits)}"
    existing_ids = [str(x) for x in df[id_col].tolist() if str(x).startswith(prefix)]
    if not existing_ids: return f"{prefix}{str(1).zfill(digits)}"
    max_num = 0
    for x in existing_ids:
        try:
            num_part = x.replace(prefix, "")
            if "-" in num_part: num_part = num_part.split("-")[0]
            val = int(num_part)
            if val > max_num: max_num = val
        except: continue
    return f"{prefix}{str(max_num + 1).zfill(digits)}"

def generate_child_id(df, parent_id, id_col):
    prefix = f"{parent_id}-"
    if df.empty: return f"{prefix}001"
    existing_ids = [str(x) for x in df[id_col].tolist() if str(x).startswith(prefix)]
    if not existing_ids: return f"{prefix}001"
    max_num = 0
    for x in existing_ids:
        try:
            suffix = x.split("-")[-1]
            val = int(suffix)
            if val > max_num: max_num = val
        except: continue
    return f"{prefix}{str(max_num + 1).zfill(3)}"

def clean_numeric(val):
    if pd.isna(val) or str(val).strip() == "": return None
    try: return float(str(val).replace(",", ""))
    except: return None

# --- TOMBOL SCROLL TO TOP (VERSI MANUAL ANCHOR) ---
def add_scroll_to_top():
    st.markdown("""
        <style>
            .scroll-to-top {
                position: fixed;
                bottom: 25px;
                right: 25px;
                background-color: #2563EB;
                color: white !important;
                width: 50px;
                height: 50px;
                border-radius: 50%;
                border: 2px solid #1D4ED8;
                text-align: center;
                box-shadow: 2px 2px 10px rgba(0,0,0,0.3);
                cursor: pointer;
                z-index: 99999;
                transition: all 0.3s ease-in-out;
                display: flex;
                align-items: center;
                justify-content: center;
                text-decoration: none; /* Hilangkan garis bawah link */
            }
            .scroll-to-top:hover {
                background-color: #FCA568; 
                border-color: #e38d4a;
                transform: translateY(-5px);
                box-shadow: 2px 5px 15px rgba(252, 165, 104, 0.5);
                color: white !important;
            }
            .scroll-to-top-icon {
                font-size: 24px;
                font-weight: 900;
                line-height: 1;
                margin-bottom: 4px; 
            }
        </style>
        
        <a href="#top-page" class="scroll-to-top" target="_self">
            <span class="scroll-to-top-icon">↑</span>
        </a>
    """, unsafe_allow_html=True)
    
# --- FUNGSI KIRIM EMAIL (UPDATED: WITH DUE DATE) ---
def send_invitation_email(to_email, vendor_name, load_type, validity, origins, password):
    # 1. Cek Config
    if "email_config" not in st.secrets:
        st.warning("Konfigurasi email belum disetting di Secrets. Email tidak terkirim.")
        return False

    sender_email = st.secrets["email_config"]["sender_email"]
    sender_password = st.secrets["email_config"]["sender_password"]
    
    # 2. Hitung Due Date (Hari ini + 14 Hari)
    today = datetime.now()
    due_date = today + timedelta(days=14)
    
    # 3. Format Tanggal ke Indonesia (dd MMMM yyyy)
    months_id = {
        1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
        7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"
    }
    due_date_str = f"{due_date.day} {months_id[due_date.month]} {due_date.year}"

    subject = f"Undangan Tender Transport {load_type} - {validity} (TACO Group)"
    origins_str = ", ".join(origins)
    
    # 4. Desain Isi Email (HTML) - Ditambahkan Due Date
    body = f"""
    <html>
    <body>
        <h3>Dear {vendor_name},</h3>
        <p>Anda telah diundang untuk berpartisipasi dalam Tender Transport <b>TACO Group</b>.</p>
        
        <p><b>Detail Tender:</b></p>
        <ul>
            <li><b>Periode:</b> {validity}</li>
            <li><b>Tipe Armada:</b> {load_type}</li>
            <li><b>Area/Origin:</b> {origins_str}</li>
            <li style="color: #d9534f;"><b>Batas Akhir Pengisian Penawaran Harga: {due_date_str}</b></li>
        </ul>
        
        <p>Silakan login ke sistem kami untuk memasukkan penawaran harga:</p>
        <p>
            <b>Link App:</b> <a href="https://taco-rfq.streamlit.app">https://taco-rfq.streamlit.app</a><br>
            <b>Email Login:</b> {to_email}<br>
            <b>Password:</b> {password}
        </p>
        
        <p>Terima Kasih,<br>Procurement Team TACO</p>
    </body>
    </html>
    """

    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html'))

        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"Gagal kirim email: {e}")
        return False
        

# --- FUNGSI GENERATE WORD (OPTIMIZED: SUPER CEPAT & RAPI) ---
def create_docx_sk(template_file, nomor_surat, validity, load_type, df_data):
    doc = DocxTemplate(template_file)
    
    # --- 1. SIAPKAN DATA ---
    unique_origins = sorted(df_data['origin'].unique())
    origin_list_str = ", ".join(unique_origins) 

    # --- HELPER 1: SET LEBAR KOLOM (DIPANGGIL SEKALI SAJA NANTI) ---
    def set_col_widths(table, widths):
        """Mengatur lebar kolom tabel dalam satuan Centimeter (Cm)"""
        # Matikan autofit agar ukuran kolom patuh
        table.autofit = False 
        table.allow_autofit = False
        
        for row in table.rows:
            for idx, width in enumerate(widths):
                if idx < len(row.cells):
                    row.cells[idx].width = width

    # --- HELPER 2: WARNA BACKGROUND CELL ---
    def set_cell_background(cell, color_hex):
        shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color_hex))
        cell._tc.get_or_add_tcPr().append(shading_elm)

    # --- HELPER 3: FORMAT PARAGRAF ---
    def format_paragraph(paragraph, size, bold=False, align=None):
        paragraph_format = paragraph.paragraph_format
        paragraph_format.space_after = Pt(0)
        paragraph_format.space_before = Pt(0)
        paragraph_format.line_spacing = 1
        if align is not None: paragraph.alignment = align
        run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.name = 'Arial'

    # --- HELPER 4: REPEAT HEADER ROW ---
    def set_repeat_table_header(row):
        tr = row._tr
        trPr = tr.get_or_add_trPr()
        tblHeader = OxmlElement('w:tblHeader')
        tblHeader.set(qn('w:val'), "true")
        trPr.append(tblHeader)

    # --- BAGIAN A: TABEL HARGA ---
    sd = doc.new_subdoc()
    winning_vendors_data = [] 
    
    for org in unique_origins:
        # Judul Origin
        p = sd.add_paragraph(f"Origin: {org}")
        p.paragraph_format.space_after = Pt(2)
        run = p.runs[0]; run.bold = True; run.font.size = Pt(10)
        
        # Filter & Top 3 Logic
        df_sub = df_data[df_data['origin'] == org].copy()
        df_sub = df_sub.sort_values(by=['kota_asal', 'kota_tujuan', 'unit_type', 'price'])
        df_sub['Ranking'] = df_sub.groupby(['kota_tujuan', 'unit_type']).cumcount() + 1
        df_sub = df_sub[df_sub['Ranking'] <= 3].copy()
        winning_vendors_data.append(df_sub)
        
        # Buat Tabel (8 Kolom)
        headers = ['Asal', 'Tujuan', 'Unit', 'Rank', 'Vendor', 'Biaya/unit', 'LeadTime', 'Term of Payment']
        table = sd.add_table(rows=1, cols=len(headers))
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER 
        
        # Header
        hdr_row = table.rows[0]; set_repeat_table_header(hdr_row)
        hdr_cells = hdr_row.cells
        for i, h in enumerate(headers):
            hdr_cells[i].text = h
            set_cell_background(hdr_cells[i], "ED7D31")
            hdr_cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER 
            format_paragraph(hdr_cells[i].paragraphs[0], size=8, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
            
        # Data Loop
        for _, row in df_sub.iterrows():
            row_cells = table.add_row().cells
            try: harga = f"Rp {int(row['price']):,}".replace(",", ".")
            except: harga = "Rp 0"
              
            # --- LOGIC LEAD TIME (+ Hari) ---
            lt_raw = str(row['lead_time'])
            if lt_raw.isdigit() or (lt_raw.replace('.','',1).isdigit()):
                lt_fmt = f"{lt_raw} Hari"
            elif lt_raw in ["-", "", "0"]:
                lt_fmt = "-"
            else:
                lt_fmt = f"{lt_raw} Hari"
            
            data_map = [
                str(row.get('kota_asal', '-')), 
                str(row['kota_tujuan']), 
                str(row['unit_type']),
                str(row['Ranking']), 
                str(row['vendor_name']), 
                harga,
                lt_fmt,
                str(row['top'])
            ]
            
            for idx, val in enumerate(data_map):
                cell = row_cells[idx]; cell.text = val
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                
                # Alignment
                if idx in [0, 1, 2, 6]: align = WD_ALIGN_PARAGRAPH.LEFT
                elif idx == 5: align = WD_ALIGN_PARAGRAPH.RIGHT
                else: align = WD_ALIGN_PARAGRAPH.CENTER
                
                format_paragraph(cell.paragraphs[0], size=7, bold=False, align=align)
        
        # --- OPTIMASI PENTING: SET LEBAR SETELAH LOOP SELESAI ---
        # Asal(2.0), Tujuan(2.8), Unit(2.2), Rank(0.8), Vendor(3.5), Harga(2.5), L.Time(1.2), TOP(2.0)
        col_widths = [Cm(2.0), Cm(2.8), Cm(2.2), Cm(0.8), Cm(3.5), Cm(2.5), Cm(1.2), Cm(2.0)]
        set_col_widths(table, col_widths) 
        # --------------------------------------------------------

        sd.add_paragraph("") 

    # --- BAGIAN B: TABEL VENDOR ---
    sd_ven = doc.new_subdoc()
    if winning_vendors_data:
        df_all = pd.concat(winning_vendors_data)
        df_uniq = df_all.drop_duplicates(subset=['vendor_email']).sort_values('vendor_name')
    else: df_uniq = pd.DataFrame()
    
    if not df_uniq.empty:
        v_headers = ['Vendor', 'Alamat', 'Email', 'PIC', 'Telepon']
        v_table = sd_ven.add_table(rows=1, cols=len(v_headers))
        v_table.style = 'Table Grid'
        v_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        v_hdr = v_table.rows[0]; set_repeat_table_header(v_hdr)
        vh_cells = v_hdr.cells
        for i, h in enumerate(v_headers):
            vh_cells[i].text = h
            set_cell_background(vh_cells[i], "ED7D31")
            vh_cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER 
            format_paragraph(vh_cells[i].paragraphs[0], size=8, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
            
        for _, row in df_uniq.iterrows():
            v_cells = v_table.add_row().cells
            v_data = [
                str(row['vendor_name']), 
                str(row.get('address', '-')), 
                str(row['vendor_email']),
                str(row.get('contact_person', '-')), 
                str(row.get('phone', '-'))
            ]
            for idx, val in enumerate(v_data):
                cell = v_cells[idx]; cell.text = val
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                format_paragraph(cell.paragraphs[0], size=7, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT)

        # --- OPTIMASI PENTING: SET LEBAR SETELAH LOOP SELESAI ---
        # Vendor(4.0), Alamat(5.0), Email(3.5), PIC(2.5), Telepon(2.5)
        v_widths = [Cm(4.0), Cm(5.0), Cm(3.5), Cm(2.5), Cm(2.5)]
        set_col_widths(v_table, v_widths)
   
    bulan_indo = {1:'Januari', 2:'Februari', 3:'Maret', 4:'April', 5:'Mei', 6:'Juni', 7:'Juli', 8:'Agustus', 9:'September', 10:'Oktober', 11:'November', 12:'Desember'}
    today = datetime.now()
    tgl_str = f"{today.day} {bulan_indo[today.month]} {today.year}"
    
    context = {
        'no_surat': nomor_surat,
        'validity': validity,
        'load_type': load_type,
        'tanggal_sk': tgl_str,
        'daftar_origin': origin_list_str,
        'tabel_harga': sd,      
        'tabel_vendor': sd_ven  
    }
    
    doc.render(context)
    output_filename = f"SK_Result_{int(time.time())}.docx"
    doc.save(output_filename)
    return output_filename
    
# --- FUNGSI GENERATE SPK (UPDATE: ALAMAT DARI SHEET) ---
def create_docx_spk(template_file, no_spk, validity, load_type, vendor_name, pic_vendor, vendor_pass, alamat_gudang, df_data):
    doc = DocxTemplate(template_file)
    
    # Ambil Origin Utama untuk Judul
    origin_utama = df_data.iloc[0]['origin'] if not df_data.empty else "-"

    # Format Tanggal
    bulan_indo = {1:'Januari', 2:'Februari', 3:'Maret', 4:'April', 5:'Mei', 6:'Juni', 7:'Juli', 8:'Agustus', 9:'September', 10:'Oktober', 11:'November', 12:'Desember'}
    today = datetime.now()
    tgl_spk = f"{today.day} {bulan_indo[today.month]} {today.year}"

    # Helper Format Rupiah
    def fmt_rp(x):
        try: return f"Rp {int(x):,}".replace(",", ".")
        except: return "Rp 0"

    def set_col_widths(table, widths):
        table.autofit = False 
        table.allow_autofit = False
        for row in table.rows:
            for idx, width in enumerate(widths):
                if idx < len(row.cells):
                    row.cells[idx].width = width

    def format_cell(cell, text, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, size=7):
        cell.text = str(text)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = cell.paragraphs[0]
        p.alignment = align
        p.paragraph_format.space_after = Pt(0)
        run = p.runs[0] if p.runs else p.add_run()
        run.font.size = Pt(size)
        run.font.name = 'Arial'
        run.font.bold = bold

    # --- TABEL DATA ---
    sd = doc.new_subdoc()
    headers = ['Asal', 'Tujuan', 'Unit', 'Rank', 'Biaya/Unit', 'MD Inner', 'MD Outer', 'Buruh', 'LeadTime']
    table = sd.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Header Styling
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        shading_elm = parse_xml(r'<w:shd {} w:fill="ED7D31"/>'.format(nsdecls('w')))
        hdr_cells[i]._tc.get_or_add_tcPr().append(shading_elm)
        format_cell(hdr_cells[i], h, bold=True, size=7)

    # Isi Data
    for idx, row in df_data.iterrows():
        row_cells = table.add_row().cells
        
        lt_raw = str(row['lead_time'])
        lt_fmt = f"{lt_raw} Hari" if (lt_raw.isdigit() or lt_raw.replace('.','',1).isdigit()) and lt_raw not in ["-","0",""] else "-"

        vals = [
            row.get('kota_asal','-'), 
            row['kota_tujuan'], 
            row['unit_type'], 
            str(row['Ranking']), 
            fmt_rp(row['price']),
            fmt_rp(row.get('inner_city_price', 0)),
            fmt_rp(row.get('outer_city_price', 0)),
            fmt_rp(row.get('labor_cost', 0)),
            lt_fmt
        ]
        
        for i, v in enumerate(vals):
            align = WD_ALIGN_PARAGRAPH.LEFT if i in [0,1] else WD_ALIGN_PARAGRAPH.CENTER
            if i in [4, 5, 6, 7]: align = WD_ALIGN_PARAGRAPH.RIGHT
            format_cell(row_cells[i], v, align, size=6.5)

    col_widths = [Cm(2.0), Cm(3.0), Cm(1.8), Cm(0.8), Cm(2.2), Cm(2.0), Cm(2.0), Cm(1.5), Cm(1.5)]
    set_col_widths(table, col_widths)

    # Context Mapping
    context = {
        'no_spk': no_spk, 
        'tanggal_spk': tgl_spk, 
        'validity': validity,
        'load_type': load_type, 
        'vendor_name': vendor_name, 
        'contact_person': pic_vendor, 
        'password_vendor': vendor_pass,
        'origin_name': origin_utama,
        'alamat_gudang': alamat_gudang, # <--- Diambil dari parameter
        'tabel_harga_vendor': sd
    }
    doc.render(context)
    
    safe_ven = "".join(x for x in vendor_name if x.isalnum())
    fn = f"temp_spk_{safe_ven}_{int(time.time())}.docx"
    doc.save(fn)
    return fn

def main():
    st.markdown('<div id="top-page"></div>', unsafe_allow_html=True)
    init_style()
    add_scroll_to_top()
    c_logo, _ = st.columns([1, 6])
    with c_logo:
        if os.path.exists("image_2.png"): st.image("image_2.png", width=120)
        else: st.markdown("## **TACO**") 

    if 'user_info' not in st.session_state: st.session_state['user_info'] = None
    if 'vendor_step' not in st.session_state: st.session_state['vendor_step'] = "dashboard" 
    
    for k in ['selected_group_id', 'selected_validity', 'sel_origin', 'sel_load', 'focused_group_id', 'temp_success_msg']:
        if k not in st.session_state: st.session_state[k] = None

    # --- LOGIN ---
    if not st.session_state['user_info']:
        c1, c2, c3 = st.columns([1, 2, 1])
        with c2:
            st.markdown("<br><br>", unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown("### 🔐 TACO Transport RFQ")
                with st.form("login"):
                    email = st.text_input("Email")
                    pw = st.text_input("Password", type="password")
                    if st.form_submit_button("Masuk", type="primary", use_container_width=True):
                        df = get_data("Users")
                        if not df.empty:
                            u = df[(df['email']==email) & (df['password']==pw)]
                            if not u.empty:
                                st.session_state['user_info'] = u.iloc[0].to_dict()
                                st.rerun()
                            else: st.error("Gagal Login")
                        else: st.error("DB Kosong")
    
    # --- DASHBOARD ---
    else:
        user = st.session_state['user_info']
        role = user['role']

        with st.container():
            c1, c2 = st.columns([8,1])
            with c1: 
                st.markdown(f"### 👋 Halo, **{user.get('vendor_name')}**")
                st.caption(f"Role: {role.upper()}")
            with c2: 
                if st.button("Logout", type="secondary"):
                    st.session_state['user_info'] = None
                    st.session_state['vendor_step'] = "dashboard"
                    st.cache_data.clear()
                    st.rerun()
        
        st.markdown("---")

        if role == 'admin':
            admin_dashboard()
        elif role == 'vendor':
            vendor_dashboard(user['email'])
        else:
            user_dashboard()

# ================= INTERNAL USER DASHBOARD =================
def user_dashboard():
    st.subheader("🔍 Portal Pencarian Tarif")
    
    # --- LOAD DATA ---
    df_p = get_data("Price_Data")
    df_r = get_data("Master_Routes")
    df_g = get_data("Master_Groups")
    df_u = get_data("Users")
    df_prof = get_data("Vendor_Profile")
    df_md = get_data("Multidrop_Data") # Load Multidrop untuk Tab Search

    # --- PREPARE MASTER DATA ---
    df_master = pd.DataFrame()
    if not df_p.empty:
        # Cleaning ID
        df_p['route_id'] = df_p['route_id'].astype(str).str.strip()
        df_r['route_id'] = df_r['route_id'].astype(str).str.strip()
        df_g['group_id'] = df_g['group_id'].astype(str).str.strip()
        if not df_md.empty:
            df_md['vendor_email'] = df_md['vendor_email'].astype(str).str.strip()
            df_md['group_id'] = df_md['group_id'].astype(str).str.strip()
            df_md['validity'] = df_md['validity'].astype(str).str.strip()

        # Merge 1: Price + Routes
        m1 = pd.merge(df_p, df_r, on='route_id', how='left')
        # Merge 2: + Groups
        m2 = pd.merge(m1, df_g, on='group_id', how='left')
        # Merge 3: + User (Vendor Name)
        m3 = pd.merge(m2, df_u[['email', 'vendor_name']], left_on='vendor_email', right_on='email', how='left')
        m3['vendor_name'] = m3['vendor_name'].fillna(m3['vendor_email'])
        
        # Merge 4: + Vendor Profile (TOP, Address, etc)
        if not df_prof.empty:
            df_prof_clean = df_prof.sort_values('updated_at', ascending=False).drop_duplicates('email')
            m4 = pd.merge(m3, df_prof_clean[['email', 'top']], left_on='vendor_email', right_on='email', how='left')
            m4['top'] = m4['top'].fillna("-")
        else:
            m4 = m3
            m4['top'] = "-"

        m4['price'] = pd.to_numeric(m4['price'], errors='coerce').fillna(0)
        
        # Filter hanya data yang sudah Locked (Final) agar user tidak melihat harga draft? 
        # Opsional: Jika ingin semua harga tampil, hapus baris ini.
        # df_master = m4[m4['status'] == 'Locked'].copy() 
        df_master = m4.copy() # Tampilkan semua status (Open/Locked)
    
    # --- TABS ---
    tab1, tab2 = st.tabs(["📊 Summary & Ranking", "🔎 Cari Vendor per Rute"])

    # === TAB 1: SUMMARY RANKING (Mirip Admin) ===
    with tab1:
        if df_master.empty:
            st.info("Data harga belum tersedia.")
        else:
            c1, c2 = st.columns(2)
            # Filter
            avail_val = sorted(df_master['validity'].unique().tolist())
            avail_load = sorted(df_master['load_type'].unique().tolist())
            
            sel_val = c1.selectbox("Filter Periode", avail_val, key="sum_val")
            sel_load = c2.selectbox("Filter Tipe Muatan", avail_load, key="sum_load")
            
            df_view = df_master[(df_master['validity'] == sel_val) & (df_master['load_type'] == sel_load)].copy()
            
            if not df_view.empty:
                unique_origins = sorted(df_view['origin'].unique())
                for org in unique_origins:
                    with st.expander(f"📍 Origin: {org}", expanded=False):
                        sub_df = df_view[df_view['origin'] == org].copy()
                        # Ranking Logic
                        sub_df = sub_df.sort_values(by=['kota_tujuan', 'unit_type', 'price'])
                        sub_df['Ranking'] = sub_df.groupby(['kota_tujuan', 'unit_type']).cumcount() + 1
                        
                        # Filter Top 3
                        sub_df = sub_df[sub_df['Ranking'] <= 3]
                        
                        sub_df['price_fmt'] = sub_df['price'].apply(lambda x: f"Rp {int(x):,}".replace(",", "."))
                        
                        st.dataframe(
                            sub_df[['kota_tujuan', 'unit_type', 'Ranking', 'vendor_name', 'price_fmt', 'lead_time', 'top']],
                            use_container_width=True,
                            hide_index=True,
                            column_config={"kota_tujuan": "Tujuan", "price_fmt": "Harga", "vendor_name": "Vendor", "top": "TOP"}
                        )
            else:
                st.warning("Data tidak ditemukan untuk filter ini.")

    # === TAB 2: SEARCH VENDOR BY ROUTE ===
    with tab2:
        if df_master.empty:
            st.info("Data belum tersedia.")
        else:
            # 1. Filter Utama
            c1, c2, c3 = st.columns(3)
            # Filter Periode
            avail_val_s = sorted(df_master['validity'].unique().tolist())
            s_val = c1.selectbox("1. Pilih Periode", avail_val_s, key="s_val")
            
            # Filter Load Type
            avail_load_s = sorted(df_master['load_type'].unique().tolist())
            s_load = c2.selectbox("2. Pilih Muatan", avail_load_s, key="s_load")
            
            # Filter Origin (Dinamis berdasarkan 2 filter sebelumnya)
            filtered_1 = df_master[(df_master['validity'] == s_val) & (df_master['load_type'] == s_load)]
            avail_org_s = sorted(filtered_1['origin'].unique().tolist())
            s_org = c3.selectbox("3. Pilih Origin", avail_org_s, key="s_org")
            
            # 2. Input Search Kota Tujuan
            st.write("")
            search_dest = st.text_input("🔍 Cari Kota Tujuan (Ketik nama kota...)", placeholder="Contoh: Surabaya").strip()
            
            if search_dest:
                # Filter Data
                df_search = filtered_1[filtered_1['origin'] == s_org].copy()
                # Filter Fuzzy / Contains untuk Kota Tujuan
                df_search = df_search[df_search['kota_tujuan'].str.contains(search_dest, case=False, na=False)]
                
                if not df_search.empty:
                    # GABUNGKAN DATA MULTIDROP & BURUH
                    # Kita perlu merge df_search dengan df_md berdasarkan (vendor_email, validity, group_id)
                    
                    if not df_md.empty:
                        # Rename kolom agar tidak bentrok saat merge atau lebih jelas
                        md_subset = df_md[['vendor_email', 'validity', 'group_id', 'inner_city_price', 'outer_city_price', 'labor_cost']].copy()
                        
                        # Merge
                        df_result = pd.merge(
                            df_search, 
                            md_subset, 
                            on=['vendor_email', 'validity', 'group_id'], 
                            how='left'
                        )
                        
                        # Fill NaN (Jika vendor belum isi multidrop)
                        for c in ['inner_city_price', 'outer_city_price', 'labor_cost']:
                            df_result[c] = pd.to_numeric(df_result[c], errors='coerce').fillna(0)
                    else:
                        df_result = df_search.copy()
                        df_result['inner_city_price'] = 0
                        df_result['outer_city_price'] = 0
                        df_result['labor_cost'] = 0

                    # Sorting: Harga Termurah -> Termahal
                    df_result = df_result.sort_values(by=['unit_type', 'price'])
                    
                    # Format Rupiah
                    def fmt_rp(x): return f"Rp {int(x):,}".replace(",", ".")
                    
                    df_result['Harga Unit'] = df_result['price'].apply(fmt_rp)
                    df_result['Multidrop Dalam'] = df_result['inner_city_price'].apply(fmt_rp)
                    df_result['Multidrop Luar'] = df_result['outer_city_price'].apply(fmt_rp)
                    df_result['Biaya Buruh'] = df_result['labor_cost'].apply(fmt_rp)
                    
                    # Tampilkan Hasil
                    # Grouping per Unit Type agar rapi
                    unique_units = df_result['unit_type'].unique()
                    
                    st.success(f"Ditemukan {len(df_result)} penawaran untuk tujuan '{search_dest}'.")
                    
                    for unit in unique_units:
                        st.markdown(f"##### 🚛 Unit: {unit}")
                        sub_res = df_result[df_result['unit_type'] == unit]
                        
                        # Kolom yang diminta: Urutan Vendor, Harga, TOP, Biaya Multidrop, Biaya Buruh, Leadtime
                        # Kita buat tabel ranking
                        sub_res['Rank'] = range(1, len(sub_res) + 1)
                        
                        display_cols = [
                            'Rank', 'vendor_name', 'Harga Unit', 'top', 
                            'lead_time', 'Multidrop Dalam', 'Multidrop Luar', 'Biaya Buruh'
                        ]
                        
                        st.dataframe(
                            sub_res[display_cols],
                            use_container_width=True,
                            hide_index=True,
                            column_config={
                                "vendor_name": "Vendor",
                                "top": "TOP (Term of Payment)",
                                "lead_time": "Lead Time (Hari)"
                            }
                        )
                        st.markdown("---")
                else:
                    st.warning(f"Tidak ditemukan rute ke '{search_dest}' dari {s_org}.")
            else:
                st.info("Silakan ketik nama kota tujuan di atas untuk mulai mencari.")
                
# ================= ADMIN =================
def admin_dashboard():
    tabs = st.tabs(["📍Master Groups", "🛣️Master Routes", "🚛Master Units", "👥Users", "🔑Access Rights", "✅Monitoring & Approval", "📊Summary", "🖨️Print File"])
# --- LOAD DATA SEKALI UNTUK SEMUA TAB ANALISA (OPTIMASI) ---
    # Kita load di sini agar Tab 7 dan Tab 8 bisa pakai data yang sama
    df_p = get_data("Price_Data")
    df_r = get_data("Master_Routes")
    df_g = get_data("Master_Groups")
    df_u = get_data("Users")
    df_prof = get_data("Vendor_Profile")
    
    # Persiapkan Data "Big Join" (Master Data Gabungan)
    df_master = pd.DataFrame()
    if not df_p.empty:
        # Cleaning & Merge
        df_p['route_id'] = df_p['route_id'].astype(str).str.strip()
        df_r['route_id'] = df_r['route_id'].astype(str).str.strip()
        df_g['group_id'] = df_g['group_id'].astype(str).str.strip()
        
        m1 = pd.merge(df_p, df_r, on='route_id', how='left')
        m2 = pd.merge(m1, df_g, on='group_id', how='left')
        m3 = pd.merge(m2, df_u[['email', 'vendor_name']], left_on='vendor_email', right_on='email', how='left')
        m3['vendor_name'] = m3['vendor_name'].fillna(m3['vendor_email'])
        
        if not df_prof.empty:
            df_prof_clean = df_prof.sort_values('updated_at', ascending=False).drop_duplicates('email')
            
            # Kita ambil kolom alamat & kontak juga
            cols_to_merge = ['email', 'top', 'ppn', 'pph', 'address', 'contact_person', 'phone']
            
            m4 = pd.merge(m3, df_prof_clean[cols_to_merge], left_on='vendor_email', right_on='email', how='left')
            
            # Fill NA dengan "-"
            fill_cols = ['top', 'ppn', 'pph', 'address', 'contact_person', 'phone']
            for c in fill_cols:
                if c in m4.columns: m4[c] = m4[c].fillna("-")
                
            m4['ppn_status'] = m4['ppn']
            m4['pph_status'] = m4['pph']
        else:
            m4 = m3
            for c in ['top', 'ppn_status', 'pph_status', 'address', 'contact_person', 'phone']:
                m4[c] = "-"
        

        m4['price'] = pd.to_numeric(m4['price'], errors='coerce').fillna(0)
        df_master = m4
    
    # --- TAB 1: GROUPS ---
    with tabs[0]:
        st.caption("Buat Group Baru (ID Otomatis)")
        with st.form("add_grp"):
            c1, c2, c3 = st.columns(3)
            lt = c1.selectbox("Load Type", ["FTL", "FCL"])
            org = c2.text_input("Origin (Area)") 
            gn = c3.text_input("Nama Route Group")
            
            if st.form_submit_button("Simpan Group", type="primary"):
                df_g = get_data("Master_Groups")
                duplicate = False
                if not df_g.empty:
                    check = df_g[
                        (df_g['load_type'] == lt) & 
                        (df_g['origin'].str.lower() == org.lower()) & 
                        (df_g['route_group'].str.lower() == gn.lower())
                    ]
                    if not check.empty: duplicate = True
                
                if duplicate: st.error(f"Gagal: Group '{gn}' dengan Origin '{org}' dan tipe '{lt}' sudah ada!")
                elif not org or not gn: st.warning("Lengkapi data.")
                else:
                    new_gid = generate_next_id(df_g, 'group_id', 'R-', 3)
                    save_data("Master_Groups", [[new_gid, lt, gn, org]])
                    st.success(f"Berhasil! ID: {new_gid}")
                    time.sleep(1); st.rerun()
        st.dataframe(get_data("Master_Groups"), use_container_width=True)

    # --- TAB 2: ROUTES ---
    with tabs[1]:
        st.caption("Tambah Rute")
        df_g = get_data("Master_Groups")
        
        c_f1, c_f2 = st.columns(2)
        f_lt = c_f1.selectbox("Filter Load Type", ["All"] + df_g['load_type'].unique().tolist() if not df_g.empty else [])
        avail_origins = []
        if not df_g.empty:
            if f_lt != "All": avail_origins = df_g[df_g['load_type']==f_lt]['origin'].unique().tolist()
            else: avail_origins = df_g['origin'].unique().tolist()
        f_org = c_f2.selectbox("Filter Origin", ["All"] + avail_origins)
        
        grp_opts = {}
        if not df_g.empty:
            filtered_g = df_g.copy()
            if f_lt != "All": filtered_g = filtered_g[filtered_g['load_type'] == f_lt]
            if f_org != "All": filtered_g = filtered_g[filtered_g['origin'] == f_org]
            for _, r in filtered_g.iterrows():
                label = f"{r['group_id']} | {r['route_group']} ({r['load_type']})"
                grp_opts[label] = r['group_id']

        with st.form("add_rt"):
            c1, c2, c3 = st.columns([2, 1, 1])
            sel_label = c1.selectbox("Pilih Group", list(grp_opts.keys()) if grp_opts else [])
            ka = c2.text_input("Kota Asal")
            kt = c3.text_input("Kota Tujuan")
            ket = st.text_input("Keterangan (Opsional)")
            
            if st.form_submit_button("Simpan Rute", type="primary"):
                if sel_label and ka and kt:
                    gid = grp_opts[sel_label]
                    df_r = get_data("Master_Routes")
                    new_rid = generate_child_id(df_r, gid, 'route_id')
                    save_data("Master_Routes", [[new_rid, gid, ka, kt, ket]])
                    st.success(f"Tersimpan! ID: {new_rid}")
                    time.sleep(1); st.rerun()
                else: st.warning("Data belum lengkap.")
        st.dataframe(get_data("Master_Routes"), use_container_width=True)

    # --- TAB 3: UNITS ---
    with tabs[2]:
        st.caption("Setting Unit per Group (Tanpa Unit ID)")
        all_grp_opts = {}
        if not df_g.empty:
             for _, r in df_g.iterrows():
                label = f"{r['group_id']} - {r['route_group']} ({r['origin']})"
                all_grp_opts[label] = r['group_id']

        with st.form("add_ut"):
            c1, c2 = st.columns(2)
            sel_g = c1.selectbox("Pilih Group", list(all_grp_opts.keys()))
            ut = c2.text_input("Jenis Unit (ex: Tronton)")
            
            if st.form_submit_button("Tambah Unit", type="primary"):
                if sel_g and ut:
                    gid = all_grp_opts[sel_g]
                    sh = connect_to_gsheet()
                    if sh:
                        ws = sh.worksheet("Master_Units")
                        existing = ws.get_all_values()
                        is_exist = False
                        if len(existing) > 1:
                            for row in existing[1:]:
                                if len(row) >= 2 and row[0] == gid and row[1].lower() == ut.lower():
                                    is_exist = True; break
                        if is_exist: st.error(f"Unit '{ut}' sudah ada di Group {gid}.")
                        else:
                            ws.append_rows([[gid, ut]])
                            get_data.clear() 
                            st.success("Unit tersimpan."); time.sleep(0.5); st.rerun()
        st.dataframe(get_data("Master_Units"), use_container_width=True)

    # --- TAB 4: USERS ---
    with tabs[3]:
        with st.form("add_usr"):
            c1, c2, c3 = st.columns(3)
            em = c1.text_input("Email"); pw = c2.text_input("Pass"); nm = c3.text_input("PT Name")
            if st.form_submit_button("Add User", type="primary"):
                save_data("Users", [[em, pw, "vendor", nm]])
                st.success("Saved")
        st.dataframe(get_data("Users"), use_container_width=True)

    # --- TAB 5: ACCESS ---
    with tabs[4]:
        st.write("Grant Access (Batch per Origin)")
        df_u = get_data("Users"); df_g = get_data("Master_Groups")
        df_rights = get_data("Access_Rights")
        
        if not df_u.empty and not df_g.empty:
            c1, c2 = st.columns(2)
            ven = c1.selectbox("Pilih Vendor", df_u[df_u['role']=='vendor']['email'].unique())
            sel_lt = c2.selectbox("Pilih Load Type", ["FTL", "FCL"])
            
            unique_origins = []
            if not df_g.empty:
                subset_g = df_g[df_g['load_type'] == sel_lt]
                unique_origins = sorted(subset_g['origin'].unique().tolist())
            
            if not unique_origins:
                st.info(f"Belum ada data Origin untuk {sel_lt}.")
            else:
                st.write(f"**Pilih Origin / Area ({sel_lt}):**")
                
                existing_gids = set()
                if not df_rights.empty:
                    existing_gids = set(df_rights[df_rights['vendor_email'] == ven]['group_id'].tolist())
                
                with st.form("acc_origin_batch"):
                    cols = st.columns(3)
                    selected_origins = []
                    
                    for idx, org in enumerate(unique_origins):
                        col_ptr = cols[idx % 3]
                        org_groups = df_g[(df_g['origin'] == org) & (df_g['load_type'] == sel_lt)]
                        org_gids = set(org_groups['group_id'].tolist())
                        is_checked = not org_gids.isdisjoint(existing_gids)
                        
                        if col_ptr.checkbox(org, value=False, key=f"chk_org_{org}_{sel_lt}"):
                            selected_origins.append(org)
                    
                    st.divider()
                    c_val1, c_val2 = st.columns([2, 1])
                    with c_val1:
                        val_period = st.selectbox("Pilih Periode", ["Januari - Juni", "Juli - Desember", "Januari - Desember"])
                    with c_val2:
                        val_year = st.text_input("Tahun", value=str(datetime.now().year))
                    
                    # ... (Bagian atas Tab 5 tetap sama) ...
                    
                    if st.form_submit_button("Grant Access", type="primary"):
                        val = f"{val_period} {val_year}"
                        
                        if selected_origins and val_year:
                            # ... (Logika filter target_groups tetap sama) ...
                            target_groups = df_g[(df_g['load_type'] == sel_lt) & (df_g['origin'].isin(selected_origins))]
                            
                            if not target_groups.empty:
                                target_gids = target_groups['group_id'].unique().tolist()
                                sh = connect_to_gsheet()
                                if sh:
                                    ws = sh.worksheet("Access_Rights")
                                    # ... (Logika existing_keys tetap sama) ...
                                    existing_data = ws.get_all_values()
                                    existing_keys = set()
                                    if len(existing_data) > 1:
                                        for row in existing_data[1:]:
                                            if len(row) >= 3:
                                                key = f"{row[0]}_{row[1]}_{row[2]}".lower()
                                                existing_keys.add(key)
                                    
                                    new_rows_to_add = []
                                    skipped_count = 0
                                    
                                    for gid in target_gids:
                                        key_check = f"{ven}_{val}_{gid}".lower()
                                        if key_check not in existing_keys:
                                            new_rows_to_add.append([ven, val, gid, "Active"])
                                        else:
                                            skipped_count += 1
                                    
                                    if new_rows_to_add:
                                        # 1. Simpan ke Google Sheet
                                        ws.append_rows(new_rows_to_add)
                                        get_data.clear()
                                        
                                        # 2. KIRIM EMAIL NOTIFIKASI
                                        # Ambil password user dari df_u
                                        try:
                                            user_row = df_u[df_u['email'] == ven].iloc[0]
                                            ven_name = user_row['vendor_name']
                                            ven_pass = user_row['password']
                                            
                                            with st.spinner("Mengirim email undangan ke vendor..."):
                                                email_status = send_invitation_email(ven, ven_name, sel_lt, val, selected_origins, ven_pass)
                                                
                                            if email_status:
                                                st.success(f"✅ Sukses! Akses diberikan & Email undangan terkirim ke {ven}.")
                                            else:
                                                st.warning("⚠️ Akses tersimpan, tapi Email GAGAL terkirim.")
                                                
                                        except Exception as e:
                                            st.warning(f"Akses tersimpan, tapi gagal mengambil data user untuk email: {e}")

                                        time.sleep(2); st.rerun()
                                    else:
                                        st.warning(f"Semua data sudah ada ({skipped_count} skip). Tidak ada update.")
                            else:
                                st.warning("Data Group error.")
                        else:
                            st.warning("Pilih Origin dan isi Tahun.")
                            
        st.dataframe(get_data("Access_Rights"), use_container_width=True)

# --- TAB 6: MONITORING (UPDATE: ADA LOAD TYPE DI JUDUL) ---
    with tabs[5]:
        st.subheader("Monitoring Harga & Approval")
        df_price = get_data("Price_Data")
        df_routes = get_data("Master_Routes")
        df_md = get_data("Multidrop_Data")
        df_g = get_data("Master_Groups")
        
        if df_price.empty:
            st.info("Belum ada data harga masuk.")
        else:
            df_price['route_id'] = df_price['route_id'].astype(str).str.strip()
            df_routes['route_id'] = df_routes['route_id'].astype(str).str.strip()
            
            merged_pr = pd.merge(df_price, df_routes[['route_id', 'group_id', 'kota_asal', 'kota_tujuan']], on='route_id', how='left')
            merged_pr['group_id'] = merged_pr['group_id'].fillna('Unknown')
            
            if not df_g.empty:
                # UPDATE 1: Ambil kolom 'load_type' juga dari Master Group
                merged_pr = pd.merge(merged_pr, df_g[['group_id', 'route_group', 'load_type']], on='group_id', how='left')
            else: 
                merged_pr['route_group'] = 'Unknown Group'
                merged_pr['load_type'] = '-'
                
            merged_pr['route_group'] = merged_pr['route_group'].fillna('Unknown Group')
            merged_pr['load_type'] = merged_pr['load_type'].fillna('-') # Handle jika kosong
            merged_pr['kota_asal'] = merged_pr['kota_asal'].fillna('Unknown')
            merged_pr['kota_tujuan'] = merged_pr['kota_tujuan'].fillna('Unknown')

            merged_pr['key_group'] = merged_pr['vendor_email'] + " | " + merged_pr['validity'] + " | " + merged_pr['route_group'] + " | " + merged_pr['group_id']
            unique_keys = merged_pr['key_group'].unique()
            
            for key in unique_keys:
                parts = key.split(" | ")
                vendor, validity, g_name, g_id = parts[0], parts[1], parts[2], parts[3]
                subset_pr = merged_pr[merged_pr['key_group'] == key]
                if subset_pr.empty: continue
                
                is_locked = "Locked" in subset_pr['status'].values
                status_icon = "🔒 LOCKED" if is_locked else "🟢 OPEN"
                
                # UPDATE 2: Ambil Load Type dari data baris pertama grup ini
                l_type = subset_pr.iloc[0]['load_type']
                
                # Masukkan ke judul Expander
                with st.expander(f"{status_icon} - {l_type} - {vendor} ({validity}) - {g_name}"):
                    st.markdown("**A. Spesifikasi Armada**")
                    if {'unit_type', 'weight_capacity', 'cubic_capacity'}.issubset(subset_pr.columns):
                        df_specs = subset_pr[['unit_type', 'weight_capacity', 'cubic_capacity']].drop_duplicates().reset_index(drop=True)
                        st.dataframe(df_specs, use_container_width=True, hide_index=True)
                    
                    st.markdown("**B. Matriks Harga**")
                    try:
                        subset_pr['price'] = pd.to_numeric(subset_pr['price'], errors='coerce')
                        pivot_df = subset_pr.pivot_table(index=['kota_asal', 'kota_tujuan'], columns='unit_type', values='price', aggfunc='first').reset_index()
                        st.dataframe(pivot_df, use_container_width=True, hide_index=True)
                    except: st.dataframe(subset_pr, use_container_width=True)

                    st.markdown("**C. Biaya Lain (Multidrop & Buruh)**")
                    if not df_md.empty:
                        df_md['vendor_email'] = df_md['vendor_email'].astype(str).str.strip()
                        df_md['validity'] = df_md['validity'].astype(str).str.strip()
                        df_md['group_id'] = df_md['group_id'].astype(str).str.strip()
                        
                        curr_ven = str(vendor).strip()
                        curr_val = str(validity).strip()
                        curr_gid = str(g_id).strip()

                        sub_md = df_md[
                            (df_md['vendor_email'] == curr_ven) &
                            (df_md['validity'] == curr_val) &
                            (df_md['group_id'] == curr_gid)
                        ]

                        if not sub_md.empty:
                            cols_to_show = ['inner_city_price', 'outer_city_price']
                            header_names = ["Multidrop Dalam Kota", " Multidrop Luar Kota"]
                            
                            if 'labor_cost' in sub_md.columns:
                                cols_to_show.append('labor_cost')
                                header_names.append("Biaya Buruh")
                            
                            disp_md = sub_md[cols_to_show].reset_index(drop=True)
                            disp_md.columns = header_names
                            st.dataframe(disp_md, use_container_width=True, hide_index=True)
                        else:
                            st.info("Data Multidrop belum diinput oleh vendor.")
                    else:
                        st.info("Database Multidrop masih kosong.")
                    
                    st.divider()
                    c1, c2 = st.columns([1, 4])
                    ids = subset_pr['id_transaksi'].tolist()
                    if is_locked:
                        if c1.button("🔓 UNLOCK DATA", key=f"ul_{key}"):
                            update_status_locked(ids, "Open")
                            st.success("Unlocked!"); time.sleep(0.5); st.rerun()
                    else:
                        if c1.button("🔒 LOCK DATA", key=f"lk_{key}", type="primary"):
                            update_status_locked(ids, "Locked")
                            st.success("Locked!"); time.sleep(0.5); st.rerun()
    
    # --- TAB 7: SUMMARY ---   
    with tabs[6]:
        st.subheader("📊 Summary & Ranking Vendor")
        if df_master.empty:
            st.info("Belum ada data harga masuk.")
        else:
            # UI Filter
            c1, c2 = st.columns(2)
            avail_val = sorted(df_master['validity'].unique().tolist())
            avail_load = sorted(df_master['load_type'].unique().tolist())
            
            sel_val = c1.selectbox("Filter Periode", avail_val, key="es_val")
            sel_load = c2.selectbox("Filter Tipe Muatan", avail_load, key="es_load")
            
            # Filter Data
            df_view = df_master[(df_master['validity'] == sel_val) & (df_master['load_type'] == sel_load)].copy()
            
            if not df_view.empty:
                unique_origins = sorted(df_view['origin'].unique())
                
                for org in unique_origins:
                    with st.expander(f"📍 Origin: {org}", expanded=True):
                        sub_df = df_view[df_view['origin'] == org].copy() # Pakai .copy()
                        
                        # Ranking Logic
                        sub_df = sub_df.sort_values(by=['kota_tujuan', 'unit_type', 'price'])
                        sub_df['Ranking'] = sub_df.groupby(['kota_tujuan', 'unit_type']).cumcount() + 1
                        
                        # ▼▼▼ FILTER HANYA TOP 3 ▼▼▼
                        sub_df = sub_df[sub_df['Ranking'] <= 3]
                        # ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
                        
                        sub_df['price_fmt'] = sub_df['price'].apply(lambda x: f"Rp {int(x):,}".replace(",", "."))
                        
                        st.dataframe(
                            sub_df[['kota_tujuan', 'unit_type', 'Ranking', 'vendor_name', 'price_fmt', 'lead_time', 'top']],
                            use_container_width=True,
                            column_config={
                                "kota_tujuan": "Tujuan",
                                "unit_type": "Unit",
                                "price_fmt": "Harga",
                                "vendor_name": "Vendor"
                            },
                            hide_index=True
                        )
            else:
                st.warning("Data tidak ditemukan.")

# --- TAB 8: PRINT FILE (SK & SPK TERPISAH) ---
    with tabs[7]:
        st.subheader("🖨️ Print Dokumen")
        
        if df_master.empty:
            st.info("Data belum tersedia.")
        else:
            avail_val = sorted(df_master['validity'].unique().tolist())
            avail_load = sorted(df_master['load_type'].unique().tolist())

            # ==========================================
            # BAGIAN 1: SURAT KEPUTUSAN (SK)
            # ==========================================
            with st.container(border=True):
                st.markdown("### 1. Surat Keputusan (SK)")
                st.caption("Dokumen rekapitulasi pemenang tender (Top 3).")
                
                c1, c2 = st.columns(2)
                sk_val = c1.selectbox("Periode SK", avail_val, key="sk_val")
                sk_load = c2.selectbox("Muatan SK", avail_load, key="sk_load")
                
                # Filter Data SK
                df_sk = df_master[(df_master['validity'] == sk_val) & (df_master['load_type'] == sk_load)].copy()
                
                if not df_sk.empty:
                    avail_org = sorted(df_sk['origin'].unique())
                    sel_orgs = st.multiselect("Pilih Origin (SK):", avail_org, default=avail_org, key="sk_orgs")
                    
                    if sel_orgs:
                        df_final_sk = df_sk[df_sk['origin'].isin(sel_orgs)].copy()
                        df_final_sk = df_final_sk.sort_values(by=['origin', 'kota_tujuan', 'unit_type', 'price'])
                        df_final_sk['Ranking'] = df_final_sk.groupby(['origin', 'kota_tujuan', 'unit_type']).cumcount() + 1
                        
                        col_a, col_b = st.columns(2)
                        upl_sk = col_a.file_uploader("Upload Template SK", type="docx", key="upl_sk")
                        no_sk = col_b.text_input("Nomor Surat SK:", "001/SK/LOG/2026", key="no_sk")
                        
                        if st.button("📄 Generate File SK", type="primary"):
                            tpl_sk = "template_sk.docx"
                            if upl_sk: tpl_sk = upl_sk
                            elif not os.path.exists(tpl_sk): st.error("Template SK tidak ditemukan."); st.stop()
                            
                            try:
                                f_sk = create_docx_sk(tpl_sk, no_sk, sk_val, sk_load, df_final_sk)
                                fn_sk = f"SK_{sk_load}_{sk_val}.docx"
                                with open(f_sk, "rb") as f:
                                    st.download_button("⬇️ Download SK", f, file_name=fn_sk)
                                os.remove(f_sk)
                            except Exception as e: st.error(f"Gagal: {e}")
                    else: st.warning("Pilih minimal 1 origin.")
                else: st.warning("Data tidak ditemukan.")

            st.write("") # Jarak

# ==========================================
            # BAGIAN 2: SURAT PERINTAH KERJA (SPK)
            # ==========================================
            with st.container(border=True):
                st.markdown("### 2. Surat Perintah Kerja (SPK)")
                st.caption("Dokumen perintah kerja spesifik untuk satu vendor.")
                
                c3, c4 = st.columns(2)
                spk_val = c3.selectbox("Periode SPK", avail_val, key="spk_val")
                spk_load = c4.selectbox("Muatan SPK", avail_load, key="spk_load")
                
                # Filter Data SPK
                df_spk_raw = df_master[(df_master['validity'] == spk_val) & (df_master['load_type'] == spk_load)].copy()
                
                if not df_spk_raw.empty:
                    # Pilih Vendor
                    avail_vens = sorted(df_spk_raw['vendor_name'].unique().tolist())
                    sel_ven = st.selectbox("Pilih Vendor (SPK):", avail_vens, key="spk_ven")
                    
                    # Filter Data Vendor
                    df_final_spk = df_spk_raw[df_spk_raw['vendor_name'] == sel_ven].copy()
                    
                    if not df_final_spk.empty:
                        st.info(f"Vendor **{sel_ven}** memiliki **{len(df_final_spk)}** rute di periode ini.")
                        
                        df_final_spk = df_final_spk.sort_values(by=['kota_asal', 'kota_tujuan', 'unit_type', 'price'])
                        if 'Ranking' not in df_final_spk.columns: df_final_spk['Ranking'] = 1

                        col_c, col_d = st.columns(2)
                        upl_spk = col_c.file_uploader("Upload Template SPK", type="docx", key="upl_spk")
                        no_spk = col_d.text_input("Nomor Surat SPK:", f"001/SPK/{sel_ven[:3].upper()}/2026", key="no_spk")
                        
if st.button("📄 Generate File SPK", type="primary"):
                            tpl_spk = "template_spk.docx"
                            if upl_spk: tpl_spk = upl_spk
                            elif not os.path.exists(tpl_spk): st.error("Template SPK tidak ditemukan."); st.stop()
                            
                            # 1. Ambil Data PIC
                            pic = df_final_spk.iloc[0].get('contact_person', 'Pimpinan Perusahaan')
                            if pd.isna(pic) or pic == "-": pic = "Pimpinan Perusahaan"
                            
                            # 2. Ambil Password
                            try:
                                user_row = df_u[df_u['vendor_name'] == sel_ven]
                                if not user_row.empty:
                                    raw_pass = str(user_row.iloc[0]['password'])
                                    final_pass = raw_pass[-5:] if len(raw_pass) >= 5 else raw_pass
                                else: final_pass = "XXXXX"
                            except: final_pass = "XXXXX"

                            # 3. AMBIL ALAMAT GUDANG DARI SHEET 'Gudang'
                            alamat_gudang_text = "Alamat Gudang Belum Disetting"
                            current_origin = df_final_spk.iloc[0]['origin']
                            
                            try:
                                df_gudang = get_data("Gudang") # Load sheet Gudang
                                if not df_gudang.empty:
                                    # Cari baris yang origin-nya cocok (case insensitive)
                                    # Pastikan di sheet Gudang ada kolom 'origin' dan 'alamat'
                                    res_addr = df_gudang[df_gudang['origin'].astype(str).str.lower() == str(current_origin).lower()]
                                    if not res_addr.empty:
                                        alamat_gudang_text = res_addr.iloc[0]['alamat']
                                    else:
                                        st.warning(f"Origin '{current_origin}' tidak ditemukan di sheet Gudang.")
                                else:
                                    st.warning("Sheet 'Gudang' kosong/tidak ditemukan.")
                            except Exception as e:
                                st.warning(f"Gagal mengambil data Gudang: {e}")

                            # 4. MERGE DATA MULTIDROP
                            df_spk_merged = df_final_spk.copy()
                            # (Default value)
                            df_spk_merged['inner_city_price'] = 0
                            df_spk_merged['outer_city_price'] = 0
                            df_spk_merged['labor_cost'] = 0

                            if not df_md.empty:
                                try:
                                    md_dict = {}
                                    for _, rmd in df_md.iterrows():
                                        k = (str(rmd['vendor_email']).strip(), str(rmd['validity']).strip(), str(rmd['group_id']).strip())
                                        md_dict[k] = {
                                            'in': rmd.get('inner_city_price', 0),
                                            'out': rmd.get('outer_city_price', 0),
                                            'lab': rmd.get('labor_cost', 0)
                                        }
                                    
                                    def get_md_val(row, kind):
                                        key = (str(row['vendor_email']).strip(), str(row['validity']).strip(), str(row['group_id']).strip())
                                        res = md_dict.get(key, {'in':0, 'out':0, 'lab':0})
                                        return res[kind]

                                    df_spk_merged['inner_city_price'] = df_spk_merged.apply(lambda x: get_md_val(x, 'in'), axis=1)
                                    df_spk_merged['outer_city_price'] = df_spk_merged.apply(lambda x: get_md_val(x, 'out'), axis=1)
                                    df_spk_merged['labor_cost'] = df_spk_merged.apply(lambda x: get_md_val(x, 'lab'), axis=1)
                                except Exception as e: st.warning(f"Gagal merge multidrop: {e}")

                            # 5. Generate Docx
                            with st.spinner(f"Memproses SPK {sel_ven}..."):
                                try:
                                    # Pass alamat_gudang_text ke fungsi
                                    f_spk = create_docx_spk(tpl_spk, no_spk, spk_val, spk_load, sel_ven, pic, final_pass, alamat_gudang_text, df_spk_merged)
                                    
                                    # Custom Filename
                                    safe_val = str(spk_val).replace(" - ", "-").replace(" ", "_")
                                    safe_load = str(spk_load).replace(" ", "")
                                    safe_ven_file = str(sel_ven).replace(" ", "_").replace(".", "").replace(",", "")
                                    
                                    custom_filename = f"SPK_{safe_load}_{safe_val}_{safe_ven_file}.docx"
                                    
                                    with open(f_spk, "rb") as f:
                                        st.download_button("⬇️ Download SPK", f, file_name=custom_filename)
                                    os.remove(f_spk)
                                except Exception as e: st.error(f"Gagal generate: {e}")
                                    
                    else: st.warning("Vendor ini tidak memiliki data.")
                else: st.warning("Data tidak ditemukan.")
                
# ================= VENDOR DASHBOARD (FULL FIX) =================
def vendor_dashboard(email):
    step = st.session_state['vendor_step']
    
    # --- STEP 1: DASHBOARD / PROFIL ---
    if step == "dashboard":
        t1, t2 = st.tabs(["🛣️ Pilih Rute & Isi Harga", "📋 Isi Data Perusahaan"])
        
        # Tab 2: Profil
        with t2:
            df_p = get_data("Vendor_Profile")
            curr = {}
            if not df_p.empty:
                m = df_p[df_p['email']==email]
                if not m.empty: curr = m.iloc[-1].to_dict()
            
            with st.container(border=True):
                with st.form("prof"):
                    c1, c2 = st.columns(2)
                    with c1:
                        ad = st.text_area("Alamat Perusahaan", value=curr.get('address',''))
                        cp = st.text_input("Nama PIC", value=curr.get('contact_person',''))
                    with c2:
                        ph = st.text_input("No. Telepon", value=curr.get('phone',''))
                        top = st.selectbox("Term of Payment", ["7 hari","14 Hari", "30 Hari"])
                    
                    ppn = st.selectbox("PPN", ["11%", "1,1%","0%"])
                    pph = st.selectbox("PPh", ["Include", "Exclude"])
                    
                    if st.form_submit_button("Simpan Data", type="primary"):
                        save_data("Vendor_Profile", [[email, ad, cp, ph, top, ppn, pph, datetime.now().strftime("%Y-%m-%d")]])
                        st.success("Saved")
        
        # Tab 1: List Rute
        with t1:
            df_acc = get_data("Access_Rights")
            df_grps = get_data("Master_Groups")
            df_routes = get_data("Master_Routes")
            df_price = get_data("Price_Data")

            if df_acc.empty: 
                st.warning("Belum ada akses.")
                return
            
            my_access = df_acc[df_acc['vendor_email'] == email]
            if my_access.empty: 
                st.info("Anda belum diberikan akses ke project manapun.")
                return

            data_list = []
            for _, acc in my_access.iterrows():
                gid = acc['group_id']
                val = acc['validity']
                g_info = df_grps[df_grps['group_id'] == gid]
                if not g_info.empty:
                    row = g_info.iloc[0]
                    data_list.append({'validity': val, 'group_id': gid, 'origin': row.get('origin','-'), 'route_group': row.get('route_group','-'), 'load_type': row.get('load_type','-')})
            
            df_disp = pd.DataFrame(data_list)
            if df_disp.empty: 
                st.warning("Konfigurasi Group tidak ditemukan.")
                return

            avail_validity = sorted(df_disp['validity'].unique().tolist())
            sel_val = st.selectbox("Pilih Periode / Validity:", avail_validity)
            df_view = df_disp[df_disp['validity'] == sel_val]
            
            if df_view.empty: 
                st.info("Tidak ada rute.")
                return
            
            t_ftl, t_fcl = st.tabs(["🚛 FTL (Full Truck Load)", "🚢 FCL (Full Container Load)"])
            
            for t_code, t_ui in [('FTL', t_ftl), ('FCL', t_fcl)]:
                with t_ui:
                    df_sub = df_view[df_view['load_type'] == t_code]
                    if df_sub.empty: 
                        st.caption(f"Tidak ada akses {t_code}.")
                    else:
                        for org in sorted(df_sub['origin'].unique()):
                            with st.container(border=True):
                                st.markdown(f"#### 📍 {org}")
                                org_groups = df_sub[df_sub['origin'] == org]
                                c1, c2, c3, c4 = st.columns([3, 4, 2, 2])
                                c1.caption(""); c2.caption("Kota Tujuan"); c3.caption(""); c4.caption("Status Pengisian")
                                st.divider()
                                
                                for _, row in org_groups.iterrows():
                                    gid = row['group_id']
                                    grp_name = row['route_group']
                                    
                                    # Status Check
                                    r_data = df_routes[df_routes['group_id'] == gid] if not df_routes.empty else pd.DataFrame()
                                    status_ui = '<span class="status-pending">❌ Belum Ada Data</span>'
                                    is_locked_btn = False
                                    
                                    if not df_price.empty and not r_data.empty:
                                        sub_p = df_price[
                                            (df_price['vendor_email']==email) & 
                                            (df_price['validity']==sel_val) & 
                                            (df_price['route_id'].isin(r_data['route_id']))
                                        ]
                                        if not sub_p.empty:
                                            status_ui = '<span class="status-done">✅Sudah Terisi</span>'
                                            if "Locked" in sub_p['status'].values or "Waiting" in sub_p['status'].values:
                                                is_locked_btn = True
                                    
                                    c1, c2, c3, c4 = st.columns([3, 4, 2, 2])
                                    c1.write(f"**{grp_name}**")
                                    dests = r_data['kota_tujuan'].unique().tolist() if not r_data.empty else []
                                    if len(dests) > 5: preview_txt = f"{', '.join(dests[:5])}, +{len(dests)-5} kota lainnya"
                                    else: preview_txt = ", ".join(dests)
                                    c2.markdown(f"<span class='route-dest-list'>{preview_txt}</span>", unsafe_allow_html=True)

                                    if is_locked_btn:
                                        c3.button("🔒 Locked", key=f"btn_lk_{gid}", disabled=True)
                                    else:
                                        if c3.button("📌 Isi Harga", key=f"btn_{t_code}_{gid}", type="primary"):
                                            st.session_state.update({
                                                'sel_origin': org, 
                                                'sel_validity': sel_val, 
                                                'sel_load': t_code, 
                                                'vendor_step': 'input',
                                                'focused_group_id': gid
                                            })
                                            st.rerun()
                                            
                                    c4.markdown(status_ui, unsafe_allow_html=True)
                                    st.markdown("<hr>", unsafe_allow_html=True)

    # --- STEP 2: INPUT HARGA ---
    elif step == "input":
        if st.session_state.get('temp_success_msg'):
            st.success(st.session_state['temp_success_msg'])
            st.session_state['temp_success_msg'] = None

        if st.button("⬅️ Kembali ke Menu Utama", type="secondary"):
            st.session_state['vendor_step'] = "dashboard"; st.rerun()

        cur_org = st.session_state.get('sel_origin')
        cur_val = st.session_state.get('sel_validity')
        cur_load = st.session_state.get('sel_load')
        focused_gid = st.session_state.get('focused_group_id')

        st.markdown(f"### Input Penawaran Harga {cur_load}: {cur_org}")
        st.caption(f"Periode: {cur_val}")

        df_acc = get_data("Access_Rights"); df_grps = get_data("Master_Groups")
        my_acc = df_acc[(df_acc['vendor_email']==email) & (df_acc['validity']==cur_val)]
        
        target_gids = []
        grp_names = {}
        for gid in my_acc['group_id'].unique():
            r = df_grps[df_grps['group_id']==gid]
            if not r.empty:
                rr = r.iloc[0]
                if rr['origin']==cur_org and rr['load_type']==cur_load:
                    target_gids.append(gid); grp_names[gid]=rr['route_group']
        
        if not target_gids: 
            st.error("Data error.")
            return

        target_gids = sorted(target_gids)
        if focused_gid and focused_gid in target_gids:
            target_gids.remove(focused_gid)
            target_gids.insert(0, focused_gid)
        
        tabs = st.tabs([grp_names[g] for g in target_gids])
        
        df_r = get_data("Master_Routes"); df_u = get_data("Master_Units"); df_p = get_data("Price_Data"); df_m = get_data("Multidrop_Data")

        for i, gid in enumerate(target_gids):
            with tabs[i]:
                g_name = grp_names[gid]
                my_r = df_r[df_r['group_id']==gid]
                my_u = df_u[df_u['group_id']==gid]
                u_types = my_u['unit_type'].unique().tolist()
                
                if my_r.empty or not u_types: 
                    st.warning("Data belum lengkap.")
                    continue

                ex_price = {}; ex_spec = {}
                is_lock = False
                if not df_p.empty:
                    my_p = df_p[(df_p['vendor_email']==email) & (df_p['validity']==cur_val) & (df_p['route_id'].isin(my_r['route_id']))]
                    if not my_p.empty:
                        if "Locked" in my_p['status'].values or "Waiting" in my_p['status'].values: is_lock = True
                        for _, row in my_p.iterrows():
                            ex_price[(row['route_id'], row['unit_type'])] = row['price']
                            ex_spec[row['unit_type']] = {'w': row.get('weight_capacity'), 'c': row.get('cubic_capacity')}

                with st.form(key=f"f_{gid}"):
                    # 1. SPEC
                    with st.container(border=True):
                        st.markdown(f"#### 🛻 Spesifikasi Armada")
                        sp_data = []
                        for u in u_types:
                            sp_data.append({
                                "Jenis Unit": u,
                                "Kapasitas Berat Bersih (Kg)": clean_numeric(ex_spec.get(u,{}).get('w')),
                                "Kapasitas Kubikasi Dalam (CBM)": clean_numeric(ex_spec.get(u,{}).get('c'))
                            })
                        
                        df_sp = pd.DataFrame(sp_data)
                        cf_sp = {
                            "Jenis Unit": st.column_config.TextColumn(disabled=True),
                            "Kapasitas Berat Bersih (Kg)": st.column_config.NumberColumn(min_value=0, format="%d", step=1),
                            "Kapasitas Kubikasi Dalam (CBM)": st.column_config.NumberColumn(min_value=0, format="%.2f", step=0.1)
                        }
                        ed_sp = st.data_editor(df_sp, hide_index=True, use_container_width=True, disabled=is_lock, column_config=cf_sp, height=(len(df_sp)+1)*36+30)

                    # 2. PRICE
                    with st.container(border=True):
                        st.markdown(f"#### 💰 Penawaran Harga")
                        p_data = []
                        for _, row in my_r.iterrows():
                            rid = row['route_id']
                            rd = {
                                "Route ID": rid, "Kota Asal": row['kota_asal'], "Kota Tujuan": row['kota_tujuan'],
                                "Lead Time (Hari)": 0
                            }
                            for u in u_types: rd[f"Harga {u}"] = ex_price.get((rid, u), 0)
                            rd["Keterangan"] = row.get('keterangan','-')
                            p_data.append(rd)
                        
                        df_pr = pd.DataFrame(p_data)
                        for c in [f"Harga {u}" for u in u_types]: 
                            if c in df_pr.columns: df_pr[c] = pd.to_numeric(df_pr[c], errors='coerce').fillna(0)
                        
                        cf_pr = {
                            "Route ID": None,
                            "Kota Asal": st.column_config.TextColumn(disabled=True),
                            "Kota Tujuan": st.column_config.TextColumn(disabled=True),
                            "Keterangan": st.column_config.TextColumn(disabled=True),
                            "Lead Time (Hari)": st.column_config.NumberColumn(min_value=0, step=1)
                        }
                        for u in u_types:
                            cf_pr[f"Harga {u}"] = st.column_config.NumberColumn(min_value=0, step=1000, format="Rp %d")

                        ed_pr = st.data_editor(df_pr, hide_index=True, use_container_width=True, disabled=is_lock, column_config=cf_pr, height=(len(df_pr)+1)*36+30)

                    # 3. MULTIDROP
                    with st.container(border=True):
                        st.markdown("#### 📦 Biaya Multidrop & Buruh")
                        ic, oc, lc = 0, 0, 0
                        if not df_m.empty:
                            mr = df_m[(df_m['vendor_email']==email) & (df_m['validity']==cur_val) & (df_m['group_id']==gid)]
                            if not mr.empty:
                                ic = clean_numeric(mr.iloc[0].get('inner_city_price')) or 0
                                oc = clean_numeric(mr.iloc[0].get('outer_city_price')) or 0
                                lc = clean_numeric(mr.iloc[0].get('labor_cost')) or 0
                        
                        df_md = pd.DataFrame([{"Multidrop Dalam Kota": ic, "Multidrop Luar Kota": oc, "Biaya Buruh": lc}])
                        cf_md = {
                            "Multidrop Dalam Kota": st.column_config.NumberColumn(min_value=0, step=1000, format="Rp %d"),
                            "Multidrop Luar Kota": st.column_config.NumberColumn(min_value=0, step=1000, format="Rp %d"),
                            "Biaya Buruh": st.column_config.NumberColumn(min_value=0, step=1000, format="Rp %d")
                        }
                        ed_md = st.data_editor(df_md, hide_index=True, use_container_width=True, disabled=is_lock, column_config=cf_md)

# SAVE
                    st.write("")
                    if st.form_submit_button(f"Simpan Data {cur_load} {g_name}", type="primary") and not is_lock:
                        c_spec = {r['Jenis Unit']: {'w': r['Kapasitas Berat Bersih (Kg)'], 'c': r['Kapasitas Kubikasi Dalam (CBM)']} for _, r in ed_sp.iterrows()}
                        
                        f_data = []
                        ts = (datetime.utcnow() + timedelta(hours=7)).strftime("%Y-%m-%d %H:%M:%S")
                        
                        for _, r in ed_pr.iterrows():
                            rid = str(r['Route ID']); lt = int(r['Lead Time (Hari)'])
                            for u in u_types:
                                pr = int(r[f"Harga {u}"])
                                w = str(c_spec.get(u,{}).get('w','')); c = str(c_spec.get(u,{}).get('c',''))
                                tid = f"{email}_{cur_val}_{rid}_{u}".replace(" ","")
                                f_data.append([tid, email, "Open", cur_val, rid, u, lt, pr, w, c, "", ts])
                        
                        mi = int(ed_md.iloc[0]["Multidrop Dalam Kota"])
                        mo = int(ed_md.iloc[0]["Multidrop Luar Kota"])
                        ml = int(ed_md.iloc[0]["Biaya Buruh"])
                        mid = f"M_{email}_{gid}_{cur_val}"

                        save_data("Price_Data", f_data)
                        save_data("Multidrop_Data", [[mid, email, cur_val, gid, mi, mo, ml, ts]])
                        
                        st.session_state['temp_success_msg'] = f"Sukses! Data untuk {g_name} tersimpan."
                        st.cache_data.clear()
                        st.rerun()


if __name__ == "__main__":
    main()












