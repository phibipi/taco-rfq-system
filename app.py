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
import urllib.parse
import io
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

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
            font-size: 24px !important;        /* Huruf Lebih Besar */
            padding: 0.8rem 2rem !important;  /* Ukuran Tombol Lebih Besar */
            box-shadow: 0 4px 8px rgba(249, 115, 22, 0.3);
            width: 100% !important;            /* Tombol Memanjang Penuh */
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
            color: #111827 !important;             /* Paksa Teks Hitam */
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

@st.cache_data(ttl=60, show_spinner=False)
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

# --- FUNGSI KONEKSI & UPLOAD GOOGLE DRIVE ---
@st.cache_resource
def get_drive_service():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = st.secrets["gcp_service_account"]
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        elif os.path.exists("kunci_rahasia.json"):
            creds = ServiceAccountCredentials.from_json_keyfile_name("kunci_rahasia.json", scope)
        else:
            return None
        service = build('drive', 'v3', credentials=creds)
        return service
    except Exception as e:
        return None

def upload_to_drive(file_buffer, filename, mimetype, folder_id):
    service = get_drive_service()
    if not service: return None
    
    file_metadata = {
        'name': filename,
        'parents': [folder_id]
    }
    media = MediaIoBaseUpload(io.BytesIO(file_buffer.getvalue()), mimetype=mimetype, resumable=True)
    
    try:
        # Upload dan minta Google mengembalikan link webViewLink
        file = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink',supportsAllDrives=True).execute()
        return file.get('webViewLink')
    except Exception as e:
        st.error(f"Error API Drive: {e}")
        return None

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
    
# --- FUNGSI KIRIM EMAIL (UPDATE: ADA INFO ROUND) ---
def send_invitation_email(to_email, vendor_name, load_type, validity, origins, password, round_num="1"):
    # Cek Config
    if "email_config" not in st.secrets:
        st.warning("Konfigurasi email belum disetting di Secrets. Email tidak terkirim.")
        return False

    sender_email = st.secrets["email_config"]["sender_email"]
    sender_password = st.secrets["email_config"]["sender_password"]

    cc_list = ["firli.mandaras@taco.co.id", "budhi.yuono@taco.co.id"]
    cc_string = ", ".join(cc_list)
    
    # Hitung Due Date
    today = datetime.now()
    due_date = today + timedelta(days=6)
    
    months_id = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}
    due_date_str = f"{due_date.day} {months_id[due_date.month]} {due_date.year}"

    # Update Subject agar terlihat Tahap keberapa
    subject = f"Undangan Tender {load_type} - {validity} (Tahap Penawaran {round_num}) - TACO Group"
    origins_str = ", ".join(origins)
    
    body = f"""
    <html>
    <body>
        <h3>Dear {vendor_name},</h3>
        <p>Anda telah diundang untuk berpartisipasi dalam Tender Transport <b>TACO Group</b>.</p>
        <p><b>Detail Tender:</b></p>
        <ul>
            <li><b>Periode:</b> {validity}</li>
            <li><b>Tahap Penawaran:</b> {round_num}</li>
            <li><b>Tipe Armada:</b> {load_type}</li>
            <li><b>Area/Origin:</b> {origins_str}</li>
            <li style="color: #d9534f;"><b>Batas Akhir Pengisian: {due_date_str}</b></li>
        </ul>
        <p>Silakan login ke sistem kami untuk memasukkan penawaran harga:</p>
        <p>
            <b>Link App:</b> <a href="https://taco-transport.streamlit.app/">http://bit.ly/TACOtender</a><br>
            <b>Email Login:</b> {to_email}<br>
            <b>Password:</b> {password}<br>
            <b>Tutorial:</b> <a href="https://drive.google.com/file/d/1M5QnSGibg2s9LiQXlNaiCLWxaA7jebJM/view">https://bit.ly/TutorialRFQTACO</a><br>
        </p>
        <p>Terima Kasih,<br>Procurement Team TACO</p>
    </body>
    </html>
    """

    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = to_email
        msg['Cc'] = cc_string
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

# --- FUNGSI KIRIM EMAIL REMINDER ---
def send_reminder_email(to_email, vendor_name, load_type, validity, round_num, pending_groups, password):
    if "email_config" not in st.secrets: return False
    sender_email = st.secrets["email_config"]["sender_email"]
    sender_password = st.secrets["email_config"]["sender_password"]
    
    cc_list = ["firli.mandaras@taco.co.id", "budhi.yuono@taco.co.id"]
    cc_string = ", ".join(cc_list)
    
    subject = f"REMINDER: Pengisian Penawaran Harga Tender {load_type} - {validity} (Tahap {round_num})"
    
    # PENGAMAN: Ubah semua item ke string sebelum di-join agar tidak TypeError
    pending_groups_str = ", ".join([str(g) for g in pending_groups])
    
    body = f"""
    <html>
    <body>
        <h3 style="color: #d9534f;">⚠️ Reminder Pengisian Penawaran Harga</h3>
        <p>Dear <b>{vendor_name}</b>,</p>
        <p>Melalui email ini, kami ingin mengingatkan bahwa Anda <b>belum menyelesaikan</b> pengisian penawaran harga pada sistem kami untuk detail berikut:</p>
        <ul>
            <li><b>Periode:</b> {validity}</li>
            <li><b>Tahap Penawaran:</b> {round_num}</li>
            <li><b>Tipe Armada:</b> {load_type}</li>
            <li><b style="color: #d9534f;">Origin yang belum diisi: {pending_groups_str}</b></li>
        </ul>
        <p>Mohon segera login dan melengkapi form harga pada area (Origin) yang belum terselesaikan.</p>
        <p>
            <b>Link App:</b> <a href="https://taco-transport.streamlit.app/">http://bit.ly/TACOtender</a><br>
            <b>Email Login:</b> {to_email}<br>
            <b>Password:</b> {password}<br>
            <b>Tutorial:</b> <a href="https://drive.google.com/file/d/1M5QnSGibg2s9LiQXlNaiCLWxaA7jebJM/view">https://bit.ly/TutorialRFQTACO</a><br>
        </p>
        <p>Terima Kasih,<br>Procurement Team TACO</p>
    </body>
    </html>
    """
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = to_email
        msg['Cc'] = cc_string 
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html'))
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(sender_email, sender_password)
        server.send_message(msg) 
        server.quit()
        return True
    except Exception as e:
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
        run.font.name = 'Calibri'

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
    
# --- FUNGSI GENERATE SPH VENDOR ---
def create_docx_sph(template_file, vendor_name, vendor_address, validity, load_type, round_num, df_data):
    doc = DocxTemplate(template_file)
    
    # Format Tanggal
    bulan_indo = {1:'Januari', 2:'Februari', 3:'Maret', 4:'April', 5:'Mei', 6:'Juni', 7:'Juli', 8:'Agustus', 9:'September', 10:'Oktober', 11:'November', 12:'Desember'}
    today = datetime.now()
    tgl_sph = f"{today.day} {bulan_indo[today.month]} {today.year}"

    def fmt_rp(x):
        try: return f"Rp {int(x):,}".replace(",", ".")
        except: return "Rp 0"

    def set_col_widths(table, widths):
        table.autofit = False 
        table.allow_autofit = False
        for row in table.rows:
            for idx, width in enumerate(widths):
                if idx < len(row.cells): row.cells[idx].width = width

    def format_cell(cell, text, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, size=8):
        cell.text = str(text)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = cell.paragraphs[0]
        p.alignment = align
        p.paragraph_format.space_after = Pt(0)
        run = p.runs[0] if p.runs else p.add_run()
        run.font.size = Pt(size)
        run.font.name = 'Calibri'
        run.font.bold = bold

    def set_repeat_table_header(row):
        tr = row._tr
        trPr = tr.get_or_add_trPr()
        tblHeader = OxmlElement('w:tblHeader')
        tblHeader.set(qn('w:val'), "true")
        trPr.append(tblHeader)

    sd = doc.new_subdoc()
    
    # --- TABEL 1: HARGA UTAMA ---
    headers = ['No', 'Origin', 'Tujuan', 'Unit', 'Lead Time', 'Harga Penawaran', 'Keterangan']
    table = sd.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_repeat_table_header(table.rows[0])
    
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        shading_elm = parse_xml(r'<w:shd {} w:fill="ED7D31"/>'.format(nsdecls('w')))
        hdr_cells[i]._tc.get_or_add_tcPr().append(shading_elm)
        format_cell(hdr_cells[i], h, bold=True, size=8)

    for idx, row in df_data.iterrows():
        row_cells = table.add_row().cells
        
        lt_raw = str(row.get('lead_time', '-'))
        lt_fmt = f"{lt_raw} Hari" if (lt_raw.isdigit() or lt_raw.replace('.','',1).isdigit()) and lt_raw not in ["-","0",""] else "-"
        
        ket = str(row.get('keterangan', '-'))
        if ket.lower() in ["nan", "none", ""]: ket = "-"

        vals = [
            str(idx + 1),
            str(row.get('origin', '-')),
            str(row.get('kota_tujuan', '-')),
            str(row.get('unit_type', '-')),
            lt_fmt,
            fmt_rp(row.get('price', 0)),
            ket
        ]
        for i, v in enumerate(vals):
            align = WD_ALIGN_PARAGRAPH.LEFT if i in [1, 2, 6] else WD_ALIGN_PARAGRAPH.CENTER
            if i == 5: align = WD_ALIGN_PARAGRAPH.RIGHT
            format_cell(row_cells[i], v, align, size=7.5)

    # Lebar Kolom Tabel 1
    col_widths = [Cm(0.8), Cm(2.0), Cm(2.5), Cm(2.0), Cm(1.5), Cm(2.5), Cm(3.0)]
    set_col_widths(table, col_widths)

    sd.add_paragraph("") # Spasi Antar Tabel

    # --- TABEL 2: MULTIDROP & CATATAN ---
    p = sd.add_paragraph("Tabel Biaya Multidrop & Keterangan Tambahan:")
    p.runs[0].bold = True
    p.runs[0].font.name = 'Calibri'
    p.runs[0].font.size = Pt(9)
    
    headers_md = ['Origin', 'MD Dalam', 'MD Luar', 'Biaya Buruh', 'Catatan Tambahan Vendor']
    table_md = sd.add_table(rows=1, cols=len(headers_md))
    table_md.style = 'Table Grid'
    table_md.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_repeat_table_header(table_md.rows[0])
    
    hdr_cells_md = table_md.rows[0].cells
    for i, h in enumerate(headers_md):
        shading_elm = parse_xml(r'<w:shd {} w:fill="ED7D31"/>'.format(nsdecls('w')))
        hdr_cells_md[i]._tc.get_or_add_tcPr().append(shading_elm)
        format_cell(hdr_cells_md[i], h, bold=True, size=8)

    # Filter Multidrop agar unik (muncul per Origin saja)
    df_md_uniq = df_data[['origin', 'inner_city_price', 'outer_city_price', 'labor_cost', 'catatan_tambahan']].drop_duplicates(subset=['origin'])
    
    for _, row in df_md_uniq.iterrows():
        row_cells = table_md.add_row().cells
        ctt = str(row.get('catatan_tambahan', '-'))
        if ctt.lower() in ["nan", "none", ""]: ctt = "-"

        vals_md = [
            str(row.get('origin', '-')),
            fmt_rp(row.get('inner_city_price', 0)),
            fmt_rp(row.get('outer_city_price', 0)),
            fmt_rp(row.get('labor_cost', 0)),
            ctt
        ]
        for i, v in enumerate(vals_md):
            align = WD_ALIGN_PARAGRAPH.LEFT if i in [0, 4] else WD_ALIGN_PARAGRAPH.RIGHT
            format_cell(row_cells[i], v, align, size=7.5)
            
    # Lebar Kolom Tabel 2
    col_widths_md = [Cm(2.5), Cm(2.2), Cm(2.2), Cm(2.2), Cm(5.2)]
    set_col_widths(table_md, col_widths_md)

    context = {
        'tanggal': tgl_sph,
        'vendor_name': vendor_name,
        'vendor_address': vendor_address,
        'validity': validity,
        'load_type': load_type,
        'round_num': round_num,
        'tabel_sph': sd
    }
    doc.render(context)
    
    safe_ven = "".join(x for x in vendor_name if x.isalnum())
    fn = f"SPH_{safe_ven}_Tahap{round_num}_{int(time.time())}.docx"
    doc.save(fn)
    return fn
    
# --- FUNGSI GENERATE SPK (UPDATE: MULTI ORIGIN & LEBAR KOLOM) ---
def create_docx_spk(template_file, no_spk, validity, load_type, vendor_name, pic_vendor, vendor_pass, origin_name_str, alamat_gudang, df_data):
    doc = DocxTemplate(template_file)
    
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
        run.font.name = 'Calibri'
        run.font.bold = bold

    # --- TABEL DATA ---
    sd = doc.new_subdoc()
    headers = ['Asal', 'Tujuan', 'Unit', 'Rank', 'Biaya/Unit', 'Multidrop dalam Kota', 'Multidrop luar Kota', 'Biaya Buruh', 'Lead Time']
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

    # --- UPDATE LEBAR KOLOM (Tujuan dikecilkan, Unit dilebarkan) ---
    # Lama: Tujuan(3.0), Unit(1.8)
    # Baru: Tujuan(2.5), Unit(2.3)
    # Asal(2.0), Tujuan(2.5), Unit(2.3), Rank(0.8), Biaya(2.2), MD_In(2.0), MD_Out(2.0), Buruh(1.5), LT(1.5)
    col_widths = [Cm(2.0), Cm(2.5), Cm(2.3), Cm(0.8), Cm(2.2), Cm(2.0), Cm(2.0), Cm(1.5), Cm(1.5)]
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
        'origin_name': origin_name_str, 
        'alamat_gudang': alamat_gudang,
        'tabel_harga_vendor': sd
    }
    doc.render(context)
    
    safe_ven = "".join(x for x in vendor_name if x.isalnum())
    fn = f"temp_spk_{safe_ven}_{int(time.time())}.docx"
    doc.save(fn)
    return fn

# --- HELPER: HITUNG TARGET PRICE (FIX: TYPE ERROR) ---
def get_target_price(df_all, route_id, unit_type, cur_validity):
    # SAFETY: Pastikan kolom price dibaca sebagai angka
    # Kita buat copy agar tidak merusak dataframe asli
    df_safe = df_all.copy()
    if 'price' in df_safe.columns:
        df_safe['price'] = pd.to_numeric(df_safe['price'], errors='coerce').fillna(0)
    else:
        return 0

    # 1. Ambil Harga Terendah Periode SAAT INI
    df_curr = df_safe[
        (df_safe['validity'] == cur_validity) & 
        (df_safe['route_id'] == str(route_id)) & 
        (df_safe['unit_type'] == unit_type)
    ]
    
    if df_curr.empty: return 0 
    
    min_curr = df_curr['price'].min()
    if min_curr == 0: return 0

    # 2. Tentukan Periode "SEBELUMNYA"
    target_price = 0
    df_prev = pd.DataFrame()
    
    try:
        parts = cur_validity.split(" ") 
        cur_year_str = parts[-1]
        cur_year_int = int(cur_year_str)
        
        is_semester_2 = "juli" in cur_validity.lower() or "july" in cur_validity.lower()
        
        if is_semester_2:
            # SKENARIO A: Periode Juli-Desember -> Cari Jan-Jun tahun sama
            df_prev = df_safe[
                (df_safe['validity'].str.contains(cur_year_str, na=False)) & 
                (df_safe['validity'].str.contains("Jan", case=False, na=False)) & 
                (df_safe['route_id'] == str(route_id)) & 
                (df_safe['unit_type'] == unit_type)
            ]
        else:
            # SKENARIO B: Periode Januari-Juni -> Cari Tahun Sebelumnya
            prev_year_str = str(cur_year_int - 1)
            df_prev = df_safe[
                (df_safe['validity'].str.contains(prev_year_str, na=False)) & 
                (df_safe['route_id'] == str(route_id)) & 
                (df_safe['unit_type'] == unit_type)
            ]

        # 3. Bandingkan Harga
        if not df_prev.empty:
            min_prev = df_prev['price'].min()
            
            if min_prev > 0:
                if min_prev < min_curr:
                    target_price = min_prev * 0.95 
                else:
                    target_price = min_curr * 0.92 
            else:
                target_price = min_curr * 0.92
        else:
            target_price = min_curr * 0.92
            
    except:
        target_price = min_curr * 0.92
        
    return int(target_price)

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
    if 'admin_step' not in st.session_state: st.session_state['admin_step'] = "home" # <--- TAMBAHAN BARU
    
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
        
        df_master = m4.copy() 
        
        # --- TAMBAHAN FILTER HARGA 0 ---
        df_master = df_master[df_master['price'] > 0]
    
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
    step = st.session_state.get('admin_step', 'home')
    
    # --- HALAMAN UTAMA (HOME) ---
    if step == 'home':
        st.markdown("## 🎛️ Admin Portal")
        st.markdown("<br>", unsafe_allow_html=True)
        
        c1, c2 = st.columns(2)
        with c1:
            with st.container(border=True):
                st.markdown("### 📂 Akses & Master Data")
                st.write("Manage Origin, Rute, Unit, User Vendor, dan Akses Pengisian Harga.")
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("Masuk ke Master Data ➡️", type="primary", use_container_width=True):
                    st.session_state['admin_step'] = 'master'
                    st.rerun()
                    
        with c2:
            with st.container(border=True):
                st.markdown("### 📊 Monitoring & Summary")
                st.write("Monitor progres submit vendor, Lock/Unlock data, Ranking Harga, dan Cetak SK/SPK.")
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("Masuk ke Monitoring ➡️", type="primary", use_container_width=True):
                    st.session_state['admin_step'] = 'monitoring'
                    st.rerun()   
                    
    # --- HALAMAN 1: MASTER DATA ---
    elif step == 'master':
        if st.button("⬅️ Kembali ke Menu Utama"):
            st.session_state['admin_step'] = 'home'
            st.rerun() 
        st.markdown("### 📂 Akses & Master Data")
        tabs = st.tabs(["📍Master Groups", "🛣️Master Routes", "🚛Master Units", "👥Users", "🔑Access Rights"])

    
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
                            if is_exist: 
                                st.error(f"Unit '{ut}' sudah ada di Group {gid}.")
                            else:
                                ws.append_rows([[gid, ut]])
                                get_data.clear() 
                                st.success("Unit tersimpan."); time.sleep(0.5); st.rerun()
            st.dataframe(get_data("Master_Units"), use_container_width=True)

        # --- TAB 4: USERS ---
        with tabs[3]:
            with st.form("add_usr"):
                c1, c2, c3 = st.columns(3)
                em = c1.text_input("Email")
                pw = c2.text_input("Pass")
                nm = c3.text_input("PT Name")
                if st.form_submit_button("Add User", type="primary"):
                    save_data("Users", [[em, pw, "vendor", nm]])
                    st.success("Saved")
            
            # Baris ini sekarang sejajar persis dengan 'with st.form'
            st.dataframe(get_data("Users"), use_container_width=True)

# --- TAB 5: ACCESS RIGHTS (UPDATE: FORMAT VENDOR DROPDOWN) ---
        with tabs[4]:
            st.write("Grant Access (Batch per Origin)")
            df_u = get_data("Users"); df_g = get_data("Master_Groups"); df_rights = get_data("Access_Rights")
        
            if not df_u.empty and not df_g.empty:
                # --- HELPER FORMAT DROPDOWN ---
                df_vendors = df_u[df_u['role']=='vendor']
                vendor_emails = df_vendors['email'].unique()
                
                def fmt_vendor(eml):
                    nm = df_vendors[df_vendors['email'] == eml]['vendor_name']
                    if not nm.empty: return f"{nm.iloc[0]} - {eml}"
                    return eml
                # ------------------------------

                c1, c2, c3 = st.columns(3)
                ven = c1.selectbox("Pilih Vendor", vendor_emails, format_func=fmt_vendor) # <-- UPDATE DI SINI
                sel_lt = c2.selectbox("Pilih Load Type", ["FTL", "FCL"])
                sel_round = c3.selectbox("Tahap Penawaran", ["1", "2", "3"]) 
            
                unique_origins = []
                subset_g = df_g[df_g['load_type'] == sel_lt]
                unique_origins = sorted(subset_g['origin'].unique().tolist())
                
                if unique_origins:
                    existing_gids = set()
                    if not df_rights.empty: 
                        if 'round' in df_rights.columns:
                            sub_r = df_rights[(df_rights['vendor_email'] == ven) & (df_rights['round'] == sel_round)]
                        else:
                            sub_r = df_rights[df_rights['vendor_email'] == ven]
                        existing_gids = set(sub_r['group_id'].tolist())

                    with st.form("acc_origin_batch"):
                        cols = st.columns(3)
                        selected_origins = []
                        for idx, org in enumerate(unique_origins):
                            org_groups = df_g[(df_g['origin'] == org) & (df_g['load_type'] == sel_lt)]
                            org_gids = set(org_groups['group_id'].tolist())
                            is_checked = not org_gids.isdisjoint(existing_gids)
                            if cols[idx % 3].checkbox(org, value=is_checked, key=f"chk_{org}"): 
                                selected_origins.append(org)
                        
                        st.divider()
                        c_val1, c_val2 = st.columns([2, 1])
                        
                        opt_validity = ["Januari - Juni", "Juli - Desember"] if sel_lt == "FCL" else ["Januari - Desember"]
                        val_period = c_val1.selectbox("Periode", opt_validity)
                        val_year = c_val2.text_input("Tahun", value=str(datetime.now().year))
                        
                        if st.form_submit_button("Grant Access", type="primary"):
                            val = f"{val_period} {val_year}"
                            if selected_origins:
                                target_groups = df_g[(df_g['load_type'] == sel_lt) & (df_g['origin'].isin(selected_origins))]
                                target_gids = target_groups['group_id'].unique().tolist()
                                
                                sh = connect_to_gsheet()
                                if sh:
                                    ws = sh.worksheet("Access_Rights")
                                    existing_data = ws.get_all_values()
                                    existing_keys = set()
                                    if len(existing_data) > 1:
                                        header = existing_data[0]
                                        has_round_col = 'round' in [h.lower() for h in header]
                                        for row in existing_data[1:]:
                                            if len(row) >= 3:
                                                base_key = f"{row[0]}_{row[1]}_{row[2]}".lower()
                                                if has_round_col and len(row) >= 5: base_key += f"_{row[4]}".lower()
                                                elif not has_round_col: base_key += "_1"
                                                existing_keys.add(base_key)
                                    
                                    new_rows = []
                                    skipped_count = 0
                                    for gid in target_gids:
                                        key_check = f"{ven}_{val}_{gid}_{sel_round}".lower()
                                        if key_check not in existing_keys:
                                            new_rows.append([ven, val, gid, "Active", sel_round])
                                        else: skipped_count += 1
                                    
                                    if new_rows:
                                        ws.append_rows(new_rows)
                                        get_data.clear()
                                        try:
                                            user_row = df_u[df_u['email'] == ven].iloc[0]
                                            with st.spinner("Mengirim email..."):
                                                send_invitation_email(ven, user_row['vendor_name'], sel_lt, val, selected_origins, user_row['password'], sel_round)
                                            st.success(f"Sukses! Akses Tahap {sel_round} diberikan.")
                                        except Exception as e: st.warning(f"Akses OK, Email Gagal: {e}")
                                        time.sleep(1); st.rerun()
                                    else: st.warning(f"Data sudah ada ({skipped_count} skip).")
                            else: st.warning("Pilih minimal 1 origin.")
                    
                    st.markdown("---")
                    # FITUR RESET AKSES
                    with st.expander("🗑️ Area Berbahaya: Reset/Hapus Akses Vendor", expanded=False):
                        c_del1, c_del2 = st.columns(2)
                        del_ven = c_del1.selectbox("Pilih Vendor (Hapus)", vendor_emails, format_func=fmt_vendor, key="del_ven") # <-- UPDATE DI SINI JUGA
                        del_lt = c_del2.selectbox("Pilih Tipe Muatan (Hapus)", ["FTL", "FCL"], key="del_lt")
                        if st.button("⚠️ Hapus Semua Akses Vendor Ini", type="primary"):
                            target_groups_del = df_g[df_g['load_type'] == del_lt]
                            gids_to_remove = set(target_groups_del['group_id'].tolist())
                            sh = connect_to_gsheet()
                            if sh:
                                try:
                                    ws = sh.worksheet("Access_Rights")
                                    all_rows = ws.get_all_values()
                                    if len(all_rows) > 1:
                                        header = all_rows[0]; data_rows = all_rows[1:]
                                        new_data_rows, deleted_count = [], 0
                                        for row in data_rows:
                                            if row[0] == del_ven and row[2] in gids_to_remove: deleted_count += 1
                                            else: new_data_rows.append(row)
                                        if deleted_count > 0:
                                            ws.clear(); ws.append_rows([header] + new_data_rows); get_data.clear()
                                            st.success(f"Dihapus {deleted_count} akses."); time.sleep(1); st.rerun()
                                        else: st.info("Tidak ada data dihapus.")
                                except: st.error("Error Google API")
            
            st.dataframe(get_data("Access_Rights"), use_container_width=True)

    # --- HALAMAN 2: MONITORING & SUMMARY ---
    elif step == 'monitoring':
        if st.button("⬅️ Kembali ke Menu Utama"):
            st.session_state['admin_step'] = 'home'
            st.rerun()
    
        st.markdown("### 📊 Monitoring & Summary")        

        # --- LOAD DATA SEKALI UNTUK SEMUA TAB ANALISA (OPTIMASI) ---

        df_p = get_data("Price_Data")
        df_r = get_data("Master_Routes")
        df_g = get_data("Master_Groups")
        df_u = get_data("Users")
        df_prof = get_data("Vendor_Profile")
        df_md = get_data("Multidrop_Data")
        df_acc = get_data("Access_Rights")
        df_units = get_data("Master_Units")
    
        # BIG MERGE MASTER (Untuk Tab 2, 3, 4)
        df_master = pd.DataFrame()
        if not df_p.empty and not df_g.empty:
            df_p['route_id'] = df_p['route_id'].astype(str).str.strip()
            df_r['route_id'] = df_r['route_id'].astype(str).str.strip()
            df_g['group_id'] = df_g['group_id'].astype(str).str.strip()
            m1 = pd.merge(df_p, df_r, on='route_id', how='left')
            m2 = pd.merge(m1, df_g, on='group_id', how='left')
            m3 = pd.merge(m2, df_u[['email', 'vendor_name']], left_on='vendor_email', right_on='email', how='left')
            m3['vendor_name'] = m3['vendor_name'].fillna(m3['vendor_email'])
            if not df_prof.empty:
                df_prof_clean = df_prof.sort_values('updated_at', ascending=False).drop_duplicates('email')
                m4 = pd.merge(m3, df_prof_clean[['email', 'top', 'ppn', 'pph', 'address', 'contact_person', 'phone']], left_on='vendor_email', right_on='email', how='left')
                for c in ['top', 'ppn', 'pph', 'address', 'contact_person', 'phone']:
                    if c in m4.columns: m4[c] = m4[c].fillna("-")
            else:
                m4 = m3
                for c in ['top', 'address', 'contact_person', 'phone']: m4[c] = "-"
            m4['price'] = pd.to_numeric(m4['price'], errors='coerce').fillna(0)
            df_master = m4

        tabs = st.tabs(["⏳ Submit Monitor", "✅ Lock Data", "📊 Summary", "🖨️ Print Dokumen", "📥 SPH Uploads"])
        
# --- TAB 1: SUBMIT MONITOR (UPDATE: STATS & SEARCH BAR) ---
        with tabs[0]:
                        
            if not df_acc.empty and not df_g.empty:
                # Merge Access dengan Group untuk tahu origin, route_group & load type
                acc_merge = pd.merge(df_acc, df_g[['group_id', 'origin', 'route_group', 'load_type']], on='group_id', how='left')
                if 'round' not in acc_merge.columns: acc_merge['round'] = '1'
                
                c1, c2, c3 = st.columns(3)
                sel_sm_lt = c1.selectbox("Filter Tipe Muatan", acc_merge['load_type'].dropna().unique().tolist())
                sel_sm_val = c2.selectbox("Filter Periode", acc_merge['validity'].dropna().unique().tolist())
                sel_sm_rnd = c3.selectbox("Filter Tahap", sorted(acc_merge['round'].dropna().unique().tolist()))
                
                # Filter Data Target (Unik berdasarkan Vendor dan Route Group)
                acc_target = acc_merge[
                    (acc_merge['load_type'] == sel_sm_lt) & 
                    (acc_merge['validity'] == sel_sm_val) & 
                    (acc_merge['round'] == sel_sm_rnd)
                ].drop_duplicates(subset=['vendor_email', 'route_group'])
                
                # Ambil Data yg sudah disubmit (dari df_master)
                sub_master = df_master[
                    (df_master['load_type'] == sel_sm_lt) & 
                    (df_master['validity'] == sel_sm_val) & 
                    (df_master['round'] == sel_sm_rnd)
                ] if not df_master.empty else pd.DataFrame()
                
                if not acc_target.empty:
                    # --- 1. TAHAP PRE-CALCULATION & PENGUMPULAN DATA ---
                    vendor_data_list = []
                    
                    # Variabel untuk Statistik
                    total_vendors = 0
                    completed_vendors = 0  # Untuk yang selesai FULL
                    started_vendors = 0    # Untuk yang sudah mulai (Minimal 1)
                    total_groups_assigned = 0
                    total_groups_filled = 0
                    
                    for vendor in acc_target['vendor_email'].unique():
                        v_name = df_u[df_u['email']==vendor]['vendor_name'].iloc[0] if not df_u[df_u['email']==vendor].empty else vendor
                        
                        v_acc_subset = acc_target[acc_target['vendor_email'] == vendor]
                        assigned_groups = sorted(v_acc_subset['route_group'].dropna().tolist())
                        assigned_origins = v_acc_subset['origin'].dropna().unique().tolist()
                        
                        submitted_groups = []
                        if not sub_master.empty:
                            submitted_groups = sub_master[sub_master['vendor_email'] == vendor]['route_group'].dropna().unique().tolist()
                        
                        pending_groups = [grp for grp in assigned_groups if grp not in submitted_groups]
                        filled_groups = [grp for grp in assigned_groups if grp in submitted_groups]
                        
                        # Hitung Statistik Global
                        total_vendors += 1
                        
                        # Jika selesai FULL
                        if len(pending_groups) == 0 and len(assigned_groups) > 0:
                            completed_vendors += 1
                            
                        # Jika sudah mengisi MINIMAL 1
                        if len(filled_groups) > 0:
                            started_vendors += 1
                            
                        total_groups_assigned += len(assigned_groups)
                        total_groups_filled += len(filled_groups)
                        
                        vendor_data_list.append({
                            'email': vendor,
                            'name': v_name,
                            'assigned_groups': assigned_groups,
                            'submitted_groups': submitted_groups,
                            'pending_groups': pending_groups,
                            'origins': assigned_origins
                        })
                    
                    # Sort berdasarkan Alfabet Nama Vendor
                    vendor_data_list.sort(key=lambda x: str(x['name']).strip().lower())
                    
                    # --- 2. TAMPILKAN UI STATISTIK ---
                    st.divider()
                    col_stat1, col_stat2, col_stat3 = st.columns(3)
                    with col_stat1:
                        st.info(f"🏆 **Selesai (Full):** {completed_vendors} / {total_vendors} Vendor")
                    with col_stat2:
                        st.info(f"🏃 **Sudah Isi (Min. 1):** {started_vendors} / {total_vendors} Vendor")
                    with col_stat3:
                        st.info(f"📝 **Grup Rute Terisi:** {total_groups_filled} / {total_groups_assigned} Grup")
                    
                    # --- 3. TAMPILKAN SEARCH BAR ---
                    search_query = st.text_input("🔍 Cari berdasarkan Nama Vendor atau Origin (Area)...", placeholder="Contoh: Logistik atau Jakarta").strip().lower()
                    st.write("")
                    
                    # --- 4. FILTER & RENDER LIST VENDOR ---
                    filtered_vendors = []
                    for v in vendor_data_list:
                        # Logika Pencarian: Cocokkan dengan Nama ATAU salah satu Origin-nya
                        match_name = search_query in str(v['name']).lower()
                        match_origin = any(search_query in str(org).lower() for org in v['origins'])
                        
                        if search_query == "" or match_name or match_origin:
                            filtered_vendors.append(v)
                            
                    if not filtered_vendors:
                        st.warning("Tidak ada vendor atau origin yang cocok dengan pencarian.")
                    else:
                        for v in filtered_vendors:
                            v_name = v['name']
                            vendor_email = v['email']
                            assigned_groups = v['assigned_groups']
                            submitted_groups = v['submitted_groups']
                            pending_groups = v['pending_groups']
                            
                            total_g = len(assigned_groups)
                            done_g = total_g - len(pending_groups)
                            
                            if len(pending_groups) > 0:
                                header_text = f"⚠️ {v_name} — (Selesai: {done_g}/{total_g})"
                                is_expanded = False 
                            else:
                                header_text = f"✅ {v_name} — (Lengkap: {total_g}/{total_g})"
                                is_expanded = False 
                                
                            # --- BUAT UI COLLAPSE ---
                            with st.expander(header_text, expanded=is_expanded):
                                # (Optional Info) Tampilkan origin apa saja yg dipegang vendor ini
                                st.caption(f"📍 Area: {', '.join(v['origins'])}")
                                
                                c_h1, c_h2 = st.columns([4, 2])
                                c_h1.write("**Group Rute**")
                                c_h2.write("**Status Pengisian**")
                                st.divider()
                                
                                for grp in assigned_groups:
                                    c_r1, c_r2 = st.columns([4, 2])
                                    c_r1.write(f"{grp}")
                                    if grp in submitted_groups: 
                                        c_r2.markdown('<span class="status-done">✅ Sudah Terisi</span>', unsafe_allow_html=True)
                                    else:
                                        c_r2.markdown('<span class="status-pending">❌ Belum Ada Data</span>', unsafe_allow_html=True)
                                    st.markdown("<hr style='margin: 0.5em 0;'>", unsafe_allow_html=True)
                                
                                if pending_groups:
                                    st.write("") # Memberi sedikit jarak
                                    c_btn1, c_btn2 = st.columns(2)
                                    
                                    # --- TOMBOL EMAIL ---
                                    with c_btn1:
                                        if st.button(f"📨 Kirim Email", key=f"remind_{vendor_email}_{sel_sm_rnd}", type="primary", use_container_width=True):
                                            with st.spinner(f"Mengirim email ke {v_name}..."):
                                                vendor_pw = "Hubungi Admin"
                                                if not df_u[df_u['email']==vendor_email].empty:
                                                    vendor_pw = df_u[df_u['email']==vendor_email].iloc[0]['password']
                                                    
                                                res = send_reminder_email(vendor_email, v_name, sel_sm_lt, sel_sm_val, sel_sm_rnd, pending_groups, vendor_pw)
                                                if res: st.success("Email terkirim!")
                                                else: st.error("Gagal mengirim email.")
                                    
                                    # --- TOMBOL WHATSAPP ---
                                    with c_btn2:
                                        vendor_phone = ""
                                        # Cari nomor HP dari tabel Vendor Profile
                                        if not df_prof.empty:
                                            prof_subset = df_prof[df_prof['email'] == vendor_email]
                                            if not prof_subset.empty:
                                                vendor_phone = str(prof_subset.iloc[0].get('phone', '')).strip()
                                        
                                        # Jika nomor HP ada dan valid
                                        if vendor_phone and vendor_phone != "-" and vendor_phone.lower() != "nan":
                                            # Ubah awalan "08" menjadi "628" sesuai standar API WA
                                            if vendor_phone.startswith("0"):
                                                vendor_phone = "62" + vendor_phone[1:]
                                            
                                            # Siapkan teks draf WA
                                            pending_str = ", ".join([str(g) for g in pending_groups])
                                            wa_text = f"Halo *{v_name}*,\n\nKami dari TACO Group ingin mengingatkan bahwa Anda *belum menyelesaikan* pengisian harga Tender {sel_sm_lt} ({sel_sm_val}) Tahap {sel_sm_rnd} untuk area:\n\n📌 {pending_str}\n\nMohon segera melengkapi penawaran Anda di sistem.\nLink: https://taco-transport.streamlit.app/\n\n*BATAS PENGISIAN: RABU, 11 MARET 2026*\nJika tidak dilakukan pengisian, kami akan anggap dari {v_name} *TIDAK* akan mengikuti tender rute tersebut dan *TIDAK* dapat menyusul.\n\nTerima Kasih."
                                            
                                            # Encode teks agar aman di URL
                                            wa_text_encoded = urllib.parse.quote(wa_text)
                                            wa_url = f"https://wa.me/{vendor_phone}?text={wa_text_encoded}"
                                            
                                            # Render tombol hijau ala WhatsApp
                                            st.markdown(f'<a href="{wa_url}" target="_blank" style="background-color:#25D366; color:white; padding:10px 16px; border-radius:8px; text-decoration:none; font-weight:bold; display:inline-block; text-align:center; width:100%; border: 1px solid #1eaa50; box-shadow: 0 2px 4px rgba(37, 211, 102, 0.2);">💬 Kirim WA</a>', unsafe_allow_html=True)
                                        else:
                                            # Render tombol abu-abu jika nomor tidak ada
                                            st.markdown(f'<div style="background-color:#E5E7EB; color:#6B7280; padding:10px 16px; border-radius:8px; text-decoration:none; font-weight:bold; display:inline-block; text-align:center; width:100%; cursor:not-allowed; border: 1px solid #D1D5DB;">❌ No WA Tidak Ada</div>', unsafe_allow_html=True)
                                else:
                                    st.success("🎉 Pengisian Harga Lengkap!")
                else:
                    st.info("Tidak ada data akses untuk filter ini.")
                    
        # --- TAB 2: LOCK DATA ---
        with tabs[1]:
            st.subheader("Lock Data")
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
            
            merged_pr['price'] = pd.to_numeric(merged_pr['price'], errors='coerce').fillna(0)
            merged_pr = merged_pr[merged_pr['price'] > 0]
                        
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

                            if 'catatan_tambahan' in sub_md.columns:
                                v_note = str(sub_md.iloc[0]['catatan_tambahan']).strip()
                                if v_note and v_note.lower() != "nan" and v_note != "None":
                                    st.info(f"📝 **Catatan Tambahan Vendor:**\n{v_note}")
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
        with tabs[2]:
            st.subheader("📊 Summary & Ranking Vendor")
            
            if df_master.empty:
                st.info("Belum ada data harga masuk.")
            else:
                # --- FILTER HARUS DI ATAS ---
                c1, c2, c3, c4 = st.columns(4)
                avail_val = sorted(df_master['validity'].unique().tolist())
                avail_load = sorted(df_master['load_type'].unique().tolist())
            
                sel_val = c1.selectbox("Filter Periode", avail_val, key="es_val")
                sel_load = c2.selectbox("Filter Tipe Muatan", avail_load, key="es_load")
            
                # 1. Filter Awal (Periode & Muatan)
                df_view = df_master[(df_master['validity'] == sel_val) & (df_master['load_type'] == sel_load)].copy()
                df_view = df_view[df_view['price'] > 0]
                
                # 2. Tambahan Filter Kota Asal & Search Bar
                avail_asal = ["Semua Kota Asal"] + sorted(df_view['kota_asal'].dropna().unique().tolist())
                sel_asal = c3.selectbox("Filter Kota Asal", avail_asal, key="es_asal")
                
                search_keyword = c4.text_input("🔍 Cari Lokasi", placeholder="Ketik Asal/Tujuan...", key="es_dest").strip().lower()
                
                # --- TOMBOL DOWNLOAD BARU DI SINI ---
                with st.expander("📥 Download Master Summary (Excel)", expanded=False):
                    st.write("Unduh rekap seluruh rute sesuai filter di atas. Rute yang belum diisi vendor akan tetap muncul dengan harga Rp 0.")
                    
                    if not df_r.empty and not df_g.empty and not df_units.empty:
                        # Bersihkan ID
                        df_r_clean = df_r.copy()
                        df_r_clean['route_id'] = df_r_clean['route_id'].astype(str).str.strip()
                        df_r_clean['group_id'] = df_r_clean['group_id'].astype(str).str.strip()
                        df_g_clean = df_g.copy()
                        df_g_clean['group_id'] = df_g_clean['group_id'].astype(str).str.strip()
                        df_u_clean = df_units.copy()
                        df_u_clean['group_id'] = df_u_clean['group_id'].astype(str).str.strip()

                        # Gabungkan Master
                        base_df = pd.merge(df_r_clean, df_g_clean, on='group_id', how='left')
                        base_df = pd.merge(base_df, df_u_clean, on='group_id', how='left')
                        
                        # Terapkan Filter UI ke Base Excel
                        base_df = base_df[base_df['load_type'] == sel_load]
                        if sel_asal != "Semua Kota Asal":
                            base_df = base_df[base_df['kota_asal'] == sel_asal]
                        if search_keyword:
                            match_org = base_df['origin'].fillna("").str.lower().str.contains(search_keyword)
                            match_asal = base_df['kota_asal'].fillna("").str.lower().str.contains(search_keyword)
                            match_tujuan = base_df['kota_tujuan'].fillna("").str.lower().str.contains(search_keyword)
                            base_df = base_df[match_org | match_asal | match_tujuan]

                        # Siapkan Harga & Vendor
                        prices_clean = df_p.copy() if not df_p.empty else pd.DataFrame(columns=['route_id', 'unit_type', 'vendor_email', 'price', 'validity', 'round', 'lead_time'])
                        if not prices_clean.empty:
                            prices_clean['route_id'] = prices_clean['route_id'].astype(str).str.strip()
                            prices_clean['price'] = pd.to_numeric(prices_clean['price'], errors='coerce').fillna(0)
                            prices_clean = prices_clean[(prices_clean['price'] > 0) & (prices_clean['validity'] == sel_val)]
                            
                            v_names = df_u[df_u['role'] == 'vendor'][['email', 'vendor_name']]
                            prices_clean = pd.merge(prices_clean, v_names, left_on='vendor_email', right_on='email', how='left')
                            prices_clean['vendor_name'] = prices_clean['vendor_name'].fillna(prices_clean['vendor_email'])

                        # Merge Left Join
                        if not prices_clean.empty:
                            summary_df = pd.merge(
                                base_df, 
                                prices_clean[['route_id', 'unit_type', 'vendor_name', 'price', 'validity', 'round', 'lead_time']], 
                                on=['route_id', 'unit_type'], 
                                how='left'
                            )
                        else:
                            summary_df = base_df.copy()
                            summary_df['vendor_name'] = '-'
                            summary_df['price'] = 0
                            summary_df['validity'] = sel_val
                            summary_df['round'] = '-'
                            summary_df['lead_time'] = '-'

                        # Prioritas, Rapikan Kosong & Format
                        summary_df['price_sort'] = summary_df['price'].replace(0, float('inf'))
                        summary_df = summary_df.sort_values(by=['origin', 'kota_asal', 'kota_tujuan', 'unit_type', 'price_sort'])
                        summary_df['Prioritas'] = summary_df.groupby(['origin', 'kota_asal', 'kota_tujuan', 'unit_type']).cumcount() + 1
                        summary_df['Prioritas'] = summary_df.apply(lambda x: x['Prioritas'] if x['price'] > 0 else '-', axis=1)

                        summary_df['price'] = summary_df['price'].fillna(0)
                        summary_df['vendor_name'] = summary_df['vendor_name'].fillna('Belum Ada Penawaran')
                        summary_df['validity'] = summary_df['validity'].fillna(sel_val).replace('nan', sel_val)
                        summary_df['round'] = summary_df['round'].fillna('-').replace('nan', '-')
                        summary_df['lead_time'] = summary_df['lead_time'].fillna('-').replace('nan', '-')
                        summary_df['Harga Penawaran'] = summary_df['price'].apply(lambda x: f"Rp {int(x):,}".replace(",", "."))
                        
                        cols_to_keep = ['origin', 'kota_asal', 'kota_tujuan', 'route_group', 'load_type', 'unit_type', 'Prioritas', 'vendor_name', 'Harga Penawaran', 'lead_time', 'validity', 'round']
                        for c in cols_to_keep:
                            if c not in summary_df.columns: summary_df[c] = '-'
                        summary_df = summary_df[cols_to_keep]
                        summary_df.columns = ['Origin', 'Kota Asal', 'Kota Tujuan', 'Nama Grup Rute', 'Tipe Muatan', 'Unit', 'Prioritas', 'Nama Vendor', 'Harga Penawaran', 'Lead Time', 'Periode', 'Tahap']
                        
                        # Bikin Excel
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            summary_df.to_excel(writer, index=False, sheet_name='Master Summary')
                        excel_data = output.getvalue()
                        
                        safe_val_name = str(sel_val).replace(" - ", "-").replace(" ", "_")
                        st.download_button(
                            label="📊 Download Master Summary (.xlsx)",
                            data=excel_data,
                            file_name=f"Master_Summary_{sel_load}_{safe_val_name}_{int(time.time())}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True
                        )
                    else:
                        st.warning("Data Master Rute, Group, atau Unit masih kosong. Tidak bisa mengunduh summary.")
                
                st.write("")
                
                # --- LANJUTAN FILTER TAMPILAN UI ---
                if sel_asal != "Semua Kota Asal":
                    df_view = df_view[df_view['kota_asal'] == sel_asal]
                    
                if search_keyword:
                    match_org = df_view['origin'].fillna("").str.lower().str.contains(search_keyword)
                    match_asal = df_view['kota_asal'].fillna("").str.lower().str.contains(search_keyword)
                    match_tujuan = df_view['kota_tujuan'].fillna("").str.lower().str.contains(search_keyword)
                    df_view = df_view[match_org | match_asal | match_tujuan]

                # 4. Tampilkan Hasilnya ke Layar
                if not df_view.empty:
                    unique_origins = sorted(df_view['origin'].unique())
                
                    for org in unique_origins:
                        # Expander tetap per Origin Area agar tetap rapi pengelompokannya
                        with st.expander(f"📍 Origin Area: {org}", expanded=True):
                            sub_df = df_view[df_view['origin'] == org].copy()
                        
                            # Ranking Logic (Diperbarui: Tambah kota_asal agar Top 3 dihitung per rute spesifik)
                            sub_df = sub_df.sort_values(by=['kota_asal', 'kota_tujuan', 'unit_type', 'price'])
                            sub_df['Ranking'] = sub_df.groupby(['kota_asal', 'kota_tujuan', 'unit_type']).cumcount() + 1
                        
                            # ▼▼▼ FILTER HANYA TOP 3 ▼▼▼
                            sub_df = sub_df[sub_df['Ranking'] <= 3]
                            # ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
                        
                            sub_df['price_fmt'] = sub_df['price'].apply(lambda x: f"Rp {int(x):,}".replace(",", "."))
                        
                            # Menampilkan 'kota_asal' di dalam tabel
                            st.dataframe(
                                sub_df[['kota_asal', 'kota_tujuan', 'unit_type', 'Ranking', 'vendor_name', 'price_fmt', 'lead_time', 'top']],
                                use_container_width=True,
                                column_config={
                                    "kota_asal": "Kota Asal",
                                    "kota_tujuan": "Kota Tujuan",
                                    "unit_type": "Unit",
                                    "price_fmt": "Harga",
                                    "vendor_name": "Vendor"
                                },
                                hide_index=True
                            )
                else:
                    st.warning("Data tidak ditemukan.")
                    
        # --- TAB 8: PRINT FILE (SK & SPK TERPISAH) ---
        with tabs[3]:
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
                    
                    # --- TAMBAHAN FILTER HARGA 0 ---
                    df_sk = df_sk[df_sk['price'] > 0]
                
                    if not df_sk.empty:
                        avail_org = sorted(df_sk['origin'].unique())
                        sel_orgs = st.multiselect("Pilih Origin (SK):", avail_org, default=avail_org, key="sk_orgs")
                    
                        if sel_orgs:
                            df_final_sk = df_sk[df_sk['origin'].isin(sel_orgs)].copy()
                            df_final_sk = df_final_sk.sort_values(by=['origin', 'kota_tujuan', 'unit_type', 'price'])
                            df_final_sk['Ranking'] = df_final_sk.groupby(['origin', 'kota_tujuan', 'unit_type']).cumcount() + 1
                        
                            col_a, col_b = st.columns(2)
                            upl_sk = col_a.file_uploader("Upload Template SK", type="docx", key="upl_sk")
                            no_sk = col_b.text_input("Nomor Surat SK:", "", key="no_sk")
                        
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
                
                # --- TAMBAHAN FILTER HARGA 0 ---
                df_spk_raw = df_spk_raw[df_spk_raw['price'] > 0]
                
                if not df_spk_raw.empty:
                    # Pilih Vendor
                    avail_vens = sorted(df_spk_raw['vendor_name'].unique().tolist())
                    sel_ven = st.selectbox("Pilih Vendor (SPK):", avail_vens, key="spk_ven")
                    
                    # Filter Data Vendor
                    df_final_spk = df_spk_raw[df_spk_raw['vendor_name'] == sel_ven].copy()
                    
                    if not df_final_spk.empty:
                        # Preview kecil
                        st.info(f"Vendor **{sel_ven}** memiliki **{len(df_final_spk)}** rute di periode ini.")
                        
                        # Hitung Ranking Ulang (Just in case)
                        df_final_spk = df_final_spk.sort_values(by=['kota_asal', 'kota_tujuan', 'unit_type', 'price'])
                        if 'Ranking' not in df_final_spk.columns: df_final_spk['Ranking'] = 1

                        col_c, col_d = st.columns(2)
                        upl_spk = col_c.file_uploader("Upload Template SPK", type="docx", key="upl_spk")
                        no_spk = col_d.text_input("Nomor Surat SPK:", f"", key="no_spk")
                        
                        # --- START BUTTON BLOCK ---
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

                            # 3. KUMPULKAN ORIGIN & ALAMAT (Multi Origin Logic)
                            list_origin = sorted(df_final_spk['origin'].unique().tolist())
                            origin_str_combined = ", ".join(list_origin)
                            
                            alamat_list = []
                            try:
                                df_gudang = get_data("Gudang") 
                                if not df_gudang.empty:
                                    for org in list_origin:
                                        res_addr = df_gudang[df_gudang['origin'].astype(str).str.lower() == str(org).lower()]
                                        if not res_addr.empty:
                                            alamat_found = res_addr.iloc[0]['alamat']
                                            if len(list_origin) > 1: alamat_list.append(f"{org}: {alamat_found}")
                                            else: alamat_list.append(alamat_found)
                                        else: alamat_list.append(f"{org}: -")
                                else: alamat_list.append("(Sheet Gudang Kosong)")
                            except Exception as e: alamat_list.append(f"Error: {e}")
                            
                            alamat_str_combined = "\n".join(alamat_list)

                            # 4. MERGE MULTIDROP
                            df_spk_merged = df_final_spk.copy()
                            # Default columns
                            df_spk_merged['inner_city_price'] = 0
                            df_spk_merged['outer_city_price'] = 0
                            df_spk_merged['labor_cost'] = 0

                            if not df_md.empty:
                                try:
                                    md_dict = {}
                                    for _, rmd in df_md.iterrows():
                                        k = (str(rmd['vendor_email']).strip(), str(rmd['validity']).strip(), str(rmd['group_id']).strip())
                                        md_dict[k] = {'in': rmd.get('inner_city_price',0), 'out': rmd.get('outer_city_price',0), 'lab': rmd.get('labor_cost',0)}
                                    
                                    def get_md_val(row, kind):
                                        key = (str(row['vendor_email']).strip(), str(row['validity']).strip(), str(row['group_id']).strip())
                                        res = md_dict.get(key, {'in':0, 'out':0, 'lab':0})
                                        return res[kind]

                                    df_spk_merged['inner_city_price'] = df_spk_merged.apply(lambda x: get_md_val(x, 'in'), axis=1)
                                    df_spk_merged['outer_city_price'] = df_spk_merged.apply(lambda x: get_md_val(x, 'out'), axis=1)
                                    df_spk_merged['labor_cost'] = df_spk_merged.apply(lambda x: get_md_val(x, 'lab'), axis=1)
                                except: pass

                            # 5. GENERATE
                            with st.spinner(f"Memproses SPK {sel_ven}..."):
                                try:
                                    f_spk = create_docx_spk(tpl_spk, no_spk, spk_val, spk_load, sel_ven, pic, final_pass, origin_str_combined, alamat_str_combined, df_spk_merged)
                                    
                                    safe_val = str(spk_val).replace(" - ", "-").replace(" ", "_")
                                    safe_load = str(spk_load).replace(" ", "")
                                    safe_ven_file = str(sel_ven).replace(" ", "_").replace(".", "").replace(",", "")
                                    custom_filename = f"SPK_{safe_load}_{safe_val}_{safe_ven_file}.docx"
                                    
                                    with open(f_spk, "rb") as f:
                                        st.download_button("⬇️ Download SPK", f, file_name=custom_filename)
                                    os.remove(f_spk)
                                except Exception as e: st.error(f"Gagal generate: {e}")
                        # --- END BUTTON BLOCK ---
                    else: st.warning("Vendor ini tidak memiliki data.")
                else: st.warning("Data tidak ditemukan.")
        # --- TAB 5: SPH UPLOADS (FITUR BARU) ---
        with tabs[4]:
            st.subheader("📥 Dokumen SPH Vendor (Signed)")
            st.caption("Daftar dokumen Surat Penawaran Harga yang sudah ditandatangani dan di-upload oleh Vendor.")
            
            df_uploads = get_data("SPH_Uploads")
            
            if df_uploads.empty:
                st.info("Belum ada vendor yang mengupload dokumen SPH.")
            else:
                # Tampilkan tabel riwayat upload (terbaru di atas)
                df_uploads = df_uploads.sort_values(by='timestamp', ascending=False)
                
                # Buat tampilan list per dokumen
                for _, row in df_uploads.iterrows():
                    with st.container(border=True):
                        col1, col2, col3 = st.columns([4, 2, 2])
                        col1.markdown(f"**🏢 {row.get('vendor_name', '-')}**")
                        col1.caption(f"Tipe: {row.get('load_type', '-')} | Tahap: {row.get('round', '-')} | Periode: {row.get('validity', '-')}")
                        col2.write(f"🕒 {row.get('timestamp', '-')}")
                        
                        file_url = row.get('filename', '')
                        
                        with col3:
                            if str(file_url).startswith("http"):
                                st.markdown(f'<a href="{file_url}" target="_blank" style="background-color:#2563EB; color:white; padding:8px 12px; border-radius:8px; text-decoration:none; font-weight:bold; display:inline-block; text-align:center; width:100%;">🔗 Buka di Drive</a>', unsafe_allow_html=True)
                            else:
                                st.error("❌ Link Error")               
                                
# ================= VENDOR DASHBOARD (UPDATE: DYNAMIC TABS) =================
def vendor_dashboard(email):
    step = st.session_state['vendor_step']
    
    # --- STEP 1: DASHBOARD / PROFIL ---
    if step == "dashboard":
        t1, t2, t3 = st.tabs(["🛣️ Pilih Rute & Isi Harga", "📋 Isi Data Perusahaan", "📄 Surat Penawaran Harga"])
        
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
            df_gudang = get_data("Gudang")

            if df_acc.empty: 
                st.warning("Belum ada akses.")
                return
            
            my_access = df_acc[df_acc['vendor_email'] == email]
            if my_access.empty: 
                st.info("Anda belum diberikan akses ke project manapun.")
                return

            if 'round' not in my_access.columns: my_access['round'] = '1'
            avail_rounds = sorted(my_access['round'].unique().tolist())
            
            c_filter1, c_filter2 = st.columns(2)
            sel_round = c_filter1.selectbox("Pilih Tahap Penawaran:", avail_rounds)
            my_access = my_access[my_access['round'] == sel_round]

            data_list = []
            for _, acc in my_access.iterrows():
                # PENGAMAN: Gunakan .get()
                gid = acc.get('group_id')
                val = acc.get('validity')
                if not gid or not val: continue
                
                g_info = df_grps[df_grps['group_id'] == gid]
                if not g_info.empty:
                    row = g_info.iloc[0]
                    data_list.append({'validity': val, 'group_id': gid, 'origin': row.get('origin','-'), 'route_group': row.get('route_group','-'), 'load_type': row.get('load_type','-')})
            
            df_disp = pd.DataFrame(data_list)
            if df_disp.empty: 
                st.warning(f"Tidak ada rute di Penawaran Tahap {sel_round}.")
                return

            avail_validity = sorted(df_disp['validity'].unique().tolist())
            sel_val = c_filter2.selectbox("Pilih Periode / Validity:", avail_validity)
            df_view = df_disp[df_disp['validity'] == sel_val]
            
            if df_view.empty: 
                st.info("Tidak ada rute.")
                return
            
            # --- LOGIKA TAB DINAMIS (UPDATE) ---
            # 1. Cek tipe apa saja yang tersedia di data
            avail_types_raw = df_view['load_type'].unique().tolist()
            
            # 2. Definisikan urutan dan nama tab
            type_map = {
                "FTL": "🚛 FTL (Full Truck Load)",
                "FCL": "🚢 FCL (Full Container Load)"
            }
            
            # 3. Filter hanya tipe yang ada datanya (FTL duluan jika ada)
            final_types = [t for t in ["FTL", "FCL"] if t in avail_types_raw]
            
            if not final_types:
                st.warning("Tipe muatan tidak dikenali.")
            else:
                # 4. Buat Tab sesuai isi final_types
                tab_labels = [type_map.get(t, t) for t in final_types]
                created_tabs = st.tabs(tab_labels)
                
                # 5. Loop Render (zip menggabungkan list tipe dan list tab object)
                for t_code, t_ui in zip(final_types, created_tabs):
                    with t_ui:
                        df_sub = df_view[df_view['load_type'] == t_code]
                        # Disini tidak perlu cek empty lagi, karena tab dibuat hanya jika data ada
                        
                        # PENGAMAN: Cegah error origin kosong/tidak terbaca
                        unique_orgs = [o for o in df_sub['origin'].dropna().unique() if str(o).strip() != ""]
                        for org in sorted(unique_orgs, key=lambda x: str(x).strip().lower()):
                            with st.container(border=True):
                                st.markdown(f"#### 📍 {org}")

                                # --- TAMPILKAN ALAMAT GUDANG ---
                                if not df_gudang.empty and 'origin' in df_gudang.columns and 'alamat' in df_gudang.columns:
                                    res_addr = df_gudang[df_gudang['origin'].astype(str).str.lower() == str(org).lower()]
                                    if not res_addr.empty:
                                        st.caption(f"🏢 **Alamat:** {res_addr.iloc[0]['alamat']}")
                                # -------------------------------
                                
                                org_groups = df_sub[df_sub['origin'] == org]
                                c1, c2, c3, c4 = st.columns([3, 4, 2, 2])
                                c1.caption(""); c2.caption("Kota Tujuan"); c3.caption(""); c4.caption("Status Pengisian")
                                st.divider()
                                
                                for _, row in org_groups.iterrows():
                                    gid = row['group_id']
                                    grp_name = row['route_group']
                                    
                                    r_data = df_routes[df_routes['group_id'] == gid] if not df_routes.empty else pd.DataFrame()
                                    status_ui = '<span class="status-pending">❌ Belum Ada Data</span>'
                                    is_locked_btn = False
                                    
                                    # PENGAMAN: Pastikan kolom vendor_email benar-benar ada
                                    if not df_price.empty and not r_data.empty and 'vendor_email' in df_price.columns:
                                        if 'round' not in df_price.columns: df_price['round'] = '1'
                                        sub_p = df_price[
                                            (df_price['vendor_email'] == email) & 
                                            (df_price['validity'] == sel_val) & 
                                            (df_price['route_id'].isin(r_data['route_id'])) &
                                            (df_price['round'] == sel_round)
                                        ]
                                        if not sub_p.empty:
                                            status_ui = '<span class="status-done">✅Sudah Terisi</span>'
                                            if "Locked" in sub_p['status'].values: is_locked_btn = True
                                    
                                    c1, c2, c3, c4 = st.columns([3, 4, 2, 2])
                                    c1.write(f"**{grp_name}**")
                                    
                                    # PENGAMAN: Cek apakah kolom kota_tujuan benar-benar ada
                                    dests = r_data['kota_tujuan'].unique().tolist() if (not r_data.empty and 'kota_tujuan' in r_data.columns) else []
                                    
                                    if len(dests) > 5: preview_txt = f"{', '.join(dests[:5])}, +{len(dests)-5} kota lainnya"
                                    else: preview_txt = ", ".join(dests)
                                    c2.markdown(f"<span class='route-dest-list'>{preview_txt}</span>", unsafe_allow_html=True)

                                    if is_locked_btn:
                                        c3.button("🔒 Locked", key=f"btn_lk_{gid}_{sel_round}", disabled=True)
                                    else:
                                        if c3.button("📌 Isi Harga", key=f"btn_{t_code}_{gid}_{sel_round}", type="primary"):
                                            st.session_state.update({
                                                'sel_origin': org, 'sel_validity': sel_val, 'sel_load': t_code, 
                                                'vendor_step': 'input', 'focused_group_id': gid, 'sel_round': sel_round
                                            })
                                            st.rerun()
                                    c4.markdown(status_ui, unsafe_allow_html=True)
                                    st.markdown("<hr>", unsafe_allow_html=True)
# --- TAB 3: DOWNLOAD & UPLOAD SPH RESMI ---
        with t3:
            st.markdown("### 📄 SPH (Surat Penawaran Harga)")
            st.info("Anda dapat mendownload draf SPH kapan saja. Namun, **fitur Upload baru terbuka setelah SEMUA rute Anda di-Lock oleh Admin**.")
            
            df_p = get_data("Price_Data")
            df_r = get_data("Master_Routes")
            df_g = get_data("Master_Groups")
            df_prof = get_data("Vendor_Profile")
            
            if df_p.empty or df_g.empty:
                st.warning("Belum ada data penawaran.")
            else:
                # 1. Filter Milik Vendor Ini (SEMUA STATUS: Open & Locked) dan Harga > 0
                my_prices = df_p[df_p['vendor_email'] == email].copy()
                my_prices['price'] = pd.to_numeric(my_prices['price'], errors='coerce').fillna(0)
                my_prices = my_prices[my_prices['price'] > 0]
                
                if my_prices.empty:
                    st.warning("Belum ada data harga Anda, atau semua harga Anda masih Rp 0.")
                else:
                    # 2. Merge Data
                    my_prices['route_id'] = my_prices['route_id'].astype(str).str.strip()
                    df_r['route_id'] = df_r['route_id'].astype(str).str.strip()
                    df_g['group_id'] = df_g['group_id'].astype(str).str.strip()
                    if 'round' not in my_prices.columns: my_prices['round'] = '1'
                    
                    m1 = pd.merge(my_prices, df_r[['route_id', 'group_id', 'kota_tujuan']], on='route_id', how='left')
                    df_final = pd.merge(m1, df_g[['group_id', 'origin', 'load_type']], on='group_id', how='left')
                    # --- TAMBAHAN MERGE MULTIDROP & CATATAN ---
                    df_m = get_data("Multidrop_Data")
                    if not df_m.empty:
                        df_m['group_id'] = df_m['group_id'].astype(str).str.strip()
                        df_m['md_round'] = df_m['id_multidrop'].apply(lambda x: str(x).split('_')[-1] if '_' in str(x) else '1')
                        df_m_sub = df_m[['vendor_email', 'validity', 'group_id', 'md_round', 'inner_city_price', 'outer_city_price', 'labor_cost', 'catatan_tambahan']]
                        
                        df_final = pd.merge(
                            df_final, df_m_sub, 
                            left_on=['vendor_email', 'validity', 'group_id', 'round'],
                            right_on=['vendor_email', 'validity', 'group_id', 'md_round'],
                            how='left'
                        )
                    else:
                        for c in ['inner_city_price', 'outer_city_price', 'labor_cost', 'catatan_tambahan']: df_final[c] = 0
                    # 3. UI Filter (Periode, Load Type, Tahap)
                    c1, c2, c3 = st.columns(3)
                    avail_val = sorted(df_final['validity'].unique().tolist())
                    sel_val = c1.selectbox("Pilih Periode", avail_val, key="sph_val")
                    
                    avail_lt = sorted(df_final['load_type'].dropna().unique().tolist())
                    sel_lt = c2.selectbox("Pilih Tipe Armada", avail_lt, key="sph_lt")
                    
                    avail_rnd = sorted(df_final['round'].unique().tolist())
                    sel_rnd = c3.selectbox("Pilih Tahap Penawaran", avail_rnd, key="sph_rnd")
                    
                    df_print = df_final[(df_final['validity'] == sel_val) & (df_final['load_type'] == sel_lt) & (df_final['round'] == sel_rnd)].copy()
                    v_name = st.session_state['user_info'].get('vendor_name', email)
                    
                    # --- CEK APAKAH SEMUA STATUS SUDAH LOCKED ---
                    is_all_locked = False
                    if not df_print.empty:
                        # Cek apakah seluruh baris di tabel df_print statusnya 'Locked'
                        is_all_locked = (df_print['status'] == 'Locked').all()
                    
                    # --- BAGIAN A: DOWNLOAD SPH ---
                    with st.container(border=True):
                        st.markdown("#### Step 1: Download SPH")
                        if df_print.empty:
                            st.info("Tidak ada data untuk kombinasi filter ini.")
                        else:
                            st.write(f"Ditemukan **{len(df_print)} rute** yang siap dicetak SPH-nya. Mohon dapat dicap dan ditandatangani, lalu diupload pada langkah selanjutnya.")
                            
                            # --- MEMBUAT 2 TOMBOL BERSEBELAHAN ---
                            c_btn1, c_btn2 = st.columns(2)
                            
                            # TOMBOL KIRI (WORD SPH RESMI)
                            with c_btn1:
                                if st.button("📄 Buat Dokumen SPH (Word)", type="primary", use_container_width=True):
                                    tpl_sph = "template_sph.docx"
                                    if not os.path.exists(tpl_sph): 
                                        st.error("Sistem error: Template SPH tidak ditemukan di server. Hubungi Admin!")
                                        st.stop()
                                        
                                    with st.spinner("Merakit Dokumen SPH..."):
                                        try:
                                            v_addr = "-"
                                            if not df_prof.empty:
                                                vp = df_prof[df_prof['email'] == email]
                                                if not vp.empty: v_addr = str(vp.iloc[0].get('address', '-'))
                                            
                                            df_print_word = df_print.sort_values(by=['origin', 'kota_tujuan', 'unit_type']).reset_index(drop=True)
                                            file_sph = create_docx_sph(tpl_sph, v_name, v_addr, sel_val, sel_lt, sel_rnd, df_print_word)
                                            
                                            with open(file_sph, "rb") as f:
                                                st.download_button(
                                                    label="⬇️ Klik di sini untuk Download Word", 
                                                    data=f, 
                                                    file_name=f"SPH_{v_name}_{sel_lt}_Tahap{sel_rnd}.docx", 
                                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                                                    type="primary",
                                                    use_container_width=True
                                                )
                                            os.remove(file_sph)
                                        except Exception as e: st.error(f"Gagal: {e}")
                            
                            # TOMBOL KANAN (EXCEL TABEL SAJA)
                            with c_btn2:
                                df_excel = df_print.sort_values(by=['origin', 'kota_tujuan', 'unit_type']).reset_index(drop=True)
                                df_excel['No'] = range(1, len(df_excel) + 1)
                                
                                # Rapikan isi kolom sebelum di-export
                                def fmt_rp(x):
                                    try: return f"Rp {int(x):,}".replace(",", ".")
                                    except: return "Rp 0"
                                df_excel['Harga Penawaran'] = df_excel['price'].apply(fmt_rp)
                                df_excel['Multidrop Dalam Kota'] = df_excel['inner_city_price'].apply(fmt_rp)
                                df_excel['Multidrop Luar Kota'] = df_excel['outer_city_price'].apply(fmt_rp)
                                df_excel['Biaya Buruh'] = df_excel['labor_cost'].apply(fmt_rp)
                                
                                def fmt_lt(x):
                                    x_str = str(x)
                                    if x_str.isdigit() or x_str.replace('.','',1).isdigit(): return f"{x_str} Hari"
                                    return "-" if x_str in ["-", "", "nan", "None"] else x_str
                                df_excel['Lead Time (Hari)'] = df_excel['lead_time'].apply(fmt_lt)
                                
                                # Bersihkan teks kosong (PENGAMAN BARU)
                                if 'keterangan' not in df_excel.columns: df_excel['keterangan'] = '-'
                                df_excel['keterangan'] = df_excel['keterangan'].fillna('-').astype(str).replace(['nan', 'None', ''], '-')
                                
                                if 'catatan_tambahan' not in df_excel.columns: df_excel['catatan_tambahan'] = '-'
                                df_excel['catatan_tambahan'] = df_excel['catatan_tambahan'].fillna('-').astype(str).replace(['nan', 'None', ''], '-')
                                
                                # Ambil kolom yang diperlukan saja dan ganti judulnya
                                cols_to_keep = ['No', 'origin', 'kota_tujuan', 'unit_type', 'Lead Time (Hari)', 'Harga Penawaran', 'keterangan', 'Multidrop Dalam Kota', 'Multidrop Luar Kota', 'Biaya Buruh', 'catatan_tambahan']
                                df_excel = df_excel[cols_to_keep]
                                df_excel.columns = ['No', 'Origin', 'Tujuan', 'Unit', 'Lead Time', 'Harga Penawaran', 'Keterangan Rute', 'MD Dalam Kota', 'MD Luar Kota', 'Biaya Buruh', 'Catatan Tambahan Vendor']
                                
                                # Ubah ke format file Excel (BytesIO)
                                output = io.BytesIO()
                                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                    df_excel.to_excel(writer, index=False, sheet_name='Tabel SPH')
                                excel_data = output.getvalue()
                                
                                safe_ven_xls = "".join(x for x in v_name if x.isalnum())
                                
                                # Munculkan tombol Download langsung
                                st.download_button(
                                    label="📊 Download Tabel (Excel)", 
                                    data=excel_data, 
                                    file_name=f"Tabel_SPH_{safe_ven_xls}_{sel_lt}_Tahap{sel_rnd}.xlsx", 
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    type="primary",
                                    use_container_width=True
                                )

                    # --- BAGIAN B: UPLOAD SPH ---
                    with st.container(border=True):
                        st.markdown("#### Step 2: Upload SPH yang Sudah di Cap & Tanda Tangan")
                        
                        if not is_all_locked:
                            # 🛑 Jika belum locked semua, fitur upload disembunyikan dan muncul peringatan
                            st.warning("⏳ **Fitur Upload Terkunci!**\nAdmin TACO harus me-Lock SEMUA penawaran Anda di periode dan tahap ini terlebih dahulu sebelum Anda bisa mengupload dokumen SPH final.")
                        else:
                            # ✅ Jika sudah locked semua, tampilkan uploader
                            st.write(f"Upload untuk: **{sel_lt} | Periode {sel_val} | Tahap {sel_rnd}**")
                            
                            uploaded_file = st.file_uploader("Pilih file SPH (PDF)", type=['pdf', 'png', 'jpg', 'jpeg'])
                            
                            if st.button("📤 Upload Dokumen SPH", type="primary", use_container_width=True):
                                if uploaded_file is not None:
                                    # MASUKKAN ID FOLDER GOOGLE DRIVE ANDA DI SINI
                                    DRIVE_FOLDER_ID = "0AMIguJ49asOLUk9PVA" 
                                    
                                    safe_ven = "".join(x for x in v_name if x.isalnum())
                                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                    ext = uploaded_file.name.split(".")[-1]
                                    new_filename = f"SPH_{safe_ven}_{sel_lt}_T{sel_rnd}_{timestamp}.{ext}"
                                    
                                    with st.spinner("⏳ Sedang mengirim file ke Google Drive... Mohon tunggu..."):
                                        file_url = upload_to_drive(uploaded_file, new_filename, uploaded_file.type, DRIVE_FOLDER_ID)
                                        
                                        if file_url:
                                            id_up = f"UPL_{safe_ven}_{timestamp}"
                                            save_data("SPH_Uploads", [[id_up, email, v_name, sel_val, sel_lt, sel_rnd, file_url, timestamp]])
                                            st.success("✅ Sukses! Dokumen SPH telah tersimpan aman di Server TACO.")
                                        else:
                                            st.error("⚠️ Gagal mengirim dokumen ke server. Coba lagi.")
                                else:
                                    st.error("⚠️ Silakan pilih file terlebih dahulu sebelum klik Upload.")
    
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
        cur_round = str(st.session_state.get('sel_round'))
        focused_gid = st.session_state.get('focused_group_id')

        try: prev_round = str(int(cur_round) - 1)
        except: prev_round = "0"

        st.markdown(f"### Input Penawaran Harga {cur_load}: {cur_org}")
        st.caption(f"Periode: {cur_val} | **Tahap Penawaran: {cur_round}**")

        df_acc = get_data("Access_Rights"); df_grps = get_data("Master_Groups")
        if 'round' not in df_acc.columns: df_acc['round'] = '1'
        my_acc = df_acc[(df_acc['vendor_email']==email) & (df_acc['validity']==cur_val) & (df_acc['round']==cur_round)]
        
        target_gids = []
        grp_names = {}
        for gid in my_acc['group_id'].unique():
            r = df_grps[df_grps['group_id']==gid]
            if not r.empty:
                rr = r.iloc[0]
                if rr['origin']==cur_org and rr['load_type']==cur_load:
                    target_gids.append(gid); grp_names[gid]=rr['route_group']
        
        if not target_gids: 
            st.error("Data error."); return

        target_gids = sorted(target_gids)
        if focused_gid and focused_gid in target_gids:
            target_gids.remove(focused_gid)
            target_gids.insert(0, focused_gid)
        
        tabs = st.tabs([grp_names[g] for g in target_gids])
        
        df_r = get_data("Master_Routes"); df_u = get_data("Master_Units"); df_p = get_data("Price_Data"); df_m = get_data("Multidrop_Data")
        if not df_p.empty and 'round' not in df_p.columns: df_p['round'] = '1'
        if not df_m.empty and 'group_id' in df_m.columns: df_m['group_id'] = df_m['group_id'].astype(str).str.strip()

        for i, gid in enumerate(target_gids):
            with tabs[i]:
                g_name = grp_names[gid]
                my_r = df_r[df_r['group_id']==gid]
                my_u = df_u[df_u['group_id']==gid]
                u_types = my_u['unit_type'].unique().tolist()
                
                if my_r.empty or not u_types: st.warning("Data belum lengkap."); continue

                ex_price = {}; ex_spec = {}
                is_lock = False
                
                # --- LOGIKA PRE-FILL ---
                current_p_data = pd.DataFrame()
                if not df_p.empty:
                    current_p_data = df_p[
                        (df_p['vendor_email']==email) & (df_p['validity']==cur_val) & 
                        (df_p['route_id'].isin(my_r['route_id'])) & (df_p['round'] == cur_round)
                    ]
                
                source_p_data = current_p_data
                is_using_prev_data = False
                if current_p_data.empty and cur_round != "1":
                    if not df_p.empty:
                        source_p_data = df_p[
                            (df_p['vendor_email']==email) & (df_p['validity']==cur_val) & 
                            (df_p['route_id'].isin(my_r['route_id'])) & (df_p['round'] == prev_round)
                        ]
                        is_using_prev_data = True 

                if not source_p_data.empty:
                    if not is_using_prev_data and "Locked" in source_p_data['status'].values: is_lock = True
                    for _, row in source_p_data.iterrows():
                        harga_bersih = clean_numeric(row['price'])
                        ex_price[(row['route_id'], row['unit_type'])] = int(harga_bersih) if harga_bersih else 0
                        ex_spec[row['unit_type']] = {'w': row.get('weight_capacity'), 'c': row.get('cubic_capacity')}

                with st.form(key=f"f_{gid}_{cur_round}"):
                    # 1. SPEC
                    with st.container(border=True):
                        st.markdown(f"#### 🛻 Spesifikasi Armada")
                        if is_using_prev_data: st.info(f"ℹ️ Data disalin dari Tahap {prev_round}.")
                        
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
                        ed_sp = st.data_editor(df_sp, hide_index=True, use_container_width=True, disabled=is_lock, column_config=cf_sp)

                    # 2. PRICE
                    with st.container(border=True):
                        st.markdown(f"#### 💰 Penawaran Harga")
                        p_data = []
                        for _, row in my_r.iterrows():
                            rid = row['route_id']
                            rd = {
                                "Route ID": rid, "Kota Asal": row['kota_asal'], "Kota Tujuan": row['kota_tujuan'],
                                "Keterangan": row.get('keterangan', '-'),
                                "Lead Time (Hari)": 0 
                            }
                            if not source_p_data.empty:
                                temp_lt = source_p_data[source_p_data['route_id']==rid]['lead_time']
                                if not temp_lt.empty: rd["Lead Time (Hari)"] = clean_numeric(temp_lt.iloc[0]) or 0

                            for u in u_types: 
                                if cur_round == "2":
                                    tgt = get_target_price(df_p, rid, u, cur_val)
                                    rd[f"Target {u}"] = tgt
                                rd[f"Harga {u} per trip"] = ex_price.get((rid, u), 0)
                            p_data.append(rd)
                        
                        df_pr = pd.DataFrame(p_data)
                        
                        # Config
                        cf_pr = {
                            "Route ID": None,
                            "Kota Asal": st.column_config.TextColumn(disabled=True, width="small"),
                            "Kota Tujuan": st.column_config.TextColumn(disabled=True, width="small"),
                            "Keterangan": st.column_config.TextColumn(width="medium"),
                            "Lead Time (Hari)": st.column_config.NumberColumn(min_value=0, step=1, width="small")
                        }
                        
                        cols_order = ["Route ID", "Kota Asal", "Kota Tujuan", "Lead Time (Hari)"]
                        for u in u_types:
                            cf_pr[f"Harga {u} per trip"] = st.column_config.NumberColumn(min_value=0, step=1000, format="Rp %d", required=True, width="medium")
                            target_col = f"Target {u}"
                            if target_col in df_pr.columns:
                                cf_pr[target_col] = st.column_config.NumberColumn(format="Rp %d", disabled=True, width="medium")
                                cols_order.append(target_col)
                            cols_order.append(f"Harga {u} per trip")
                        
                        cols_order.append("Keterangan")

                        df_pr = df_pr[cols_order]
                        ed_pr = st.data_editor(df_pr, hide_index=True, use_container_width=True, disabled=is_lock, column_config=cf_pr)

                    # 3. MULTIDROP
                    with st.container(border=True):
                        st.markdown("#### 📦 Biaya Multidrop & Buruh")
                        ic, oc, lc = 0, 0, 0
                        md_source = pd.DataFrame()
                        
                        if not df_m.empty:
                            base_m = df_m[(df_m['vendor_email']==email) & (df_m['validity']==cur_val) & (df_m['group_id']==gid)]
                            if not base_m.empty:
                                md_curr = base_m[base_m['id_multidrop'].astype(str).str.endswith(f"_{cur_round}")]
                                if not md_curr.empty: md_source = md_curr
                                elif cur_round != "1":
                                    md_prev = base_m[base_m['id_multidrop'].astype(str).str.endswith(f"_{prev_round}")]
                                    if not md_prev.empty: md_source = md_prev

                        if not md_source.empty:
                            ic = clean_numeric(md_source.iloc[0].get('inner_city_price')) or 0
                            oc = clean_numeric(md_source.iloc[0].get('outer_city_price')) or 0
                            lc = clean_numeric(md_source.iloc[0].get('labor_cost')) or 0
                        
                        df_md_ui = pd.DataFrame([{"Multidrop Dalam Kota": ic, "Multidrop Luar Kota": oc, "Biaya Buruh": lc}])
                        
                        cf_md = {
                            "Multidrop Dalam Kota": st.column_config.NumberColumn(min_value=0, step=1000, format="Rp %d"),
                            "Multidrop Luar Kota": st.column_config.NumberColumn(min_value=0, step=1000, format="Rp %d"),
                            "Biaya Buruh": st.column_config.NumberColumn(min_value=0, step=1000, format="Rp %d")
                        }
                        ed_md = st.data_editor(df_md_ui, hide_index=True, use_container_width=True, disabled=is_lock, column_config=cf_md)
                        st.markdown("<br><b>📝 Catatan Tambahan (Opsional)</b>", unsafe_allow_html=True)
                        
                        prev_note = ""
                        if not md_source.empty and 'catatan_tambahan' in md_source.columns:
                            prev_note = str(md_source.iloc[0]['catatan_tambahan'])
                            if prev_note.lower() == "nan": prev_note = ""
                            
                        vendor_note = st.text_area(
                            "Keterangan lain yang perlu diinfokan (jika ada):", 
                            value=prev_note, 
                            disabled=is_lock,
                            height=100
                        )
                    # SAVE BUTTON
                    st.write("")
                    if st.form_submit_button(f"Simpan Data {cur_load} {g_name} (Tahap {cur_round})", type="primary") and not is_lock:
                        c_spec = {r['Jenis Unit']: {'w': r['Kapasitas Berat Bersih (Kg)'], 'c': r['Kapasitas Kubikasi Dalam (CBM)']} for _, r in ed_sp.iterrows()}
                        
                        f_data = []
                        ts = (datetime.utcnow() + timedelta(hours=7)).strftime("%Y-%m-%d %H:%M:%S")
                        
                        for _, r in ed_pr.iterrows():
                            rid = str(r['Route ID']); lt = int(r['Lead Time (Hari)'])
                            ket = str(r['Keterangan'])
                            
                            for u in u_types:
                                pr = int(r[f"Harga {u} per trip"])
                                w = str(c_spec.get(u,{}).get('w','')); c = str(c_spec.get(u,{}).get('c',''))
                                tid = f"{email}_{cur_val}_{rid}_{u}_{cur_round}".replace(" ","")
                                f_data.append([tid, email, "Open", cur_val, rid, u, lt, pr, w, c, ket, ts, cur_round])
                        
                        mi = int(ed_md.iloc[0]["Multidrop Dalam Kota"])
                        mo = int(ed_md.iloc[0]["Multidrop Luar Kota"])
                        ml = int(ed_md.iloc[0]["Biaya Buruh"])
                        mid = f"M_{email}_{gid}_{cur_val}_{cur_round}"

                        save_data("Price_Data", f_data)
                        save_data("Multidrop_Data", [[mid, email, cur_val, gid, mi, mo, ml, ts, vendor_note]])
                        
                        st.session_state['temp_success_msg'] = f"Sukses! Data Tahap {cur_round} tersimpan."
                        st.cache_data.clear()
                        st.rerun()
                        
                        
if __name__ == "__main__":
    main()





