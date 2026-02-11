import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import os

# --- KONFIGURASI ---
SPREADSHEET_ID = "1j9GCq8Wwm-MM8hOamsH26qlmjNwuDBuEMnbw6ORzTQk"

# --- CUSTOM UI STYLE ---
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

def update_status_locked(ids_to_lock, status_value="Locked"):
    sh = connect_to_gsheet()
    if sh:
        try:
            ws = sh.worksheet("Price_Data")
            vals = ws.get_all_values()
            for i, row in enumerate(vals):
                if i==0: continue
                if row[0] in ids_to_lock: ws.update_cell(i+1, 3, status_value)
            get_data.clear()
            return True
        except: return False
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

# --- MAIN APP ---
st.set_page_config(page_title="TACO Procurement", layout="wide", page_icon="🚛")

def main():
    init_style()
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

        if role == 'admin': admin_dashboard()
        else: vendor_dashboard(user['email'])

# ================= ADMIN =================
def admin_dashboard():
    tabs = st.tabs(["📍Master Groups", "🛣️Master Routes", "🚛Master Units", "👥Users", "🔑Access Rights", "✅Monitoring & Approval"])
    
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
                    
                    if st.form_submit_button("Grant Access", type="primary"):
                        val = f"{val_period} {val_year}"
                        if selected_origins and val:
                            target_groups = df_g[
                                (df_g['load_type'] == sel_lt) & 
                                (df_g['origin'].isin(selected_origins))
                            ]
                            
                            if not target_groups.empty:
                                target_gids = target_groups['group_id'].unique().tolist()
                                sh = connect_to_gsheet()
                                if sh:
                                    ws = sh.worksheet("Access_Rights")
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
                                        ws.append_rows(new_rows_to_add)
                                        get_data.clear()
                                        msg = f"Sukses! {len(new_rows_to_add)} akses baru."
                                        if skipped_count > 0: msg += f" ({skipped_count} skip)."
                                        st.success(msg)
                                        time.sleep(1); st.rerun()
                                    else:
                                        st.warning(f"Semua data sudah ada ({skipped_count} skip). Tidak ada update.")
                            else:
                                st.warning("Data Group error.")
                        else:
                            st.warning("Pilih Origin dan isi Validity.")
                            
        st.dataframe(get_data("Access_Rights"), use_container_width=True)

    # --- TAB 6: MONITORING ---
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
                merged_pr = pd.merge(merged_pr, df_g[['group_id', 'route_group']], on='group_id', how='left')
            else: merged_pr['route_group'] = 'Unknown Group'
                
            merged_pr['route_group'] = merged_pr['route_group'].fillna('Unknown Group')
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
                
                with st.expander(f"{status_icon} - {vendor} ({validity}) - {g_name}"):
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

                    st.markdown("**C. Biaya Lain (Multidrop, Buruh)**")
                    if not df_md.empty:
                        sub_md = df_md[(df_md['vendor_email'] == vendor) & (df_md['validity'] == validity) & (df_md['group_id'] == g_id)]
                        if not sub_md.empty:
                            disp_md = sub_md[['inner_city_price', 'outer_city_price','labor_cost']].reset_index(drop=True)
                            disp_md.columns = ["Dalam Kota", "Luar Kota","Biaya Buruh"]
                            st.dataframe(disp_md, use_container_width=True, hide_index=True)
                        else: st.info("Data Multidrop tidak ditemukan.")
                    
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

# ================= VENDOR =================
def vendor_dashboard(email):
    step = st.session_state['vendor_step']
    
    if step == "dashboard":
        t1, t2 = st.tabs(["🛣️Pilih Rute & Isi Harga", "📋Isi Data Perusahaan"])
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
                    ppn = st.selectbox("PPN", ["11%", "1,1%","0%"]); pph = st.selectbox("PPh", ["Include", "Exclude"])
                    if st.form_submit_button("Simpan Data", type="primary"):
                        save_data("Vendor_Profile", [[email, ad, cp, ph, top, ppn, pph, datetime.now().strftime("%Y-%m-%d")]])
                        st.success("Saved")
        
        with t1:
            df_acc = get_data("Access_Rights")
            df_grps = get_data("Master_Groups")
            df_routes = get_data("Master_Routes")
            df_price = get_data("Price_Data")

            if df_acc.empty: st.warning("Belum ada akses."); return
            my_access = df_acc[df_acc['vendor_email'] == email]
            if my_access.empty: st.info("Anda belum diberikan akses ke project manapun."); return

            data_list = []
            for _, acc in my_access.iterrows():
                gid = acc['group_id']
                val = acc['validity']
                g_info = df_grps[df_grps['group_id'] == gid]
                if not g_info.empty:
                    row = g_info.iloc[0]
                    data_list.append({'validity': val, 'group_id': gid, 'origin': row.get('origin','-'), 'route_group': row.get('route_group','-'), 'load_type': row.get('load_type','-')})
            
            df_disp = pd.DataFrame(data_list)
            if df_disp.empty: st.warning("Konfigurasi Group tidak ditemukan."); return

            avail_validity = sorted(df_disp['validity'].unique().tolist())
            sel_val = st.selectbox("Pilih Periode / Validity:", avail_validity)
            df_view = df_disp[df_disp['validity'] == sel_val]
            
            if df_view.empty: st.info("Tidak ada rute."); return
            
            t_ftl, t_fcl = st.tabs(["🚛 FTL (Full Truck Load)", "🚢 FCL (Full Container Load)"])
            
            for t_code, t_ui in [('FTL', t_ftl), ('FCL', t_fcl)]:
                with t_ui:
                    df_sub = df_view[df_view['load_type'] == t_code]
                    if df_sub.empty: st.caption(f"Tidak ada akses {t_code}.")
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
                                    
                                    # CHECK STATUS PER VALIDITY
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
                                            if "Locked" in sub_p['status'].values:
                                                is_locked_btn = True
                                    
                                    c1, c2, c3, c4 = st.columns([3, 4, 2, 2])
                                    c1.write(f"**{grp_name}**")
                                    dests = r_data['kota_tujuan'].unique().tolist() if not r_data.empty else []
                                    # LOGIC PREVIEW KOTA (+X LAINNYA)
                                    if len(dests) > 5:
                                        preview_txt = f"{', '.join(dests[:5])}, +{len(dests)-5} kota lainnya"
                                    else:
                                        preview_txt = ", ".join(dests)
                                    c2.markdown(f"<span class='route-dest-list'>{preview_txt}</span>", unsafe_allow_html=True)

                                    
                                    # BUTTON LOGIC (DISABLED IF LOCKED)
                                    if is_locked_btn:
                                        c3.button("🔒Harga Dikunci", key=f"btn_lk_{gid}", disabled=True)
                                    else:
                                        if c3.button("📌Isi Harga", key=f"btn_{t_code}_{gid}", type="primary"):
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

    # --- INPUT PAGE ---
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
        
        if not target_gids: st.error("Data error."); return

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
                
                if my_r.empty or not u_types: st.warning("Data belum lengkap."); continue

                ex_price = {}; ex_spec = {}
                is_lock = False
                if not df_p.empty:
                    my_p = df_p[(df_p['vendor_email']==email) & (df_p['validity']==cur_val) & (df_p['route_id'].isin(my_r['route_id']))]
                    if not my_p.empty:
                        if "Locked" in my_p['status'].values: is_lock = True
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
                                "Lead Time (Hari)": 0 # Lead time
                            }
                            # Harga
                            for u in u_types:
                                rd[f"Harga {u}"] = ex_price.get((rid, u), 0)
                            
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





