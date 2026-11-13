"""
import_tab.py - Tab admin "📤 Import Excel" (1 file xlsx = 1 vendor).

Cara pakai di app.py (lihat juga INTEGRASI.md):
    from import_tab import render_import_tab
    ...
    if sel_tab == "📤 Import Excel":
        render_import_tab(get_data, save_data, connect_to_gsheet)
"""
import re
import secrets
import hashlib
from datetime import datetime, timedelta

import pandas as pd

from excel_import import (
    parse_vendor_workbook, match_sheet_rows, guess_from_list, _norm,
)

NEW_VENDOR = "➕ Vendor baru..."


# ----------------------------------------------------------------------------
# Logic murni (tanpa Streamlit) -> gampang dites
# ----------------------------------------------------------------------------
def build_payload(frames, email, validity, rnd, status, ts, route_to_gid, units_by_gid, existing_md=None,
                  with_shipping=False):
    """
    frames: list of dict {df: DataFrame hasil editor, specs: {unit:{w,c}}, units: [nama unit di file]}
    with_shipping: True kalau sheet Price_Data sudah punya kolom 'pelayaran' & 'frekuensi' (di kolom 14-15,
                   setelah 'round'). Kalau False, info itu dititipkan di kolom keterangan.
    Return dict: price_rows, md_rows, gids, warnings, n_routes
    """
    existing_md = existing_md or {}
    price_rows, warnings, labor = [], [], {}
    routes_ok = 0
    for fr in frames:
        df, specs = fr["df"], fr["specs"]
        for _, row in df.iterrows():
            if not row["Import"] or not str(row["Route ID"]).strip():
                continue
            rid = str(row["Route ID"]).split("|")[0].strip()
            gid = route_to_gid.get(rid)
            if not gid:
                warnings.append(f"Route {rid} tidak punya group, dilewati")
                continue
            line = str(row.get("Pelayaran", "")).strip()
            freq = str(row.get("Frek (x/mgg)", "")).strip()
            if with_shipping:
                ket = "-"
            else:
                ket = []
                if line:
                    ket.append(f"Pelayaran: {line}")
                if freq:
                    ket.append(f"Frek: {freq}x/minggu")
                ket = "; ".join(ket) or "-"
            lt = int(row.get("Lead Time (hari)", 0) or 0)
            added = False
            for unit in fr["units"]:
                price = int(pd.to_numeric(row.get(unit, 0), errors="coerce") or 0)
                if price <= 0:
                    continue
                canon = units_by_gid.get(gid, {}).get(_norm(unit))
                if not canon:
                    warnings.append(f"{rid}: unit '{unit}' belum ada di Master_Units, dilewati")
                    continue
                sp = specs.get(unit, {})
                w = str(sp.get("w") or "")
                c = str(sp.get("c") or "")
                tid = f"{email}_{validity}_{rid}_{canon}_{rnd}".replace(" ", "")
                prow = [tid, email, status, validity, rid, canon, lt, price, w, c, ket, ts, int(rnd)]
                if with_shipping:
                    prow += [line, freq]
                price_rows.append(prow)
                added = True
            if added:
                routes_ok += 1
                for n in (20, 40):
                    v = int(pd.to_numeric(row.get(f"Buruh {n}ft", 0), errors="coerce") or 0)
                    if v > 0:
                        labor.setdefault(gid, {}).setdefault(n, v)  # nilai pertama per group

    md_rows = []
    for gid in sorted({route_to_gid[r[4]] for r in price_rows}):
        mid = f"M_{email}_{gid}_{str(validity).replace(' ', '')}_{int(rnd)}"
        mi, mo = existing_md.get(mid, (0, 0))  # jaga multidrop lama kalau sudah ada
        lab = labor.get(gid, {})
        note = f"Biaya bongkar buruh plafon 40ft: Rp {lab[40]:,}".replace(",", ".") if lab.get(40) else ""
        md_rows.append([mid, email, validity, gid, mi, mo, lab.get(20, 0), ts, note])

    return {"price_rows": price_rows, "md_rows": md_rows,
            "gids": sorted({route_to_gid[r[4]] for r in price_rows}),
            "warnings": warnings, "n_routes": routes_ok}


# ----------------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------------
def _backup(sh, names, st):
    """Salin sheet ke tab 'bak_<nama>_<tgl>' sebelum import (jaga-jaga save_data yang ws.clear())."""
    stamp = datetime.now().strftime("%m%d_%H%M")
    for n in names:
        try:
            sh.worksheet(n).duplicate(new_sheet_name=f"bak_{n}_{stamp}"[:99])
        except Exception as e:
            st.warning(f"Backup {n} gagal ({e}). Import tetap dilanjutkan.")


def render_import_tab(get_data, save_data, connect_to_gsheet):
    import streamlit as st

    st.subheader("📤 Import Penawaran dari Excel")
    st.caption("1 file xlsx = 1 vendor. Setiap sheet = 1 origin. Cek preview dulu sebelum disimpan.")

    df_u, df_g = get_data("Users"), get_data("Master_Groups")
    df_r, df_units = get_data("Master_Routes"), get_data("Master_Units")
    df_prof, df_md = get_data("Vendor_Profile"), get_data("Multidrop_Data")
    df_p = get_data("Price_Data")
    if df_g.empty or df_r.empty or df_units.empty:
        st.error("Master Groups/Routes/Units masih kosong.")
        return

    up = st.file_uploader("Upload template penawaran vendor (.xlsx)", type=["xlsx"], key="imp_file")
    if not up:
        return
    data = up.getvalue()
    fkey = hashlib.md5(data).hexdigest()[:10]
    if st.session_state.get("imp_key") != fkey:
        try:
            st.session_state["imp_parsed"] = parse_vendor_workbook(data)
            st.session_state["imp_key"] = fkey
        except Exception as e:
            st.error(f"File tidak bisa dibaca: {e}")
            return
    parsed = st.session_state["imp_parsed"]
    ven = parsed["vendor"]

    # ---------------- 1. Vendor & parameter ----------------
    with st.container(border=True):
        st.markdown("#### 1. Vendor & Periode")
        vendors = df_u[df_u["role"] == "vendor"] if not df_u.empty else pd.DataFrame(columns=["email", "vendor_name"])
        name_by_email = dict(zip(vendors["email"], vendors["vendor_name"]))
        guess = guess_from_list(ven.get("name", ""), list(name_by_email.values()))
        opts = [NEW_VENDOR] + sorted(name_by_email.keys(), key=lambda e: str(name_by_email[e]).lower())
        default_idx = 0
        if guess:
            default_idx = opts.index(next(e for e, n in name_by_email.items() if n == guess))
        c1, c2 = st.columns(2)
        sel = c1.selectbox("Vendor di database", opts, index=default_idx, key=f"imp_ven_{fkey}",
                           format_func=lambda e: e if e == NEW_VENDOR else f"{name_by_email[e]} - {e}")
        c1.caption(f"Nama di file: **{ven.get('name') or '(kosong)'}**")
        if sel == NEW_VENDOR:
            new_email = c2.text_input("Email vendor baru", key=f"imp_newem_{fkey}").strip().lower()
            new_name = c2.text_input("Nama PT", value=ven.get("name", ""), key=f"imp_newnm_{fkey}").strip()
            email, vname = new_email, new_name
        else:
            email, vname = sel, name_by_email[sel]

        c3, c4, c5, c6 = st.columns(4)
        validity = c3.text_input("Periode (validity)", value=ven.get("validity", ""), key=f"imp_val_{fkey}").strip()
        existing_vals = sorted(df_p["validity"].dropna().unique().tolist()) if not df_p.empty else []
        if existing_vals:
            c3.caption("Periode yang sudah ada: " + " | ".join(existing_vals))
        rnd = c4.selectbox("Tahap", ["1", "2", "3"], key=f"imp_rnd_{fkey}")
        load_type = c5.selectbox("Tipe muatan", ["FCL", "FTL", "FTL Plafon"], key=f"imp_lt_{fkey}")
        status = c6.selectbox("Status data", ["Open", "Locked"], key=f"imp_st_{fkey}")
        opt_prof = st.checkbox("Perbarui profil vendor (alamat, PIC, telp, TOP, PPN, PPh) dari file", value=True, key=f"imp_prof_{fkey}")
        opt_acc = st.checkbox("Catat juga ke Access_Rights (supaya muncul di Submit Monitor)", value=True, key=f"imp_acc_{fkey}")

    # ---------------- 2. Preview per sheet ----------------
    st.markdown("#### 2. Preview per Origin")
    origins = sorted(df_g[df_g["load_type"] == load_type]["origin"].dropna().unique().tolist())
    if not origins:
        st.error(f"Belum ada origin untuk tipe {load_type} di Master_Groups.")
        return

    frames, all_ok = [], True
    route_to_gid = dict(zip(df_r["route_id"].astype(str).str.strip(), df_r["group_id"].astype(str).str.strip()))
    units_by_gid = {}
    for gid, sub in df_units.assign(group_id=df_units["group_id"].astype(str).str.strip()).groupby("group_id"):
        units_by_gid[gid] = {_norm(u): u for u in sub["unit_type"]}

    for sh in parsed["sheets"]:
        title = f"📍 {sh['sheet']} — {len(sh['rows'])} rute terisi, {sh['skipped']} dilewati (kosong)"
        with st.expander(title, expanded=bool(sh["rows"])):
            if not sh["rows"]:
                st.info("Tidak ada harga yang diisi di sheet ini.")
                continue
            g_org = guess_from_list(sh["sheet"], origins)
            origin = st.selectbox("Origin di database", origins, key=f"imp_org_{fkey}_{sh['sheet']}",
                                  index=origins.index(g_org) if g_org else 0)
            matches, cand, _ = match_sheet_rows(sh["rows"], origin, load_type, df_r, df_g, df_units)
            route_opts = [""] + [f"{x.route_id} | {x.kota_asal} → {x.kota_tujuan}" for x in cand.itertuples()]
            opt_by_id = {o.split(" | ")[0]: o for o in route_opts if o}

            recs = []
            for row, m in zip(sh["rows"], matches):
                rec = {"Import": m["import_default"], "Baris": row["excel_row"],
                       "Asal": row["asal"], "Tujuan": row["tujuan"],
                       "Route ID": opt_by_id.get(m["route_id"], ""), "Status": m["status"]}
                for u in sh["units"]:
                    rec[u] = row["prices"].get(u, 0)
                rec.update({"Lead Time (hari)": row["lead_days"], "Pelayaran": row["line"], "Frek (x/mgg)": row["freq"],
                            "Buruh 20ft": row["labor"].get(20, 0), "Buruh 40ft": row["labor"].get(40, 0)})
                recs.append(rec)
            df_prev = pd.DataFrame(recs)
            if df_prev[["Buruh 20ft", "Buruh 40ft"]].sum().sum() == 0:  # sheet tanpa kolom buruh: sembunyikan
                df_prev = df_prev.drop(columns=["Buruh 20ft", "Buruh 40ft"])
            cfg = {
                "Import": st.column_config.CheckboxColumn(),
                "Baris": st.column_config.NumberColumn(disabled=True, width="small"),
                "Asal": st.column_config.TextColumn(disabled=True),
                "Tujuan": st.column_config.TextColumn(disabled=True),
                "Route ID": st.column_config.SelectboxColumn("Rute Master", options=route_opts, width="medium"),
                "Status": st.column_config.TextColumn(disabled=True),
            }
            for u in sh["units"]:
                cfg[u] = st.column_config.NumberColumn(format="Rp %d", min_value=0)
            edited = st.data_editor(df_prev, hide_index=True, use_container_width=True, column_config=cfg,
                                    key=f"imp_ed_{fkey}_{sh['sheet']}_{origin}_{load_type}")
            bad = edited[(edited["Import"]) & (edited["Route ID"].astype(str).str.strip() == "")]
            if len(bad):
                all_ok = False
                st.warning(f"{len(bad)} baris dicentang tapi Rute Master belum dipilih.")
            for col in ("Buruh 20ft", "Buruh 40ft"):
                if col not in edited.columns:
                    edited[col] = 0
            frames.append({"df": edited, "specs": sh["specs"], "units": sh["units"]})

    if not frames:
        st.warning("Tidak ada sheet berisi harga.")
        return

    # ---------------- 3. Simpan ----------------
    existing_md = {}
    if not df_md.empty and "id_multidrop" in df_md.columns:
        for _, r in df_md.iterrows():
            existing_md[str(r["id_multidrop"])] = (r.get("inner_city_price", 0) or 0, r.get("outer_city_price", 0) or 0)
    ts = (datetime.utcnow() + timedelta(hours=7)).strftime("%Y-%m-%d %H:%M:%S")
    with_shipping = {"pelayaran", "frekuensi"} <= set(df_p.columns)
    payload = build_payload(frames, email, validity, rnd, status, ts, route_to_gid, units_by_gid, existing_md,
                            with_shipping=with_shipping)
    if not with_shipping:
        st.info("Kolom `pelayaran` & `frekuensi` belum ada di sheet Price_Data, jadi info ini dititipkan di kolom keterangan.")

    with st.container(border=True):
        st.markdown("#### 3. Simpan")
        m1, m2, m3 = st.columns(3)
        m1.metric("Rute akan diimport", payload["n_routes"])
        m2.metric("Baris harga (rute × unit)", len(payload["price_rows"]))
        m3.metric("Group terdampak", len(payload["gids"]))
        for w in payload["warnings"][:10]:
            st.caption("⚠️ " + w)
        st.caption("Catatan: template tidak punya kolom Multidrop; hanya Biaya Bongkar Buruh Plafon (→ Biaya Buruh, nilai 20ft; "
                   "40ft masuk catatan). Nilai multidrop yang sudah ada di database dipertahankan.")

        problems = []
        if not email or "@" not in email:
            problems.append("Email vendor belum valid.")
        if sel == NEW_VENDOR and not vname:
            problems.append("Nama PT vendor baru belum diisi.")
        if not validity:
            problems.append("Periode belum diisi.")
        if not all_ok:
            problems.append("Ada baris tercentang tanpa Rute Master.")
        if not payload["price_rows"]:
            problems.append("Belum ada baris harga yang bisa disimpan.")
        for p in problems:
            st.error(p)

        if st.button("💾 Simpan ke Database", type="primary", use_container_width=True, disabled=bool(problems),
                     key=f"imp_save_{fkey}"):
            with st.spinner("Backup & menyimpan..."):
                sh_obj = connect_to_gsheet()
                if sh_obj:
                    _backup(sh_obj, ["Price_Data", "Multidrop_Data"], st)
                ok = True
                if sel == NEW_VENDOR:
                    # password acak: akun ini hanya penanda vendor, bukan untuk login
                    ok &= bool(save_data("Users", [[email, secrets.token_urlsafe(16), "vendor", vname]]))
                if opt_prof:
                    old = df_prof[df_prof["email"] == email].iloc[-1].to_dict() if (not df_prof.empty and (df_prof["email"] == email).any()) else {}
                    pick = lambda new, key: new or old.get(key, "")
                    ok &= bool(save_data("Vendor_Profile", [[
                        email, pick(ven.get("address"), "address"), pick(ven.get("contact_person"), "contact_person"),
                        pick(ven.get("phone"), "phone"), pick(ven.get("top_norm"), "top"),
                        pick(ven.get("ppn"), "ppn"), pick(ven.get("pph"), "pph"),
                        datetime.now().strftime("%Y-%m-%d"), old.get("direktur", ""), old.get("jabatan", "Direktur Utama"),
                    ]]))
                ok &= bool(save_data("Price_Data", payload["price_rows"]))
                ok &= bool(save_data("Multidrop_Data", payload["md_rows"]))
                if opt_acc and sh_obj:
                    try:
                        ws = sh_obj.worksheet("Access_Rights")
                        have = {f"{r[0]}_{r[1]}_{r[2]}_{r[4] if len(r) > 4 else '1'}".lower() for r in ws.get_all_values()[1:] if len(r) >= 3}
                        new = [[email, validity, g, "Active", rnd] for g in payload["gids"]
                               if f"{email}_{validity}_{g}_{rnd}".lower() not in have]
                        if new:
                            ws.append_rows(new)
                        get_data.clear()
                    except Exception as e:
                        st.warning(f"Access_Rights tidak terupdate: {e}")
            if ok:
                st.success(f"Berhasil! {len(payload['price_rows'])} baris harga dari {vname} tersimpan (Tahap {rnd}, {validity}).")
            else:
                st.error("Sebagian data gagal tersimpan. Cek pesan error di atas lalu coba lagi (aman diulang, data akan menimpa ID yang sama).")
