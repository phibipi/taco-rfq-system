"""
cost_saving.py - Rate periode sebelumnya (Tahap 2) + sheet Potential Cost Saving (khusus FCL).

Dipakai oleh Master Summary di app.py.
  add_prev_rate()      -> tambah kolom prev_rate / prev_period ke summary_df
  build_saving_df()    -> tabel potential saving 3 bulan (1 baris per rute + unit)
  write_saving_sheet() -> tulis ke sheet 'Potential Cost Saving' (selisih & saving pakai rumus Excel)
"""
import re
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from target_price import get_target_info

SAVING_SHEET = "Potential Cost Saving"
MONTHS_AHEAD = 3


def fmt_rp(x):
    try:
        x = int(x)
    except (TypeError, ValueError):
        return "Rp 0"
    return f"Rp {x:,}".replace(",", ".") if x > 0 else "Rp 0"


def parse_volume(raw):
    """'4-6 unit / bulan' -> 6 ; '1-2 rit per bulan' -> 2 ; '1 unit / bulan' -> 1 ; kosong/'-' -> 0."""
    nums = [int(n) for n in re.findall(r"\d+", str(raw or ""))]
    return max(nums) if nums else 0


def add_prev_rate(summary_df, df_p, cur_validity):
    """
    Rate termurah (urutan 1) periode SEBELUMNYA, khusus Tahap 2, per rute + unit.
    Tidak ada periode sebelumnya / tidak ada Tahap 2 -> 0.
    """
    cache = {}

    def one(rid, unit):
        k = (str(rid).strip(), str(unit).strip().lower())
        if k not in cache:
            ti = get_target_info(df_p, rid, unit, cur_validity, prev_round=2, prev_round_strict=True)
            cache[k] = (int(ti["min_prev"]), ti["prev_period"] if ti["min_prev"] > 0 else "-")
        return cache[k]

    res = [one(r, u) for r, u in zip(summary_df["route_id"], summary_df["unit_type"])]
    summary_df["prev_rate"] = [x[0] for x in res]
    summary_df["prev_period"] = [x[1] for x in res]
    return summary_df


def build_saving_df(summary_df, route_df):
    """
    summary_df : sudah punya route_id, unit_type, price (0 = belum ada), vendor_name, prev_rate, prev_period,
                 dan terurut termurah dulu per rute+unit.
    route_df   : Master_Routes (route_id, keterangan).
    """
    keys = ["route_id", "unit_type"]
    s = summary_df[summary_df["unit_type"].notna()]
    uni = s.drop_duplicates(keys)[["route_id", "origin", "kota_asal", "kota_tujuan", "unit_type",
                                   "prev_rate", "prev_period"]]
    best = (s[s["price"] > 0].sort_values("price").drop_duplicates(keys)
            [["route_id", "unit_type", "price", "vendor_name"]]
            .rename(columns={"price": "cur_rate", "vendor_name": "cur_vendor"}))
    df = uni.merge(best, on=keys, how="left")
    df["cur_rate"] = df["cur_rate"].fillna(0).astype(int)
    df["cur_vendor"] = df["cur_vendor"].fillna("-")

    ket = {}
    if "keterangan" in route_df.columns:
        ket = dict(zip(route_df["route_id"].astype(str).str.strip(), route_df["keterangan"]))
    df["ket"] = df["route_id"].astype(str).str.strip().map(ket).fillna("-")
    df["vol"] = df["ket"].map(parse_volume)

    def status(r):
        if r["vol"] <= 0:
            return "Volume kosong di Master_Routes"
        if r["prev_rate"] <= 0:
            return "Tidak ada rate periode sebelumnya"
        if r["cur_rate"] <= 0:
            return "Belum ada penawaran periode ini"
        return "OK"
    df["status"] = df.apply(status, axis=1)
    return df.reset_index(drop=True)


def write_saving_sheet(writer, df_sav, sel_round, cur_validity):
    cols = ["Origin", "Kota Asal", "Kota Tujuan", "Unit", "Keterangan Volume (Master_Routes)",
            "Volume / Bulan (dipakai)", f"Volume {MONTHS_AHEAD} Bulan",
            "Rate SK Sebelumnya (Tahap 2)", "Periode Sebelumnya",
            "Vendor Urutan 1 (Periode Ini)", f"Rate Urutan 1 Periode Ini (Tahap {sel_round})",
            "Selisih per Unit", f"Potential Saving {MONTHS_AHEAD} Bulan", "Status"]
    out = pd.DataFrame({
        cols[0]: df_sav["origin"], cols[1]: df_sav["kota_asal"], cols[2]: df_sav["kota_tujuan"],
        cols[3]: df_sav["unit_type"], cols[4]: df_sav["ket"], cols[5]: df_sav["vol"],
        cols[6]: 0, cols[7]: df_sav["prev_rate"], cols[8]: df_sav["prev_period"],
        cols[9]: df_sav["cur_vendor"], cols[10]: df_sav["cur_rate"], cols[11]: 0, cols[12]: 0,
        cols[13]: df_sav["status"],
    })
    out.to_excel(writer, index=False, sheet_name=SAVING_SHEET)
    ws = writer.sheets[SAVING_SHEET]
    n = len(out)

    for i in range(n):
        r = i + 2
        ws[f"G{r}"] = f"=F{r}*{MONTHS_AHEAD}"
        ws[f"L{r}"] = f"=IF(AND(H{r}>0,K{r}>0),H{r}-K{r},0)"
        ws[f"M{r}"] = f"=L{r}*G{r}"
        for c in "HKLM":
            ws[f"{c}{r}"].number_format = '#,##0;[Red]-#,##0'

    t = n + 2
    ws[f"A{t}"] = "TOTAL"
    ws[f"M{t}"] = f"=SUM(M2:M{n + 1})"
    ws[f"M{t}"].number_format = '#,##0;[Red]-#,##0'
    for c in range(1, len(cols) + 1):
        ws.cell(t, c).font = Font(bold=True)
    notes = [
        f"Periode ini: {cur_validity} (Tahap {sel_round}). Saving = (Rate SK Sebelumnya - Rate urutan 1 periode ini) x volume {MONTHS_AHEAD} bulan.",
        "Volume diambil dari Master_Routes kolom keterangan; kalau berupa rentang (mis. 4-6) dipakai angka terbesar.",
        "Hanya baris berstatus OK yang menghasilkan saving. Nilai negatif (merah) = harga naik dibanding SK sebelumnya.",
        "Volume berlaku per rute; bila satu rute punya >1 unit (mis. 20ft & 40ft) volume yang sama terhitung di tiap unit.",
    ]
    for i, txt in enumerate(notes):
        ws[f"A{t + 2 + i}"] = txt
        ws[f"A{t + 2 + i}"].font = Font(italic=True, color="666666")

    for c in range(1, len(cols) + 1):
        h = ws.cell(1, c)
        h.font = Font(bold=True, color="FFFFFF")
        h.fill = PatternFill("solid", fgColor="1F4E78")
        h.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(c)].width = 16
    for c, w in {"A": 14, "B": 16, "C": 24, "E": 24, "I": 24, "J": 26, "N": 34}.items():
        ws.column_dimensions[c].width = w
    ws.freeze_panes = "A2"
    return out
