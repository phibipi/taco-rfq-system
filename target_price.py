"""
target_price.py - hitung Target Price dengan membandingkan periode ini vs periode SEBELUMNYA.

Aturan (sama seperti sebelumnya, tapi periode sebelumnya sekarang dicari berdasarkan
URUTAN WAKTU, bukan tebakan tahun):
  base A = harga termurah periode INI, Tahap 1 (semua vendor, rute + unit yang sama)
  base B = harga termurah periode SEBELUMNYA, Tahap 2 (harga final), rute + unit yang sama
  - B ada dan A ada : B < A -> target = B x 0.95 (turun 5%)   selain itu -> A x 0.85 (turun 15%)
  - hanya B ada     : B x 0.95
  - hanya A ada     : A x 0.85
  - tidak ada       : 0

Periode sebelumnya = periode yang PUNYA harga untuk rute + unit itu dan mulai paling dekat
sebelum periode ini. Lompat 1 bulan / kontrak 3 bulan / kontrak 12 bulan tidak masalah.
Format periode yang dikenali (dengan/tanpa spasi):
  '1Desember2026-28Februari2027', 'Agustus-Oktober2026', 'Januari-Juni2026',
  'September2026-Agustus2027', 'Juli 2026 - Desember 2026'
"""
import re
from functools import lru_cache

import pandas as pd

_MONTHS = {
    "januari": 1, "jan": 1, "februari": 2, "pebruari": 2, "feb": 2, "maret": 3, "mar": 3,
    "april": 4, "apr": 4, "mei": 5, "juni": 6, "jun": 6, "juli": 7, "jul": 7,
    "agustus": 8, "agu": 8, "agt": 8, "aug": 8, "september": 9, "sept": 9, "sep": 9,
    "oktober": 10, "okt": 10, "oct": 10, "november": 11, "nopember": 11, "nov": 11,
    "desember": 12, "des": 12, "dec": 12,
}
_MONTH_RE = re.compile(
    r"(\d{1,2})?(" + "|".join(sorted(_MONTHS, key=len, reverse=True)) + r")(\d{4})?"
)


def _cell(v):
    t = "" if v is None else str(v).strip()
    return "-" if t.lower() in ("", "nan", "none") else t


def _key(v):
    return re.sub(r"\s+", "", str(v)).lower().strip()


def pretty_period(v):
    s = str(v).strip()
    s = re.sub(r"(?<=\d)(?=[A-Za-z])|(?<=[A-Za-z])(?=\d)", " ", s)
    s = re.sub(r"\s*-\s*", " - ", s)
    return re.sub(r"\s+", " ", s)


@lru_cache(maxsize=1024)
def parse_period(v):
    """'Agustus-Oktober2026' -> (2026*12+8, 2026*12+10). Return None kalau tidak terbaca."""
    s = re.sub(r"[^a-z0-9]", "", str(v).lower())
    toks = [(_MONTHS[m.group(2)], int(m.group(3)) if m.group(3) else None) for m in _MONTH_RE.finditer(s)]
    if not toks:
        return None
    end_m, end_y = toks[-1]
    if end_y is None:
        years = [y for _, y in toks if y] or [int(y) for y in re.findall(r"20\d{2}", s)]
        if not years:
            return None
        end_y = years[-1]
    start_m, start_y = toks[0]
    if start_y is None:
        start_y = end_y if (len(toks) == 1 or start_m <= end_m) else end_y - 1
    start_i, end_i = start_y * 12 + start_m, end_y * 12 + end_m
    if end_i < start_i:  # mis. 'Juli 2026 - Juni 2026' -> anggap lewat tahun
        end_i += 12
    return start_i, end_i


def get_target_info(df_all, route_id, unit_type, cur_validity, prev_round=2, prev_round_strict=False):
    """
    Return dict: target, base, base_source, pct, min_curr, min_prev, prev_period, prev_round,
                 pelayaran, frekuensi  (milik penawaran termurah yang dijadikan base; '-' bila kolomnya belum ada)
    prev_round_strict=False -> kalau periode sebelumnya belum punya Tahap 2 untuk rute+unit itu,
                               pakai tahap tertinggi yang ada. True -> hanya Tahap 2, kalau tidak ada dianggap kosong.
    """
    info = {"target": 0, "base": 0, "base_source": "-", "pct": 0, "min_curr": 0, "min_prev": 0,
            "prev_period": "-", "prev_round": "-", "pelayaran": "-", "frekuensi": "-"}
    if df_all is None or df_all.empty or "price" not in df_all.columns:
        return info

    rid = str(route_id).strip()
    d = df_all[df_all["route_id"].astype(str).str.strip() == rid]
    if d.empty:
        return info
    d = d.copy()
    d["price"] = pd.to_numeric(d["price"], errors="coerce").fillna(0)
    d = d[(d["price"] > 0) & (d["unit_type"].astype(str).str.strip().str.lower() == str(unit_type).strip().lower())]
    if d.empty:
        return info
    d["val_key"] = d["validity"].map(_key)
    d["rnd"] = pd.to_numeric(d["round"], errors="coerce").fillna(1).astype(int)

    cur_key = _key(cur_validity)

    # A. termurah periode ini, Tahap 1
    cur = d[(d["val_key"] == cur_key) & (d["rnd"] == 1)]
    min_curr = float(cur["price"].min()) if not cur.empty else 0.0
    row_cur = cur.loc[cur["price"].idxmin()] if not cur.empty else None

    # B. periode sebelumnya (urutan waktu), harga Tahap 2
    min_prev, prev_label, prev_rnd = 0.0, "-", "-"
    row_prev = None
    p_cur = parse_period(cur_key)
    if p_cur:
        best = None
        for raw in d["validity"].astype(str).unique():
            if _key(raw) == cur_key:
                continue
            p = parse_period(raw)
            if p and p[0] < p_cur[0] and (best is None or p > best[0]):
                best = (p, raw)
        if best:
            prev_rows = d[d["val_key"] == _key(best[1])]
            rounds = set(prev_rows["rnd"])
            use = prev_round if prev_round in rounds else (None if prev_round_strict else max(rounds))
            if use is not None:
                sel_rows = prev_rows[prev_rows["rnd"] == use]
                min_prev = float(sel_rows["price"].min())
                row_prev = sel_rows.loc[sel_rows["price"].idxmin()]
                prev_label, prev_rnd = pretty_period(best[1]), use

    # Keputusan
    if min_prev > 0 and min_curr > 0:
        if min_prev < min_curr:
            base, factor, src = min_prev, 0.95, "prev"
        else:
            base, factor, src = min_curr, 0.85, "cur"
    elif min_prev > 0:
        base, factor, src = min_prev, 0.95, "prev"
    elif min_curr > 0:
        base, factor, src = min_curr, 0.85, "cur"
    else:
        return {**info, "min_curr": min_curr, "min_prev": min_prev, "prev_period": prev_label, "prev_round": prev_rnd}

    base_row = row_prev if src == "prev" else row_cur
    info.update({
        "pelayaran": _cell(base_row.get("pelayaran")) if base_row is not None else "-",
        "frekuensi": _cell(base_row.get("frekuensi")) if base_row is not None else "-",
        "target": int(base * factor + 1e-9),
        "base": int(base),
        "pct": int(round((1 - factor) * 100)),
        "min_curr": int(min_curr), "min_prev": int(min_prev),
        "prev_period": prev_label, "prev_round": prev_rnd,
        "base_source": ("Periode ini (Tahap 1)" if src == "cur"
                        else f"Periode sebelumnya: {prev_label} (Tahap {prev_rnd})"),
    })
    return info


def get_target_price(df_all, route_id, unit_type, cur_validity):
    """Kompatibel dengan pemanggilan lama: hanya mengembalikan angka target."""
    return get_target_info(df_all, route_id, unit_type, cur_validity)["target"]
