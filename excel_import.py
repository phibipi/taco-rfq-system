"""
excel_import.py  -  parser template penawaran vendor (xlsx) -> struktur data siap import.

Tanpa dependensi Streamlit, jadi bisa dites terpisah.
Satu file = satu vendor; satu sheet = satu origin (sheet FAQ dilewati otomatis).

Layout tiap sheet BEDA-BEDA (kolom Pelayaran ada di H / F / J, dst), jadi kolom
dicari lewat TEKS HEADER, bukan lewat alamat sel tetap.
"""
import io
import re
import difflib
from collections import Counter

import openpyxl
from openpyxl.utils import column_index_from_string


# ----------------------------------------------------------------------------
# Helper kecil
# ----------------------------------------------------------------------------
def _norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def parse_rupiah(v):
    """'Rp 1.500.000' / 1500000 / '1.500.000,00' -> 1500000 (int). Kosong -> 0."""
    if v is None:
        return 0
    if isinstance(v, (int, float)):
        return int(round(v)) if v > 0 else 0
    s = re.sub(r"rp|idr|\s", "", str(v).lower())
    if s in ("", "-", "nan", "none"):
        return 0
    if "." in s and re.search(r",\d{1,2}$", s):  # desimal gaya Indonesia
        s = s.rsplit(",", 1)[0]
    s = re.sub(r"[.,]", "", s)
    m = re.search(r"\d+", s)
    return int(m.group()) if m else 0


def parse_days(v):
    """'14-16 hari' -> 16 (ambil angka terbesar). Kosong -> 0."""
    if v is None:
        return 0
    if isinstance(v, (int, float)):
        return int(v)
    nums = [int(x) for x in re.findall(r"\d+", str(v))]
    return max(nums) if nums else 0


def clean_validity(raw):
    """'1 Desember 2026 -  28 Februari 2027 (3 bulan)' -> '1 Desember 2026 - 28 Februari 2027'."""
    s = re.sub(r"\(.*?\)", "", str(raw or ""))
    s = re.sub(r"\s*[-–—]\s*", " - ", s)
    return re.sub(r"\s+", " ", s).strip(" :")


def norm_top(raw):
    m = re.search(r"\d+", str(raw or ""))
    if not m:
        return ""
    n = int(m.group())
    return "7 hari" if n == 7 else f"{n} Hari"  # persis opsi selectbox di form profil lama


def norm_ppn(raw):
    s = str(raw or "").lower().replace(" ", "")
    if not s or s in ("nan", "none"):
        return ""
    if "1.1" in s or "1,1" in s:
        return "1,1%"
    if "11" in s:
        return "11%"
    if "non" in s or s.startswith("0"):
        return "0%"
    return ""


def norm_pph(raw):
    s = str(raw or "").lower()
    if "excl" in s:
        return "Exclude"
    if "incl" in s:
        return "Include"
    return ""


def _mode(values):
    values = [v for v in values if v]
    return Counter(values).most_common(1)[0][0] if values else ""


def guess_from_list(text, options):
    """Tebak item terdekat dari daftar (nama sheet -> origin, nama vendor -> vendor)."""
    options = [o for o in options if str(o).strip()]
    n = _norm(text)
    if not n:
        return None
    for o in options:
        if _norm(o) == n:
            return o
    for o in options:
        if _norm(o) and (_norm(o) in n or n in _norm(o)):
            return o
    m = difflib.get_close_matches(str(text), [str(o) for o in options], 1, 0.6)
    return m[0] if m else None


# ----------------------------------------------------------------------------
# Baca workbook (nilai cache + fallback formula =Sheet!Cell)
# ----------------------------------------------------------------------------
class _Book:
    def __init__(self, data: bytes):
        self.wf = openpyxl.load_workbook(io.BytesIO(data))
        self.wv = openpyxl.load_workbook(io.BytesIO(data), data_only=True)

    def val(self, sheet, row, col, _depth=0):
        v = self.wv[sheet].cell(row, col).value
        if v is None and _depth < 3:
            f = self.wf[sheet].cell(row, col).value
            if isinstance(f, str) and f.startswith("="):
                m = re.fullmatch(r"=\s*'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+)", f.strip())
                if m and m.group(1) in self.wf.sheetnames:
                    return self.val(
                        m.group(1), int(m.group(3)), column_index_from_string(m.group(2)), _depth + 1
                    )
        return v

    def txt(self, sheet, row, col):
        v = self.val(sheet, row, col)
        return "" if v is None else str(v).strip()


HEADER_LABELS = {
    "nama vendor": "name",
    "alamat vendor": "address",
    "contact person": "contact_person",
    "no. telp": "phone",
    "no telp": "phone",
    "term of payment": "top",
    "masa berlaku": "validity_raw",
}


def _parse_header_block(bk: _Book, sheet):
    """Blok identitas vendor di baris 1-14 (label di kolom B, isi di kolom C)."""
    info = {}
    ws = bk.wv[sheet]
    for r in range(1, 15):
        for c in range(1, 5):
            label = bk.txt(sheet, r, c).lower().rstrip(": ")
            for k, field in HEADER_LABELS.items():
                if label.startswith(k):
                    for cc in range(c + 1, min(c + 4, ws.max_column) + 1):
                        v = bk.txt(sheet, r, cc).lstrip(": ").strip()
                        if v and not v.startswith("="):
                            info[field] = v
                            break
    return info


def _find_header_row(bk, sheet):
    ws = bk.wv[sheet]
    for r in range(1, min(ws.max_row, 40) + 1):
        for c in range(1, 7):
            if "daerah asal" in bk.txt(sheet, r, c).lower():
                return r
    return None


def _parse_sheet(bk: _Book, sheet):
    ws = bk.wv[sheet]
    hr = _find_header_row(bk, sheet)
    if hr is None:
        return None

    cols = {"price": {}, "labor": {}}
    for c in range(1, ws.max_column + 1):
        top = bk.txt(sheet, hr, c).lower()
        sub = bk.txt(sheet, hr + 1, c).lower()
        if "daerah asal" in top:
            cols["asal"] = c
        elif "daerah tujuan" in top:
            cols["tujuan"] = c
        elif "pelayaran" in top:
            cols["line"] = c
        elif "frekuensi" in top:
            cols["freq"] = c
        elif "leadtime" in top or "lead time" in top:
            cols["lead"] = c
        elif "ppn" in top:
            cols["ppn"] = c
        elif "pph" in top:
            cols["pph"] = c
        m = re.search(r"container\s*(\d+)\s*ft", sub)
        if m and "bongkar" not in sub:
            cols["price"][f"Container {m.group(1)}ft"] = c
        elif "bongkar" in sub and "buruh" in sub:
            m2 = re.search(r"(\d+)\s*ft", sub)
            if m2:
                cols["labor"][int(m2.group(1))] = c

    warnings = []
    if "asal" not in cols or "tujuan" not in cols or not cols["price"]:
        return {"sheet": sheet, "rows": [], "specs": {}, "warnings": ["Header tabel tidak terbaca"], "skipped": 0,
                "units": [], "ppn": "", "pph": ""}

    rows, skipped, last_asal = [], 0, ""
    for r in range(hr + 3, ws.max_row + 1):
        b = bk.txt(sheet, r, cols["asal"])
        t = bk.txt(sheet, r, cols["tujuan"])
        if b.lower().startswith("alamat stuffing") or (not b and not t):
            break
        if b:
            last_asal = b
        if not t:
            continue
        prices = {u: parse_rupiah(bk.val(sheet, r, c)) for u, c in cols["price"].items()}
        prices = {u: p for u, p in prices.items() if p > 0}
        if not prices:
            skipped += 1  # vendor tidak menawar rute ini
            continue
        g = lambda k: bk.txt(sheet, r, cols[k]) if k in cols else ""
        rows.append({
            "excel_row": r,
            "asal": b or last_asal,
            "tujuan": t,
            "prices": prices,
            "labor": {n: parse_rupiah(bk.val(sheet, r, c)) for n, c in cols["labor"].items()},
            "lead_raw": g("lead"),
            "lead_days": parse_days(bk.val(sheet, r, cols["lead"])) if "lead" in cols else 0,
            "line": g("line"),
            "freq": g("freq"),
            "ppn_raw": g("ppn"),
            "pph_raw": g("pph"),
        })

    # Blok "SPESIFIKASI UNIT"
    specs = {}
    for r in range(hr + 3, ws.max_row + 1):
        hdr_cells = {c: bk.txt(sheet, r, c).lower() for c in range(1, ws.max_column + 1)}
        if any(v == "type truk" for v in hdr_cells.values()):
            c_type = [c for c, v in hdr_cells.items() if v == "type truk"][0]
            c_w = next((c for c, v in hdr_cells.items() if "kapasitas berat" in v), None)
            c_c = next((c for c, v in hdr_cells.items() if "kubikasi" in v), None)
            for rr in range(r + 1, min(r + 9, ws.max_row) + 1):
                m = re.search(r"container\s*(\d+)\s*ft", bk.txt(sheet, rr, c_type).lower())
                if m:
                    w = parse_rupiah(bk.val(sheet, rr, c_w)) if c_w else 0
                    cv = bk.val(sheet, rr, c_c) if c_c else None
                    try:
                        cv = float(str(cv).replace(",", ".")) if cv not in (None, "") else 0
                    except ValueError:
                        cv = 0
                    specs[f"Container {m.group(1)}ft"] = {"w": w, "c": cv}
            break

    return {
        "sheet": sheet,
        "rows": rows,
        "specs": specs,
        "skipped": skipped,
        "units": list(cols["price"].keys()),
        "ppn": _mode([norm_ppn(x["ppn_raw"]) for x in rows]),
        "pph": _mode([norm_pph(x["pph_raw"]) for x in rows]),
        "warnings": warnings,
    }


def parse_vendor_workbook(data: bytes):
    """Return {'vendor': {...}, 'sheets': [ ... ]}"""
    bk = _Book(data)
    vendor, sheets = {}, []
    for name in bk.wf.sheetnames:
        res = _parse_sheet(bk, name)
        if res is None:  # mis. sheet FAQ
            continue
        for k, v in _parse_header_block(bk, name).items():
            vendor.setdefault(k, v)
        sheets.append(res)
    vendor["validity"] = clean_validity(vendor.get("validity_raw", ""))
    vendor["top_norm"] = norm_top(vendor.get("top", ""))
    vendor["ppn"] = _mode([s["ppn"] for s in sheets])
    vendor["pph"] = _mode([s["pph"] for s in sheets])
    return {"vendor": vendor, "sheets": sheets}


# ----------------------------------------------------------------------------
# Cocokkan baris Excel dengan Master_Routes / Master_Units
# ----------------------------------------------------------------------------
def match_sheet_rows(rows, origin, load_type, df_r, df_g, df_units):
    """
    Return list of dict (satu per baris Excel) berisi route_id, group_id, status, import_default.
    - Nama kota dinormalisasi (huruf kecil, tanpa spasi/tanda baca).
    - '(multidrop)' / '(crossdock)' dicoba exact dulu, lalu tanpa tanda kurung.
    - Rute kembar (mis. Batam 2x) dipetakan berurutan ke rute master yang kembar;
      kalau master kurang, baris berikutnya ditandai Duplikat dan tidak di-import.
    """
    g = df_g[(df_g["load_type"] == load_type) & (df_g["origin"].map(_norm) == _norm(origin))]
    gids = set(g["group_id"].astype(str).str.strip())
    r = df_r[df_r["group_id"].astype(str).str.strip().isin(gids)].copy()
    r["route_id"] = r["route_id"].astype(str).str.strip()
    r["group_id"] = r["group_id"].astype(str).str.strip()
    r["k_asal"] = r["kota_asal"].map(_norm)
    r["k_tuj"] = r["kota_tujuan"].map(_norm)

    units = {}
    for gid, sub in df_units.assign(group_id=df_units["group_id"].astype(str).str.strip()).groupby("group_id"):
        units[gid] = {_norm(u): u for u in sub["unit_type"]}

    seen, out = Counter(), []
    for row in rows:
        ka = _norm(row["asal"])
        cands = r[(r.k_asal == ka) & (r.k_tuj == _norm(row["tujuan"]))]
        if cands.empty:
            cands = r[(r.k_asal == ka) & (r.k_tuj == _norm(re.sub(r"\(.*?\)", "", row["tujuan"])))]
        res = {"route_id": "", "group_id": "", "status": "", "import_default": False}
        if cands.empty:
            res["status"] = "❌ Rute tidak ditemukan di Master_Routes"
        else:
            key = (ka, cands.iloc[0]["k_tuj"])
            n = seen[key]
            seen[key] += 1
            pick = cands.iloc[min(n, len(cands) - 1)]
            res["route_id"], res["group_id"] = pick["route_id"], pick["group_id"]
            if n >= len(cands):
                res["status"] = f"⚠️ Duplikat (master hanya punya {len(cands)} rute ini)"
            else:
                missing = [u for u in row["prices"] if _norm(u) not in units.get(pick["group_id"], {})]
                if missing:
                    res["status"] = "⚠️ Unit belum ada di Master_Units: " + ", ".join(missing)
                    res["import_default"] = True  # unit lain tetap masuk
                else:
                    res["status"], res["import_default"] = "OK", True
        out.append(res)
    return out, r, units
