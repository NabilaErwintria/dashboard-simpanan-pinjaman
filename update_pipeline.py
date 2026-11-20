"""SEKDA BI -> Google Sheets (sumber Tableau Public).

Jalankan:
  python update_pipeline.py --inspect   # lihat 15 baris pertama tiap file, untuk menyetel parser
  python update_pipeline.py             # unduh, validasi, tulis ke Google Sheets bila ada perubahan

Environment:
  GSHEET_ID, GCP_SA_JSON (isi JSON service account), SEKDA_II04_URL, SEKDA_II15_URL
"""
import os, io, json, hashlib, sys, datetime as dt
import requests, pandas as pd, gspread
from google.oauth2.service_account import Credentials

KAB = ["Bojonegoro", "Gresik", "Lamongan", "Tuban"]
# label_wil / label_kat = indeks kolom Wilayah dan kategori di sheet BI. SESUAIKAN setelah --inspect.
TABLES = {
    "Simpanan": dict(url_env="SEKDA_II04_URL", sheet="ii04", label_wil=0, label_kat=1,
                     kat_name="Jenis Simpanan", val_name="Nilai Simpanan"),
    "Pinjaman": dict(url_env="SEKDA_II15_URL", sheet="ii15", label_wil=0, label_kat=1,
                     kat_name="Lapangan Usaha", val_name="Nilai Pinjaman"),
}
HEADERS = {"User-Agent": "Mozilla/5.0 (dashboard riset regional)"}
BULAN = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "mei": 5, "may": 5, "jun": 6, "jul": 7,
         "agu": 8, "aug": 8, "ags": 8, "sep": 9, "okt": 10, "oct": 10, "nov": 11, "des": 12, "dec": 12}


def to_period(x):
    """Ubah sel header (datetime / 'Jan 2024' / 'Januari 2024') menjadi Timestamp awal bulan."""
    if isinstance(x, (dt.datetime, pd.Timestamp)):
        return pd.Timestamp(x.year, x.month, 1)
    if isinstance(x, str):
        p = x.replace("*", "").strip().lower().split()
        if len(p) == 2 and p[1].isdigit() and p[0][:3] in BULAN:
            return pd.Timestamp(int(p[1]), BULAN[p[0][:3]], 1)
    return None


def clean_text(s):
    return " ".join(str(s).split()) if pd.notna(s) else None   # buang spasi ganda/trailing


def parse_bi_table(raw, cfg):
    hdr = next((i for i in range(len(raw)) if sum(to_period(v) is not None for v in raw.iloc[i]) >= 6), None)
    if hdr is None:
        raise ValueError("Baris header periode tidak ditemukan - format file BI mungkin berubah.")
    periods = {c: to_period(v) for c, v in raw.iloc[hdr].items() if to_period(v) is not None}
    body = raw.iloc[hdr + 1:].copy()
    body["Wilayah"] = body.iloc[:, cfg["label_wil"]].map(clean_text).ffill()
    body[cfg["kat_name"]] = body.iloc[:, cfg["label_kat"]].map(clean_text)
    body = body[body[cfg["kat_name"]].notna()]
    long = body.melt(id_vars=["Wilayah", cfg["kat_name"]], value_vars=list(periods),
                     var_name="col", value_name=cfg["val_name"])
    long["Periode"] = long["col"].map(periods)
    long[cfg["val_name"]] = pd.to_numeric(long[cfg["val_name"]], errors="coerce")
    long = long.dropna(subset=[cfg["val_name"]]).drop(columns="col")
    long = long[long["Wilayah"].isin(KAB)]
    long["Satuan"] = "Rp Juta"     # cek satuan di file asli
    long["Granularitas"] = "Bulanan"
    long["Tahun"] = long["Periode"].dt.year
    long["Bulan"] = long["Periode"].dt.month
    long["Kuartal"] = "Q" + long["Periode"].dt.quarter.astype(str)
    long["Max_Bulan"] = long.groupby("Tahun")["Bulan"].transform("max")
    long["Periode"] = long["Periode"].dt.strftime("%Y-%m-%d")
    cols = ["Wilayah", cfg["kat_name"], "Satuan", "Periode", cfg["val_name"], "Granularitas",
            "Bulan", "Tahun", "Max_Bulan", "Kuartal"]
    return long[cols].reset_index(drop=True)


def validate(df, cfg, prev_rows):
    if df.empty:
        raise ValueError("Hasil parse kosong.")
    missing = set(KAB) - set(df["Wilayah"])
    if missing:
        raise ValueError(f"Wilayah hilang: {missing}")
    if (df[cfg["val_name"]] < 0).any():
        raise ValueError("Ada nilai negatif.")
    if prev_rows and len(df) < 0.95 * prev_rows:
        raise ValueError(f"Baris turun dari {prev_rows} ke {len(df)} - cek perubahan format.")


def client():
    info = json.loads(os.environ["GCP_SA_JSON"])
    creds = Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return gspread.authorize(creds)


def get_ws(sh, title):
    try:
        return sh.worksheet(title)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(title, rows=100, cols=12)


def write_df(ws, df):
    vals = [df.columns.tolist()] + df.astype(object).where(df.notna(), "").values.tolist()
    ws.clear()
    ws.resize(rows=len(vals), cols=len(df.columns))
    for i in range(0, len(vals), 5000):
        ws.update(values=vals[i:i + 5000], range_name=f"A{i + 1}")


def main(inspect=False):
    sh = None if inspect else client().open_by_key(os.environ["GSHEET_ID"])
    meta_ws = None if inspect else get_ws(sh, "_meta")
    meta = {} if inspect else {r[0]: r[1] for r in meta_ws.get_all_values() if len(r) >= 2}
    for name, cfg in TABLES.items():
        url = os.environ.get(cfg["url_env"])
        if not url:
            print(f"[{name}] {cfg['url_env']} belum diisi, dilewati"); continue
        resp = requests.get(url, headers=HEADERS, timeout=120); resp.raise_for_status()
        h = hashlib.sha256(resp.content).hexdigest()
        raw = pd.read_excel(io.BytesIO(resp.content), sheet_name=cfg["sheet"], header=None)
        if inspect:
            print(f"== {name} ==\n", raw.head(15).to_string()); continue
        if meta.get(f"{name}_hash") == h:
            print(f"[{name}] tidak ada perubahan"); continue
        df = parse_bi_table(raw, cfg)
        ws = get_ws(sh, name)
        validate(df, cfg, max(len(ws.col_values(1)) - 1, 0))
        write_df(ws, df)
        meta[f"{name}_hash"] = h
        meta[f"{name}_periode_terakhir"] = df["Periode"].max()
        meta[f"{name}_diperbarui"] = dt.datetime.utcnow().isoformat(timespec="seconds")
        print(f"[{name}] diperbarui: {len(df)} baris, periode terakhir {df['Periode'].max()}")
    if not inspect:
        meta_ws.clear(); meta_ws.update(values=[[k, v] for k, v in meta.items()], range_name="A1")


if __name__ == "__main__":
    main(inspect="--inspect" in sys.argv)
