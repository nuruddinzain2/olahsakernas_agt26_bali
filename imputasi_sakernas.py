# -*- coding: utf-8 -*-
"""
IMPUTASI SAKERNAS: terapkan hasil konfirmasi lapangan (file Excel di sebuah folder) ke data SPSS awal.

Alur:
  1. Baca data awal (.sav)  -> TIDAK diubah.
  2. Baca semua file Excel konfirmasi di FOLDER_KONFIRMASI (boleh ditambah kapan saja).
  3. Responden yang dikonfirmasi "Bekerja" -> variabel pekerjaan di data diisi dari Excel.
  4. Simpan data baru (.sav) + file log Excel (apa saja yang berubah).
Data hasil (OUT_SAV) lalu dipakai oleh olah_sakernas.py.

Cara pakai:  pip install pandas pyreadstat xlsxwriter openpyxl   lalu   python imputasi_sakernas.py
"""
import glob, importlib, os, re, sys
from collections import Counter


def _tutup():
    try:
        input("\nTekan Enter untuk menutup jendela ini...")
    except Exception:
        pass


_kurang = []
for _m in ("numpy", "pandas", "pyreadstat", "xlsxwriter", "openpyxl"):
    try:
        importlib.import_module(_m)
    except ImportError:
        _kurang.append(_m)
if _kurang:
    print("\nPustaka berikut belum terpasang: " + ", ".join(_kurang) +
          "\n\nInstal lewat Command Prompt:\n    pip install " + " ".join(_kurang) +
          "\n(bila 'pip' tidak dikenali: python -m pip install " + " ".join(_kurang) + ")\n")
    _tutup()
    sys.exit(1)

import warnings
warnings.filterwarnings("ignore", message="Unknown extension")

import numpy as np
import pandas as pd
import pyreadstat

# ============================ PENGATURAN ============================
SAV_AWAL = r"D:\REKON II AGT 2026\Data Awal_51.sav"            # data awal (tidak diubah)
FOLDER_KONFIRMASI = r"D:\REKON II AGT 2026\Konfirmasi"          # taruh file Excel konfirmasi di sini
OUT_SAV = r"D:\REKON II AGT 2026\Data Awal_51_imputasi.sav"     # data imputasi (dipakai olah_sakernas.py)
OUT_LOG = r"D:\REKON II AGT 2026\Log_Imputasi.xlsx"             # catatan semua perubahan

KUNCI = "urutan"                                    # kolom pengenal responden (ada di Excel dan di .sav)
CEK_COCOK = ["kdkab", "klas", "dem_sex", "dem_age"]  # harus sama antara Excel dan .sav, bila tidak baris dilewati

# Kolom tujuan di data SPSS
VAR_MULT = "mjj_mult"                # penanda punya pekerjaan (1/2 = bekerja)
VAR_KBLI, VAR_KBJI = "mjj_kbli", "mjj_kbji"
JAM_KE = ["wkt_jml_u", "wkt_jml_s"]  # jam kerja seminggu (pekerjaan utama, seluruh pekerjaan)
# Penghasilan masuk ke kolom mana, menurut status pekerjaan (1-7). None = tidak diisi
PENGHASILAN_KE = {1: "mjj_laba_u", 2: "mjj_laba_u", 3: "mjj_laba_u",
                  4: "mjj_upah_u", 5: "mjj_upah_u", 6: "mjj_upah_u", 7: None}
# Kolom Excel tambahan -> variabel SPSS, contoh: {"apa yang dikerjakan": "nama_variabel_spss"}
MAP_TAMBAHAN = {}
# Kode status dipelajari otomatis dari responden yang sudah bekerja di data. Isi bila ingin memaksa, contoh:
# KODE_STATUS_MANUAL = {7: {"mjj_emprel": 5}}
KODE_STATUS_MANUAL = {}
# ====================================================================

STATUS_NOMOR = {"berusaha sendiri": 1, "dibantu pekerja tidak tetap": 2, "dibantu pekerja tetap": 3,
                "buruh karyawan pegawai": 4, "pekerja bebas pertanian": 5,
                "pekerja bebas bukan pertanian": 6, "pekerja keluarga": 7}
KODE_DEFAULT = {1: {"mjjemprel2": 1}, 2: {"mjjemprel2": 2}, 3: {"mjjemprel2": 3},
                4: {"mjj_emprel": 4, "mjjemprel1": 1}, 5: {"mjjemprel1": 2}, 6: {"mjjemprel1": 2},
                7: {"mjj_emprel": 3}}                # dipakai hanya bila status itu belum ada di data
KOLOM_EXCEL = {"kerja": "apakah bekerja", "kbli": "kbli", "kbji": "kbji", "status": "status pekerja",
               "penghasilan": "penghasilan", "jam": "jumlah jam kerja"}


def cek_file_keluaran(path):
    folder = os.path.dirname(path) or "."
    if not os.path.isdir(folder):
        sys.exit(f"\n[GAGAL] Folder tujuan tidak ditemukan:\n    {folder}\n")
    if os.path.exists(path):
        try:
            with open(path, "ab"):
                pass
        except PermissionError:
            sys.exit(f"\n[GAGAL] File hasil sedang terbuka di program lain (Excel/SPSS):\n    {path}\n"
                     f"Tutup file itu, lalu jalankan lagi.\n")


def angka(x):
    """Ubah isian Excel ke angka. '2.500.000' / '3 000.000' / 'Rp 1500000' -> bilangan bulat."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return np.nan
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)
    t = re.sub(r"[^\d]", "", str(x))
    return float(t) if t else np.nan


def angka_jam(x):
    if isinstance(x, (int, float, np.integer, np.floating)):
        return np.nan if pd.isna(x) else float(x)
    m = re.search(r"\d+(?:[.,]\d+)?", str(x) if x is not None else "")
    return float(m.group().replace(",", ".")) if m else np.nan


def sama(a, b):
    try:
        return float(a) == float(b)
    except (TypeError, ValueError):
        return str(a).strip().lower() == str(b).strip().lower()


def baca_konfirmasi(folder):
    if not os.path.isdir(folder):
        os.makedirs(folder, exist_ok=True)
        sys.exit(f"\nFolder konfirmasi belum ada, sudah dibuat:\n    {folder}\n"
                 f"Taruh file Excel konfirmasi di folder itu, lalu jalankan skrip lagi.\n")
    files = [f for f in glob.glob(os.path.join(folder, "*.xls*")) if not os.path.basename(f).startswith("~$")]
    if not files:
        sys.exit(f"\n[GAGAL] Tidak ada file Excel di folder:\n    {folder}\nTaruh file konfirmasi di sana dulu.\n")
    files.sort(key=os.path.getmtime)                 # file lebih baru menimpa yang lebih lama
    parts, info = [], []
    for f in files:
        nm = os.path.basename(f)
        book = pd.read_excel(f, sheet_name=None, dtype=object)
        dipakai = False
        for sh, d in book.items():
            low = {str(c).strip().lower(): c for c in d.columns}
            if KUNCI not in low:
                continue
            col = {}
            for k, awal in KOLOM_EXCEL.items():
                c = next((o for l, o in low.items() if l.startswith(awal)), None)
                col[k] = c
            if col["kerja"] is None:
                continue
            out = pd.DataFrame({"kunci_id": pd.to_numeric(d[low[KUNCI]], errors="coerce")})
            for k, c in col.items():
                out[k] = d[c] if c is not None else np.nan
            for c in CEK_COCOK:
                out["cek_" + c] = d[low[c]] if c in low else np.nan
            for awal, var in MAP_TAMBAHAN.items():
                c = next((o for l, o in low.items() if l.startswith(awal.lower())), None)
                out["x_" + var] = d[c] if c is not None else np.nan
            out["nama_x"] = d[low["dem_name"]] if "dem_name" in low else np.nan
            out["sumber_x"] = f"{nm} [{sh}]"
            parts.append(out.dropna(subset=["kunci_id"]))
            info.append((f"{nm} [{sh}]", len(parts[-1])))
            dipakai = True
        if not dipakai:
            print(f"Peringatan: '{nm}' dilewati (tidak ada sheet dengan kolom '{KUNCI}' dan 'Apakah bekerja...').")
    if not parts:
        sys.exit("\n[GAGAL] Tidak ada sheet konfirmasi yang bisa dibaca.\n")
    k = pd.concat(parts, ignore_index=True)
    ganda = k["kunci_id"].duplicated(keep=False)
    n_ganda = int(k.loc[ganda, "kunci_id"].nunique())
    k = k.drop_duplicates("kunci_id", keep="last").reset_index(drop=True)
    return k, info, n_ganda


def hitung_status(df, L):
    """Status pekerjaan 1-7, logika sama dengan syntax SPSS."""
    g = lambda n: df[L[n]] if n in L else pd.Series(np.nan, index=df.index)
    k = g("mjj_mult").isin([1, 2])
    st = pd.Series(np.nan, index=df.index)
    st[k & (g("mjjemprel2") == 1)] = 1
    st[k & (g("mjjemprel2") == 2)] = 2
    st[k & (g("mjjemprel2") == 3)] = 3
    st[k & ((g("mjj_emprel") == 4) | (g("mjjemprel1") == 1))] = 4
    st[k & (g("mjjemprel1") == 2) & (g("mjj_kbli") <= 3300)] = 5
    st[k & (g("mjjemprel1") == 2) & (g("mjj_kbli") >= 5101)] = 6
    st[k & g("mjj_emprel").isin([3, 5])] = 7
    return st


def pelajari_kode(df, L, st):
    """Cari kombinasi kode status yang paling umum pada responden yang SUDAH bekerja di data."""
    kol = [c for c in ("mjj_emprel", "mjjemprel1", "mjjemprel2") if c in L]
    hasil = {s: dict(v) for s, v in KODE_DEFAULT.items()}
    sumber = {s: "bawaan (belum ada contoh di data)" for s in hasil}
    for s in range(1, 8):
        sub = df[st == s]
        if sub.empty or not kol:
            continue
        c = Counter(tuple(None if pd.isna(v) else float(v) for v in r)
                    for r in sub[[L[x] for x in kol]].itertuples(index=False, name=None))
        best, n = c.most_common(1)[0]
        hasil[s] = {x: v for x, v in zip(kol, best) if v is not None}
        sumber[s] = f"dipelajari dari {len(sub)} responden"
    for s, v in KODE_STATUS_MANUAL.items():
        hasil[s] = dict(v); sumber[s] = "manual"
    return hasil, sumber


def main():
    if os.path.abspath(OUT_SAV).lower() == os.path.abspath(SAV_AWAL).lower():
        sys.exit("\n[GAGAL] OUT_SAV sama dengan SAV_AWAL. Data awal akan tertimpa, jadi proses dihentikan.\n"
                 "Beri nama berbeda untuk OUT_SAV (contoh: Data Awal_51_imputasi.sav).\n")
    cek_file_keluaran(OUT_SAV)
    cek_file_keluaran(OUT_LOG)
    path = SAV_AWAL
    if not os.path.exists(path):
        cand = [f for f in glob.glob(os.path.join(os.path.dirname(path), "Data Awal*51*.sav"))
                if not re.search(r"imputasi|usia15", f, re.I)]
        if not cand:
            sys.exit(f"\n[GAGAL] Data awal tidak ditemukan: {path}\n")
        path = cand[0]
    print("Membaca data awal:", path)
    df, meta = pyreadstat.read_sav(path, apply_value_formats=False)
    L = {c.lower(): c for c in df.columns}
    if KUNCI not in L:
        sys.exit(f"\n[GAGAL] Kolom '{KUNCI}' tidak ada di data .sav.\n")
    if VAR_MULT not in L:
        sys.exit(f"\n[GAGAL] Kolom '{VAR_MULT}' tidak ada di data .sav (penanda bekerja).\n")
    uk = pd.to_numeric(df[L[KUNCI]], errors="coerce")
    if uk.dropna().duplicated().any():
        sys.exit(f"\n[GAGAL] Kolom '{KUNCI}' di .sav tidak unik ({int(uk.dropna().duplicated().sum())} ganda). "
                 f"Beri tahu pengelola skrip.\n")
    pos = {int(round(v)): i for i, v in enumerate(uk) if pd.notna(v)}

    print("Membaca file konfirmasi di:", FOLDER_KONFIRMASI)
    kon, info, n_ganda = baca_konfirmasi(FOLDER_KONFIRMASI)
    for nm, n in info:
        print(f"  - {nm}: {n:,} baris")

    st0 = hitung_status(df, L)
    kode, sumber_kode = pelajari_kode(df, L, st0)
    sub = df[L[VAR_MULT]][df[L[VAR_MULT]].isin([1, 2])]
    mult_val = float(sub.mode().iloc[0]) if len(sub) else 1.0
    print("Kode status yang dipakai:", {s: kode[s] for s in sorted(kode)})

    tujuan_ada = lambda v: v is not None and v.lower() in L
    dilewati = set()

    perubahan, tdk_ketemu, tdk_cocok, tdk_lengkap = [], [], [], []
    n_bekerja = n_tidak = n_belum = 0
    diterapkan = {}                      # posisi baris -> status yang dituju
    flag = np.zeros(len(df))

    for r in kon.itertuples(index=False):
        t = str(r.kerja).strip().lower() if pd.notna(r.kerja) else ""
        if t == "":
            n_belum += 1; continue
        if t.startswith("tidak"):
            n_tidak += 1; continue
        if t != "bekerja":
            n_belum += 1; continue
        u = int(round(r.kunci_id))
        i = pos.get(u)
        if i is None:
            tdk_ketemu.append({"urutan": u, "nama": r.nama_x, "sumber": r.sumber_x}); continue
        salah = [c for c in CEK_COCOK if c in L and pd.notna(getattr(r, "cek_" + c))
                 and not sama(df.at[i, L[c]], getattr(r, "cek_" + c))]
        if salah:
            tdk_cocok.append({"urutan": u, "nama": r.nama_x, "kolom_beda": ", ".join(
                f"{c}: sav={df.at[i, L[c]]} excel={getattr(r, 'cek_' + c)}" for c in salah), "sumber": r.sumber_x})
            continue

        txt = str(r.status).strip().lower() if pd.notna(r.status) else ""
        sno = STATUS_NOMOR.get(txt)
        kbli, kbji = angka(r.kbli), angka(r.kbji)
        peng, jam = angka(r.penghasilan), angka_jam(r.jam)
        kurang = []
        if sno is None: kurang.append("status" + ("" if not txt else f" tidak dikenal ('{txt}')"))
        if np.isnan(kbli): kurang.append("KBLI")
        if np.isnan(jam): kurang.append("jam kerja")
        if pd.notna(r.penghasilan) and np.isnan(peng): kurang.append("penghasilan tidak terbaca")

        isi = {L[VAR_MULT]: mult_val}
        if sno is not None:
            for c, v in kode[sno].items():
                if c in L: isi[L[c]] = v
                else: dilewati.add(c)
            kolp = PENGHASILAN_KE.get(sno)
            if kolp and not np.isnan(peng):
                if tujuan_ada(kolp): isi[L[kolp.lower()]] = peng
                else: dilewati.add(kolp)
        for var, v in ((VAR_KBLI, kbli), (VAR_KBJI, kbji)):
            if not np.isnan(v):
                if tujuan_ada(var): isi[L[var.lower()]] = v
                else: dilewati.add(var)
        if not np.isnan(jam):
            for var in JAM_KE:
                if tujuan_ada(var): isi[L[var.lower()]] = jam
                else: dilewati.add(var)
        for var in MAP_TAMBAHAN.values():
            v = getattr(r, "x_" + var)
            if pd.notna(v):
                if tujuan_ada(var): isi[L[var.lower()]] = v
                else: dilewati.add(var)

        nama = df.at[i, L["dem_name"]] if "dem_name" in L else r.nama_x
        for c, v in isi.items():
            lama = df.at[i, c]
            if df[c].dtype == object:
                v = str(v)
            df.at[i, c] = v
            perubahan.append({"urutan": u, "nama": nama, "variabel": c,
                              "nilai_lama": lama, "nilai_baru": v, "sumber": r.sumber_x})
        flag[i] = 1
        diterapkan[i] = sno
        n_bekerja += 1
        if kurang:
            tdk_lengkap.append({"urutan": u, "nama": nama, "yang_kurang": ", ".join(kurang), "sumber": r.sumber_x})

    # verifikasi status hasil imputasi
    st1 = hitung_status(df, L)
    tdk_sesuai = []
    for i, sno in diterapkan.items():
        if sno is not None and st1.iloc[i] != sno:
            tdk_sesuai.append({"urutan": int(round(uk.iloc[i])),
                               "status_di_excel": sno, "status_hasil_hitung": st1.iloc[i],
                               "kbli": df.at[i, L[VAR_KBLI]] if VAR_KBLI in L else np.nan})

    df["flag_imputasi"] = flag
    labels = list(meta.column_labels) + ["1 = diubah berdasarkan konfirmasi lapangan"]
    vl = dict(meta.variable_value_labels); vl["flag_imputasi"] = {0.0: "0 Tidak diubah", 1.0: "1 Diubah (konfirmasi)"}
    try:
        pyreadstat.write_sav(df, OUT_SAV, column_labels=labels, variable_value_labels=vl,
                             missing_ranges=meta.missing_ranges, variable_display_width=meta.variable_display_width,
                             variable_measure=meta.variable_measure)
    except Exception as e:
        print(f"(catatan: metadata tambahan tidak bisa dibawa [{e}], menyimpan dengan label saja)")
        pyreadstat.write_sav(df, OUT_SAV, column_labels=labels, variable_value_labels=vl)
    print(f"Data imputasi ({len(df):,} baris) -> {OUT_SAV}")

    ring = [("Data awal", path), ("File konfirmasi dibaca", len(info)), ("Baris konfirmasi unik", len(kon)),
            ("urutan ganda antar file (yang terbaru dipakai)", n_ganda),
            ("Dikonfirmasi BEKERJA & diterapkan", n_bekerja), ("Dikonfirmasi tidak bekerja (tanpa perubahan)", n_tidak),
            ("Belum diisi / tidak jelas (tanpa perubahan)", n_belum),
            ("Bekerja tapi urutan tidak ada di .sav", len(tdk_ketemu)),
            ("Bekerja tapi data Excel tidak cocok dengan .sav (dilewati)", len(tdk_cocok)),
            ("Diterapkan tapi isian kurang lengkap", len(tdk_lengkap)),
            ("Diterapkan tapi status hasil hitung beda dgn Excel", len(tdk_sesuai)),
            ("Variabel tujuan tidak ada di .sav (dilewati)", ", ".join(sorted(dilewati)) or "-")]
    with pd.ExcelWriter(OUT_LOG, engine="xlsxwriter") as xw:
        hd = xw.book.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1})
        def tulis(nm, d):
            d.to_excel(xw, sheet_name=nm, index=False)
            ws = xw.sheets[nm]
            for j, c in enumerate(d.columns):
                ws.write(0, j, c, hd)
                ws.set_column(j, j, min(60, max(14, (max(len(str(v)) for v in d[c]) if len(d) else 14) + 2)))
        tulis("Ringkasan", pd.DataFrame(ring, columns=["Butir", "Jumlah"]))
        tulis("Kode_status", pd.DataFrame([(s, str(kode[s]), sumber_kode[s]) for s in sorted(kode)],
                                          columns=["Status", "Kode di data", "Asal"]))
        tulis("Perubahan", pd.DataFrame(perubahan, columns=["urutan", "nama", "variabel", "nilai_lama", "nilai_baru", "sumber"]))
        tulis("Tidak_lengkap", pd.DataFrame(tdk_lengkap, columns=["urutan", "nama", "yang_kurang", "sumber"]))
        tulis("Status_tdk_sesuai", pd.DataFrame(tdk_sesuai, columns=["urutan", "status_di_excel", "status_hasil_hitung", "kbli"]))
        tulis("Tidak_ketemu", pd.DataFrame(tdk_ketemu, columns=["urutan", "nama", "sumber"]))
        tulis("Tidak_cocok", pd.DataFrame(tdk_cocok, columns=["urutan", "nama", "kolom_beda", "sumber"]))
    print(f"Log -> {OUT_LOG}\n")
    for b, j in ring[4:]:
        print(f"  {b}: {j}")
    print("\nSelesai. Pakai file imputasi di olah_sakernas.py (SAV_PATH).")


if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        if e.code not in (None, 0):
            print(e.code)
    except PermissionError as e:
        print(f"\n[GAGAL] Akses file ditolak (mungkin masih terbuka di Excel/SPSS):\n    {e.filename}\n")
    except FileNotFoundError as e:
        print(f"\n[GAGAL] File/folder tidak ditemukan:\n    {e.filename}\n")
    except Exception:
        import traceback
        print("\n[GAGAL] Kesalahan tak terduga. Salin pesan di bawah dan kirim ke pengelola:\n")
        traceback.print_exc()
    _tutup()
