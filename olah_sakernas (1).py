# -*- coding: utf-8 -*-
"""
Olah Sakernas: TPT, TPAK, Status Pekerjaan, Lapangan Usaha (17 kategori KBLI 2025), Upah
Dirinci: Total, Jenis Kelamin, Kota/Desa, Kabupaten/Kota. Semua angka PAKAI PENIMBANG.
Logika variabel mengikuti file syntax "B. Variabel SAK202608.sps".

Cara pakai:  pip install pandas pyreadstat xlsxwriter   lalu   python olah_sakernas.py
"""
import glob, os, re, sys
import importlib


def _tutup():
    """Tahan jendela supaya pesan sempat dibaca saat skrip diklik dua kali."""
    try:
        input("\nTekan Enter untuk menutup jendela ini...")
    except Exception:
        pass


# Cek pustaka dulu, beri pesan jelas bila ada yang belum terpasang
_butuh = {"numpy": "numpy", "pandas": "pandas", "pyreadstat": "pyreadstat", "xlsxwriter": "xlsxwriter"}
_kurang = []
for _modul, _paket in _butuh.items():
    try:
        importlib.import_module(_modul)
    except ImportError:
        _kurang.append(_paket)
if _kurang:
    print("\nPustaka berikut belum terpasang di komputer ini: " + ", ".join(_kurang) +
             "\n\nInstal dulu lewat Command Prompt dengan perintah:\n"
             "    pip install " + " ".join(_kurang) +
             "\n\nBila 'pip' tidak dikenali, coba:\n"
             "    python -m pip install " + " ".join(_kurang) + "\n")
    _tutup()
    sys.exit(1)

import numpy as np
import pandas as pd
import pyreadstat

# ============================ PENGATURAN ============================
SAV_PATH = r"D:\REKON II AGT 2026\Data Awal_51_imputasi.sav"   # hasil imputasi_sakernas.py
OUT_XLSX = r"D:\REKON II AGT 2026\Hasil_Sakernas.xlsx"
OUT_SAV  = r"D:\REKON II AGT 2026\Data Awal 51_usia15plus.sav"   # SPSS baru: usia 15+ & variabel baru

# Isi manual bila auto-deteksi gagal (huruf besar/kecil bebas), contoh: KOL["kab"]="r102"
KOL = {"weight": None, "kab": None, "sex": None, "klas": None}
KANDIDAT = {
    "weight": ["weightr", "weight", "wt", "fwt", "bobot", "penimbang", "weight_final", "w_final", "wi", "faktor"],
    "kab":    ["kab", "kabkot", "kab_kota", "kd_kab", "kode_kab", "kabupaten", "r102", "kdkab"],
    "sex":    ["dem_sex", "sex", "jk", "dem_jk", "jenis_kelamin", "kelamin"],
    "klas":   ["klas", "klasifikasi", "kd_klas", "kota_desa", "klasdes", "daerah", "r105", "area"],
}
# Upah: "karyawan" = hanya buruh/karyawan/pegawai (status 4)  -> standar BPS
#       "semua"    = semua pekerja (upah + laba)
MODE_UPAH = "karyawan"
UMUR_MIN = 15

# Angka acuan 2025 (rilis) untuk tabel perbandingan: "Wilayah": (TPT, TPAK). Kosongkan {} bila tak dipakai.
# Nama wilayah harus mengandung nama kabupaten/kota seperti di variabel nmkab; "Bali" = total provinsi.
REF_2025 = {"Jembrana": (1.64, 78.69), "Tabanan": (1.58, 76.88), "Badung": (1.57, 74.61),
            "Gianyar": (1.55, 75.74), "Klungkung": (1.15, 78.66), "Bangli": (0.73, 81.73),
            "Karangasem": (1.80, 84.69), "Buleleng": (1.51, 79.80), "Denpasar": (1.41, 70.98),
            "Bali": (1.49, 77.31)}
JUDUL_BANDING = ("Tabel TPT dan TPAK Menurut Kabupaten/Kota di Provinsi Bali, "
                 "Tahun 2025 (Rilis) dan Tahun 2026 (Angka Sangat Sementara)")
# ====================================================================

LABEL_DEFAULT = {
    "sex": {1: "Laki-laki", 2: "Perempuan"},
    "klas": {1: "Perkotaan", 2: "Perdesaan"},
}
STATUS = {1: "1. Berusaha sendiri",
          2: "2. Berusaha dibantu buruh tidak tetap/pekerja keluarga/tidak dibayar",
          3: "3. Berusaha dibantu buruh tetap dan dibayar",
          4: "4. Buruh/karyawan/pegawai",
          5: "5. Pekerja bebas pertanian",
          6: "6. Pekerja bebas non-pertanian",
          7: "7. Pekerja keluarga/tidak dibayar"}
LU = {1: "A Pertanian, Kehutanan, dan Perikanan", 2: "B Pertambangan dan Penggalian", 3: "C Industri",
      4: "D Listrik, Gas, Uap/Air Panas, Udara Dingin", 5: "E Air, Pengelolaan Limbah, Remediasi",
      6: "F Konstruksi", 7: "G Perdagangan Besar dan Eceran", 8: "H Transportasi dan Pergudangan",
      9: "I Akomodasi dan Makan Minum", 10: "JK Informasi, Komunikasi (J+K)", 11: "L Keuangan dan Asuransi",
      12: "M Real Estat", 13: "NO Jasa Profesional & Administratif (N+O)",
      14: "P Administrasi Pemerintahan", 15: "Q Pendidikan", 16: "R Kesehatan dan Kegiatan Sosial",
      17: "STUV Jasa Lainnya (S+T+U+V)"}
KBLI_RANGE = [(1111, 3300, 1), (5101, 9900, 2), (10111, 33203, 3), (35111, 35402, 4), (36001, 39009, 5),
              (41011, 43909, 6), (46100, 47909, 7), (49111, 53309, 8), (55101, 56400, 9), (58110, 63900, 10),
              (64110, 66309, 11), (68111, 68299, 12), (69101, 82990, 13), (84111, 84300, 14),
              (85101, 85699, 15), (86101, 88907, 16), (90111, 93299, 17), (94110, 95299, 17),
              (95311, 95320, 17), (95400, 96900, 17), (97000, 98200, 17), (99000, 99000, 17)]


def cari_kolom(df, kunci):
    """Kembalikan nama kolom, atau None bila tidak ketemu otomatis."""
    if KOL[kunci]:
        c = KOL[kunci].lower()
        if c not in df.columns:
            sys.exit(f"\n[GAGAL] Kolom '{c}' (untuk {kunci}) yang Anda isi di KOL tidak ada di data.\n")
        return c
    for c in KANDIDAT[kunci]:
        if c in df.columns:
            return c
    if kunci == "weight":                      # cadangan: kolom berawalan "weight"
        for c in df.columns:
            if c.startswith("weight"):
                return c
    return None


KATA_CARI = {"weight": ["weight", "bobot", "penimbang", "faktor", "wt"],
             "kab": ["kab", "kota", "wilayah", "regency"],
             "sex": ["sex", "kelamin", "jk", "gender"],
             "klas": ["klas", "desa", "daerah", "perkotaan", "urban"]}
NAMA_TAMPIL = {"weight": "PENIMBANG (weight)", "kab": "KABUPATEN/KOTA", "sex": "JENIS KELAMIN",
               "klas": "KOTA/DESA (klasifikasi)"}


def laporkan_hilang(df, var_label, hilang):
    """Tampilkan ringkas variabel mana yang tidak ketemu + kolom mirip, simpan daftar lengkap ke file."""
    daftar = os.path.join(os.path.dirname(OUT_XLSX) or ".", "daftar_variabel.txt")
    try:
        with open(daftar, "w", encoding="utf-8") as f:
            for c in df.columns:
                f.write(f"{c}\t{var_label.get(c, '')}\n")
        info = f"Daftar lengkap seluruh variabel (nama + label) disimpan di:\n    {daftar}\n"
    except Exception:
        info = ""
    print("\n[GAGAL] Variabel berikut tidak ketemu otomatis di data:")
    for k in hilang:
        print(f"\n  * {NAMA_TAMPIL[k]}")
        mirip = [c for c in df.columns
                 if any(x in c or x in var_label.get(c, "").lower() for x in KATA_CARI[k])][:12]
        if mirip:
            print("    Kolom yang mungkin cocok:")
            for c in mirip:
                print(f"      {c:<18} {var_label.get(c, '')[:60]}")
        else:
            print("    (tidak ada kolom yang mirip, cek daftar_variabel.txt)")
    print("\nCara memperbaiki: buka skrip, di bagian PENGATURAN isi nama kolom yang benar, contoh:")
    print('    KOL = {"weight": "nama_kolom_penimbang", "kab": None, "sex": None, "klas": None}')
    print(info)
    sys.exit(None)


def baca_data(path):
    if not os.path.exists(path):
        cand = glob.glob(os.path.join(os.path.dirname(path), "Data Awal*51*imputasi*.sav"))
        if not cand:
            sys.exit(f"File tidak ditemukan: {path}")
        path = cand[0]
    print("Membaca:", path)
    df, meta = pyreadstat.read_sav(path, apply_value_formats=False)
    asli = {c.lower(): c for c in df.columns}                       # nama kolom asli (huruf besar/kecil)
    var_label = {k.lower(): v for k, v in meta.column_names_to_labels.items() if v}
    df.columns = [c.lower() for c in df.columns]
    labels = {k.lower(): v for k, v in meta.variable_value_labels.items()}
    return df, labels, asli, var_label


def hitung_variabel(d):
    """Replikasi syntax SPSS."""
    g = lambda c: d[c] if c in d.columns else pd.Series(np.nan, index=d.index)
    bekerja = np.where(g("mjj_mult").isin([1, 2]), 1, 2)
    peng = np.full(len(d), 2)
    peng[(g("srh_kerja") == 1) | (g("srh_usaha") == 1)] = 1
    peng[(g("srh_kerja") == 2) & ((g("srh_usaha") == 2) | g("srh_usaha").isna()) & (g("srh_alsn") <= 3)] = 1
    jk = np.full(len(d), np.nan)
    jk[bekerja == 1] = 1
    m = (bekerja == 2) & (peng == 1); jk[m] = 2
    m = (bekerja == 2) & (peng == 2)
    jk[m & (g("mkl_utama") == 1)] = 4
    jk[m & (g("mkl_utama") == 2)] = 5
    jk[m & (g("mkl_utama") >= 3)] = 6
    d["jeniskegiatan_imputasi"] = jk
    d["ak3"] = pd.Series(jk, index=d.index).map({1: 1, 2: 2, 4: 3, 5: 3, 6: 3})   # 1 Bekerja 2 Penganggur 3 BAK

    # status pekerjaan (urutan sama dgn syntax: yang belakang menimpa)
    st = pd.Series(np.nan, index=d.index)
    k = d["jeniskegiatan_imputasi"] == 1
    st[k & (g("mjjemprel2") == 1)] = 1
    st[k & (g("mjjemprel2") == 2)] = 2
    st[k & (g("mjjemprel2") == 3)] = 3
    st[k & ((g("mjj_emprel") == 4) | (g("mjjemprel1") == 1))] = 4
    st[k & (g("mjjemprel1") == 2) & (g("mjj_kbli") <= 3300)] = 5
    st[k & (g("mjjemprel1") == 2) & (g("mjj_kbli") >= 5101)] = 6
    st[k & g("mjj_emprel").isin([3, 5])] = 7
    d["status_pekerjaan"] = st

    lu = pd.Series(np.nan, index=d.index)
    kb = g("mjj_kbli")
    for lo, hi, c in KBLI_RANGE:
        lu[(kb >= lo) & (kb <= hi)] = c
    d["lu17"] = lu

    d["upah"] = d[[c for c in ("mjj_upah_u", "mjj_upah_b") if c in d.columns]].sum(axis=1, min_count=1) \
        if any(c in d.columns for c in ("mjj_upah_u", "mjj_upah_b")) else np.nan
    d["laba"] = d[[c for c in ("mjj_laba_u", "mjj_laba_b") if c in d.columns]].sum(axis=1, min_count=1) \
        if any(c in d.columns for c in ("mjj_laba_u", "mjj_laba_b")) else np.nan
    return d


def wmedian(x, w):
    o = np.argsort(x.values); x = x.values[o]; w = w.values[o]
    cw = np.cumsum(w)
    return x[np.searchsorted(cw, cw[-1] / 2)]


# variabel baru -> (nama di SPSS, label variabel, label nilai)
VAR_BARU = {
    "jeniskegiatan_imputasi":    ("JenisKegiatan_imputasi", "Jenis Kegiatan (rinci) - hasil hitung skrip",
                         {1: "1 Bekerja", 2: "2 Pengangguran", 4: "4 Sekolah", 5: "5 Mengurus Rumah Tangga", 6: "6 Lainnya"}),
    "ak3":              ("AK3", "Jenis Kegiatan", {1: "1 Bekerja", 2: "2 Pengangguran", 3: "3 Bukan Angkatan Kerja"}),
    "akbak":            ("AKBAK", "AKBAK", {1: "1 Angkatan Kerja", 2: "2 Bukan Angkatan Kerja"}),
    "status_pekerjaan": ("status_pekerjaan", "Status Pekerjaan (Konsep Lama)", STATUS),
    "lu17":             ("kategori172025", "Lapangan Usaha 17 Kategori KBLI 2025", LU),
    "upah":             ("upah_gaji", "Upah/gaji sebulan (uang + barang)", {}),
    "laba":             ("laba_usaha", "Laba/keuntungan usaha sebulan (uang + barang)", {}),
}


def simpan_sav(df, labels, asli, var_label, path):
    out = df.drop(columns=["w"], errors="ignore").copy()
    out["akbak"] = out["jeniskegiatan_imputasi"].map({1: 1, 2: 1, 4: 2, 5: 2, 6: 2})
    # Bila variabel dengan nama SPSS yang sama sudah ada di data asli, isinya ditimpa (tidak dobel)
    final = {}
    for c, (spss, _, _) in VAR_BARU.items():
        tgt = spss.lower()
        if c != tgt and tgt in out.columns:
            out[tgt] = out[c]
            out = out.drop(columns=[c])
            final[c] = tgt
        else:
            final[c] = c
    nama = {c: (asli.get(c) or next((v[0] for kk, v in VAR_BARU.items() if final[kk] == c), c))
            for c in out.columns}
    out = out.rename(columns=nama)
    col_labels = {nama[c]: v for c, v in var_label.items() if c in nama}
    val_labels = {nama[c]: v for c, v in labels.items() if c in nama}
    for c, (_, lbl, vl) in VAR_BARU.items():
        col = final[c]
        if col in nama:
            col_labels[nama[col]] = lbl
            if vl:
                val_labels[nama[col]] = {float(k): v for k, v in vl.items()}
    assert out.columns.is_unique, "Ada nama kolom ganda: " + str(out.columns[out.columns.duplicated()].tolist())
    pyreadstat.write_sav(out, path, column_labels=[col_labels.get(c) or "" for c in out.columns],
                         variable_value_labels=val_labels)
    print(f"File SPSS baru ({len(out):,} baris, {len(out.columns)} variabel) -> {path}")


def cek_file_keluaran(path):
    """Pastikan folder ada dan file hasil tidak sedang terbuka di program lain."""
    folder = os.path.dirname(path) or "."
    if not os.path.isdir(folder):
        sys.exit(f"\n[GAGAL] Folder tujuan tidak ditemukan:\n    {folder}\n"
                 f"Periksa penulisan folder pada OUT_XLSX / OUT_SAV di bagian PENGATURAN.\n")
    if os.path.exists(path):
        try:
            with open(path, "ab"):
                pass
        except PermissionError:
            sys.exit(f"\n[GAGAL] File hasil tidak bisa ditimpa karena sedang terbuka di program lain "
                     f"(Excel/SPSS):\n    {path}\nTutup file tersebut, lalu jalankan skrip lagi.\n")
    elif not os.access(folder, os.W_OK):
        sys.exit(f"\n[GAGAL] Tidak punya izin menulis di folder:\n    {folder}\n")


def norm_wil(x):
    x = re.sub(r"\[.*?\]", "", str(x)).lower()
    for w in ("kabupaten", "kab.", "kota"):
        x = x.replace(w, "")
    return re.sub(r"[^a-z]", "", x)


def buat_perbandingan(t_ind):
    """Gabungkan angka acuan 2025 dengan hasil hitung 2026 (TPT & TPAK)."""
    kab = t_ind[t_ind["Dimensi"] == "Kabupaten/Kota"]
    idx = {norm_wil(k): r for k, (_, r) in zip(kab["Kelompok"], kab.iterrows())}
    total = t_ind.iloc[0]
    rows = []
    for nama, (t25, k25) in REF_2025.items():
        r = total if norm_wil(nama) == "bali" else idx.get(norm_wil(nama))
        if r is None:
            print(f"Peringatan: '{nama}' pada acuan 2025 tidak ditemukan di data 2026 (cek nama di nmkab).")
            t26 = k26 = np.nan
        else:
            t26, k26 = round(float(r["TPT (%)"]), 2), round(float(r["TPAK (%)"]), 2)
        rows.append((nama, t25, t26, round(t26 - t25, 2), k25, k26, round(k26 - k25, 2)))
    cols = pd.MultiIndex.from_tuples([("TPT", "2025"), ("TPT", "2026"), ("TPT", "Perubahan"),
                                      ("TPAK", "2025"), ("TPAK", "2026"), ("TPAK", "Perubahan")])
    df = pd.DataFrame([r[1:] for r in rows], columns=cols, index=pd.Index([r[0] for r in rows], name="Wilayah"))
    return df


def main():
    cek_file_keluaran(OUT_XLSX)      # cek di awal, supaya tidak gagal setelah proses lama
    cek_file_keluaran(OUT_SAV)
    df, labels, asli, var_label = baca_data(SAV_PATH)
    ketemu = {k: cari_kolom(df, k) for k in ("weight", "kab", "sex", "klas")}
    hilang = [k for k, v in ketemu.items() if v is None]
    if hilang:
        laporkan_hilang(df, var_label, hilang)
    c_w, c_kab, c_sex, c_kl = ketemu["weight"], ketemu["kab"], ketemu["sex"], ketemu["klas"]
    print(f"Weight={c_w}  Kab={c_kab}  JK={c_sex}  Kota/Desa={c_kl}")

    df = hitung_variabel(df)
    df = df[df["dem_age"] >= UMUR_MIN].copy()          # penduduk usia kerja
    simpan_sav(df, labels, asli, var_label, OUT_SAV)   # SPSS usia 15+ + variabel baru
    df["w"] = df[c_w]

    def lab(col, key):
        m = labels.get(col, {})
        base = LABEL_DEFAULT.get(key, {})

        def f(v):
            try:
                nv = float(v); nv = int(nv) if nv.is_integer() else nv
            except (TypeError, ValueError):
                nv = v
            nama = m.get(v, m.get(nv, base.get(nv)))
            if key == "kab" and nama is None and "nmkab" in df.columns:
                nm = nm_kab.get(v)
                if nm is not None:
                    nm = str(nm).strip()
                    nama = nm if nm.startswith("[") else f"{nv} - {nm}"   # "[01] JEMBRANA" sudah memuat kode
            return str(nama if nama is not None else nv)
        return f

    nm_kab = df.dropna(subset=[c_kab]).groupby(c_kab)["nmkab"].first().to_dict() if "nmkab" in df.columns else {}

    # daftar kelompok: (dimensi, label, mask)
    grup = [("Total", "Total", pd.Series(True, index=df.index))]
    for dim, col, key in (("Jenis Kelamin", c_sex, "sex"), ("Kota/Desa", c_kl, "klas"),
                          ("Kabupaten/Kota", c_kab, "kab")):
        f = lab(col, key)
        for v in sorted(df[col].dropna().unique()):
            grup.append((dim, f(v), df[col] == v))
    fk, fl = lab(c_kab, "kab"), lab(c_kl, "klas")
    for vk in sorted(df[c_kab].dropna().unique()):
        for vl in sorted(df[c_kl].dropna().unique()):
            grup.append(("Kabupaten/Kota x Kota/Desa", f"{fk(vk)} - {fl(vl)}",
                         (df[c_kab] == vk) & (df[c_kl] == vl)))
    print("Jumlah kelompok:", len(grup))

    # ---------- 1. TPT & TPAK ----------
    rows = []
    for dim, nm, m in grup:
        d = df[m]; w = d["w"]
        puk = w.sum()
        bek = w[d["ak3"] == 1].sum(); png = w[d["ak3"] == 2].sum(); bak = w[d["ak3"] == 3].sum()
        ak = bek + png
        rows.append({"Dimensi": dim, "Kelompok": nm, "Sampel (n)": len(d),
                     f"Penduduk {UMUR_MIN}+": puk, "Bekerja": bek, "Pengangguran": png,
                     "Angkatan Kerja": ak, "Bukan Angkatan Kerja": bak,
                     "TPAK (%)": ak / puk * 100 if puk else np.nan,
                     "TPT (%)": png / ak * 100 if ak else np.nan})
    t_ind = pd.DataFrame(rows)

    # ---------- 2 & 3. Status pekerjaan, Lapangan usaha ----------
    def sebaran(var, kamus):
        jml = {}
        for dim, nm, m in grup:
            d = df[m & (df["ak3"] == 1)]
            jml[(dim, nm)] = d.groupby(var)["w"].sum().reindex(list(kamus)).fillna(0)
        J = pd.DataFrame(jml); J.index = [kamus[i] for i in J.index]
        J.loc["Total"] = J.sum()
        P = J / J.loc["Total"] * 100
        return J, P
    st_j, st_p = sebaran("status_pekerjaan", STATUS)
    lu_j, lu_p = sebaran("lu17", LU)

    # ---------- 4. Upah ----------
    if MODE_UPAH == "karyawan":
        mask_up = (df["status_pekerjaan"] == 4) & (df["upah"] > 0)
        nilai = "upah"; judul = "Rata-rata Upah/Gaji Sebulan - Buruh/Karyawan/Pegawai"
    else:
        df["pendapatan"] = df[["upah", "laba"]].sum(axis=1, min_count=1)
        mask_up = (df["ak3"] == 1) & (df["pendapatan"] > 0)
        nilai = "pendapatan"; judul = "Rata-rata Pendapatan Sebulan - Semua Pekerja"
    du = df[mask_up]

    def stat_upah(d):
        w = d["w"]
        if len(d) == 0:
            return pd.Series({"Sampel (n)": 0, "Pekerja (terbobot)": 0, "Rata-rata": np.nan, "Median": np.nan})
        return pd.Series({"Sampel (n)": len(d), "Pekerja (terbobot)": w.sum(),
                          "Rata-rata": np.average(d[nilai], weights=w), "Median": wmedian(d[nilai], w)})
    rows = []
    for dim, nm, m in grup:
        r = stat_upah(du[m.loc[du.index]]); r["Dimensi"] = dim; r["Kelompok"] = nm; rows.append(r)
    t_up = pd.DataFrame(rows)[["Dimensi", "Kelompok", "Sampel (n)", "Pekerja (terbobot)", "Rata-rata", "Median"]]
    rows = []
    for k, nm in LU.items():
        d = du[du["lu17"] == k]; r = stat_upah(d); r["Lapangan Usaha"] = nm; rows.append(r)
    t_up_lu = pd.DataFrame(rows)[["Lapangan Usaha", "Sampel (n)", "Pekerja (terbobot)", "Rata-rata", "Median"]]

    # ---------- 5. Cek: upah karyawan saja atau plus pengusaha? ----------
    rows = []
    for k, nm in STATUS.items():
        d = df[(df["status_pekerjaan"] == k)]
        w = d["w"]
        def wm(col):
            x = d[d[col] > 0]
            return np.average(x[col], weights=x["w"]) if len(x) else np.nan
        rows.append({"Status Pekerjaan": nm, "Pekerja (terbobot)": w.sum(), "Sampel (n)": len(d),
                     "n upah terisi >0": int((d["upah"] > 0).sum()),
                     "% upah terisi": (d["upah"] > 0).mean() * 100 if len(d) else np.nan,
                     "Rata-rata upah": wm("upah"),
                     "n laba terisi >0": int((d["laba"] > 0).sum()),
                     "% laba terisi": (d["laba"] > 0).mean() * 100 if len(d) else np.nan,
                     "Rata-rata laba": wm("laba")})
    t_cek = pd.DataFrame(rows)

    # ---------- Tulis Excel ----------
    t_ind[["TPAK (%)", "TPT (%)"]] = t_ind[["TPAK (%)", "TPT (%)"]].round(2)
    st_p, lu_p = st_p.round(2), lu_p.round(2)
    t_cek[["% upah terisi", "% laba terisi"]] = t_cek[["% upah terisi", "% laba terisi"]].round(2)

    with pd.ExcelWriter(OUT_XLSX, engine="xlsxwriter") as xw:
        wb = xw.book
        F = {"judul": wb.add_format({"bold": True, "font_size": 12}),
             "head": wb.add_format({"bold": True, "text_wrap": True, "align": "center", "valign": "vcenter",
                                    "top": 1, "bottom": 1, "bg_color": "#DDEBF7"}),
             "int": wb.add_format({"num_format": "#,##0"}),
             "dec": wb.add_format({"num_format": "0.00"})}       # persentase: 2 digit desimal
        lembar = {}

        def sheet(nm):
            if nm not in lembar:
                lembar[nm] = wb.add_worksheet(nm)
            return lembar[nm]

        def tabel(nm, row, judul_, df_, index=False, pct_cols=(), pct_all=False):
            ws = sheet(nm)
            ws.write(row, 0, judul_, F["judul"])
            r = row + 1
            multi = isinstance(df_.columns, pd.MultiIndex)
            c0 = 1 if index else 0
            nm_idx = str(df_.index.name or "")
            if index:
                ws.write(r, 0, "" if multi else nm_idx, F["head"])
            if multi:
                if index:
                    ws.write(r + 1, 0, nm_idx, F["head"])
                prev = None
                for k, col in enumerate(df_.columns):
                    ws.write(r, c0 + k, col[0] if col[0] != prev else "", F["head"])
                    ws.write(r + 1, c0 + k, col[1], F["head"])
                    prev = col[0]
                ws.set_row(r + 1, 32)
                r += 2
            else:
                for k, col in enumerate(df_.columns):
                    ws.write(r, c0 + k, str(col), F["head"])
                ws.set_row(r, 32)
                r += 1
            for a in range(len(df_)):
                if index:
                    ws.write(r + a, 0, str(df_.index[a]))
                for k, col in enumerate(df_.columns):
                    v = df_.iat[a, k]
                    pct = pct_all or (col in pct_cols)
                    if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool):
                        if pd.isna(v) or np.isinf(v):
                            ws.write_blank(r + a, c0 + k, None)
                        else:
                            ws.write_number(r + a, c0 + k, float(v), F["dec"] if pct else F["int"])
                    elif pd.isna(v):
                        ws.write_blank(r + a, c0 + k, None)
                    else:
                        ws.write(r + a, c0 + k, str(v))
            return r + len(df_) + 2

        tabel("TPT_TPAK", 0, f"TPT dan TPAK (penimbang; penduduk {UMUR_MIN}+ tahun)", t_ind,
              pct_cols=("TPAK (%)", "TPT (%)"))
        if REF_2025:
            tabel("Perbandingan_2025_2026", 0, JUDUL_BANDING, buat_perbandingan(t_ind), index=True, pct_all=True)
        r = tabel("Status_Pekerjaan", 0, "Jumlah Penduduk Bekerja menurut Status Pekerjaan (jiwa, terbobot)", st_j, index=True)
        tabel("Status_Pekerjaan", r, "Persentase menurut Status Pekerjaan (%)", st_p, index=True, pct_all=True)
        r = tabel("Lapangan_Usaha", 0, "Jumlah Penduduk Bekerja menurut Lapangan Usaha (jiwa, terbobot)", lu_j, index=True)
        tabel("Lapangan_Usaha", r, "Persentase menurut Lapangan Usaha (%)", lu_p, index=True, pct_all=True)
        r = tabel("Upah", 0, judul + " (Rp)", t_up)
        tabel("Upah", r, judul + " menurut Lapangan Usaha (Rp)", t_up_lu)
        tabel("Cek_Upah", 0, "Cek: variabel upah vs laba menurut status pekerjaan", t_cek,
              pct_cols=("% upah terisi", "% laba terisi"))

        for nm, ws in lembar.items():
            ws.set_column(0, 0, 55 if nm != "TPT_TPAK" else 28)
            ws.set_column(1, 1, 40 if nm in ("TPT_TPAK", "Upah") else 16)
            ws.set_column(2, 80, 16)
    print("Selesai ->", OUT_XLSX)


if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:                       # pesan dari sys.exit(...) di dalam skrip
        if e.code not in (None, 0):
            print(e.code)
    except PermissionError as e:
        print(f"\n[GAGAL] Akses file ditolak (kemungkinan file masih terbuka di Excel/SPSS):\n    {e.filename}\n"
              f"Tutup file tersebut, lalu jalankan skrip lagi.\n")
    except FileNotFoundError as e:
        print(f"\n[GAGAL] File/folder tidak ditemukan:\n    {e.filename}\n"
              f"Periksa SAV_PATH, OUT_XLSX, dan OUT_SAV di bagian PENGATURAN.\n")
    except Exception:
        import traceback
        print("\n[GAGAL] Terjadi kesalahan tak terduga. Salin pesan di bawah ini dan kirim ke yang membantu:\n")
        traceback.print_exc()
    _tutup()
