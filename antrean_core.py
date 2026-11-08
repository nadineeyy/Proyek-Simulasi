"""
antrean_core.py  -  Inti perhitungan untuk dua program (aplikasi jendela dan web).

Alur:
  data mentah (Excel/CSV) -> waktu antar kedatangan (WAK), waktu pelayanan 1 (WL1),
  waktu pelayanan 2 (WL2) -> rata-rata -> uji distribusi (Kolmogorov-Smirnov dengan
  koreksi Lilliefors, Monte Carlo) -> pilih distribusi terbaik -> simulasi antrean
  dua fase -> ukuran kinerja (rho, P0, Lq, L, Wq, W).

Data mentah yang dibutuhkan per pelanggan (4 kolom waktu, format jam HH:MM:SS):
  WD   = waktu kedatangan
  WML1 = waktu mulai dilayani di kasir
  WSL1 = waktu selesai di kasir (= masuk fase 2)
  WSL2 = waktu selesai fase 2 (pesanan jadi)
Dari sini: WAK = selisih WD berurutan, WL1 = WSL1 - WML1, WL2 = WSL2 - WSL1.
"""

import datetime as dt
import heapq
import io
import math
import re

import numpy as np
import pandas as pd
from scipy import special, stats, optimize

PERAN = ["WD", "WML1", "WSL1", "WSL2"]
KETERANGAN_PERAN = {
    "WD": "Waktu kedatangan",
    "WML1": "Waktu mulai dilayani di kasir",
    "WSL1": "Waktu selesai di kasir (masuk fase 2)",
    "WSL2": "Waktu selesai fase 2 (pesanan jadi)",
}
ALPHA = 0.05


# ======================================================================
# 1. MEMBACA FILE
# ======================================================================
def daftar_sheet(sumber, nama_file):
    """Nama sheet jika file Excel, atau ['(CSV)'] jika CSV."""
    if nama_file.lower().endswith((".xlsx", ".xlsm", ".xls")):
        _rewind(sumber)
        return pd.ExcelFile(sumber).sheet_names
    return ["(CSV)"]


def _rewind(sumber):
    if hasattr(sumber, "seek"):
        sumber.seek(0)


def baca_mentah(sumber, nama_file, sheet=None):
    """Baca file apa adanya (tanpa menentukan baris judul). Mengembalikan DataFrame tanpa header."""
    _rewind(sumber)
    if nama_file.lower().endswith((".xlsx", ".xlsm", ".xls")):
        return pd.read_excel(sumber, sheet_name=sheet or 0, header=None)
    if isinstance(sumber, (str, bytes)) or not hasattr(sumber, "read"):
        return pd.read_csv(sumber, header=None, sep=None, engine="python")
    isi = sumber.read()
    if isinstance(isi, bytes):
        isi = isi.decode("utf-8-sig", errors="replace")
    return pd.read_csv(io.StringIO(isi), header=None, sep=None, engine="python")


def tebak_baris_judul(mentah):
    """Baris (mulai 1) yang paling mirip baris judul kolom: yang paling banyak sel teksnya."""
    terbaik, skor_terbaik = 1, -1
    for i in range(min(15, len(mentah))):
        skor = sum(isinstance(v, str) and not _mirip_waktu(v) for v in mentah.iloc[i].tolist())
        if skor > skor_terbaik:
            terbaik, skor_terbaik = i + 1, skor
    return terbaik


def _mirip_waktu(teks):
    return bool(re.fullmatch(r"\s*\d{1,2}[:.]\d{2}([:.]\d{2})?\s*", str(teks)))


def terapkan_baris_judul(mentah, baris_judul):
    """Jadikan baris ke-`baris_judul` (mulai 1) sebagai judul kolom, sisanya data."""
    judul = []
    for j, v in enumerate(mentah.iloc[baris_judul - 1].tolist()):
        judul.append(str(v).strip() if pd.notna(v) else f"(kolom {j + 1})")
    df = mentah.iloc[baris_judul:].copy()
    df.columns = _unik(judul)
    return df.reset_index(drop=True)


def _unik(nama):
    hasil, hitung = [], {}
    for n in nama:
        hitung[n] = hitung.get(n, 0) + 1
        hasil.append(n if hitung[n] == 1 else f"{n} ({hitung[n]})")
    return hasil


def tebak_kolom(df):
    """Tebak kolom mana untuk WD, WML1, WSL1, WSL2 dari judul kolom."""
    kolom = list(df.columns)
    peta = {}
    # a) dari kode di awal judul, misalnya "WSL1 (Waktu Selesai ...)"
    for k in kolom:
        kode = re.split(r"[\s(\[]", str(k).strip().upper())[0]
        if kode in PERAN and kode not in peta:
            peta[kode] = k
    if len(peta) == 4:
        return peta
    # b) dari kata kunci
    peta = {}
    selesai = []
    for k in kolom:
        s = str(k).lower()
        if "datang" in s and "WD" not in peta:
            peta["WD"] = k
        elif "mulai" in s and "WML1" not in peta:
            peta["WML1"] = k
        elif "selesai" in s:
            selesai.append(k)
    if len(selesai) >= 2:
        peta["WSL1"], peta["WSL2"] = selesai[0], selesai[1]
    if len(peta) == 4:
        return peta
    # c) 4 kolom pertama yang isinya berupa waktu
    kandidat = [k for k in kolom if _persen_waktu(df[k]) > 0.8]
    return {p: (kandidat[i] if i < len(kandidat) else None) for i, p in enumerate(PERAN)}


def _persen_waktu(seri):
    nilai = seri.dropna()
    if len(nilai) == 0:
        return 0.0
    return float(np.mean([_satu_detik(v) is not None for v in nilai]))


def _satu_detik(v):
    """Ubah satu nilai waktu menjadi detik sejak tengah malam. None jika bukan waktu."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    if isinstance(v, dt.datetime):
        return v.hour * 3600 + v.minute * 60 + v.second + v.microsecond / 1e6
    if isinstance(v, dt.time):
        return v.hour * 3600 + v.minute * 60 + v.second + v.microsecond / 1e6
    if isinstance(v, (pd.Timedelta, dt.timedelta)):
        return pd.Timedelta(v).total_seconds() % 86400
    if isinstance(v, (int, float, np.integer, np.floating)):
        v = float(v)
        return v * 86400 if 0 <= v < 1 else None   # pecahan hari ala Excel
    m = re.fullmatch(r"\s*(\d{1,2})[:.](\d{2})(?:[:.](\d{2}))?\s*", str(v))
    if m:
        return int(m.group(1)) * 3600 + int(m.group(2)) * 60 + int(m.group(3) or 0)
    return None


def ke_detik(seri):
    return pd.Series([_satu_detik(v) for v in seri], index=seri.index, dtype="float64")


# ======================================================================
# 2. MENGHITUNG WAK, WL1, WL2 DAN RATA-RATA
# ======================================================================
def hitung_waktu(df, peta, ganti_nol=True):
    """
    Dari data mentah ke WAK, WL1, WL2 (detik).
    Mengembalikan (data, catatan). `data` berisi array numpy dan ringkasan observasi.
    """
    catatan = []
    for p in PERAN:
        if peta.get(p) is None:
            raise ValueError(f"Kolom untuk {p} ({KETERANGAN_PERAN[p]}) belum dipilih.")
    if len({peta[p] for p in PERAN}) < 4:
        raise ValueError("Keempat peran (WD, WML1, WSL1, WSL2) harus memakai kolom yang berbeda.")
    t = pd.DataFrame({p: ke_detik(df[peta[p]]) for p in PERAN})
    sebelum = len(t)
    t = t.dropna().reset_index(drop=True)
    if len(t) < sebelum:
        catatan.append(f"{sebelum - len(t)} baris dibuang karena ada waktu yang kosong/tidak terbaca.")
    if len(t) < 10:
        raise ValueError(
            f"Hanya {len(t)} baris data yang terbaca. Cek pilihan kolom dan baris judul "
            "(format waktu harus seperti 17:02:47).")

    urut = t.sort_values("WD", kind="stable").reset_index(drop=True)
    if not urut["WD"].equals(t["WD"]):
        catatan.append("Data tidak berurutan menurut waktu kedatangan, jadi diurutkan otomatis.")
    t = urut

    wak = np.diff(t["WD"].to_numpy())
    wl1 = (t["WSL1"] - t["WML1"]).to_numpy()
    wl2 = (t["WSL2"] - t["WSL1"]).to_numpy()
    wt = (t["WML1"] - t["WD"]).to_numpy()
    total = (t["WSL2"] - t["WD"]).to_numpy()

    for nama, arr in (("WAK", wak), ("WL1", wl1), ("WL2", wl2)):
        if np.ptp(arr) == 0:
            raise ValueError(f"Semua nilai {nama} sama, jadi distribusinya tidak bisa diuji. "
                             "Cek pilihan kolom.")
    for nama, arr in (("WL1", wl1), ("WL2", wl2), ("WT", wt)):
        if (arr < 0).any():
            catatan.append(f"PERHATIAN: {int((arr < 0).sum())} nilai {nama} negatif. "
                           "Cek apakah kolom waktu sudah benar.")

    if ganti_nol:
        for nama, arr in (("WAK", wak), ("WL1", wl1), ("WL2", wl2)):
            nol = arr <= 0
            if nol.any():
                catatan.append(f"{int(nol.sum())} nilai {nama} yang 0 (atau negatif) diganti menjadi 1 detik "
                               "(resolusi pencatatan 1 detik; dibutuhkan distribusi yang hanya menerima nilai positif).")
                arr[nol] = 1.0

    data = {
        "n_pelanggan": len(t),
        "WAK": wak, "WL1": wl1, "WL2": wl2,
        "WT": wt, "TOTAL": total,
        "total_wak": float(wak.sum()),
    }
    return data, catatan


def ringkas(arr):
    a = np.asarray(arr, dtype=float)
    return {"n": len(a), "rata-rata": a.mean(), "SD": a.std(ddof=1) if len(a) > 1 else 0.0,
            "min": a.min(), "median": float(np.median(a)), "maks": a.max()}


def tabel_rata_rata(data):
    baris = []
    for kunci, nama in (("WAK", "WAK (waktu antar kedatangan)"),
                        ("WL1", "WL1 (waktu pelayanan 1, kasir)"),
                        ("WL2", "WL2 (waktu pelayanan 2, sampai pesanan jadi)"),
                        ("WT", "WT (waktu tunggu di antrean kasir)"),
                        ("TOTAL", "Total waktu di sistem")):
        r = ringkas(data[kunci])
        baris.append({"Variabel": nama, "N": r["n"], "Rata-rata (detik)": round(r["rata-rata"], 2),
                      "SD": round(r["SD"], 2), "Min": round(r["min"], 2),
                      "Median": round(r["median"], 2), "Maks": round(r["maks"], 2)})
    return pd.DataFrame(baris)


def tabel_laju(data):
    """Hitung manual seperti artikel acuan: lambda, mu, rho (per jam)."""
    lam = 3600 / data["WAK"].mean()
    mu1 = 3600 / data["WL1"].mean()
    mu2 = 3600 / data["WL2"].mean()
    return pd.DataFrame([
        {"Ukuran": "Laju kedatangan (lambda), pelanggan/jam", "Nilai": round(lam, 4)},
        {"Ukuran": "Laju pelayanan 1 (mu1), pelanggan/jam", "Nilai": round(mu1, 4)},
        {"Ukuran": "Laju pelayanan 2 (mu2), pelanggan/jam", "Nilai": round(mu2, 4)},
        {"Ukuran": "Tingkat pemanfaatan kasir (rho1 = lambda/mu1)", "Nilai": round(lam / mu1, 4)},
    ])


# ======================================================================
# 3. UJI DISTRIBUSI  (Kolmogorov-Smirnov + koreksi Lilliefors, Monte Carlo)
# ======================================================================
NAMA_DISTRIBUSI = ["Normal", "Eksponensial", "Lognormal", "Gamma", "Weibull", "Uniform"]


def _fit_normal(x):
    return {"mean": float(np.mean(x)), "sd": float(np.std(x, ddof=1))}


def _fit_eksp(x):
    return {"mean": float(np.mean(x))}


def _fit_lognormal(x):
    lx = np.log(x)
    return {"mu": float(np.mean(lx)), "sigma": float(np.std(lx, ddof=1))}


def _fit_gamma(x):
    rata, s = np.mean(x), np.log(np.mean(x)) - np.mean(np.log(x))
    if s <= 1e-12:
        s = 1e-12
    k = (3 - s + math.sqrt((s - 3) ** 2 + 24 * s)) / (12 * s)
    for _ in range(8):   # perbaikan Newton
        k = max(k - (math.log(k) - special.digamma(k) - s) / (1 / k - special.polygamma(1, k)), 1e-6)
    return {"shape": float(k), "scale": float(rata / k)}


def _fit_weibull(x):
    y = x / np.mean(x)
    ly = np.log(y)

    def f(k):
        yk = y ** k
        return np.sum(yk * ly) / np.sum(yk) - 1 / k - np.mean(ly)

    try:
        k = optimize.brentq(f, 0.02, 200)
    except ValueError:
        k = 1.0
    skala = float(np.mean(x)) * float(np.mean(y ** k)) ** (1 / k)
    return {"shape": float(k), "scale": skala}


def _fit_uniform(x):
    return {"min": float(np.min(x)), "maks": float(np.max(x))}


def _cdf(nama, p, x):
    if nama == "Normal":
        return stats.norm.cdf(x, p["mean"], max(p["sd"], 1e-12))
    if nama == "Eksponensial":
        return 1 - np.exp(-x / p["mean"])
    if nama == "Lognormal":
        return stats.norm.cdf((np.log(x) - p["mu"]) / max(p["sigma"], 1e-12))
    if nama == "Gamma":
        return special.gammainc(p["shape"], x / p["scale"])
    if nama == "Weibull":
        return 1 - np.exp(-((x / p["scale"]) ** p["shape"]))
    if nama == "Uniform":
        return np.clip((x - p["min"]) / max(p["maks"] - p["min"], 1e-12), 0, 1)
    raise ValueError(nama)


def _sampel(nama, p, n, rng):
    if nama == "Normal":
        return rng.normal(p["mean"], max(p["sd"], 1e-12), n)
    if nama == "Eksponensial":
        return rng.exponential(p["mean"], n)
    if nama == "Lognormal":
        return np.exp(rng.normal(p["mu"], max(p["sigma"], 1e-12), n))
    if nama == "Gamma":
        return rng.gamma(p["shape"], p["scale"], n)
    if nama == "Weibull":
        return p["scale"] * rng.weibull(p["shape"], n)
    if nama == "Uniform":
        return rng.uniform(p["min"], p["maks"], n)
    raise ValueError(nama)


_FIT = {"Normal": _fit_normal, "Eksponensial": _fit_eksp, "Lognormal": _fit_lognormal,
        "Gamma": _fit_gamma, "Weibull": _fit_weibull, "Uniform": _fit_uniform}
_BUTUH_POSITIF = {"Eksponensial", "Lognormal", "Gamma", "Weibull"}


def _statistik(nama, p, x):
    """Mengembalikan (D Kolmogorov-Smirnov, A2 Anderson-Darling)."""
    xs = np.sort(x)
    n = len(xs)
    F = np.clip(_cdf(nama, p, xs), 1e-12, 1 - 1e-12)
    i = np.arange(1, n + 1)
    d = float(max(np.max(i / n - F), np.max(F - (i - 1) / n)))
    a2 = float(-n - np.sum((2 * i - 1) * (np.log(F) + np.log(1 - F[::-1]))) / n)
    return d, a2


def uji_distribusi(x, n_mc=1000, seed=2026, kandidat=None):
    """
    Uji kecocokan x terhadap beberapa distribusi dengan dua statistik:
    Kolmogorov-Smirnov (seperti SPSS) dan Anderson-Darling (seperti Minitab).
    p-value dihitung dengan Monte Carlo (koreksi Lilliefors: parameter diestimasi ulang
    di tiap sampel), sama prinsipnya dengan kolom 'Monte Carlo Sig.' di SPSS.
    """
    x = np.asarray(x, dtype=float)
    rng = np.random.default_rng(seed)
    hasil = []
    for nama in (kandidat or NAMA_DISTRIBUSI):
        if nama in _BUTUH_POSITIF and np.min(x) <= 0:
            hasil.append({"distribusi": nama, "ks": None, "ad": None, "p_ks": None, "p_ad": None,
                          "param": None, "catatan": "dilewati: data berisi nilai <= 0"})
            continue
        param = _FIT[nama](x)
        d_obs, a_obs = _statistik(nama, param, x)
        lebih_d = lebih_a = 0
        sah = 0
        for _ in range(n_mc):
            xb = _sampel(nama, param, len(x), rng)
            if nama in _BUTUH_POSITIF:
                xb = np.maximum(xb, 1e-9)
            try:
                pb = _FIT[nama](xb)
                db, ab = _statistik(nama, pb, xb)
            except Exception:
                continue
            sah += 1
            lebih_d += db >= d_obs
            lebih_a += ab >= a_obs
        hasil.append({"distribusi": nama, "ks": d_obs, "ad": a_obs,
                      "p_ks": (lebih_d + 1) / (sah + 1), "p_ad": (lebih_a + 1) / (sah + 1),
                      "param": param, "catatan": ""})
    return hasil


TAHAP1 = ("Normal", "Eksponensial", "Uniform")   # diuji dengan K-S (seperti di SPSS)
METODE = {
    "BERTAHAP": "Bertahap seperti artikel acuan: K-S untuk Normal/Eksponensial/Uniform; "
                "kalau tidak ada yang lolos, Anderson-Darling untuk Lognormal/Gamma/Weibull",
    "KS": "Semua distribusi diuji dengan Kolmogorov-Smirnov",
    "AD": "Semua distribusi diuji dengan Anderson-Darling",
}


def _p_penentu(h, metode):
    """(p-value yang dipakai memutuskan, nama uji) untuk satu hasil uji."""
    if metode == "KS" or (metode == "BERTAHAP" and h["distribusi"] in TAHAP1):
        return h["p_ks"], "KS"
    return h["p_ad"], "AD"


def pilih_distribusi(hasil_uji, data_asli, metode="BERTAHAP"):
    """
    Pilih distribusi yang cocok (p > 0,05) dengan p-value terbesar.
    Jika tidak ada yang lolos, pakai data empiris (ambil acak dari data asli).
    """
    def terbaik(daftar):
        lolos = [(h, _p_penentu(h, metode)[0]) for h in daftar]
        lolos = [(h, p) for h, p in lolos if p is not None and p > ALPHA]
        return max(lolos, key=lambda hp: hp[1]) if lolos else None

    if metode == "BERTAHAP":
        t = terbaik([h for h in hasil_uji if h["distribusi"] in TAHAP1]) or \
            terbaik([h for h in hasil_uji if h["distribusi"] not in TAHAP1])
    else:
        t = terbaik(hasil_uji)
    if t:
        return {"distribusi": t[0]["distribusi"], "param": t[0]["param"], "p": t[1],
                "lolos": True, "catatan": ""}
    return {"distribusi": "Empiris", "param": {"data": list(map(float, data_asli))}, "p": None,
            "lolos": False,
            "catatan": "Tidak ada distribusi yang lolos uji (p > 0,05). Simulasi memakai data empiris "
                       "(mengambil nilai acak dari data observasi)."}


def teks_param(pilihan):
    p, d = pilihan["param"], pilihan["distribusi"]
    if d == "Normal":
        return f"mean {p['mean']:.2f}; SD {p['sd']:.3f}"
    if d == "Eksponensial":
        return f"mean {p['mean']:.2f}"
    if d == "Lognormal":
        return f"mu {p['mu']:.5f}; sigma {p['sigma']:.5f}"
    if d in ("Gamma", "Weibull"):
        return f"shape {p['shape']:.4f}; scale {p['scale']:.2f}"
    if d == "Uniform":
        return f"min {p['min']:.2f}; maks {p['maks']:.2f}"
    return f"{len(p['data'])} data observasi"


def tabel_uji(hasil_uji, pilihan, metode="BERTAHAP"):
    baris = []
    for h in hasil_uji:
        p, uji = _p_penentu(h, metode)
        if p is None:
            baris.append({"Distribusi": h["distribusi"], "Statistik KS": "-", "p-value KS": "-",
                          "Statistik AD": "-", "p-value AD": "-", "Uji penentu": "-",
                          "Keputusan": h["catatan"], "Parameter": "-"})
        else:
            ket = "cocok" if p > ALPHA else "ditolak"
            if pilihan["lolos"] and h["distribusi"] == pilihan["distribusi"]:
                ket += "  <- DIPILIH"
            baris.append({"Distribusi": h["distribusi"],
                          "Statistik KS": round(h["ks"], 4), "p-value KS": round(h["p_ks"], 4),
                          "Statistik AD": round(h["ad"], 4), "p-value AD": round(h["p_ad"], 4),
                          "Uji penentu": uji, "Keputusan": ket,
                          "Parameter": teks_param({"distribusi": h["distribusi"], "param": h["param"]})})
    return pd.DataFrame(baris)


# ======================================================================
# 4. SIMULASI ANTREAN DUA FASE
# ======================================================================
def buat_pembangkit(pilihan, rng):
    d, p = pilihan["distribusi"], pilihan["param"]
    if d == "Normal":
        return lambda: max(1.0, rng.normal(p["mean"], p["sd"]))
    if d == "Eksponensial":
        return lambda: rng.exponential(p["mean"])
    if d == "Lognormal":
        return lambda: float(np.exp(rng.normal(p["mu"], p["sigma"])))
    if d == "Gamma":
        return lambda: rng.gamma(p["shape"], p["scale"])
    if d == "Weibull":
        return lambda: p["scale"] * rng.weibull(p["shape"])
    if d == "Uniform":
        return lambda: rng.uniform(p["min"], p["maks"])
    if d == "Empiris":
        arr = np.asarray(p["data"], dtype=float)
        return lambda: float(arr[rng.integers(len(arr))])
    raise ValueError(d)


def _satu_replikasi(gen_wak, gen_wl1, gen_wl2, T, c1, c2, wl2_sudah_tunggu):
    # kedatangan
    arr, t = [], 0.0
    while True:
        t += gen_wak()
        if t > T:
            break
        arr.append(t)
    n = len(arr)
    if n < 2:
        return None

    # fase 1: c1 server, FIFO
    bebas = [0.0] * c1
    heapq.heapify(bebas)
    wl1, mulai1, selesai1 = [0.0] * n, [0.0] * n, [0.0] * n
    for i in range(n):
        wl1[i] = gen_wl1()
        mulai1[i] = max(arr[i], heapq.heappop(bebas))
        selesai1[i] = mulai1[i] + wl1[i]
        heapq.heappush(bebas, selesai1[i])

    # fase 2: masuk sesuai urutan selesai fase 1
    urut = sorted(range(n), key=lambda i: selesai1[i])
    wl2, mulai2, selesai2 = [0.0] * n, [0.0] * n, [0.0] * n
    if wl2_sudah_tunggu:
        for i in urut:
            wl2[i] = gen_wl2()
            mulai2[i] = selesai1[i]
            selesai2[i] = selesai1[i] + wl2[i]
    else:
        bebas2 = [0.0] * c2
        heapq.heapify(bebas2)
        for i in urut:
            wl2[i] = gen_wl2()
            mulai2[i] = max(selesai1[i], heapq.heappop(bebas2))
            selesai2[i] = mulai2[i] + wl2[i]
            heapq.heappush(bebas2, selesai2[i])

    arr_, m1, s1 = np.array(arr), np.array(mulai1), np.array(selesai1)
    m2, s2 = np.array(mulai2), np.array(selesai2)
    wl1_, wl2_ = np.array(wl1), np.array(wl2)

    t_akhir = s2.max()
    win_sis = t_akhir - arr_[0]
    win_ksr = s1.max() - arr_[0]
    win_f2 = t_akhir - s1.min()
    wq1 = m1 - arr_
    wq2 = m2 - s1
    wsis = s2 - arr_

    rho1 = wl1_.sum() / (c1 * win_ksr)
    return {
        "rho1": rho1,
        "rho2": None if wl2_sudah_tunggu else wl2_.sum() / (c2 * win_f2),
        "p0": 1 - rho1,
        "lq1": wq1.sum() / win_sis,
        "lq2": 0.0 if wl2_sudah_tunggu else wq2.sum() / win_sis,
        "l": wsis.sum() / win_sis,
        "wq1": wq1.mean(),
        "wq2": 0.0 if wl2_sudah_tunggu else wq2.mean(),
        "wl1": wl1_.mean(),
        "wl2": wl2_.mean(),
        "w": wsis.mean(),
        "n": n,
    }


LABEL_UKURAN = [
    ("rho1", "Utilisasi kasir (rho1)", ""),
    ("rho2", "Utilisasi tiap pekerja fase 2 (rho2)", ""),
    ("p0", "Peluang kasir menganggur (P0)", ""),
    ("lq1", "Rata-rata pelanggan antre di kasir (Lq1)", "orang"),
    ("lq2", "Rata-rata pelanggan antre menunggu pekerja fase 2 (Lq2)", "orang"),
    ("l", "Rata-rata pelanggan di seluruh sistem (L)", "orang"),
    ("wq1", "Rata-rata waktu tunggu di antrean kasir (Wq1)", "detik"),
    ("wq2", "Rata-rata waktu tunggu pekerja fase 2 (Wq2)", "detik"),
    ("wl1", "Rata-rata lama di kasir (WL1)", "detik"),
    ("wl2", "Rata-rata lama pembuatan pesanan (WL2)", "detik"),
    ("w", "Rata-rata total waktu di sistem (W)", "detik"),
    ("n", "Rata-rata jumlah pelanggan per replikasi", "orang"),
]


def simulasi(pil_wak, pil_wl1, pil_wl2, T, c1=1, c2=4, replikasi=100,
             wl2_sudah_tunggu=False, seed=2026):
    rng = np.random.default_rng(seed)
    gw, g1, g2 = (buat_pembangkit(pil_wak, rng), buat_pembangkit(pil_wl1, rng),
                  buat_pembangkit(pil_wl2, rng))
    semua = []
    for _ in range(replikasi):
        r = _satu_replikasi(gw, g1, g2, T, c1, c2, wl2_sudah_tunggu)
        if r is not None:
            semua.append(r)
    if not semua:
        raise ValueError("Simulasi tidak menghasilkan pelanggan. Naikkan lama simulasi.")
    baris = []
    for kunci, label, satuan in LABEL_UKURAN:
        if kunci == "rho2" and wl2_sudah_tunggu:
            baris.append({"Ukuran kinerja": label, "Rata-rata": "-", "CI 95% (+/-)": "-", "Satuan": satuan})
            continue
        v = np.array([r[kunci] for r in semua], dtype=float)
        ci = 1.96 * v.std(ddof=1) / math.sqrt(len(v)) if len(v) > 1 else 0.0
        baris.append({"Ukuran kinerja": label, "Rata-rata": round(float(v.mean()), 4),
                      "CI 95% (+/-)": round(float(ci), 4), "Satuan": satuan})
    return pd.DataFrame(baris), len(semua)


# ======================================================================
# 5. SEMUA LANGKAH SEKALIGUS + RINGKASAN KATA-KATA
# ======================================================================
def jalankan_semua(df, peta, opsi, progres=None):
    """
    opsi: dict dengan kunci ganti_nol, n_mc, metode ("BERTAHAP", "KS" atau "AD"), kandidat (daftar nama distribusi
          atau None = semua), c1, c2, replikasi, wl2_sudah_tunggu, T (None = otomatis), seed
    Mengembalikan dict hasil siap tampil.
    """
    def lapor(teks):
        if progres:
            progres(teks)

    lapor("Menghitung WAK, WL1, WL2 ...")
    data, catatan = hitung_waktu(df, peta, ganti_nol=opsi.get("ganti_nol", True))

    pilihan, tabel_ujian = {}, {}
    for kunci in ("WAK", "WL1", "WL2"):
        lapor(f"Menguji distribusi {kunci} ...")
        metode = opsi.get("metode", "BERTAHAP")
        h = uji_distribusi(data[kunci], n_mc=opsi.get("n_mc", 1000), seed=opsi.get("seed", 2026),
                           kandidat=opsi.get("kandidat"))
        p = pilih_distribusi(h, data[kunci], metode)
        pilihan[kunci] = p
        tabel_ujian[kunci] = tabel_uji(h, p, metode)
        if p["catatan"]:
            catatan.append(f"{kunci}: {p['catatan']}")

    T = opsi.get("T") or data["total_wak"]
    lapor("Menjalankan simulasi ...")
    hasil_sim, n_rep = simulasi(pilihan["WAK"], pilihan["WL1"], pilihan["WL2"], T,
                                c1=opsi.get("c1", 1), c2=opsi.get("c2", 4),
                                replikasi=opsi.get("replikasi", 100),
                                wl2_sudah_tunggu=opsi.get("wl2_sudah_tunggu", False),
                                seed=opsi.get("seed", 2026))

    tabel_pilihan = pd.DataFrame([
        {"Variabel": k, "Distribusi terpilih": pilihan[k]["distribusi"],
         "Parameter": teks_param(pilihan[k]),
         "p-value": "-" if pilihan[k]["p"] is None else round(pilihan[k]["p"], 4)}
        for k in ("WAK", "WL1", "WL2")])

    obs = pd.DataFrame([
        {"Ukuran": "Rata-rata waktu tunggu di kasir (WT)", "Observasi (detik)": round(data["WT"].mean(), 2),
         "Simulasi (detik)": float(hasil_sim.loc[hasil_sim["Ukuran kinerja"].str.contains("Wq1"), "Rata-rata"].iloc[0])},
        {"Ukuran": "Rata-rata total waktu di sistem", "Observasi (detik)": round(data["TOTAL"].mean(), 2),
         "Simulasi (detik)": float(hasil_sim.loc[hasil_sim["Ukuran kinerja"].str.contains(r"\(W\)"), "Rata-rata"].iloc[0])},
    ])

    return {
        "data": data, "catatan": catatan,
        "rata_rata": tabel_rata_rata(data), "laju": tabel_laju(data),
        "uji": tabel_ujian, "pilihan": tabel_pilihan, "pilihan_mentah": pilihan,
        "simulasi": hasil_sim, "replikasi_dipakai": n_rep, "T": T,
        "pembanding": obs, "ringkasan": buat_ringkasan(data, hasil_sim, opsi, pilihan),
    }


def _ambil(tabel, potongan):
    v = tabel.loc[tabel["Ukuran kinerja"].str.contains(potongan, regex=False), "Rata-rata"].iloc[0]
    return None if v == "-" else float(v)


def buat_ringkasan(data, hasil_sim, opsi, pilihan):
    wq1 = _ambil(hasil_sim, "(Wq1)")
    wq2 = _ambil(hasil_sim, "(Wq2)")
    w = _ambil(hasil_sim, "(W)")
    rho1 = _ambil(hasil_sim, "(rho1)")
    wl1, wl2 = data["WL1"].mean(), data["WL2"].mean()
    kalimat = [
        f"Dari {data['n_pelanggan']} pelanggan yang diamati, rata-rata waktu antar kedatangan adalah "
        f"{data['WAK'].mean():.1f} detik, rata-rata pelayanan di kasir {wl1:.1f} detik, dan rata-rata "
        f"waktu sampai pesanan jadi {wl2:.1f} detik.",
        f"Pada data observasi, pelanggan menunggu rata-rata {data['WT'].mean():.1f} detik di antrean kasir "
        f"dan menghabiskan total {data['TOTAL'].mean():.1f} detik di sistem.",
        f"Hasil simulasi menunjukkan kasir sibuk sekitar {rho1 * 100:.1f}% waktu, dengan rata-rata waktu "
        f"tunggu di antrean kasir {wq1:.1f} detik dan rata-rata total waktu di sistem {w:.1f} detik.",
    ]
    if wq2 is not None and not opsi.get("wl2_sudah_tunggu", False):
        kalimat.append(f"Di fase 2, rata-rata pelanggan masih menunggu {wq2:.1f} detik sampai ada pekerja kosong "
                       "sebelum pesanannya mulai dibuat.")
    bagian = wl2 / data["TOTAL"].mean()
    if bagian > 0.5:
        kalimat.append(f"Sekitar {bagian * 100:.0f}% waktu pelanggan di sistem dihabiskan untuk menunggu pesanan "
                       "selesai (fase 2), jauh lebih besar daripada waktu mengantre di kasir.")
    return " ".join(kalimat)


def ekspor_excel(hasil, jalur):
    """Simpan semua tabel hasil ke satu file Excel."""
    with pd.ExcelWriter(jalur, engine="openpyxl") as w:
        hasil["rata_rata"].to_excel(w, sheet_name="Rata-rata", index=False)
        hasil["laju"].to_excel(w, sheet_name="Laju", index=False)
        for k, tabel in hasil["uji"].items():
            tabel.to_excel(w, sheet_name=f"Uji {k}", index=False)
        hasil["pilihan"].to_excel(w, sheet_name="Distribusi terpilih", index=False)
        hasil["simulasi"].to_excel(w, sheet_name="Simulasi", index=False)
        hasil["pembanding"].to_excel(w, sheet_name="Pembanding", index=False)
        pd.DataFrame({"Ringkasan": [hasil["ringkasan"]] + [""] + hasil["catatan"]}).to_excel(
            w, sheet_name="Ringkasan", index=False)
