"""
app_web.py  -  Web sederhana: simulasi antrean dua fase (tampilan berwarna).

Cara menjalankan di komputer sendiri:
    pip install -r requirements.txt
    streamlit run app_web.py
Browser akan terbuka otomatis. Untuk dipakai orang lain lewat tautan, program ini bisa
diunggah (bersama antrean_core.py dan requirements.txt) ke Streamlit Community Cloud (gratis).
"""

import io
import os

import numpy as np
import pandas as pd
import streamlit as st

import antrean_core as ac

FOLDER = os.path.dirname(os.path.abspath(__file__))

# Palet warna (tema kopi)
KOPI_TUA = "#4E342E"
KOPI = "#8D5A3B"
KARAMEL = "#C68B59"
KRIM = "#FFF8F0"
HIJAU = "#2E7D32"
MERAH = "#C62828"
BIRU = "#1565C0"

st.set_page_config(page_title="Simulasi Antrean Dua Fase", page_icon="☕", layout="wide")

st.markdown(f"""
<style>
.stApp {{ background-color: {KRIM}; color: {KOPI_TUA}; }}
[data-testid="stSidebar"] {{ background-color: #F3E5D8; }}
[data-testid="stSidebar"] * {{ color: {KOPI_TUA}; }}
.banner {{
    background: linear-gradient(120deg, {KOPI_TUA} 0%, {KOPI} 55%, {KARAMEL} 100%);
    color: white; padding: 26px 30px; border-radius: 16px; margin-bottom: 18px;
    box-shadow: 0 4px 14px rgba(78,52,46,0.25);
}}
.banner h1 {{ color: white; margin: 0 0 6px 0; font-size: 2rem; }}
.banner p {{ color: #FBE9D7; margin: 0; font-size: 1rem; }}
.kartu {{
    background: white; border-radius: 14px; padding: 16px 20px; margin: 6px 0 14px 0;
    border-left: 6px solid {KARAMEL}; box-shadow: 0 2px 8px rgba(0,0,0,0.07); color: {KOPI_TUA};
}}
.kartu.hijau {{ border-left-color: {HIJAU}; }}
.kartu.biru {{ border-left-color: {BIRU}; }}
.kartu h4 {{ margin: 0 0 6px 0; color: {KOPI}; }}
.langkah {{
    display: inline-block; background: {KOPI}; color: white; border-radius: 50%;
    width: 28px; height: 28px; line-height: 28px; text-align: center; font-weight: bold; margin-right: 8px;
}}
.judul-bagian {{ color: {KOPI}; font-size: 1.35rem; font-weight: 700; margin: 18px 0 6px 0; }}
[data-testid="stMetric"] {{
    background: white; border-radius: 14px; padding: 14px 16px; border-top: 5px solid {KARAMEL};
    box-shadow: 0 2px 8px rgba(0,0,0,0.07);
}}
[data-testid="stMetricLabel"] p {{ color: {KOPI}; font-weight: 600; }}
[data-testid="stMetricValue"] {{ color: {KOPI_TUA}; }}
button[kind="primary"] {{ background-color: {KOPI}; border-color: {KOPI}; color: white; }}
button[kind="primary"]:hover {{ background-color: {KOPI_TUA}; border-color: {KOPI_TUA}; color: white; }}
.stTabs [data-baseweb="tab-list"] {{ gap: 6px; }}
.stTabs [data-baseweb="tab"] {{
    background: #F3E5D8; border-radius: 10px 10px 0 0; padding: 8px 16px; color: {KOPI_TUA};
}}
.stTabs [aria-selected="true"] {{ background: {KOPI}; color: white; }}
.stTabs [aria-selected="true"] p {{ color: white; }}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="banner">
  <h1>☕ Simulasi Antrean Dua Fase</h1>
  <p>Unggah data mentah (Excel/CSV): program menghitung rata-rata, menguji distribusi,
  memilih yang paling cocok, lalu menyimulasikan antrean kasir dan pembuatan pesanan.</p>
</div>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------- fungsi bantu
def warnai_uji(df):
    """Beri warna pada tabel uji distribusi: hijau = dipilih, merah muda = ditolak."""
    def gaya(baris):
        ket = str(baris["Keputusan"])
        if "DIPILIH" in ket:
            return ["background-color: #C8E6C9; font-weight: 600"] * len(baris)
        if "cocok" in ket:
            return ["background-color: #F1F8E9"] * len(baris)
        if "ditolak" in ket:
            return ["background-color: #FDECEA; color: #8A2B2B"] * len(baris)
        return [""] * len(baris)
    return df.style.apply(gaya, axis=1)


def ambil(tabel, potongan):
    cocok = tabel.loc[tabel["Ukuran kinerja"].str.contains(potongan, regex=False), "Rata-rata"]
    if cocok.empty or cocok.iloc[0] == "-":
        return None
    return float(cocok.iloc[0])


def histogram(nilai, jumlah_bin=8):
    cacah, batas = np.histogram(nilai, bins=jumlah_bin)
    label = [f"{batas[i]:.0f}-{batas[i + 1]:.0f}" for i in range(len(cacah))]
    return pd.DataFrame({"Jumlah": cacah}, index=label)


def bagian(judul):
    st.markdown(f'<div class="judul-bagian">{judul}</div>', unsafe_allow_html=True)


# ---------------------------------------------------------------- sidebar
berkas = None
with st.sidebar:
    st.markdown('<span class="langkah">1</span><b>Data mentah</b>', unsafe_allow_html=True)
    berkas = st.file_uploader("Unggah file Excel (.xlsx) atau CSV", type=["xlsx", "xlsm", "xls", "csv"])
    with st.expander("Format data yang dibutuhkan"):
        st.markdown(
            "Satu baris per pelanggan, dengan 4 kolom waktu berformat jam (misalnya 17:02:47):\n\n"
            "- **WD**: waktu kedatangan\n- **WML1**: waktu mulai dilayani di kasir\n"
            "- **WSL1**: waktu selesai di kasir (masuk fase 2)\n- **WSL2**: waktu selesai fase 2 (pesanan jadi)\n\n"
            "Dari sini dihitung WAK (selisih kedatangan), WL1 = WSL1 - WML1, dan WL2 = WSL2 - WSL1.")

    st.markdown('<span class="langkah">2</span><b>Pengaturan</b>', unsafe_allow_html=True)
    ganti_nol = st.checkbox("Ganti nilai waktu 0 menjadi 1 detik", value=True)
    metode = st.selectbox("Metode uji distribusi", list(ac.METODE.keys()),
                          format_func=lambda k: {"BERTAHAP": "Bertahap",
                                                 "KS": "Kolmogorov-Smirnov saja",
                                                 "AD": "Anderson-Darling saja"}[k])
    st.caption(ac.METODE[metode])
    c1 = st.number_input("Jumlah server fase 1 (kasir)", min_value=1, max_value=50, value=1)
    c2 = st.number_input("Jumlah server fase 2 (pembuat pesanan)", min_value=1, max_value=50, value=4)
    replikasi = st.number_input("Jumlah replikasi simulasi", min_value=1, max_value=2000, value=100, step=10)
    T_teks = st.text_input("Lama simulasi (detik)", value="", help="Kosong = total WAK dari data")
    wl2_tunggu = st.checkbox("WL2 sudah termasuk waktu menunggu pekerja (tanpa antrean tambahan di fase 2)",
                             value=False)

# ---------------------------------------------------------------- data + kolom
if berkas is None:
    st.markdown("""
<div class="kartu">
<h4>Mulai dari sini</h4>
Pilih <b>Unggah file</b> di panel kiri, atau coba dulu dengan <b>Contoh Weekday / Contoh Weekend</b>.
Setelah itu klik tombol <b>Hitung dan simulasikan</b>.
</div>
<div class="kartu hijau">
<h4>Yang dihasilkan program</h4>
1. Rata-rata, simpangan baku, laju kedatangan dan pelayanan<br>
2. Uji distribusi (Kolmogorov-Smirnov dan Anderson-Darling) dan pemilihan distribusi terbaik<br>
3. Simulasi antrean dua fase beserta ukuran kinerja (ρ, L, Lq, Wq, W)<br>
4. Perbandingan dengan data observasi dan ringkasan dalam kalimat
</div>
""", unsafe_allow_html=True)
    st.stop()

try:
    sheet_list = ac.daftar_sheet(berkas, berkas.name)
    sheet = st.sidebar.selectbox("Sheet", sheet_list, disabled=len(sheet_list) == 1)
    mentah = ac.baca_mentah(berkas, berkas.name, None if sheet == "(CSV)" else sheet)
    baris = st.sidebar.number_input("Baris judul kolom", min_value=1, max_value=30,
                                    value=ac.tebak_baris_judul(mentah))
    df = ac.terapkan_baris_judul(mentah, int(baris))
except Exception as e:
    st.error(f"File tidak bisa dibaca: {e}")
    st.stop()

st.sidebar.markdown('<span class="langkah">3</span><b>Pilih kolom waktu</b>', unsafe_allow_html=True)
tebakan = ac.tebak_kolom(df)
opsi_kolom = [""] + list(df.columns)
peta = {}
for p in ac.PERAN:
    awal = tebakan.get(p)
    peta[p] = st.sidebar.selectbox(f"{p}: {ac.KETERANGAN_PERAN[p]}", opsi_kolom,
                                   index=opsi_kolom.index(awal) if awal in opsi_kolom else 0) or None

with st.expander("Lihat data yang terbaca", expanded=False):
    st.dataframe(df.head(50))

# ---------------------------------------------------------------- jalankan
if st.button("Hitung dan simulasikan", type="primary"):
    try:
        T = float(T_teks.replace(",", ".")) if T_teks.strip() else None
        opsi = dict(ganti_nol=ganti_nol, metode=metode, n_mc=10000, c1=int(c1), c2=int(c2),
                    replikasi=int(replikasi), wl2_sudah_tunggu=wl2_tunggu, T=T, seed=2026)
        with st.spinner("Menghitung, menguji distribusi, dan menyimulasikan ..."):
            st.session_state["hasil"] = ac.jalankan_semua(df, peta, opsi)
            st.session_state["metode"] = metode
    except Exception as e:
        st.session_state.pop("hasil", None)
        st.error(str(e))

h = st.session_state.get("hasil")
if h:
    sim, data = h["simulasi"], h["data"]

    # ---- ringkasan dalam kalimat
    st.markdown(f'<div class="kartu"><h4>Ringkasan</h4>{h["ringkasan"]}</div>', unsafe_allow_html=True)
    for c in h["catatan"]:
        st.warning(c)

    # ---- kartu angka utama
    wq1, w, l, rho1 = (ambil(sim, "(Wq1)"), ambil(sim, "(W)"), ambil(sim, "(L)"), ambil(sim, "(rho1)"))
    kolom = st.columns(4)
    kolom[0].metric("Utilisasi kasir (ρ₁)", f"{rho1 * 100:.1f}%")
    kolom[1].metric("Tunggu di kasir (Wq₁)", f"{wq1:.1f} detik",
                    f"observasi {data['WT'].mean():.1f}", delta_color="off")
    kolom[2].metric("Total waktu di sistem (W)", f"{w:.1f} detik",
                    f"observasi {data['TOTAL'].mean():.1f}", delta_color="off")
    kolom[3].metric("Pelanggan di sistem (L)", f"{l:.2f} orang")
    st.caption(f"Lama simulasi {h['T']:.0f} detik, {h['replikasi_dipakai']} replikasi. "
               f"Metode uji: {ac.METODE[st.session_state['metode']]}.")

    # ---- distribusi terpilih (kartu berwarna)
    bagian("Distribusi terpilih")
    kol = st.columns(3)
    for i, k in enumerate(("WAK", "WL1", "WL2")):
        pil = h["pilihan_mentah"][k]
        warna = "hijau" if pil["lolos"] else ""
        p_teks = "" if pil["p"] is None else f"<br>p-value = {pil['p']:.4f}"
        kol[i].markdown(
            f'<div class="kartu {warna}"><h4>{k}</h4><b>{pil["distribusi"]}</b>'
            f'<br>{ac.teks_param(pil)}{p_teks}</div>', unsafe_allow_html=True)

    # ---- tab rincian
    tabs = st.tabs(["📊 Rata-rata", "🧪 Uji distribusi", "🎯 Hasil simulasi",
                    "⚖️ Observasi vs simulasi", "📈 Grafik"])
    with tabs[0]:
        st.dataframe(h["rata_rata"], hide_index=True)
        st.dataframe(h["laju"], hide_index=True)
    with tabs[1]:
        st.caption("Hijau tua = dipilih, hijau muda = cocok, merah muda = ditolak (p ≤ 0,05).")
        for k in ("WAK", "WL1", "WL2"):
            st.markdown(f"**{k}**")
            st.dataframe(warnai_uji(h["uji"][k]), hide_index=True)
    with tabs[2]:
        st.dataframe(sim, hide_index=True)
    with tabs[3]:
        st.dataframe(h["pembanding"], hide_index=True)
    with tabs[4]:
        st.caption("Sebaran data observasi (jumlah pelanggan per rentang detik).")
        g = st.columns(3)
        for col, k, warna in zip(g, ("WAK", "WL1", "WL2"), (KARAMEL, KOPI, BIRU)):
            col.markdown(f"**{k}**")
            col.bar_chart(histogram(data[k]), color=warna)
        bagian("Komposisi rata-rata waktu pelanggan di sistem (detik)")
        komposisi = pd.DataFrame(
            {"Rata-rata (detik)": [data["WT"].mean(), data["WL1"].mean(), data["WL2"].mean()]},
            index=["Tunggu di kasir (WT)", "Dilayani di kasir (WL1)", "Menunggu pesanan jadi (WL2)"])
        st.bar_chart(komposisi, color=KARAMEL)

    # ---- unduh
    buf = io.BytesIO()
    ac.ekspor_excel(h, buf)
    st.download_button("⬇️ Unduh semua hasil (Excel)", data=buf.getvalue(),
                       file_name="hasil_simulasi_antrean.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")