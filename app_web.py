"""
app_web.py  -  Web sederhana: simulasi antrean dua fase.

Cara menjalankan di komputer sendiri:
    pip install -r requirements.txt
    streamlit run app_web.py
Browser akan terbuka otomatis. Untuk dipakai orang lain lewat tautan, program ini bisa
diunggah (bersama antrean_core.py dan requirements.txt) ke Streamlit Community Cloud (gratis).
"""

import io

import streamlit as st

import antrean_core as ac

st.set_page_config(page_title="Simulasi Antrean Dua Fase", layout="wide")
st.title("Simulasi Antrean Dua Fase")
st.caption("Unggah data mentah (Excel/CSV), program menghitung rata-rata, menguji distribusi, "
           "memilih yang paling cocok, lalu menyimulasikan antrean.")

# ---------------------------------------------------------------- sidebar
with st.sidebar:
    st.header("1. Data mentah")
    berkas = st.file_uploader("Unggah file Excel (.xlsx) atau CSV", type=["xlsx", "xlsm", "xls", "csv"])
    with st.expander("Format data yang dibutuhkan"):
        st.markdown(
            "Satu baris per pelanggan, dengan 4 kolom waktu berformat jam (misalnya 17:02:47):\n\n"
            "- **WD**: waktu kedatangan\n- **WML1**: waktu mulai dilayani di kasir\n"
            "- **WSL1**: waktu selesai di kasir (masuk fase 2)\n- **WSL2**: waktu selesai fase 2 (pesanan jadi)\n\n"
            "Dari sini dihitung WAK (selisih kedatangan), WL1 = WSL1 - WML1, dan WL2 = WSL2 - WSL1.")

    st.header("3. Pengaturan")
    ganti_nol = st.checkbox("Ganti nilai waktu 0 menjadi 1 detik", value=True)
    metode = st.selectbox("Metode uji distribusi", list(ac.METODE.keys()),
                          format_func=lambda k: {"BERTAHAP": "Bertahap (seperti artikel acuan)",
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
    st.info("Unggah file data di panel kiri untuk mulai.")
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

st.sidebar.header("2. Pilih kolom waktu")
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
        opsi = dict(ganti_nol=ganti_nol, metode=metode, n_mc=1000, c1=int(c1), c2=int(c2),
                    replikasi=int(replikasi), wl2_sudah_tunggu=wl2_tunggu, T=T, seed=2026)
        with st.spinner("Menghitung, menguji distribusi, dan menyimulasikan ..."):
            st.session_state["hasil"] = ac.jalankan_semua(df, peta, opsi)
            st.session_state["metode"] = metode
    except Exception as e:
        st.session_state.pop("hasil", None)
        st.error(str(e))

h = st.session_state.get("hasil")
if h:
    st.subheader("Ringkasan")
    st.write(h["ringkasan"])
    if h["catatan"]:
        for c in h["catatan"]:
            st.warning(c)
    st.caption(f"Lama simulasi {h['T']:.0f} detik, {h['replikasi_dipakai']} replikasi. "
               f"Metode uji: {ac.METODE[st.session_state['metode']]}.")

    tabs = st.tabs(["Rata-rata", "Uji distribusi", "Distribusi terpilih", "Hasil simulasi",
                    "Observasi vs simulasi"])
    with tabs[0]:
        st.dataframe(h["rata_rata"], hide_index=True)
        st.dataframe(h["laju"], hide_index=True)
    with tabs[1]:
        for k in ("WAK", "WL1", "WL2"):
            st.markdown(f"**{k}**")
            st.dataframe(h["uji"][k], hide_index=True)
    with tabs[2]:
        st.dataframe(h["pilihan"], hide_index=True)
    with tabs[3]:
        st.dataframe(h["simulasi"], hide_index=True)
    with tabs[4]:
        st.dataframe(h["pembanding"], hide_index=True)

    buf = io.BytesIO()
    ac.ekspor_excel(h, buf)
    st.download_button("Unduh semua hasil (Excel)", data=buf.getvalue(),
                       file_name="hasil_simulasi_antrean.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
