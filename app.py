import streamlit as st
import pandas as pd
import locale
locale.setlocale(locale.LC_ALL, '')

st.set_page_config(page_title="Aplikasi Akuntansi", layout="centered")
st.title("📘 Aplikasi Akuntansi Sederhana")

# Inisialisasi data
if "jurnal" not in st.session_state:
    st.session_state["jurnal"] = []

# Form input transaksi
with st.form("transaksi_form"):
    st.subheader("Tambah Transaksi")
    
    tanggal = st.date_input("Tanggal")
    nama_akun = st.text_input("Nama Akun")
    keterangan = st.selectbox("Keterangan", ["Kas", "Modal", "Pendapatan", "Beban", "Utang"])
    jenis = st.selectbox("Jenis Transaksi", ["Debit", "Kredit"])
    jumlah = st.number_input("Jumlah", min_value=0.0)
    tambah = st.form_submit_button("Tambah Transaksi")

    if tambah:
        debit = jumlah if jenis == "Debit" else 0
        kredit = jumlah if jenis == "Kredit" else 0
        st.session_state["jurnal"].append({
            "Tanggal": str(tanggal),
            "Nama Akun": nama_akun,
            "Keterangan": keterangan,
            "Debit": debit,
            "Kredit": kredit
        })
        st.success("Transaksi berhasil ditambahkan!")

# Tampilkan jurnal umum
st.subheader("📄 Jurnal Umum")
df = pd.DataFrame(st.session_state["jurnal"])

if not df.empty and "Debit" in df.columns and "Kredit" in df.columns:
    total_debit = df["Debit"].sum()
    total_kredit = df["Kredit"].sum()

    total_row = pd.DataFrame([{
        "Tanggal": "",
        "Nama Akun": "",
        "Keterangan": "Total",
        "Debit": total_debit,
        "Kredit": total_kredit
    }])
    df_total = pd.concat([df, total_row], ignore_index=True)

    def format_rupiah(x):
        if x == 0 or x == "":
            return ""
        return f"{int(x):,}".replace(",", ".")

    df_total["Debit"] = df_total["Debit"].apply(format_rupiah)
    df_total["Kredit"] = df_total["Kredit"].apply(format_rupiah)

    st.dataframe(df_total, use_container_width=True)
else:
    st.info("Belum ada data transaksi. Silakan tambahkan dulu.")

# Tombol laporan dan ekspor (sejajar kiri, tengah, kanan)
col1, col2, col3 = st.columns([1, 1, 1])

with col1:
    if st.button("📊 Laporan Laba Rugi"):
        pendapatan = sum(t["Kredit"] for t in st.session_state["jurnal"] if t["Keterangan"] == "Pendapatan")
        beban = sum(t["Debit"] for t in st.session_state["jurnal"] if t["Keterangan"] == "Beban")
        laba_rugi = pendapatan - beban

        data_laba_rugi = {
            "Keterangan": ["Pendapatan", "Beban"],
            "Jumlah (Rp)": [format_rupiah(pendapatan), format_rupiah(beban)]
        }
        df_laba_rugi = pd.DataFrame(data_laba_rugi)
        st.table(df_laba_rugi)
        st.markdown(f"**Laba/Rugi = Pendapatan - Beban ➡️ {format_rupiah(pendapatan)} - {format_rupiah(beban)} = {format_rupiah(laba_rugi)}**")

with col2:
    if st.button("📄 Tampilkan Neraca"):
        kas = sum(t["Debit"] for t in st.session_state["jurnal"] if t["Keterangan"] == "Kas")
        utang = sum(t["Kredit"] for t in st.session_state["jurnal"] if t["Keterangan"] == "Utang")
        modal = sum(t["Kredit"] for t in st.session_state["jurnal"] if t["Keterangan"] == "Modal")

        data_neraca = {
            "Keterangan": ["Kas (Aset)", "Utang", "Modal"],
            "Jumlah (Rp)": [format_rupiah(kas), format_rupiah(utang), format_rupiah(modal)]
        }
        df_neraca = pd.DataFrame(data_neraca)
        st.table(df_neraca)
        st.markdown(f"**Aset = Utang + Modal ➡️ {format_rupiah(kas)} = {format_rupiah(utang)} + {format_rupiah(modal)}**")

with col3:
    if st.button("💾 Simpan ke Excel"):
        df.to_excel("jurnal_umum.xlsx", index=False)
        st.success("Berhasil disimpan ke jurnal_umum.xlsx")
