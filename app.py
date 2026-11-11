import streamlit as st
import pandas as pd
import io

# --- KONFIGURASI HALAMAN & HILANGKAN SIDEBAR ---
st.set_page_config(
    page_title="Auto-Check Laporan Keuangan", 
    page_icon="📊", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

hide_streamlit_style = """
    <style>
        [data-testid="collapsedControl"] {display: none;} 
        section[data-testid="stSidebar"] {display: none;} 
        #MainMenu {display: none;} 
        footer {display: none;} 
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)


# --- FUNGSI GENERATE EXCEL UJI COBA ---
def generate_sample_excel():
    data_neraca = {
        'Akun': ['Kas di Tangan', 'Piutang Usaha', 'Hutang Bank', 'Modal Awal', 'Total Aset', 'Total Pasiva'],
        'Kategori': ['Aset', 'Aset', 'Aset', 'Pasiva', 'Total Aset', 'Total Pasiva'], 
        'Saldo': [-2500000, 15000000, 20000000, 25000000, 32500000, 25000000] 
    }
    df_neraca = pd.DataFrame(data_neraca)
    
    data_tb = {
        'Akun': ['Kas di Tangan', 'Piutang Usaha', 'Hutang Bank', 'Modal Awal'],
        'Saldo': [5000000, 15000000, 20000000, 25000000] 
    }
    df_tb = pd.DataFrame(data_tb)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_neraca.to_excel(writer, sheet_name='Neraca', index=False)
        df_tb.to_excel(writer, sheet_name='Trial_Balance', index=False)
    
    return output.getvalue()


# --- UI APLIKASI UTAMA ---
st.title("📊 Validasi Otomatis Laporan Keuangan")
st.markdown("""
Aplikasi web untuk memeriksa laporan keuangan dari kesalahan umum secara otomatis.
**Alur:** Upload Excel ➡️ Cek Otomatis ➡️ Lihat Hasil ➡️ Perbaiki & Upload Ulang.
""")

st.markdown("---")
col1, col2 = st.columns([2, 1])

with col1:
    st.info("💡 **Ingin mencoba aplikasi ini?** Unduh file Excel template di bawah ini yang sudah sengaja disisipkan kesalahan (Kas minus, salah klasifikasi, dan selisih).")
with col2:
    sample_excel = generate_sample_excel()
    st.download_button(
        label="📥 Unduh Excel Uji Coba",
        data=sample_excel,
        file_name="Laporan_Keuangan_Dummy.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

st.markdown("---")

# --- AREA UPLOAD FILE ---
uploaded_file = st.file_uploader("Upload File Laporan Keuangan (Excel: .xlsx)", type=["xlsx"])

if uploaded_file is not None:
    st.success("File berhasil diunggah. Memulai pemeriksaan otomatis...")
    
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        if 'Neraca' not in xls.sheet_names or 'Trial_Balance' not in xls.sheet_names:
            st.error("❌ Format file tidak sesuai. Pastikan terdapat sheet bernama 'Neraca' dan 'Trial_Balance'.")
        else:
            df_neraca = pd.read_excel(xls, sheet_name='Neraca')
            df_tb = pd.read_excel(xls, sheet_name='Trial_Balance')
            
            errors = []
            warnings = []
            
            col_neraca = df_neraca.columns.tolist()
            
            # --- MULAI VALIDASI ---
            if 'Akun' in col_neraca and 'Saldo' in col_neraca and 'Kategori' in col_neraca:
                
                # 1. Pengecekan Saldo Kas Minus
                kas_neraca_row = df_neraca[df_neraca['Akun'].str.contains('Kas', case=False, na=False)]
                kas_neraca_saldo = 0
                if not kas_neraca_row.empty:
                    kas_neraca_saldo = kas_neraca_row['Saldo'].values[0]
                    if kas_neraca_saldo < 0:
                        errors.append(f"Saldo Kas bernilai minus (Rp {kas_neraca_saldo:,.0f}).")
                
                # 2. Pengecekan Balance Neraca
                total_aset = df_neraca[df_neraca['Kategori'] == 'Total Aset']['Saldo'].sum()
                total_pasiva = df_neraca[df_neraca['Kategori'] == 'Total Pasiva']['Saldo'].sum()
                selisih_neraca = abs(total_aset - total_pasiva)
                
                if selisih_neraca > 0:
                    if 'Akun' in df_tb.columns and 'Saldo' in df_tb.columns:
                        kas_tb = df_tb[df_tb['Akun'].str.contains('Kas', case=False, na=False)]['Saldo'].sum()
                        if kas_neraca_saldo != kas_tb:
                            errors.append(f"Neraca selisih Rp {selisih_neraca:,.0f} — kemungkinan dari akun Kas (total di neraca Rp {kas_neraca_saldo:,.0f}, tapi detail di trial balance Rp {kas_tb:,.0f}).")
                        else:
                            errors.append(f"Neraca tidak balance. Selisih: Rp {selisih_neraca:,.0f}.")
                    else:
                        errors.append(f"Neraca tidak balance. Selisih: Rp {selisih_neraca:,.0f}.")

                # 3. Pengecekan Klasifikasi Akun
                validasi_kategori = {
                    "Kas": "Aset",
                    "Piutang": "Aset",
                    "Hutang": "Pasiva",
                    "Modal": "Pasiva",
                }
                for index, row in df_neraca.iterrows():
                    nama_akun = str(row['Akun'])
                    kategori_aktual = str(row['Kategori'])
                    for kata_kunci, kategori_seharusnya in validasi_kategori.items():
                        if kata_kunci.lower() in nama_akun.lower():
                            if kategori_aktual != kategori_seharusnya and kategori_aktual.lower() != 'nan':
                                errors.append(f"Akun salah klasifikasi: '{nama_akun}' masuk ke '{kategori_aktual}', seharusnya masuk ke '{kategori_seharusnya}'.")
            else:
                warnings.append("⚠️ Kolom tidak lengkap. Pastikan sheet 'Neraca' memiliki kolom: 'Akun', 'Kategori', 'Saldo'.")

            # --- MENAMPILKAN HASIL VALIDASI ---
            st.subheader("Hasil Pemeriksaan")
            
            if len(errors) == 0 and len(warnings) == 0:
                st.success("✅ LOLOS: Tidak ditemukan kesalahan pada laporan keuangan Anda.")
            else:
                if len(errors) > 0:
                    st.error("🚨 DITEMUKAN KESALAHAN (ERROR):")
                    for err in errors:
                        st.write(f"- {err}")
                if len(warnings) > 0:
                    st.warning("⚠️ PERINGATAN (WARNING):")
                    for warn in warnings:
                        st.write(f"- {warn}")
                st.info("💡 Silakan perbaiki file Excel Anda berdasarkan catatan di atas, lalu upload ulang.")

            # --- MENAMPILKAN GRAFIK & TABEL ---
            st.markdown("---")
            st.subheader("📈 Ringkasan Data yang Diupload")
            
            # Membuat tab untuk tampilan yang rapi
            tab1, tab2, tab3 = st.tabs(["Tabel Neraca", "Tabel Trial Balance", "Grafik Saldo Akun"])
            
            with tab1:
                st.dataframe(df_neraca, use_container_width=True)
                
            with tab2:
                st.dataframe(df_tb, use_container_width=True)
                
            with tab3:
                # Memfilter baris 'Total' agar grafik hanya menampilkan akun individual
                df_chart = df_neraca[~df_neraca['Kategori'].isin(['Total Aset', 'Total Pasiva'])]
                # Menjadikan kolom 'Akun' sebagai index agar menjadi label di grafik
                df_chart = df_chart.set_index('Akun')[['Saldo']]
                st.bar_chart(df_chart)

    except Exception as e:
        st.error(f"Terjadi kesalahan sistem saat memproses file: {e}")
