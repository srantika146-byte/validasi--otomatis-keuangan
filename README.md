# Auto-Check Laporan Keuangan 📊

Aplikasi web berbasis Python & Streamlit untuk memvalidasi laporan keuangan secara otomatis sebelum diserahkan. Aplikasi ini berjalan murni tanpa sidebar untuk memberikan tampilan yang fokus dan bersih.

## Alur Kerja Aplikasi
1. **Upload** file laporan keuangan (Excel)
2. **Periksa otomatis** (Sistem mendeteksi selisih neraca, klasifikasi akun salah, dan saldo kas minus)
3. **Lihat Hasil** (Sistem menampilkan status: Lolos / Warning / Error)
4. **Perbaiki & Upload Ulang**

## Persyaratan Format Excel (.xlsx)
Agar aplikasi dapat membaca data, file Excel harus memiliki setidaknya dua sheet berikut:
1. Sheet **Neraca**: Harus memiliki kolom `Akun`, `Kategori` (isinya misal: Aset, Pasiva, Total Aset, Total Pasiva), dan `Saldo`.
2. Sheet **Trial_Balance**: Harus memiliki kolom `Akun` dan `Saldo`.

## Cara Menjalankan di Komputer Lokal
Jika ingin mengetes di komputer lokal sebelum upload ke GitHub, jalankan perintah ini di terminal:
```bash
pip install -r requirements.txt
streamlit run app.py
