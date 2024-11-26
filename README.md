# Sertifikat Generator Microsoft Excel to Microsoft Word

## Deskripsi
Sertifikat Generator adalah program Python yang memungkinkan pembuatan sertifikat secara otomatis dari data yang disimpan dalam file Excel. Program ini menggunakan dua library utama:
- `openpyxl`: untuk membaca data dari file Excel.
- `docxtpl`: untuk memanipulasi template dokumen Word dan menghasilkan sertifikat berdasarkan data yang diambil.

Dengan program ini, pengguna dapat membuat sertifikat dengan mudah dan cepat, cukup dengan menyiapkan file Excel yang berisi data acara, peserta, dan penyelenggara, kemudian menjalankan program untuk menghasilkan file sertifikat yang sudah terisi secara otomatis.

## Fitur
- Membaca data dari file Excel yang memiliki informasi tentang nama acara, tahun, jumlah peserta, tanggal, dan penyelenggara.
- Menggunakan template Word untuk menghasilkan sertifikat yang sudah terisi.
- Menyimpan setiap sertifikat dengan nama file yang dinamis sesuai dengan data yang ada (misalnya: "Sertifikat_Seminar Teknologi AI_Firdaus Firmansyah.docx").
- Mengonversi tanggal ke format yang lebih ramah pengguna (DD-MM-YYYY).

## Cara Menggunakan
1. Siapkan file Excel dengan format sebagai berikut:
   - Kolom 1: Judul Acara
   - Kolom 2: Tahun
   - Kolom 3: Jumlah Peserta
   - Kolom 4: Tanggal Acara
   - Kolom 5: Penyelenggara

2. Pastikan template dokumen Word (`Sertifikat.docx`) sudah tersedia dengan placeholder seperti `{JUDUL}`, `{TAHUN}`, `{NAMA}`, `{TANGGAL}`, dan `{PENYELENGGARA}` di dalamnya.

3. Jalankan program Python di terminal dengan perintah:
   ```bash
   python main.py

4. Sertifikat akan dibuat dan disimpan di folder yang sama dengan script Python.

Instalasi
Pastikan Python sudah terinstall di sistem kamu, kemudian instal library yang dibutuhkan dengan menjalankan perintah berikut:
pip install openpyxl 
pip install docxtpl

Terima kasih telah menggunakan Sertifikat Generator. Semoga program ini bermanfaat untuk mempermudah pembuatan sertifikat di acara-acara Anda.
Jangan lupa untuk mengikuti saya di media sosial:
Instagram: @daussauruss
Tiktok: www.tiktok.com/@firdauuussss03
