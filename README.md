# Web Autofill Data Karyawan

Web Google Apps Script ini mencari nama pada spreadsheet **Informasi Kontak (Jawaban)**, mengisi NIK, Bagian, dan Jabatan, lalu mengirim data ke Google Form **Database Karyawan**.

## Cara memasang

1. Buka https://script.google.com dan buat **Project baru**.
2. Salin isi `Code.gs` ke file `Code.gs` di Apps Script.
3. Tambahkan file HTML bernama **Index**, lalu salin isi `Index.html`.
4. Klik **Deploy > New deployment > Web app**.
5. Atur **Execute as: Me** dan pilih siapa yang boleh mengakses sesuai kebijakan perusahaan.
6. Klik **Deploy**, izinkan akses Spreadsheet dan koneksi eksternal, lalu bagikan URL web app.

## Sumber dan tujuan

- Spreadsheet ID: `17E79JTYAY9qY5EYtiG-VNyfAFqhKIxtHDGIbv-5tdJQ`
- Sheet: `Form Responses 1`
- Google Form: `1FAIpQLSeKmvRS4eHSkPywcVfxEVltvCJgO4IVnwpJapJ5LHnB8ks9yg`

Jika kolom pada spreadsheet dipindahkan, sesuaikan pembacaan rentang pada fungsi `employeeRows_()`.
