# Catatan Hutang — GitHub Pages + Google Sheets

Aplikasi sederhana untuk pencatatan hutang dengan frontend statis di GitHub Pages dan Google Sheets sebagai database realtime melalui Google Apps Script.

## Struktur — hanya 3 file

- `index.html` — seluruh tampilan + CSS + JavaScript.
- `Code.gs` — backend Google Apps Script dan akses Google Sheet.
- `README.md` — panduan instalasi.

## Instalasi

### 1. Buat Google Sheet

Buat spreadsheet baru. Buka **Extensions → Apps Script**.

### 2. Pasang backend

Hapus isi `Code.gs`, lalu salin isi file `Code.gs` dari repository ini.

Jalankan fungsi `setup()` satu kali dari editor Apps Script dan izinkan permission yang diminta.

### 3. Buat password admin

Di Apps Script, jalankan:

```text
setAdminPassword('PasswordAdminMinimal8Karakter')
```

Ganti password contoh dengan password Anda. Jangan menaruh password tersebut di GitHub.

### 4. Deploy Apps Script

Pilih **Deploy → New deployment → Web app**.

Gunakan:
- Execute as: **Me**
- Who has access: **Anyone**

Salin URL deployment yang berakhiran `/exec`.

### 5. Hubungkan frontend

Buka `index.html`, cari:

```js
const API_URL = "PASTE_URL_APPS_SCRIPT_DI_SINI";
```

Ganti dengan URL `/exec` dari Apps Script.

### 6. Upload ke GitHub

Buat repository, upload ketiga file, lalu aktifkan **Settings → Pages** dan pilih branch yang digunakan.

Setelah GitHub Pages aktif, buka alamat Pages Anda.

## Cara kerja

### User
User hanya mengetik nama. Aplikasi mengambil data yang nama-nya cocok persis setelah normalisasi spasi/huruf besar-kecil, lalu menampilkan:
- bulan
- total hutang
- jumlah angsuran
- jumlah angsuran yang sudah dibayar
- nominal angsuran
- total terbayar
- sisa hutang
- tanggal jatuh tempo
- tanggal angsuran terakhir
- status
- catatan

### Admin
Admin login menggunakan password dan dapat:
- melihat semua data
- menambah hutang
- mengedit hutang
- menghapus hutang
- memfilter bulan
- mencari nama
- memfilter status
- melihat ringkasan total hutang, terbayar, dan sisa

## Struktur data Google Sheet

Sheet `Hutang` otomatis dibuat oleh `setup()` dengan kolom:

`id, nama, bulan, tanggal_jatuh_tempo, total_hutang, jumlah_angsuran_total, jumlah_angsuran_dibayar, nominal_angsuran, angsuran_terakhir, status, catatan, created_at, updated_at`

## Catatan keamanan penting

Versi ini mengikuti permintaan "user cukup memasukkan nama". Itu berarti nama **bukan autentikasi kuat**: seseorang yang mengetahui nama pelanggan lain secara teori dapat meminta data nama tersebut melalui endpoint. Jangan gunakan versi ini untuk data yang sangat sensitif.

Password admin tidak disimpan di GitHub; hanya hash SHA-256 yang disimpan di Script Properties. Sesi admin memakai token sementara di Apps Script Cache.

Untuk privasi yang lebih kuat, tahap berikutnya sebaiknya menambahkan PIN/password per user atau login Google. 
