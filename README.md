# Catatan Hutang v8

Aplikasi catatan hutang sederhana: GitHub Pages sebagai frontend, Google Apps Script + Google Sheet sebagai backend/database.

## Fitur v8
- Login user dengan Nama + PIN.
- User hanya melihat hutang miliknya.
- Admin mengelola semua data.
- Nama barang kredit, misalnya Mesin Cuci/HP/TV.
- Jadwal angsuran otomatis berdasarkan tanggal mulai, tanggal jatuh tempo bulanan, jumlah angsuran, dan nominal per bulan.
- Admin bisa membayar langsung tanpa persetujuan.
- User punya tombol **Bayar Angsuran**.
- User wajib memilih tanggal pembayaran dan upload bukti transfer dari galeri HP.
- Bukti gambar disimpan di Google Drive dan link/metadata permanen dicatat di Google Sheet.
- Pembayaran user masuk sebagai **Menunggu Persetujuan**.
- Admin dapat klik **Ya, setujui**: angsuran otomatis masuk.
- Admin dapat klik **Tidak**: angsuran tidak berubah; alasan penolakan dapat dicatat.
- Riwayat permintaan pembayaran tetap tersimpan.
- Password admin disimpan di sheet Pengaturan dan dapat diganti dari aplikasi.

## Struktur Google Sheet
- `Hutang`
- `Pengguna`
- `Pembayaran`
- `Permintaan Pembayaran`
- `Pengaturan`

## Instalasi / update
1. Buka Google Sheet yang dipakai aplikasi.
2. Extensions > Apps Script.
3. Ganti isi `Code.gs` dengan versi v8.
4. Simpan.
5. Jalankan `setup()` sekali. Beri izin Google Drive saat diminta. Script akan membuat folder Drive `Bukti Transfer Catatan Hutang` dan menyimpan ID folder di `Pengaturan!B3`.
6. Deploy > Manage deployments > Edit > New version > Deploy.
7. Pastikan Web app: Execute as **Me**, Who has access **Anyone**.
8. Di `index.html`, isi `API_URL` dengan URL `/exec` Web App.
9. Upload `index.html` dan `README.md` ke GitHub Pages.

## Catatan bukti transfer
File gambar fisik disimpan di Google Drive milik akun yang menjalankan Apps Script. Google Sheet menyimpan ID file, URL, nama file, tanggal, nama user, hutang, dan status persetujuan. Jangan menghapus folder/file bukti jika ingin riwayat tetap dapat dibuka.

## Upload dari HP
Input bukti memakai `accept="image/*"`, sehingga pengguna dapat memilih foto/gambar dari galeri HP. Browser akan mengecilkan gambar sebelum dikirim agar lebih ringan.


### Login otomatis user
User yang sudah berhasil login akan disimpan di perangkat menggunakan localStorage. Saat aplikasi dibuka kembali, sesi user dipulihkan otomatis. Jika token server sudah kedaluwarsa, aplikasi melakukan login ulang otomatis menggunakan hash PIN yang tersimpan di perangkat. Tombol Keluar menghapus sesi tersimpan.


## Versi 11 - Banyak Kredit per Pelanggan
Satu nama pelanggan dapat memiliki banyak data kredit secara terpisah. Setiap barang/kredit mempunyai ID sendiri sehingga nominal angsuran, jumlah cicilan, tanggal mulai, tanggal jatuh tempo, pembayaran, dan bukti transfer tidak tercampur. User melihat semua kredit miliknya sebagai kartu terpisah. Admin dapat membuat kredit baru dengan mengetik nama pelanggan yang sama.
