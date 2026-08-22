# Catatan Hutang — v6 (Bayar Angsuran)

Aplikasi catatan hutang sederhana: frontend `index.html` bisa di-host di GitHub Pages, sedangkan Google Apps Script + Google Sheet menjadi backend/database.

## Fitur
- Login user dengan nama + PIN.
- User hanya melihat hutangnya sendiri.
- Login admin dengan password yang tersimpan di sheet `Pengaturan`.
- Admin dapat mengganti password dari aplikasi.
- Admin dapat membuat/reset PIN pelanggan.
- Input hutang berdasarkan:
  - tanggal mulai angsuran
  - tanggal jatuh tempo setiap bulan
  - jumlah kali angsuran
  - angsuran per bulan
- Total hutang dihitung otomatis.
- Jadwal angsuran dibuat otomatis per bulan.
- **Tombol Bayar Angsuran**: admin cukup klik tombol, pilih tanggal pembayaran, lalu sistem otomatis:
  - menambah angsuran dibayar +1
  - mencatat tanggal pembayaran terakhir
  - menandai angsuran tersebut Lunas
  - menghitung angsuran berikutnya
  - mengubah status menjadi Lunas setelah semua angsuran selesai
- Riwayat setiap pembayaran disimpan di sheet `Pembayaran`.

## Google Sheet
Jalankan `setup()` sekali. Sheet yang digunakan:

1. `Hutang` — data kontrak/jadwal hutang.
2. `Pengguna` — nama dan hash PIN user.
3. `Pengaturan` — password admin.
4. `Pembayaran` — riwayat pembayaran angsuran.

## Instalasi
1. Buat Google Sheet kosong.
2. Buka **Extensions → Apps Script**.
3. Ganti isi `Code.gs` dengan file `Code.gs` dari paket ini.
4. Jalankan fungsi `setup()` sekali dan izinkan akses.
5. **Deploy → New deployment → Web app**.
6. Pilih **Execute as: Me** dan akses **Anyone**.
7. Salin URL `/exec`.
8. Buka `index.html`, lalu ubah `API_URL` menjadi URL Web App tersebut.
9. Upload `index.html` ke repository GitHub dan aktifkan GitHub Pages.

## Catatan update dari versi lama
Jika sebelumnya sudah memakai versi v5, cukup ganti `Code.gs` dan jalankan `setup()` sekali. Sheet `Pembayaran` akan dibuat otomatis. Data `Hutang`, `Pengguna`, dan `Pengaturan` tetap digunakan.

## Cara memakai Bayar Angsuran
1. Login sebagai admin.
2. Cari pelanggan pada bagian **Data Hutang**.
3. Klik **💵 Bayar Angsuran**.
4. Aplikasi menampilkan nomor angsuran berikutnya, nominal, dan jatuh tempo.
5. Pilih tanggal pembayaran sebenarnya.
6. Klik **Konfirmasi Bayar**.
7. Jadwal langsung berubah dan angsuran berikutnya otomatis muncul.

Contoh: 12 kali × Rp150.000. Setelah tombol Bayar dipakai 3 kali, sistem otomatis menunjukkan **3/12 sudah dibayar**, angsuran berikutnya adalah ke-4, dan sisa menjadi 9 angsuran.


### Catatan pembayaran
Tanggal pembayaran terakhir **tidak perlu diisi di form Tambah/Edit Hutang**. Tanggal tersebut akan dicatat otomatis saat admin menekan tombol **💵 Bayar Angsuran** dan memilih tanggal pembayaran.


## Versi terbaru: pembayaran otomatis

Form tambah/edit hutang tidak lagi meminta "Angsuran sudah dibayar". Setiap data baru dimulai dari 0 pembayaran. Jumlah angsuran dibayar hanya berubah melalui tombol **Bayar Angsuran**. Saat edit data, jumlah pembayaran yang sudah tercatat tetap dipertahankan dan tidak bisa tertimpa oleh form.
