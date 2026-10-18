# Input Soal Lewat Google Sheets (Apps Script)

Guru mengetik soal di Google Sheets (seperti Excel, bisa dari HP/laptop mana saja), lalu aplikasi mengambilnya dengan satu tombol. Setelah tersinkron, soal tersimpan di perangkat dan **tetap bisa dipakai offline**.

## A. Pasang di Google Sheets (sekali saja, ±5 menit)
1. Buka [sheets.google.com](https://sheets.google.com) → **Kosong** (spreadsheet baru). Beri nama, misalnya "Bank Soal IPA 7".
2. Menu **Ekstensi → Apps Script**. Hapus isi file `Code.gs` bawaan, lalu **tempel seluruh isi** `apps-script/Code.gs` dari paket ini. Klik ikon **Simpan** (💾).
3. Kembali ke spreadsheet, **muat ulang halaman** (F5). Muncul menu baru **Kelas Champions**.
4. Klik **Kelas Champions → 1. Buat / reset template soal**. Saat pertama kali, Google minta izin: pilih akun → *Lanjutan* → *Buka … (tidak aman)* → **Izinkan** (wajar, karena skrip ini milik Anda sendiri).
5. Isi soal mulai **baris 2** di sheet "Soal". Kolom `type`, `grade`, `difficulty` berisi dropdown. Petunjuk kolom `answer` ada di catatan pada judul kolomnya.
6. Klik **Kelas Champions → 2. Periksa soal** — baris yang salah diberi warna merah beserta alasannya.

## B. Buat alamat (Web App) untuk aplikasi
1. Di Apps Script: **Deploy (Terapkan) → Deployment baru** → ikon roda gigi → pilih **Aplikasi web**.
2. **Jalankan sebagai:** *Saya*. **Yang memiliki akses:** *Siapa saja*.
3. Klik **Deploy**, izinkan jika diminta, lalu **salin URL Aplikasi Web** (berakhiran `/exec`).

> Keamanan: siapa pun yang memegang URL itu bisa **membaca** soal (tidak bisa mengubah). Jangan sebarkan ke siswa jika soal belum boleh bocor.
> Setelah mengubah isi `Code.gs`, buat **Deployment baru** (atau *Kelola deployment → Edit → Versi baru*) agar perubahan aktif. Mengubah **isi sheet** tidak perlu deploy ulang.

## C. Ambil soal di aplikasi
1. Buka aplikasi → **Teacher Studio → Bank Soal**.
2. Tempel URL `/exec` di kotak **Sinkron dari Google Sheets** → klik **Ambil Soal dari Sheets**. Alamat diingat otomatis di perangkat itu.
3. Edit soal di Sheets kapan pun, lalu klik tombol itu lagi: soal yang sama **diperbarui**, soal baru **ditambahkan**, soal yang dihapus di Sheets ikut hilang dari aplikasi. Soal yang diketik langsung di aplikasi (manual/Excel) tidak tersentuh.
4. Butuh internet hanya saat klik tombol sinkron. Sehari sebelum mengajar, sinkron dulu di kelas/rumah yang ada internetnya.

## Tanpa Apps Script (lebih sederhana, hanya baca)
Di Google Sheets: **File → Bagikan → Publikasikan ke web** → pilih sheet "Soal" dan format **CSV** → salin link-nya, tempel di kotak sinkron aplikasi. Judul kolom harus sama seperti template.

## Satu spreadsheet untuk banyak mapel/kelas
Tambahkan sheet lain (mis. "IPA7") dengan judul kolom sama, lalu tambahkan `?sheet=IPA7` di akhir URL `/exec`.

## Kolom `id` (opsional)
Kosongkan saja — aplikasi mengenali soal dari isi kalimatnya. Isi kode unik (mis. `ipa-001`) hanya jika Anda sering memperbaiki kalimat soal tetapi ingin soal itu tetap dianggap sama.

## Kalau gagal
- *"Balasan berupa halaman web"* → akses Web App belum **Siapa saja**, atau URL bukan yang berakhiran `/exec`.
- *"Sheet 'Soal' tidak ditemukan"* → jalankan menu **Buat template soal** atau nama sheet diubah.
- *Gagal mengambil data* → cek internet dan URL.
