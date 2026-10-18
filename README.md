# Kelas Champions SMP — Panduan Deploy

Proyek ini sudah **teruji berhasil di-build** (Vite + React + Tailwind) di lingkungan terpisah dari Claude, jadi tinggal ikuti langkah di bawah.

## 1. Jalankan di komputer sendiri (opsional, untuk coba dulu)
```bash
npm install
npm run dev
```
Buka http://localhost:5173

## 2. Build untuk produksi
```bash
npm install
npm run build
```
Hasilnya ada di folder `dist/` — folder inilah yang di-deploy (berisi file statis HTML/CSS/JS saja).

## 3. Pilihan hosting gratis (pilih salah satu)

### A. Vercel (paling mudah)
1. Buat akun di vercel.com (bisa pakai GitHub).
2. "Add New Project" → upload folder ini atau hubungkan ke repo GitHub.
3. Framework Preset: pilih **Vite**. Build Command: `npm run build`. Output Directory: `dist`.
4. Klik Deploy — selesai, dapat link `https://nama-project.vercel.app`.

### B. Netlify
1. Buat akun di netlify.com.
2. "Add new site" → "Deploy manually" → **drag & drop folder `dist/`** setelah `npm run build`.
3. Selesai, dapat link `https://nama-acak.netlify.app` (bisa diganti nama).

### C. GitHub Pages (sudah disiapkan otomatis di proyek ini)
Proyek ini sudah dilengkapi workflow GitHub Actions (`.github/workflows/deploy.yml`) dan konfigurasi `base: "./"` di `vite.config.js`, jadi tidak perlu edit apa pun lagi — tinggal push ke GitHub dan setiap perubahan otomatis ter-build & ter-deploy. Panduan langkah demi langkah lengkap ada di **`GITHUB_PAGES.md`**.

Ringkasnya:
1. Buat repo baru di GitHub, push folder ini ke branch `main`.
2. Di repo → **Settings → Pages → Build and deployment → Source**, pilih **GitHub Actions**.
3. Tunggu ±1 menit di tab **Actions** sampai workflow selesai (centang hijau).
4. Buka `https://username-anda.github.io/nama-repo/`.
5. Setiap kali push perubahan baru ke `main`, situs otomatis ter-update sendiri — tidak perlu build manual lagi.

## 3b. Mode offline
Aplikasi ini sudah mendukung **offline (PWA)**. Cukup buka alamat situsnya **sekali dengan internet** — setelah itu semua file tersimpan di perangkat dan aplikasi bisa dibuka tanpa internet (data soal/kelompok/sesi memang sudah tersimpan di browser).
- Chrome/Edge: klik ikon **Install** di address bar (atau menu ⋮ → *Install Kelas Champions SMP*) agar muncul seperti aplikasi biasa di desktop/IFP.
- Penting: buka dulu di perangkat yang sama dan browser yang sama dengan yang dipakai di kelas. Jangan hapus data browser (cache/cookies).
- Saat ada versi baru dan internet menyala, buka aplikasi lalu refresh sekali; versi baru otomatis tersimpan.
- Menggunakan gambar soal lewat URL internet tetap butuh internet untuk gambar itu saja.
- Tidak bisa dibuka dengan klik dua kali file `index.html` (file://) — browser memblokir mode offline untuk cara itu. Pakai alamat GitHub Pages, atau jalankan `npm run preview` di komputer sendiri.

## 3c. Input soal lewat Google Sheets
Guru bisa mengetik soal di Google Sheets dan menyinkronkannya ke aplikasi dengan satu tombol — panduan di **`APPS_SCRIPT.md`**, skripnya di `apps-script/Code.gs`.

## 4. Catatan penting: penyimpanan data
Aplikasi ini otomatis memakai **localStorage browser** saat dijalankan di luar Claude — artinya:
- Data (bank soal, kelompok, pengaturan, riwayat) tersimpan **di browser/komputer masing-masing**, bukan di server terpusat.
- Kalau dibuka dari laptop guru A, datanya tidak otomatis muncul di laptop guru B — tiap perangkat punya datanya sendiri.
- Kalau ingin data yang sama diakses banyak guru/laptop, perlu backend/database tambahan (di luar cakupan proyek ini saat ini).
- Untuk pemakaian di kelas dengan 1 laptop + proyektor/IFP, ini sudah cukup — data akan tetap ada meski browser ditutup/laptop dimatikan, selama tidak menghapus data browser (cache/cookies) di perangkat itu.

## 5. Struktur proyek
- `src/App.jsx` — seluruh aplikasi (game engine, Teacher Studio, Bank Soal, dll).
- `src/main.jsx` — entry point React.
- `src/index.css` — Tailwind directive.
- `index.html`, `vite.config.js`, `tailwind.config.js`, `postcss.config.js` — konfigurasi build.
