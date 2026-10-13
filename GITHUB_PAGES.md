# Deploy KELAS CHAMPIONS SMP ke GitHub Pages — Panduan Lengkap

Panduan ini mengasumsikan Anda belum pernah pakai Git/GitHub sama sekali. Ikuti urut dari atas.

Ada 2 cara. **Cara A** (pakai Git di komputer) lebih cepat kalau sudah pernah install Git.
**Cara B** (upload lewat browser) cocok kalau tidak mau install apa-apa sama sekali.

---

## Persiapan (sekali saja)

1. Buat akun di [github.com](https://github.com) kalau belum punya (gratis).
2. **Cara A saja**: install [Git](https://git-scm.com/downloads) dan [Node.js](https://nodejs.org) (versi LTS) di komputer Anda, kalau belum ada.

---

## Langkah 1 — Buat repository baru di GitHub

1. Login ke github.com, klik tombol **+** di kanan atas → **New repository**.
2. Isi **Repository name**, misalnya `kelas-champions-smp`.
3. Pilih **Public**.
4. **Jangan** centang "Add a README file" (folder kita sudah punya semuanya).
5. Klik **Create repository**. Biarkan halaman "…quick setup…" yang muncul tetap terbuka, kita butuh alamat repo-nya di langkah berikut.

---

## Cara A — Pakai Git lewat Terminal (direkomendasikan)

1. Ekstrak file `kelas-champions-smp-deploy.zip` yang sudah dikirim, lalu buka folder hasil ekstraknya di Terminal (Mac/Linux) atau Command Prompt/PowerShell (Windows):
   ```bash
   cd path/ke/folder/kelas-champions-smp-deploy
   ```
2. Jalankan perintah berikut satu per satu (ganti `USERNAME` dan `kelas-champions-smp` sesuai punya Anda — alamatnya persis sama dengan yang ditampilkan GitHub di halaman "quick setup" langkah sebelumnya):
   ```bash
   git init
   git add .
   git commit -m "Deploy awal Kelas Champions SMP"
   git branch -M main
   git remote add origin https://github.com/USERNAME/kelas-champions-smp.git
   git push -u origin main
   ```
   Kalau diminta login, ikuti instruksi di layar (biasanya lewat browser atau Personal Access Token — GitHub akan memandu).
3. Lanjut ke **Langkah 2 — Aktifkan GitHub Pages** di bawah.

---

## Cara B — Upload lewat browser (tanpa Git/Terminal)

1. Ekstrak file `kelas-champions-smp-deploy.zip` di komputer Anda.
2. Di halaman repo GitHub yang baru dibuat, klik **uploading an existing file** (atau menu **Add file → Upload files**).
3. Seret (drag & drop) **semua isi** folder `kelas-champions-smp-deploy` (bukan folder-nya, tapi isinya: `src/`, `package.json`, `.github/`, dll — pastikan file yang namanya diawali titik seperti `.gitignore` dan folder `.github` ikut ter-upload; browser kadang menyembunyikannya, jadi cek dulu di file manager apakah "hidden files" sudah ditampilkan).
4. Scroll ke bawah, isi pesan commit bebas (misal "Deploy awal"), klik **Commit changes**.
5. Lanjut ke **Langkah 2** di bawah.

> Catatan: Cara B sedikit rawan kalau file `.github/workflows/deploy.yml` tidak ikut ter-upload (karena tersembunyi oleh sistem operasi). Kalau setelah Langkah 2 workflow tidak muncul di tab **Actions**, pakai Cara A, atau upload ulang khusus folder `.github` lewat **Add file → Upload files** sambil memastikan file tersembunyi ditampilkan di jendela pemilihan file.

---

## Langkah 2 — Aktifkan GitHub Pages

1. Di repo GitHub, buka tab **Settings** (ikon gerigi, paling kanan).
2. Di sidebar kiri, klik **Pages**.
3. Di bagian **Build and deployment → Source**, pilih **GitHub Actions** (bukan "Deploy from a branch").
4. Tidak perlu isi apa-apa lagi — workflow yang sudah ada di `.github/workflows/deploy.yml` akan otomatis jalan.

---

## Langkah 3 — Tunggu proses build selesai

1. Klik tab **Actions** di bagian atas repo.
2. Akan ada 1 proses berjalan bernama "Deploy ke GitHub Pages" (otomatis terpicu dari push/upload tadi). Klik untuk melihat progresnya.
3. Tunggu sampai muncul centang hijau ✅ (biasanya 1–2 menit).

---

## Langkah 4 — Buka situsnya

Alamatnya adalah:
```
https://USERNAME.github.io/kelas-champions-smp/
```
(ganti `USERNAME` dan `kelas-champions-smp` sesuai punya Anda — alamat persisnya juga ditampilkan di tab **Settings → Pages** setelah deploy pertama selesai, dengan tulisan "Your site is live at …").

---

## Update berikutnya

- **Cara A**: setiap ada perubahan kode, cukup jalankan `git add .`, `git commit -m "pesan perubahan"`, `git push` — situs otomatis ter-update dalam 1–2 menit.
- **Cara B**: upload ulang file yang berubah lewat **Add file → Upload files**, commit — situs otomatis ter-update juga.

Tidak perlu `npm run build` manual atau upload folder `dist/` — GitHub Actions yang mengerjakan build-nya di server GitHub setiap kali ada perubahan.

---

## Troubleshooting

**Halaman putih kosong setelah dibuka** — biasanya cache browser. Coba refresh paksa (`Ctrl+Shift+R` / `Cmd+Shift+R`), atau buka di jendela incognito.

**Tab Actions menunjukkan tanda silang merah ❌** — klik ke dalam prosesnya, baca pesan error di bagian yang gagal (biasanya "Install dependencies" atau "Build"). Pesan error paling umum: ada file yang gagal ter-upload lengkap di Cara B — pastikan semua file ter-upload lalu commit ulang.

**Menu Pages tidak muncul di Settings** — pastikan repo berstatus **Public** (fitur Pages gratis GitHub Actions deploy butuh repo publik kecuali Anda berlangganan GitHub Pro/Team).

**Data soal/kelompok tidak sinkron antar perangkat** — ini normal (lihat bagian "Catatan penting: penyimpanan data" di `README.md`). GitHub Pages hanya meng-hosting file statis aplikasinya; data disimpan di browser masing-masing perangkat, bukan di server.

---

## Pakai domain sendiri (opsional)

Kalau punya domain sendiri (misal `kuis.sekolah-anda.sch.id`), tambahkan file bernama `CNAME` (tanpa ekstensi) berisi domain tersebut ke dalam folder `public/` (buat foldernya kalau belum ada), lalu push/upload. Vite otomatis menyalin isi folder `public/` apa adanya ke hasil build. Selanjutnya atur DNS domain Anda mengikuti [panduan resmi GitHub](https://docs.github.com/pages/configuring-a-custom-domain-for-your-github-pages-site).
