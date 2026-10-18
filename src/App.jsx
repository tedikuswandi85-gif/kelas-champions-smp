import { useReducer, useEffect, useRef, useMemo, useState, useCallback } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  Trophy, Zap, Shield, Flame, Star, Play, Users, BookOpen, Settings,
  HelpCircle, ChevronRight, ChevronLeft, Pause, RotateCcw, Check, X,
  Home as HomeIcon, Award, Clock, Volume2, VolumeX, Swords, Target,
  ListChecks, BarChart3, Sparkles, Crown, ArrowRight, PlayCircle,
  PauseCircle, RefreshCw, Medal, GraduationCap, AlertTriangle, ShieldCheck,
  Gauge, Skull, CheckCircle2, Circle, ChevronDown, Search, Filter, Plus,
  Pencil, Trash2, Upload, Download, Save, LayoutDashboard, History,
  SlidersHorizontal, Image as ImageIcon, ChevronUp, Info, Loader2, DatabaseZap, LayoutGrid,
} from "lucide-react";

/* ============================================================
   BANK SOAL DEMO BAWAAN (dipakai untuk hidrasi awal & tombol "Reset Demo")
   ============================================================ */
const DEFAULT_QUESTION_BANK = [
  { id: "ips-01", subject: "IPS", grade: 9, topic: "Globalisasi", learningObjective: "Memahami konsep globalisasi", indicator: "Menjelaskan pengertian globalisasi", question: "Proses menyatunya warga dunia secara mendunia melalui pertukaran informasi, barang, dan budaya disebut...", type: "multiple_choice", options: ["Urbanisasi", "Globalisasi", "Migrasi", "Modernisasi"], answer: 1, explanation: "Globalisasi adalah proses mendunia yang menghubungkan masyarakat di berbagai negara melalui pertukaran informasi, ekonomi, dan budaya.", points: 100, timeLimit: 30, difficulty: "easy", imageUrl: null },
  { id: "ips-02", subject: "IPS", grade: 9, topic: "Globalisasi", learningObjective: "Mengidentifikasi bentuk globalisasi budaya", indicator: "Memberi contoh globalisasi budaya", question: "Berikut ini yang merupakan contoh globalisasi di bidang budaya adalah...", type: "multiple_choice", options: ["Ekspor kelapa sawit ke luar negeri", "Remaja Indonesia menonton drama Korea secara streaming", "Kenaikan suku bunga bank", "Pembangunan jalan tol antarkota"], answer: 1, explanation: "Menonton konten budaya asing secara mudah melalui internet adalah contoh nyata globalisasi budaya.", points: 100, timeLimit: 30, difficulty: "easy", imageUrl: null },
  { id: "ips-03", subject: "IPS", grade: 9, topic: "Globalisasi", learningObjective: "Menganalisis dampak globalisasi ekonomi", indicator: "Menjelaskan dampak positif bagi UMKM", question: "Dampak positif globalisasi ekonomi bagi pelaku UMKM di Indonesia adalah...", type: "multiple_choice", options: ["Produk lokal sulit bersaing dengan produk impor", "Pasar produk menjadi lebih luas lewat platform digital global", "Bahan baku produksi menjadi lebih mahal", "Pelaku usaha wajib menutup toko fisik"], answer: 1, explanation: "Globalisasi membuka akses pasar internasional bagi UMKM melalui e-commerce dan platform digital lintas negara.", points: 100, timeLimit: 30, difficulty: "medium", imageUrl: null },
  { id: "ips-04", subject: "IPS", grade: 9, topic: "Globalisasi", learningObjective: "Menganalisis dampak globalisasi budaya", indicator: "Menjelaskan dampak negatif globalisasi terhadap budaya lokal", question: "Salah satu dampak negatif globalisasi terhadap budaya lokal adalah...", type: "multiple_choice", options: ["Meningkatnya kunjungan wisatawan mancanegara", "Generasi muda mulai kurang mengenal permainan tradisional", "Bertambahnya lapangan kerja di sektor pariwisata", "Meningkatnya jumlah sekolah bahasa daerah"], answer: 1, explanation: "Derasnya budaya asing dapat menggeser minat generasi muda terhadap budaya dan tradisi lokal.", points: 100, timeLimit: 30, difficulty: "medium", imageUrl: null },
  { id: "ips-05", subject: "IPS", grade: 9, topic: "Globalisasi", learningObjective: "Merumuskan upaya pelestarian budaya", indicator: "Menentukan strategi pelestarian budaya lokal", question: "Manakah upaya paling tepat untuk melestarikan budaya lokal di tengah arus globalisasi?", type: "multiple_choice", options: ["Melarang seluruh budaya asing masuk ke Indonesia", "Mengenalkan budaya lokal lewat konten media sosial", "Menutup akses internet di daerah pedesaan", "Menghentikan kegiatan pertukaran pelajar"], answer: 1, explanation: "Memanfaatkan media digital untuk mempromosikan budaya lokal adalah cara adaptif melestarikan budaya di era global.", points: 200, timeLimit: 30, difficulty: "medium", imageUrl: null },
  { id: "ips-06", subject: "IPS", grade: 9, topic: "Globalisasi", learningObjective: "Memahami ketergantungan antarruang", indicator: "Memberi contoh ketergantungan antarruang", question: "Contoh ketergantungan antarruang akibat globalisasi adalah...", type: "multiple_choice", options: ["Indonesia mengimpor gandum karena keterbatasan lahan pertanian gandum", "Setiap daerah memproduksi seluruh kebutuhannya sendiri", "Desa tidak lagi berhubungan dengan kota", "Setiap negara menutup diri dari perdagangan internasional"], answer: 0, explanation: "Ketergantungan antarruang terjadi karena perbedaan sumber daya, sehingga suatu wilayah bergantung pada wilayah lain untuk memenuhi kebutuhannya.", points: 200, timeLimit: 30, difficulty: "medium", imageUrl: null },
  { id: "ips-07", subject: "IPS", grade: 9, topic: "Globalisasi", learningObjective: "Memahami cakupan globalisasi", indicator: "Menilai pernyataan tentang globalisasi", question: "Globalisasi hanya berdampak pada bidang ekonomi saja.", type: "true_false", options: ["Benar", "Salah"], answer: 1, explanation: "Globalisasi berdampak pada banyak bidang, seperti ekonomi, budaya, politik, teknologi, dan sosial.", points: 100, timeLimit: 20, difficulty: "easy", imageUrl: null },
  { id: "ips-08", subject: "IPS", grade: 9, topic: "Globalisasi", learningObjective: "Menganalisis kasus nyata globalisasi", indicator: "Menganalisis dampak digitalisasi UMKM desa", question: "Sebuah desa wisata mulai menjual kerajinan tangan melalui marketplace daring dan pesanan datang dari luar negeri. Dampak paling tepat dari fenomena ini adalah...", type: "multiple_choice", options: ["Desa tersebut semakin tertutup dari dunia luar", "Pendapatan warga berpotensi meningkat melalui pasar global", "Kerajinan tangan menjadi tidak bernilai lagi", "Warga desa wajib pindah ke kota"], answer: 1, explanation: "Akses ke pasar global melalui digitalisasi dapat meningkatkan pendapatan dan memperluas jangkauan pemasaran produk lokal.", points: 300, timeLimit: 30, difficulty: "hard", imageUrl: null },
  { id: "ips-09", subject: "IPS", grade: 9, topic: "Globalisasi", learningObjective: "Menganalisis data terkait globalisasi", indicator: "Menghubungkan data ekspor dengan faktor globalisasi", question: "Nilai ekspor kerajinan lokal terus meningkat setiap tahun sejak pelaku usaha menjual produknya lewat platform belanja daring internasional. Faktor pendorong utama peningkatan ini adalah...", type: "multiple_choice", options: ["Kenaikan harga bahan baku", "Kemudahan akses pasar global melalui teknologi digital", "Penurunan jumlah penduduk", "Pembatasan ekspor oleh pemerintah"], answer: 1, explanation: "Teknologi digital mempermudah pelaku usaha menjangkau pembeli di berbagai negara tanpa batasan geografis.", points: 300, timeLimit: 30, difficulty: "hard", imageUrl: null },
  { id: "ips-10", subject: "IPS", grade: 9, topic: "Globalisasi", learningObjective: "Menganalisis studi kasus globalisasi", indicator: "Mengaitkan gerai waralaba asing dengan globalisasi budaya", question: "Masuknya gerai makanan cepat saji dari luar negeri ke kota-kota kecil dan mengubah kebiasaan makan warga setempat merupakan contoh dari...", type: "multiple_choice", options: ["Globalisasi politik", "Globalisasi budaya dan gaya hidup", "Globalisasi pertahanan", "Otonomi daerah"], answer: 1, explanation: "Perubahan kebiasaan konsumsi akibat masuknya produk dan budaya asing termasuk globalisasi di bidang budaya dan gaya hidup.", points: 300, timeLimit: 30, difficulty: "hard", imageUrl: null },

  { id: "inf-01", subject: "Informatika", grade: 7, topic: "Berpikir Komputasional", learningObjective: "Mengenal perangkat keras komputer", indicator: "Menjelaskan fungsi CPU", question: "Perangkat yang berfungsi sebagai pusat pemrosesan atau 'otak' komputer adalah...", type: "multiple_choice", options: ["Monitor", "CPU (Processor)", "Mouse", "Speaker"], answer: 1, explanation: "CPU (Central Processing Unit) memproses seluruh instruksi dan perhitungan dalam komputer, sehingga disebut otak komputer.", points: 100, timeLimit: 30, difficulty: "easy", imageUrl: null },
  { id: "inf-02", subject: "Informatika", grade: 7, topic: "Berpikir Komputasional", learningObjective: "Memahami konsep algoritma", indicator: "Menjelaskan pengertian algoritma", question: "Urutan langkah logis dan sistematis untuk menyelesaikan suatu masalah disebut...", type: "multiple_choice", options: ["Algoritma", "Aplikasi", "Jaringan", "Basis data"], answer: 0, explanation: "Algoritma adalah kumpulan langkah yang tersusun secara logis dan berurutan untuk menyelesaikan sebuah masalah.", points: 100, timeLimit: 30, difficulty: "easy", imageUrl: null },
  { id: "inf-03", subject: "Informatika", grade: 7, topic: "Berpikir Komputasional", learningObjective: "Memahami sifat algoritma", indicator: "Menilai pernyataan tentang algoritma", question: "Algoritma harus memiliki langkah-langkah yang jelas dan berurutan.", type: "true_false", options: ["Benar", "Salah"], answer: 0, explanation: "Salah satu ciri utama algoritma adalah langkahnya jelas, logis, dan berurutan agar dapat diikuti dan menghasilkan solusi yang tepat.", points: 100, timeLimit: 20, difficulty: "easy", imageUrl: null },
  { id: "inf-04", subject: "Informatika", grade: 7, topic: "Sistem Komputer", learningObjective: "Mengenal perangkat lunak", indicator: "Mengidentifikasi contoh sistem operasi", question: "Berikut ini yang merupakan contoh perangkat lunak sistem operasi adalah...", type: "multiple_choice", options: ["Microsoft Word", "Windows", "Google Chrome", "Adobe Photoshop"], answer: 1, explanation: "Windows adalah sistem operasi yang mengatur seluruh sumber daya perangkat keras dan lunak pada komputer.", points: 100, timeLimit: 30, difficulty: "easy", imageUrl: null },
  { id: "inf-05", subject: "Informatika", grade: 7, topic: "Berpikir Komputasional", learningObjective: "Memahami dekomposisi masalah", indicator: "Menjelaskan makna dekomposisi", question: "Dalam berpikir komputasional, memecah masalah besar menjadi bagian-bagian kecil yang lebih mudah diselesaikan disebut...", type: "multiple_choice", options: ["Abstraksi", "Dekomposisi", "Generalisasi", "Enkripsi"], answer: 1, explanation: "Dekomposisi adalah proses memecah masalah kompleks menjadi bagian-bagian yang lebih kecil dan sederhana.", points: 200, timeLimit: 30, difficulty: "medium", imageUrl: null },
  { id: "inf-06", subject: "Informatika", grade: 7, topic: "Sistem Komputer", learningObjective: "Mengenal satuan data", indicator: "Menjelaskan satuan terkecil data digital", question: "Satuan data terkecil yang dikenali komputer, berupa nilai 0 atau 1, disebut...", type: "multiple_choice", options: ["Byte", "Bit", "Pixel", "Kilobyte"], answer: 1, explanation: "Bit (binary digit) adalah satuan data terkecil dalam sistem komputer yang bernilai 0 atau 1.", points: 200, timeLimit: 30, difficulty: "medium", imageUrl: null },
  { id: "inf-07", subject: "Informatika", grade: 7, topic: "Berpikir Komputasional", learningObjective: "Menerapkan berpikir komputasional pada masalah sehari-hari", indicator: "Menganalisis solusi berbasis dekomposisi", question: "Laptop terasa lambat karena banyak aplikasi terbuka bersamaan. Konsep berpikir komputasional yang paling tepat digunakan untuk menyelesaikan hal ini adalah...", type: "multiple_choice", options: ["Mengabaikan masalah karena akan hilang sendiri", "Mengenali pola aplikasi mana yang paling banyak memakan sumber daya lalu menutupnya", "Mengganti seluruh laptop dengan yang baru", "Mematikan koneksi internet selamanya"], answer: 1, explanation: "Pengenalan pola (pattern recognition) membantu mengidentifikasi penyebab utama masalah sebelum mengambil tindakan yang tepat.", points: 300, timeLimit: 30, difficulty: "hard", imageUrl: null },
  { id: "inf-08", subject: "Informatika", grade: 7, topic: "Sistem Komputer", learningObjective: "Memahami fungsi komponen komputer", indicator: "Menjelaskan fungsi RAM", question: "Fungsi utama RAM (Random Access Memory) pada komputer adalah...", type: "multiple_choice", options: ["Menyimpan data secara permanen selamanya", "Menyimpan data sementara yang sedang diproses agar kinerja lebih cepat", "Menampilkan gambar ke layar", "Menghubungkan komputer ke internet"], answer: 1, explanation: "RAM menyimpan data sementara yang sedang diakses prosesor sehingga mempercepat kinerja komputer.", points: 200, timeLimit: 30, difficulty: "medium", imageUrl: null },
  { id: "inf-09", subject: "Informatika", grade: 7, topic: "Berpikir Komputasional", learningObjective: "Memahami manfaat pengenalan pola", indicator: "Menjelaskan manfaat pattern recognition", question: "Pengenalan pola (pattern recognition) dalam berpikir komputasional bermanfaat untuk...", type: "multiple_choice", options: ["Menghapus semua data yang mirip", "Menemukan kesamaan atau tren untuk mempermudah penyelesaian masalah", "Membuat masalah menjadi lebih rumit", "Mengganti seluruh algoritma yang sudah ada"], answer: 1, explanation: "Pengenalan pola membantu menemukan kesamaan sehingga masalah dapat diselesaikan lebih efisien dengan solusi yang sudah terbukti berhasil.", points: 200, timeLimit: 30, difficulty: "medium", imageUrl: null },
  { id: "inf-10", subject: "Informatika", grade: 7, topic: "Sistem Komputer", learningObjective: "Menyusun langkah operasional dasar", indicator: "Mengurutkan langkah menyalakan komputer dengan benar", question: "Urutan yang paling tepat saat menyalakan komputer dari kondisi mati total adalah...", type: "multiple_choice", options: ["Menekan tombol power CPU, lalu menyalakan monitor", "Menekan tombol power monitor, lalu mencabut kabel CPU", "Langsung membuka aplikasi sebelum komputer menyala", "Mencabut semua kabel lalu menyalakan monitor saja"], answer: 0, explanation: "Langkah yang benar adalah menyalakan unit CPU terlebih dahulu (atau bersamaan), kemudian memastikan monitor menyala untuk menampilkan proses booting.", points: 300, timeLimit: 30, difficulty: "hard", imageUrl: null },

  { id: "ipa-01", subject: "IPA", grade: 8, topic: "Sistem Pencernaan", learningObjective: "Mengenal organ pencernaan manusia", indicator: "Mengidentifikasi organ tempat pencernaan mekanis dimulai", question: "Proses pencernaan mekanis pertama kali terjadi di organ...", type: "multiple_choice", options: ["Lambung", "Mulut", "Usus halus", "Usus besar"], answer: 1, explanation: "Pencernaan mekanis dimulai di mulut melalui proses mengunyah dengan bantuan gigi dan lidah.", points: 100, timeLimit: 30, difficulty: "easy", imageUrl: null },
  { id: "ipa-02", subject: "IPA", grade: 8, topic: "Sistem Pencernaan", learningObjective: "Memahami fungsi enzim pencernaan", indicator: "Menjelaskan fungsi enzim pepsin", question: "Enzim pepsin yang dihasilkan lambung berfungsi untuk...", type: "multiple_choice", options: ["Mengubah amilum menjadi glukosa", "Mengubah protein menjadi pepton", "Mengubah lemak menjadi asam lemak", "Menyerap air dari sisa makanan"], answer: 1, explanation: "Pepsin bekerja memecah protein menjadi molekul lebih sederhana yang disebut pepton di dalam lambung.", points: 200, timeLimit: 30, difficulty: "medium", imageUrl: null },
  { id: "ipa-03", subject: "IPA", grade: 8, topic: "Sistem Pencernaan", learningObjective: "Memahami proses penyerapan nutrisi", indicator: "Menilai pernyataan tentang usus halus", question: "Usus halus adalah organ utama tempat penyerapan sari-sari makanan ke dalam darah.", type: "true_false", options: ["Benar", "Salah"], answer: 0, explanation: "Usus halus memiliki jonjot usus (vili) yang berfungsi memperluas permukaan penyerapan sari makanan ke pembuluh darah.", points: 100, timeLimit: 20, difficulty: "easy", imageUrl: null },
  { id: "ipa-04", subject: "IPA", grade: 8, topic: "Sistem Pencernaan", learningObjective: "Menganalisis gangguan sistem pencernaan", indicator: "Menganalisis penyebab gangguan pencernaan dari pola makan", question: "Seorang siswa sering menunda makan dan jarang makan sayur berserat. Gangguan pencernaan yang paling mungkin ia alami adalah...", type: "multiple_choice", options: ["Sembelit (konstipasi)", "Patah tulang", "Anemia akibat kekurangan zat besi", "Mimisan"], answer: 0, explanation: "Kurangnya asupan serat dan pola makan tidak teratur dapat menyebabkan sembelit karena sisa makanan sulit bergerak di usus besar.", points: 300, timeLimit: 30, difficulty: "hard", imageUrl: null },
  { id: "ipa-05", subject: "IPA", grade: 8, topic: "Sistem Pencernaan", learningObjective: "Memahami urutan saluran pencernaan", indicator: "Mengurutkan organ pencernaan dengan benar", question: "Urutan organ pencernaan manusia yang benar adalah...", type: "multiple_choice", options: ["Mulut - kerongkongan - lambung - usus halus - usus besar", "Mulut - lambung - kerongkongan - usus besar - usus halus", "Kerongkongan - mulut - usus halus - lambung - usus besar", "Mulut - usus halus - lambung - kerongkongan - usus besar"], answer: 0, explanation: "Saluran pencernaan manusia berurutan dari mulut, kerongkongan, lambung, usus halus, hingga usus besar.", points: 200, timeLimit: 30, difficulty: "medium", imageUrl: null },

  { id: "mtk-01", subject: "Matematika", grade: 7, topic: "Bilangan Bulat", learningObjective: "Melakukan operasi bilangan bulat", indicator: "Menghitung penjumlahan bilangan bulat negatif", question: "Hasil dari (-8) + 5 adalah...", type: "multiple_choice", options: ["-3", "3", "-13", "13"], answer: 0, explanation: "(-8) + 5 = -(8-5) = -3, karena bilangan dengan nilai mutlak lebih besar (8) bertanda negatif.", points: 100, timeLimit: 20, difficulty: "easy", imageUrl: null },
  { id: "mtk-02", subject: "Matematika", grade: 7, topic: "Bilangan Bulat", learningObjective: "Melakukan operasi perkalian bilangan bulat", indicator: "Menghitung perkalian dua bilangan negatif", question: "Hasil dari (-6) x (-4) adalah...", type: "multiple_choice", options: ["-24", "24", "-10", "10"], answer: 1, explanation: "Perkalian dua bilangan bertanda sama (negatif x negatif) menghasilkan bilangan positif: (-6) x (-4) = 24.", points: 100, timeLimit: 20, difficulty: "easy", imageUrl: null },
  { id: "mtk-03", subject: "Matematika", grade: 7, topic: "Pecahan", learningObjective: "Membandingkan nilai pecahan", indicator: "Menilai pernyataan tentang pecahan senilai", question: "Pecahan 2/4 senilai dengan 1/2.", type: "true_false", options: ["Benar", "Salah"], answer: 0, explanation: "2/4 jika disederhanakan dengan membagi pembilang dan penyebut dengan 2 menghasilkan 1/2, sehingga keduanya senilai.", points: 100, timeLimit: 20, difficulty: "easy", imageUrl: null },
  { id: "mtk-04", subject: "Matematika", grade: 7, topic: "Pecahan", learningObjective: "Menyelesaikan operasi hitung pecahan", indicator: "Menghitung penjumlahan pecahan berbeda penyebut", question: "Hasil dari 1/3 + 1/4 adalah...", type: "multiple_choice", options: ["2/7", "7/12", "1/12", "5/7"], answer: 1, explanation: "Samakan penyebut menjadi 12: 4/12 + 3/12 = 7/12.", points: 200, timeLimit: 30, difficulty: "medium", imageUrl: null },
  { id: "mtk-05", subject: "Matematika", grade: 7, topic: "Bilangan Bulat", learningObjective: "Menerapkan bilangan bulat pada masalah kontekstual", indicator: "Menganalisis soal cerita bilangan bulat", question: "Suhu suatu ruangan pada pagi hari 2°C. Pada siang hari suhu naik 9°C, kemudian pada malam hari turun 15°C dari suhu siang. Suhu pada malam hari adalah...", type: "multiple_choice", options: ["-4°C", "4°C", "-6°C", "6°C"], answer: 0, explanation: "Siang: 2 + 9 = 11°C. Malam: 11 - 15 = -4°C.", points: 300, timeLimit: 30, difficulty: "hard", imageUrl: null },

  { id: "bindo-01", subject: "Bahasa Indonesia", grade: 7, cp: null, topic: "Teks Deskripsi", learningObjective: "Memahami tujuan teks deskripsi", indicator: "Menjelaskan tujuan penulisan teks deskripsi", question: "Tujuan utama penulisan teks deskripsi adalah...", type: "multiple_choice", options: ["Menceritakan urutan peristiwa secara kronologis", "Menggambarkan suatu objek secara rinci agar pembaca seolah melihat, mendengar, atau merasakannya sendiri", "Meyakinkan pembaca untuk melakukan sesuatu", "Menjelaskan langkah-langkah membuat sesuatu"], answer: 1, explanation: "Teks deskripsi bertujuan menggambarkan objek secara rinci sehingga pembaca seolah dapat melihat, mendengar, atau merasakan objek tersebut.", points: 100, timeLimit: 30, difficulty: "easy", imageUrl: null },
  { id: "bindo-02", subject: "Bahasa Indonesia", grade: 7, cp: null, topic: "Teks Deskripsi", learningObjective: "Mengidentifikasi struktur teks deskripsi", indicator: "Menyebutkan bagian awal struktur teks deskripsi", question: "Bagian struktur teks deskripsi yang berisi gambaran umum objek yang akan dideskripsikan disebut...", type: "multiple_choice", options: ["Identifikasi", "Deskripsi bagian", "Penutup", "Orientasi konflik"], answer: 0, explanation: "Bagian identifikasi berisi gambaran umum tentang objek yang akan dideskripsikan, sebelum masuk ke deskripsi bagian yang lebih rinci.", points: 100, timeLimit: 30, difficulty: "medium", imageUrl: null },
  { id: "bindo-03", subject: "Bahasa Indonesia", grade: 7, cp: null, topic: "Teks Deskripsi", learningObjective: "Membedakan kalimat perincian dan kalimat umum", indicator: "Menilai pernyataan tentang ciri kalimat perincian", question: "Teks deskripsi banyak menggunakan kalimat perincian yang melibatkan pancaindra.", type: "true_false", options: ["Benar", "Salah"], answer: 0, explanation: "Ciri khas teks deskripsi adalah penggunaan kalimat perincian yang melibatkan pancaindra (penglihatan, pendengaran, penciuman, perabaan, pengecapan) agar gambaran objek terasa hidup.", points: 100, timeLimit: 20, difficulty: "easy", imageUrl: null },
  { id: "bindo-04", subject: "Bahasa Indonesia", grade: 7, cp: null, topic: "Teks Deskripsi", learningObjective: "Menganalisis kalimat perincian yang tepat", indicator: "Memilih kalimat perincian paling efektif menggambarkan objek", question: "Kalimat manakah yang paling tepat digunakan sebagai kalimat perincian dalam teks deskripsi tentang pantai?", type: "multiple_choice", options: ["Pantai itu ada di Indonesia.", "Pasir putih pantai itu terasa lembut ketika diinjak, dan deburan ombaknya terdengar tenang sepanjang sore.", "Pantai adalah tempat wisata yang populer.", "Banyak orang suka pergi ke pantai."], answer: 1, explanation: "Kalimat kedua melibatkan pancaindra (perabaan dan pendengaran) sehingga lebih hidup dan sesuai ciri kalimat perincian dalam teks deskripsi, dibanding kalimat umum lainnya.", points: 200, timeLimit: 30, difficulty: "medium", imageUrl: null },
  { id: "bindo-05", subject: "Bahasa Indonesia", grade: 7, cp: null, topic: "Teks Deskripsi", learningObjective: "Menyunting penggunaan majas dalam teks deskripsi", indicator: "Menganalisis penggunaan kata bermajas dalam kalimat deskripsi", question: "Kalimat 'Daun-daun kelapa melambai lembut seolah menyapa setiap orang yang lewat' menggunakan majas...", type: "multiple_choice", options: ["Metafora", "Personifikasi", "Hiperbola", "Simile"], answer: 1, explanation: "Kalimat tersebut menggambarkan benda mati (daun kelapa) seolah memiliki sifat manusia (menyapa), sehingga termasuk majas personifikasi.", points: 300, timeLimit: 30, difficulty: "hard", imageUrl: null },
];

const SUBJECTS_META = [
  { key: "IPS", label: "IPS", icon: "🌏", color: "#2F6FED" },
  { key: "Informatika", label: "Informatika", icon: "💻", color: "#22C55E" },
  { key: "IPA", label: "IPA", icon: "🔬", color: "#F2B705" },
  { key: "Matematika", label: "Matematika", icon: "📐", color: "#A855F7" },
  { key: "Bahasa Indonesia", label: "Bahasa Indonesia", icon: "📖", color: "#EF4444" },
  { key: "Bahasa Inggris", label: "Bahasa Inggris", icon: "🔤", color: "#06B6D4" },
  { key: "Pendidikan Pancasila", label: "Pendidikan Pancasila", icon: "🇮🇩", color: "#DC2626" },
  { key: "Pendidikan Agama", label: "Pendidikan Agama", icon: "🙏", color: "#0D9488" },
  { key: "PJOK", label: "PJOK", icon: "⚽", color: "#F97316" },
  { key: "Seni Budaya", label: "Seni Budaya", icon: "🎨", color: "#EC4899" },
  { key: "Prakarya", label: "Prakarya", icon: "🛠️", color: "#78716C" },
  { key: "Bahasa Sunda", label: "Bahasa Sunda", icon: "🗣️", color: "#8B5CF6" },
];
// Daftar nama mapel baku yang selalu ditawarkan di Bank Soal, di luar mapel
// custom yang ditambahkan guru sendiri secara bebas (opsi "Custom").
const STANDARD_SUBJECT_NAMES = SUBJECTS_META.map((s) => s.key);
function getSubjectMeta(subjectName) {
  return SUBJECTS_META.find((s) => s.key === subjectName) || { key: subjectName, label: subjectName, icon: "📘", color: "#4C86FF" };
}
function getAvailableSubjects(bank) {
  const names = [...new Set(bank.map((q) => q.subject))].sort();
  return names.map(getSubjectMeta);
}

const TIMER_OPTIONS = [15, 30, 45, 60];
const MIN_TEAM_COUNT = 2;
const MAX_TEAM_COUNT = 10;
const GROUP_ICONS = [Shield, Zap, Flame, Star, Crown, Target, Gauge, Sparkles, Swords, Medal];
const GROUP_COLORS = ["#F2B705", "#2F6FED", "#22C55E", "#F97066", "#A855F7", "#06B6D4", "#EC4899", "#84CC16", "#FB923C", "#14B8A6"];
const KELOMPOK_PALETTE = ["#F2B705", "#2F6FED", "#22C55E", "#F97066", "#A855F7", "#06B6D4", "#EC4899", "#84CC16", "#FB923C", "#14B8A6"];
const DEFAULT_GROUP_NAMES = ["Garuda", "Elang", "Rajawali", "Cakrawala", "Harimau", "Komodo", "Merpati", "Cendrawasih", "Banteng", "Singa"];

const QUESTION_TYPES = [
  { key: "multiple_choice", label: "Pilihan Ganda" },
  { key: "true_false", label: "Benar / Salah" },
  { key: "short_answer", label: "Isian Singkat" },
  { key: "matching", label: "Menjodohkan" },
  { key: "image_guess", label: "Tebak Gambar" },
  { key: "step_order", label: "Susun Langkah" },
  { key: "case_study", label: "Studi Kasus" },
  { key: "chart_table", label: "Grafik / Tabel" },
];
const CHOICE_TYPES = ["multiple_choice", "true_false"];
const DIFFICULTY_OPTIONS = [
  { key: "easy", label: "Mudah" },
  { key: "medium", label: "Sedang" },
  { key: "hard", label: "Sulit" },
];

/* ============================================================
   TOURNAMENT ENGINE — DEFINISI BABAK
   ============================================================ */
const ROUND_ORDER = ["qualification", "headtohead", "strategy", "semifinal", "final"];

const ROUND_DEFS = {
  qualification: { key: "qualification", name: "Qualification", shortName: "Kualifikasi", icon: ListChecks, description: "Semua kelompok menjawab soal secara bergiliran untuk mengumpulkan poin dasar dan menentukan posisi awal.", rules: ["Semua kelompok ikut serta di setiap soal", "Jawaban benar mendapat poin sesuai tingkat kesulitan", "Tidak ada eliminasi pada babak ini"], questionCount: 6, timerSeconds: 30, pointsMode: "Standar (100–300 sesuai kesulitan)" },
  headtohead: { key: "headtohead", name: "Head to Head", shortName: "Duel", icon: Swords, description: "Guru memilih dua kelompok untuk saling berduel menjawab soal secara langsung.", rules: ["Guru memilih 2 kelompok yang akan berduel", "Guru menentukan kelompok mana yang menjawab tiap soal", "Kelompok dengan jawaban benar terbanyak menjadi pemenang duel"], questionCount: 4, timerSeconds: 20, pointsMode: "Standar per soal, direkap sebagai kemenangan duel" },
  strategy: { key: "strategy", name: "Strategy Arena", shortName: "Strategi", icon: Target, description: "Kelompok memilih sendiri tingkat tantangan — makin tinggi risikonya, makin besar hadiah poinnya.", rules: ["Kelompok yang tampil memilih tingkat kesulitan sebelum soal ditampilkan", "Easy 100 · Medium 200 · Hard 300 · Extreme 500 poin", "Jawaban salah pada Hard/Extreme berisiko kehilangan sebagian poin"], questionCount: 4, timerSeconds: 30, pointsMode: "Sesuai tantangan yang dipilih (100–500)" },
  semifinal: { key: "semifinal", name: "Semifinal", shortName: "Semifinal", icon: Award, description: "Guru menentukan finalis dari peringkat sementara. Hanya finalis yang melanjutkan permainan.", rules: ["Guru memilih finalis berdasarkan peringkat sementara", "Soal yang tampil lebih menantang (sedang–sulit)", "Skor diperbarui sebelum memasuki Grand Final"], questionCount: 4, timerSeconds: 25, pointsMode: "Standar, bobot soal lebih sulit" },
  final: { key: "final", name: "Grand Final", shortName: "Final", icon: Crown, description: "Babak puncak untuk menentukan Juara 1, 2, dan 3 di antara para finalis.", rules: ["Hanya diikuti oleh finalis", "Soal tersulit dengan tekanan waktu lebih ketat", "Peringkat akhir menentukan Juara 1, 2, dan 3"], questionCount: 4, timerSeconds: 20, pointsMode: "Standar + bonus final" },
};

const STRATEGY_LEVELS = [
  { key: "easy", label: "Easy", points: 100, riskLabel: "Tanpa risiko", color: "var(--green)", icon: ShieldCheck },
  { key: "medium", label: "Medium", points: 200, riskLabel: "Tanpa risiko", color: "var(--blue-soft)", icon: Gauge },
  { key: "hard", label: "Hard", points: 300, riskLabel: "Risiko: -30% poin bila salah", color: "var(--gold)", icon: AlertTriangle },
  { key: "extreme", label: "Extreme", points: 500, riskLabel: "Risiko: -50% poin bila salah", color: "var(--red)", icon: Skull },
];

const BOARD_POINT_LADDER = [100, 200, 300, 400, 500];
const BOARD_SIZE_PRESETS = [6, 9, 12, 16, 20, 25];
const BOARD_MIN_SIZE = 4; // jumlah kotak minimum agar papan tetap terasa seperti papan, bukan batas atas — maksimal bebas mengikuti soal yang tersedia

const NAV_ITEMS = [
  { key: "home", label: "Beranda", icon: HomeIcon, ready: true },
  { key: "quickplay-setup", label: "Quick Play", icon: Zap, ready: true, mode: "quickplay" },
  { key: "board-setup", label: "Papan Kuis", icon: LayoutGrid, ready: true, mode: "board" },
  { key: "tournament-setup", label: "Tournament", icon: Swords, ready: true, mode: "tournament" },
  { key: "practice-setup", label: "Practice", icon: Target, ready: true, mode: "practice" },
  { key: "teacher-panel", label: "Panel Guru", icon: GraduationCap, ready: true, teacherTab: "dashboard" },
  { key: "question-bank", label: "Bank Soal", icon: BookOpen, ready: true, teacherTab: "bank" },
  { key: "groups-manage", label: "Kelompok", icon: Users, ready: true, teacherTab: "kelompok" },
  { key: "guide", label: "Panduan", icon: HelpCircle, ready: true },
  { key: "settings", label: "Pengaturan", icon: Settings, ready: true },
];

/* ============================================================
   HELPERS UMUM
   ============================================================ */
function shuffleArray(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function makeGroupAt(i, override) {
  const name = (override && override.name) || DEFAULT_GROUP_NAMES[i % DEFAULT_GROUP_NAMES.length];
  const color = (override && override.color) || GROUP_COLORS[i % GROUP_COLORS.length];
  return { id: i, name, color, iconIdx: i % GROUP_ICONS.length, score: 0, correct: 0, wrong: 0, streak: 0, bestStreak: 0 };
}

function makeDefaultGroups(count = 4) {
  const n = Math.max(MIN_TEAM_COUNT, Math.min(MAX_TEAM_COUNT, count || 4));
  return Array.from({ length: n }, (_, i) => makeGroupAt(i));
}

function makeEmptyQuestion() {
  return { id: null, subject: "", grade: 7, cp: "", topic: "", learningObjective: "", indicator: "", question: "", type: "multiple_choice", options: ["", "", "", ""], answer: 0, explanation: "", points: 100, timeLimit: 30, difficulty: "medium", imageUrl: "" };
}

function genQuestionId() {
  return `custom-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
}

function genClassId() {
  return `class-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
}

function makeClassRecord(name, teams) {
  return { id: genClassId(), name: name || "Kelas Baru", teams: teams || makeDefaultGroups(4), savedSession: null };
}

function playBeep(kind) {
  try {
    const AudioCtx = window.AudioContext || window.webkitAudioContext;
    if (!AudioCtx) return;
    const ctx = new AudioCtx();
    const o = ctx.createOscillator();
    const g = ctx.createGain();
    o.connect(g); g.connect(ctx.destination); o.type = "sine";
    if (kind === "tick") { o.frequency.value = 880; g.gain.setValueAtTime(0.06, ctx.currentTime); }
    else if (kind === "correct") { o.frequency.value = 660; g.gain.setValueAtTime(0.12, ctx.currentTime); }
    else if (kind === "wrong") { o.frequency.value = 220; g.gain.setValueAtTime(0.12, ctx.currentTime); }
    else { o.frequency.value = 440; g.gain.setValueAtTime(0.12, ctx.currentTime); }
    o.start();
    g.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.35);
    o.stop(ctx.currentTime + 0.36);
    setTimeout(() => ctx.close(), 500);
  } catch (e) { /* audio tidak tersedia, abaikan */ }
}

function pickQuestions(bank, subject, grade, count, opts = {}) {
  const { excludeIds = [], preferDifficulties = null, topic = null, difficulty = null } = opts;
  // PENTING: pool SELALU dikunci ke subject && grade yang sama lebih dulu,
  // sehingga bank soal antar-mapel tidak akan pernah tercampur, bahkan
  // ketika mekanisme fallback di bawah harus mengulang soal.
  const basePool = bank.filter((q) => q.subject === subject && q.grade === grade);

  // Filter materi (topic), dengan fallback ke seluruh materi mapel & kelas
  // yang sama (tidak pernah keluar dari basePool) jika soal materi tsb kurang.
  let pool = basePool;
  if (topic) {
    const byTopic = basePool.filter((q) => q.topic === topic);
    if (byTopic.length > 0) pool = byTopic;
  }

  // Filter tingkat kesulitan, dengan fallback serupa (tetap di dalam pool
  // mapel & kelas & materi yang sudah difilter).
  if (difficulty && difficulty !== "all") {
    const byDifficulty = pool.filter((q) => q.difficulty === difficulty);
    if (byDifficulty.length > 0) pool = byDifficulty;
  }

  const fresh = pool.filter((q) => !excludeIds.includes(q.id));
  let source = fresh.length >= count ? fresh : pool;
  if (preferDifficulties) {
    const preferred = source.filter((q) => preferDifficulties.includes(q.difficulty));
    if (preferred.length >= Math.min(count, source.length)) source = preferred;
  }
  return shuffleArray(source).slice(0, Math.min(count, source.length));
}

/* ============================================================
   SCORE ENGINE — fungsi murni terpisah dari UI
   ============================================================ */
function calcStreakBonus(streak) { if (streak >= 5) return 100; if (streak >= 3) return 50; return 0; }
function calcChallengeBonus(difficulty) { if (difficulty === "extreme") return 100; if (difficulty === "hard") return 30; return 0; }
function calcPenalty(basePoints, difficulty) { if (difficulty === "extreme") return Math.round(basePoints * 0.5); if (difficulty === "hard") return Math.round(basePoints * 0.3); return 0; }
function calcPercentage(correct, total) { return total > 0 ? Math.round((correct / total) * 100) : 0; }
function updateRanking(teams) { return [...teams].sort((a, b) => b.score - a.score); }
function applyCorrectAnswer(team, basePoints, opts = {}) {
  const streak = team.streak + 1;
  const streakBonus = calcStreakBonus(streak);
  const challengeBonus = opts.difficulty ? calcChallengeBonus(opts.difficulty) : 0;
  const gained = basePoints + streakBonus + challengeBonus;
  return { ...team, score: team.score + gained, correct: team.correct + 1, streak, bestStreak: Math.max(team.bestStreak, streak), lastGain: gained };
}
function applyWrongAnswer(team, opts = {}) {
  const penalty = opts.penaltyEnabled ? (typeof opts.penaltyPoints === "number" ? opts.penaltyPoints : calcPenalty(opts.basePoints || 0, opts.difficulty)) : 0;
  return { ...team, score: Math.max(0, team.score - penalty), wrong: team.wrong + 1, streak: 0, lastGain: penalty > 0 ? -penalty : 0 };
}
function scoresFromTeams(teams) { const map = {}; teams.forEach((t) => { map[t.id] = t.score; }); return map; }
function buildResultsSnapshot(teams, totalQuestions, roundStatus) { return { ranking: updateRanking(teams), generatedAt: Date.now(), totalQuestions, roundStatus }; }
function buildHistoryEntry(state, modeOverride) {
  return {
    id: `hist-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
    timestamp: Date.now(),
    mode: modeOverride || state.gameMode,
    subject: state.subject,
    grade: state.grade,
    totalQuestions: state.playedLog.length,
    ranking: updateRanking(state.teams).map((t) => ({ name: t.name, score: t.score, correct: t.correct, wrong: t.wrong })),
  };
}

/* ============================================================
   CSV IMPORT / EKSPOR
   ============================================================ */
function downloadTextFile(content, filename, mime) {
  try {
    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
    return true;
  } catch (e) { return false; }
}
function questionsToCsvRows(list) {
  return list.map((q) => ({
    id: q.id, subject: q.subject, grade: q.grade, cp: q.cp || "", topic: q.topic,
    learningObjective: q.learningObjective || "", indicator: q.indicator || "",
    question: q.question, type: q.type,
    option1: q.options?.[0] || "", option2: q.options?.[1] || "", option3: q.options?.[2] || "", option4: q.options?.[3] || "",
    answer: q.answer, explanation: q.explanation || "", points: q.points, timeLimit: q.timeLimit,
    difficulty: q.difficulty, imageUrl: q.imageUrl || "",
  }));
}
function exportQuestionsCsv(list) {
  const csv = Papa.unparse(questionsToCsvRows(list));
  downloadTextFile(csv, `bank-soal-kelas-champions-${Date.now()}.csv`, "text/csv;charset=utf-8;");
}
const QUESTION_XLSX_TEMPLATE_ROW = {
  subject: "IPA", grade: 7, cp: "", topic: "Contoh Materi", learningObjective: "", indicator: "",
  question: "Contoh: Planet terdekat dengan Matahari adalah?", type: "multiple_choice",
  option1: "Merkurius", option2: "Venus", option3: "Bumi", option4: "Mars",
  answer: 0, explanation: "Merkurius adalah planet terdekat dengan Matahari.", points: 100, timeLimit: 30,
  difficulty: "easy", imageUrl: "",
};
function exportQuestionsXlsx(list, filename) {
  const rows = list.length ? questionsToCsvRows(list) : [QUESTION_XLSX_TEMPLATE_ROW];
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = Object.keys(rows[0]).map((k) => ({ wch: Math.max(12, Math.min(40, k.length + 6)) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Bank Soal");
  XLSX.writeFile(wb, filename || `bank-soal-kelas-champions-${Date.now()}.xlsx`);
}
function downloadQuestionTemplateXlsx() {
  const ws = XLSX.utils.json_to_sheet([QUESTION_XLSX_TEMPLATE_ROW]);
  ws["!cols"] = Object.keys(QUESTION_XLSX_TEMPLATE_ROW).map((k) => ({ wch: Math.max(12, Math.min(40, k.length + 6)) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Template Soal");
  XLSX.writeFile(wb, "template-import-soal-kelas-champions.xlsx");
}
function exportHistoryCsv(history) {
  const rows = [];
  history.forEach((h) => {
    h.ranking.forEach((r, i) => {
      rows.push({ tanggal: new Date(h.timestamp).toLocaleString("id-ID"), mode: h.mode, subject: h.subject, grade: h.grade, peringkat: i + 1, kelompok: r.name, skor: r.score, benar: r.correct, salah: r.wrong });
    });
  });
  const csv = Papa.unparse(rows);
  downloadTextFile(csv, `riwayat-hasil-kelas-champions-${Date.now()}.csv`, "text/csv;charset=utf-8;");
}
function cellStr(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}
// Baris sumber bisa berasal dari CSV (Papa.parse) atau Excel (XLSX.utils.sheet_to_json) —
// keduanya menghasilkan array of plain objects dengan nama kolom yang sama, jadi logika
// validasi & konversi ke bentuk soal ini dibagi satu sumber kebenaran untuk kedua format.
function hashStr(str) {
  let h = 5381;
  for (let i = 0; i < str.length; i++) h = ((h << 5) + h + str.charCodeAt(i)) | 0;
  return (h >>> 0).toString(36);
}
function buildQuestionsFromRows(rows, opts) {
  const errors = [];
  const questions = [];
  const usedIds = new Set();
  // stableIds (sinkron Google Sheets): id soal dibuat tetap dari kolom 'id' atau isi soalnya,
  // supaya sinkron ulang memperbarui soal yang sama — bukan menambah duplikat.
  const makeId = (row, question) => {
    if (!opts?.stableIds) return genQuestionId();
    const rid = cellStr(row.id);
    let id = `sheet-${rid || hashStr([cellStr(row.subject), cellStr(row.grade), question].join("|"))}`;
    let n = 2;
    while (usedIds.has(id)) id = `${id.replace(/-dup\d+$/, "")}-dup${n++}`;
    usedIds.add(id);
    return id;
  };
  (rows || []).forEach((row, i) => {
    const question = cellStr(row.question);
    if (!question) { errors.push(`Baris ${i + 2}: kolom 'question' kosong, dilewati.`); return; }
    const typeRaw = cellStr(row.type);
    const type = QUESTION_TYPES.some((t) => t.key === typeRaw) ? typeRaw : "multiple_choice";
    let options = [];
    let answer = row.answer;
    if (type === "true_false") { options = ["Benar", "Salah"]; answer = parseInt(answer, 10) === 1 ? 1 : 0; }
    else if (type === "multiple_choice") {
      options = [row.option1, row.option2, row.option3, row.option4].map((v) => cellStr(v)).filter((v) => v !== "");
      if (options.length < 2) { errors.push(`Baris ${i + 2}: pilihan ganda butuh minimal 2 opsi, dilewati.`); return; }
      answer = Math.min(parseInt(answer, 10) || 0, options.length - 1);
    } else { options = []; answer = cellStr(answer); }
    questions.push({
      id: makeId(row, question), subject: cellStr(row.subject) || "Custom", grade: parseInt(row.grade, 10) || 7, cp: cellStr(row.cp),
      topic: cellStr(row.topic) || "-", learningObjective: cellStr(row.learningObjective), indicator: cellStr(row.indicator),
      question, type, options, answer, explanation: cellStr(row.explanation),
      points: parseInt(row.points, 10) || 100, timeLimit: parseInt(row.timeLimit, 10) || 30,
      difficulty: ["easy", "medium", "hard"].includes(cellStr(row.difficulty)) ? cellStr(row.difficulty) : "medium",
      imageUrl: cellStr(row.imageUrl) || null,
    });
  });
  return { questions, errors };
}
// Menerima balasan dari Google Sheets: JSON dari Apps Script ({rows:[...]}) atau CSV hasil "Publikasikan ke web".
function parseSheetPayload(text) {
  const t = String(text || "").trim();
  if (!t) return { questions: [], errors: ["Balasan dari Google Sheets kosong."] };
  if (t.startsWith("{") || t.startsWith("[")) {
    try {
      const data = JSON.parse(t);
      if (data && data.error) return { questions: [], errors: [`Google Sheets: ${data.error}`] };
      const rows = Array.isArray(data) ? data : data.rows;
      if (!Array.isArray(rows) || !rows.length) return { questions: [], errors: ["Sheet kosong atau tidak memiliki baris soal."] };
      return buildQuestionsFromRows(rows, { stableIds: true });
    } catch (e) { return { questions: [], errors: ["Balasan JSON dari Google Sheets tidak bisa dibaca."] }; }
  }
  if (t.startsWith("<")) return { questions: [], errors: ["Balasan berupa halaman web (bukan data). Pastikan Web App di-deploy dengan akses \"Siapa saja\" atau sheet sudah dipublikasikan sebagai CSV."] };
  const parsed = Papa.parse(t, { header: true, skipEmptyLines: true });
  return buildQuestionsFromRows(parsed.data, { stableIds: true });
}
function parseImportedQuestionsCsv(text) {
  const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
  return buildQuestionsFromRows(parsed.data);
}
function parseImportedQuestionsXlsx(arrayBuffer) {
  try {
    const wb = XLSX.read(arrayBuffer, { type: "array" });
    const sheetName = wb.SheetNames[0];
    if (!sheetName) return { questions: [], errors: ["File Excel tidak memiliki sheet apapun."] };
    const sheet = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    if (!rows.length) return { questions: [], errors: ["Sheet pertama pada file Excel kosong atau tidak memiliki baris data."] };
    return buildQuestionsFromRows(rows);
  } catch (e) {
    return { questions: [], errors: [`Gagal membaca file Excel: ${e?.message || "format tidak dikenali"}. Pastikan file berformat .xlsx/.xls dan baris pertama berisi nama kolom.`] };
  }
}

/* ============================================================
   PENYIMPANAN PERSISTEN
   Catatan penting: lingkungan artifact Claude TIDAK mendukung
   localStorage/sessionStorage. Sebagai gantinya kami memakai
   window.storage — API penyimpanan persisten bawaan Claude yang
   berfungsi setara (data tersimpan permanen per pengguna) dan
   memang dirancang untuk kebutuhan seperti ini di dalam artifact.
   Saat di-deploy sebagai situs mandiri (di luar Claude), window.storage
   tidak akan ada — kode di bawah otomatis beralih memakai localStorage
   browser biasa, yang didukung penuh di lingkungan hosting normal.
   ============================================================ */
const STORAGE_KEYS = { bank: "kc_question_bank_v1", teams: "kc_teams_config_v1", settings: "kc_settings_v1", history: "kc_history_v1", classes: "kc_classes_v1" };

async function loadStorageSlice(key, fallback) {
  try {
    if (typeof window !== "undefined" && window.storage && typeof window.storage.get === "function") {
      const res = await window.storage.get(key, false);
      if (res && res.value) return JSON.parse(res.value);
      return fallback;
    }
    if (typeof localStorage !== "undefined") {
      const raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : fallback;
    }
    return fallback;
  } catch (e) { return fallback; }
}
async function saveStorageSlice(key, value) {
  try {
    if (typeof window !== "undefined" && window.storage && typeof window.storage.set === "function") {
      await window.storage.set(key, JSON.stringify(value), false);
      return true;
    }
    if (typeof localStorage !== "undefined") {
      localStorage.setItem(key, JSON.stringify(value));
      return true;
    }
    return false;
  } catch (e) { return false; }
}

/* ============================================================
   GAME STATE + REDUCER (GAME ENGINE CORE)
   ============================================================ */
function screenForRound(round, finalists) {
  if ((round === "semifinal" || round === "final") && finalists.length === 0) return "finalist-select";
  return "round-transition";
}

const DEFAULT_SETTINGS = { defaultTimerSeconds: 30, quickplayDefaultCount: 10, strategyPenaltyEnabled: true, selfSelectPenaltyEnabled: true, boardDefaultSize: 12 };

// Kolom-kolom state yang membentuk satu "aktivitas/sesi kuis" yang sedang berjalan.
// Disimpan (pause) dan dipulihkan (resume) per kelas, supaya kuis bisa dilanjutkan di lain waktu.
const SESSION_FIELDS = [
  "activeScreen", "gameMode", "subject", "grade", "topic", "topicFilter", "difficultyFilter",
  "teams", "questions", "currentQuestionIndex", "currentRound", "selectedRounds", "roundStatus",
  "scores", "selectedAnswers", "playedLog", "usedQuestionIds", "timerSecondsConfig", "timer",
  "gameStatus", "duel", "finalists", "strategySelection", "answeringTeamId",
  "boardQuestions", "boardTiles", "boardTurnOrder", "boardTurnIndex", "boardActiveTileId", "boardTimerSeconds",
  "results",
];
function pickSessionSnapshot(state) {
  const snap = {};
  SESSION_FIELDS.forEach((k) => { snap[k] = state[k]; });
  snap.timer = { ...state.timer, running: false }; // jangan biarkan timer tetap berjalan di latar saat disimpan
  return snap;
}
// Apakah ada aktivitas kuis yang sedang berlangsung (bukan di Home/setup) dan layak ditawarkan untuk disimpan.
function hasActiveSession(state) {
  return !!state.gameMode && state.activeScreen !== "home" && state.gameStatus !== "setup";
}
// Menitipkan kelas yang sedang aktif ke dalam array `classes` sebelum berpindah kelas lain.
// Jika ada kuis yang sedang berlangsung TAPI belum ditekan "Simpan & Keluar" secara manual (mis. guru
// langsung berpindah kelas lewat Panel Guru di tengah permainan), progresnya tetap otomatis
// disimpan sebagai sesi tertunda milik kelas tersebut — tidak pernah hilang begitu saja.
function stashActiveClassIntoClasses(state) {
  if (hasActiveSession(state)) {
    const snapshot = pickSessionSnapshot(state);
    return state.classes.map((c) => (c.id === state.activeClassId ? { ...c, teams: state.teams, savedSession: { ...snapshot, savedAt: Date.now() } } : c));
  }
  return state.classes.map((c) => (c.id === state.activeClassId ? { ...c, teams: state.teams } : c));
}
// Nilai default untuk seluruh kolom sesi/aktivitas (dipakai saat pindah/menambah/menghapus kelas, atau setelah menyimpan sesi).
const GAME_MODE_LABELS = { quickplay: "Quick Play", practice: "Practice Mode", tournament: "Tournament Mode", board: "Papan Kuis" };
const IDLE_SESSION_DEFAULTS = {
  activeScreen: "home", gameMode: null, subject: null, grade: null, topic: null, topicFilter: null, difficultyFilter: null,
  questions: [], currentQuestionIndex: 0, currentRound: null, selectedRounds: [...ROUND_ORDER], roundStatus: {},
  scores: {}, selectedAnswers: {}, playedLog: [], usedQuestionIds: [],
  timerSecondsConfig: 30, timer: { total: 30, remaining: 30, running: false, locked: false },
  gameStatus: "setup", duel: null, finalists: [], strategySelection: null, answeringTeamId: null,
  boardQuestions: [], boardTiles: [], boardTurnOrder: [], boardTurnIndex: 0, boardActiveTileId: null, boardTimerSeconds: 30,
  results: null,
};

const initialGameState = {
  activeScreen: "home",
  gameMode: null,
  subject: null, grade: null, topic: null, topicFilter: null, difficultyFilter: null,
  teams: makeDefaultGroups(),
  classes: [makeClassRecord("Kelas Saya", makeDefaultGroups())],
  activeClassId: null, // diisi ke id kelas pertama tepat di bawah
  questionBank: DEFAULT_QUESTION_BANK,
  questions: [],
  currentQuestionIndex: 0,
  currentRound: null,
  selectedRounds: [...ROUND_ORDER],
  roundStatus: {},
  scores: {},
  selectedAnswers: {},
  playedLog: [],
  usedQuestionIds: [],
  timerSecondsConfig: 30,
  timer: { total: 30, remaining: 30, running: false, locked: false },
  gameStatus: "setup",
  duel: null,
  finalists: [],
  strategySelection: null,
  answeringTeamId: null,
  boardQuestions: [],
  boardTiles: [],
  boardTurnOrder: [],
  boardTurnIndex: 0,
  boardActiveTileId: null,
  boardTimerSeconds: 30,
  results: null,
  soundOn: true,
  settings: DEFAULT_SETTINGS,
  history: [],
  teacherTab: "dashboard",
  storageReady: false,
};
initialGameState.activeClassId = initialGameState.classes[0].id;
initialGameState.teams = initialGameState.classes[0].teams;

function gameReducer(state, action) {
  switch (action.type) {
    case "SET_SCREEN": return { ...state, activeScreen: action.payload };
    case "SET_TEACHER_TAB": return { ...state, teacherTab: action.payload };
    case "TOGGLE_SOUND": return { ...state, soundOn: !state.soundOn };
    case "UPDATE_TEAM_NAME": return { ...state, teams: state.teams.map((t) => (t.id === action.payload.id ? { ...t, name: action.payload.name } : t)) };
    case "UPDATE_TEAM_COLOR": return { ...state, teams: state.teams.map((t) => (t.id === action.payload.id ? { ...t, color: action.payload.color } : t)) };

    case "SET_TEAM_COUNT": {
      const count = Math.max(MIN_TEAM_COUNT, Math.min(MAX_TEAM_COUNT, parseInt(action.payload, 10) || state.teams.length));
      if (count === state.teams.length) return state;
      if (count < state.teams.length) return { ...state, teams: state.teams.slice(0, count) };
      const extra = Array.from({ length: count - state.teams.length }, (_, idx) => makeGroupAt(state.teams.length + idx));
      return { ...state, teams: [...state.teams, ...extra] };
    }

    case "HYDRATE": {
      const { questionBank, teamsConfig, settings, history, classesData } = action.payload;
      let classes = state.classes;
      let activeClassId = state.activeClassId;
      let teams = state.teams;
      if (classesData && Array.isArray(classesData.classes) && classesData.classes.length > 0) {
        // Data kelas (fitur baru) sudah pernah tersimpan — pakai itu.
        classes = classesData.classes.map((c) => ({
          id: c.id || genClassId(),
          name: c.name || "Kelas",
          teams: Array.isArray(c.teams) && c.teams.length >= MIN_TEAM_COUNT
            ? Array.from({ length: Math.min(c.teams.length, MAX_TEAM_COUNT) }, (_, i) => makeGroupAt(i, c.teams[i]))
            : makeDefaultGroups(4),
          savedSession: c.savedSession || null,
        }));
        activeClassId = classesData.activeClassId && classes.some((c) => c.id === classesData.activeClassId)
          ? classesData.activeClassId
          : classes[0].id;
        teams = classes.find((c) => c.id === activeClassId).teams;
      } else if (Array.isArray(teamsConfig) && teamsConfig.length >= MIN_TEAM_COUNT) {
        // Migrasi dari data lama (sebelum fitur Kelas ada): roster tunggal yang tersimpan jadi satu kelas default.
        teams = Array.from({ length: Math.min(teamsConfig.length, MAX_TEAM_COUNT) }, (_, i) => makeGroupAt(i, teamsConfig[i]));
        classes = [makeClassRecord("Kelas Saya", teams)];
        activeClassId = classes[0].id;
      }
      return {
        ...state,
        questionBank: questionBank && questionBank.length ? questionBank : state.questionBank,
        teams, classes, activeClassId,
        settings: { ...DEFAULT_SETTINGS, ...(settings || {}) },
        history: Array.isArray(history) ? history : state.history,
        storageReady: true,
      };
    }

    /* ---------- KELAS & SESI TERSIMPAN ---------- */
    case "ADD_CLASS": {
      const name = (action.payload?.name || "").trim() || `Kelas ${state.classes.length + 1}`;
      const newClass = makeClassRecord(name, makeDefaultGroups(state.teams.length || 4));
      const classes = stashActiveClassIntoClasses(state);
      return { ...state, ...IDLE_SESSION_DEFAULTS, classes: [...classes, newClass], activeClassId: newClass.id, teams: newClass.teams };
    }
    case "RENAME_CLASS": {
      const trimmed = (action.payload?.name || "").trim();
      if (!trimmed) return state;
      return { ...state, classes: state.classes.map((c) => (c.id === action.payload.id ? { ...c, name: trimmed } : c)) };
    }
    case "DELETE_CLASS": {
      if (state.classes.length <= 1) return state; // minimal 1 kelas harus selalu ada
      const remaining = state.classes.filter((c) => c.id !== action.payload.id);
      if (action.payload.id !== state.activeClassId) return { ...state, classes: remaining };
      const next = remaining[0];
      return { ...state, ...IDLE_SESSION_DEFAULTS, classes: remaining, activeClassId: next.id, teams: next.teams };
    }
    case "SET_ACTIVE_CLASS": {
      const { id } = action.payload;
      if (id === state.activeClassId) return state;
      const target = state.classes.find((c) => c.id === id);
      if (!target) return state;
      const classes = stashActiveClassIntoClasses(state);
      return { ...state, ...IDLE_SESSION_DEFAULTS, classes, activeClassId: id, teams: target.teams };
    }
    case "PAUSE_SESSION": {
      const snapshot = pickSessionSnapshot(state);
      const classes = state.classes.map((c) => (c.id === state.activeClassId ? { ...c, teams: state.teams, savedSession: { ...snapshot, savedAt: Date.now() } } : c));
      return { ...state, ...IDLE_SESSION_DEFAULTS, classes, teams: state.teams.map((t) => ({ ...t, score: 0, correct: 0, wrong: 0, streak: 0, bestStreak: 0 })) };
    }
    case "RESUME_SESSION": {
      const { classId } = action.payload;
      const target = state.classes.find((c) => c.id === classId);
      if (!target || !target.savedSession) return state;
      const classes = (classId === state.activeClassId ? state.classes : stashActiveClassIntoClasses(state))
        .map((c) => (c.id === classId ? { ...c, savedSession: null } : c));
      return { ...state, ...target.savedSession, classes, activeClassId: classId };
    }
    case "DISCARD_SESSION": {
      const { classId } = action.payload;
      return { ...state, classes: state.classes.map((c) => (c.id === classId ? { ...c, savedSession: null } : c)) };
    }

    /* ---------- BANK SOAL CRUD ---------- */
    case "ADD_QUESTION": return { ...state, questionBank: [action.payload, ...state.questionBank] };
    case "UPDATE_QUESTION": return { ...state, questionBank: state.questionBank.map((q) => (q.id === action.payload.id ? action.payload : q)) };
    case "DELETE_QUESTION": return { ...state, questionBank: state.questionBank.filter((q) => q.id !== action.payload) };
    case "IMPORT_QUESTIONS": return { ...state, questionBank: [...action.payload, ...state.questionBank] };
    case "SYNC_SHEET_QUESTIONS": {
      if (!Array.isArray(action.payload) || !action.payload.length) return state;
      const keep = state.questionBank.filter((q) => !String(q.id).startsWith("sheet-"));
      return { ...state, questionBank: [...action.payload, ...keep] };
    }
    case "RESET_QUESTION_BANK_DEMO": return { ...state, questionBank: DEFAULT_QUESTION_BANK };

    /* ---------- PENGATURAN & RIWAYAT ---------- */
    case "SET_SETTINGS": return { ...state, settings: { ...state.settings, ...action.payload } };
    case "ADD_HISTORY_ENTRY": return { ...state, history: [action.payload, ...state.history].slice(0, 50) };
    case "CLEAR_HISTORY": return { ...state, history: [] };
    case "RESET_TEAMS_CONFIG": return { ...state, teams: makeDefaultGroups(state.teams.length) };

    case "START_QUICKPLAY": {
      const { subject, grade, questionCount, timerSeconds, topic = null, difficulty = null, mode = "quickplay" } = action.payload;
      const qs = pickQuestions(state.questionBank, subject, grade, questionCount, { topic, difficulty });
      const teams = state.teams.map((t) => ({ ...t, score: 0, correct: 0, wrong: 0, streak: 0, bestStreak: 0 }));
      return {
        ...state, gameMode: mode, subject, grade, topic: qs[0]?.topic || null,
        topicFilter: topic, difficultyFilter: difficulty,
        teams, scores: scoresFromTeams(teams), questions: qs, usedQuestionIds: qs.map((q) => q.id),
        currentQuestionIndex: 0, currentRound: null, selectedRounds: [], roundStatus: {}, finalists: [],
        duel: null, strategySelection: null, answeringTeamId: null, playedLog: [], selectedAnswers: {},
        timerSecondsConfig: timerSeconds, timer: { total: timerSeconds, remaining: timerSeconds, running: true, locked: false },
        gameStatus: "playing", activeScreen: "arena", results: null,
      };
    }

    case "START_BOARD": {
      const { subject, grade, boardSize, timerSeconds, topic = null, difficulty = null } = action.payload;
      const qs = pickQuestions(state.questionBank, subject, grade, boardSize, { topic, difficulty });
      const points = Array.from({ length: qs.length }, (_, i) => BOARD_POINT_LADDER[i % BOARD_POINT_LADDER.length]);
      const shuffledPoints = shuffleArray(points);
      const tiles = qs.map((q, i) => ({ id: i, points: shuffledPoints[i], questionId: q.id, used: false, answeredBy: null, result: null }));
      const teams = state.teams.map((t) => ({ ...t, score: 0, correct: 0, wrong: 0, streak: 0, bestStreak: 0 }));
      return {
        ...state, gameMode: "board", subject, grade, topic: null, topicFilter: topic, difficultyFilter: difficulty,
        teams, scores: scoresFromTeams(teams),
        boardQuestions: qs, boardTiles: tiles, boardTurnOrder: teams.map((t) => t.id), boardTurnIndex: 0, boardActiveTileId: null, boardTimerSeconds: timerSeconds,
        questions: [], currentQuestionIndex: 0, currentRound: null, selectedRounds: [], roundStatus: {}, finalists: [],
        duel: null, strategySelection: null, answeringTeamId: null, playedLog: [], selectedAnswers: {}, usedQuestionIds: qs.map((q) => q.id),
        timerSecondsConfig: timerSeconds, timer: { total: timerSeconds, remaining: timerSeconds, running: false, locked: false },
        gameStatus: "ready", activeScreen: "board-arena", results: null,
      };
    }

    case "PICK_TILE": {
      if (state.boardActiveTileId != null) return state; // guard: sudah ada soal terbuka, tidak bisa pilih ubin lain
      const tile = state.boardTiles.find((t) => t.id === action.payload);
      if (!tile || tile.used) return state;
      const q = state.boardQuestions.find((bq) => bq.id === tile.questionId);
      if (!q) return state;
      const currentTeamId = state.boardTurnOrder[state.boardTurnIndex % state.boardTurnOrder.length];
      return {
        ...state, boardActiveTileId: tile.id, questions: [q], currentQuestionIndex: 0, answeringTeamId: currentTeamId,
        timerSecondsConfig: state.boardTimerSeconds,
        timer: { total: state.boardTimerSeconds, remaining: state.boardTimerSeconds, running: true, locked: false },
        gameStatus: "playing",
      };
    }

    case "BOARD_NEXT": {
      const activeQ = state.questions[0];
      const answer = activeQ ? state.selectedAnswers[activeQ.id] : null;
      const newTiles = state.boardTiles.map((t) => (t.id === state.boardActiveTileId ? { ...t, used: true, answeredBy: answer?.teamId ?? null, result: answer?.result ?? null } : t));
      const allUsed = newTiles.every((t) => t.used);
      const nextTurnIndex = (state.boardTurnIndex + 1) % Math.max(1, state.boardTurnOrder.length);
      if (allUsed) {
        const entry = buildHistoryEntry(state);
        return {
          ...state, boardTiles: newTiles, boardActiveTileId: null, gameStatus: "finished", activeScreen: "results",
          results: buildResultsSnapshot(state.teams, state.playedLog.length, state.roundStatus),
          history: [entry, ...state.history].slice(0, 50),
        };
      }
      return {
        ...state, boardTiles: newTiles, boardActiveTileId: null, boardTurnIndex: nextTurnIndex,
        answeringTeamId: null, gameStatus: "ready", questions: [], currentQuestionIndex: 0,
      };
    }

    case "START_TOURNAMENT": {
      const { subject, grade, selectedRounds, topic = null, difficulty = null } = action.payload;
      const order = ROUND_ORDER.filter((r) => selectedRounds.includes(r));
      const firstRound = order[0];
      const roundStatus = {};
      ROUND_ORDER.forEach((r) => { roundStatus[r] = selectedRounds.includes(r) ? "pending" : "skipped"; });
      const teams = state.teams.map((t) => ({ ...t, score: 0, correct: 0, wrong: 0, streak: 0, bestStreak: 0 }));
      return {
        ...state, gameMode: "tournament", subject, grade, topic: null, topicFilter: topic, difficultyFilter: difficulty,
        teams, scores: scoresFromTeams(teams),
        selectedRounds: order, roundStatus, currentRound: firstRound, usedQuestionIds: [], finalists: [],
        duel: null, strategySelection: null, answeringTeamId: null, playedLog: [], selectedAnswers: {}, questions: [], currentQuestionIndex: 0,
        activeScreen: firstRound ? screenForRound(firstRound, []) : "home", gameStatus: "ready", results: null,
      };
    }

    case "START_ROUND": {
      const def = ROUND_DEFS[state.currentRound];
      if (!def) return state;
      const preferDifficulties = state.currentRound === "semifinal" || state.currentRound === "final" ? ["medium", "hard"] : null;
      const qs = pickQuestions(state.questionBank, state.subject, state.grade, def.questionCount, { excludeIds: state.usedQuestionIds, preferDifficulties, topic: state.topicFilter, difficulty: state.difficultyFilter });
      const isStrategy = state.currentRound === "strategy";
      const duelTeams = action.payload && action.payload.duelTeams;
      return {
        ...state, questions: qs, usedQuestionIds: [...state.usedQuestionIds, ...qs.map((q) => q.id)], currentQuestionIndex: 0,
        topic: qs[0]?.topic || state.topic, timerSecondsConfig: def.timerSeconds,
        timer: { total: def.timerSeconds, remaining: def.timerSeconds, running: !isStrategy, locked: false },
        gameStatus: isStrategy ? "ready" : "playing", roundStatus: { ...state.roundStatus, [state.currentRound]: "active" },
        duel: state.currentRound === "headtohead" ? (duelTeams ? { teamAId: duelTeams.teamAId, teamBId: duelTeams.teamBId, wins: {} } : state.duel) : state.duel,
        strategySelection: null, answeringTeamId: null, selectedAnswers: {},
        activeScreen: state.currentRound === "headtohead" ? "duel-arena" : state.currentRound === "strategy" ? "strategy-arena" : "arena",
      };
    }

    case "SET_FINALISTS": return { ...state, finalists: action.payload };
    case "SET_ANSWERING_TEAM": if (state.timer.locked) return state; return { ...state, answeringTeamId: action.payload };
    case "STRATEGY_SELECT_TEAM": return { ...state, strategySelection: { teamId: action.payload, difficulty: null, confirmed: false } };
    case "STRATEGY_SELECT_DIFFICULTY": if (!state.strategySelection) return state; return { ...state, strategySelection: { ...state.strategySelection, difficulty: action.payload, confirmed: false } };
    case "STRATEGY_RESET_SELECTION": return { ...state, strategySelection: null };
    case "STRATEGY_CONFIRM": {
      if (!state.strategySelection || !state.strategySelection.difficulty) return state;
      return { ...state, strategySelection: { ...state.strategySelection, confirmed: true }, timer: { ...state.timer, running: true, locked: false }, gameStatus: "playing" };
    }

    case "TIMER_TICK": {
      if (!state.timer.running || state.timer.locked || state.gameStatus === "finished") return state;
      const next = state.timer.remaining - 1;
      if (next <= 0) {
        const qId = state.questions[state.currentQuestionIndex]?.id;
        return {
          ...state, timer: { ...state.timer, remaining: 0, running: false, locked: true }, gameStatus: "reviewing", answeringTeamId: null,
          selectedAnswers: { ...state.selectedAnswers, [qId]: { teamId: null, result: "timeup" } },
          playedLog: [...state.playedLog, { question: state.questions[state.currentQuestionIndex], teamId: null, result: "timeup", round: state.currentRound }],
        };
      }
      return { ...state, timer: { ...state.timer, remaining: next }, gameStatus: next <= 5 ? "answering" : "playing" };
    }
    case "TIMER_PAUSE": if (state.timer.locked) return state; return { ...state, timer: { ...state.timer, running: false }, gameStatus: "paused" };
    case "TIMER_RESUME": if (state.timer.locked) return state; return { ...state, timer: { ...state.timer, running: true }, gameStatus: state.timer.remaining <= 5 ? "answering" : "playing" };
    case "TIMER_RESET": if (state.timer.locked) return state; return { ...state, timer: { total: state.timerSecondsConfig, remaining: state.timerSecondsConfig, running: true, locked: false }, gameStatus: "playing" };

    case "ANSWER_LOCK": {
      if (state.timer.locked) return state; // guard: cegah skor dihitung dua kali
      const { teamId, result, basePoints, difficulty, penaltyEnabled, penaltyPoints, selectedIndex } = action.payload;
      let teams = state.teams;
      if (teamId != null && result === "correct") teams = teams.map((t) => (t.id === teamId ? applyCorrectAnswer(t, basePoints, { difficulty }) : t));
      else if (teamId != null && result === "wrong") teams = teams.map((t) => (t.id === teamId ? applyWrongAnswer(t, { basePoints, difficulty, penaltyEnabled, penaltyPoints }) : t));
      let duel = state.duel;
      if (state.currentRound === "headtohead" && duel && teamId != null && result === "correct") duel = { ...duel, wins: { ...duel.wins, [teamId]: (duel.wins[teamId] || 0) + 1 } };
      const qId = state.questions[state.currentQuestionIndex]?.id;
      return {
        ...state, teams, scores: scoresFromTeams(teams), duel,
        selectedAnswers: { ...state.selectedAnswers, [qId]: { teamId, result, selectedIndex: selectedIndex ?? null } },
        playedLog: [...state.playedLog, { question: state.questions[state.currentQuestionIndex], teamId, result, round: state.currentRound, selectedIndex: selectedIndex ?? null }],
        timer: { ...state.timer, running: false, locked: true }, gameStatus: "reviewing", answeringTeamId: null,
      };
    }

    // Penyesuaian poin manual oleh guru setelah jawaban keluar — misal bonus jawaban kreatif,
    // koreksi salah input, atau pengurangan kedisiplinan. Tidak mengubah statistik benar/salah/streak,
    // murni menambah/mengurangi skor. Skor tidak pernah turun di bawah 0.
    case "ADJUST_SCORE": {
      const { teamId, delta } = action.payload;
      if (teamId == null || !delta) return state;
      const teams = state.teams.map((t) => (t.id === teamId ? { ...t, score: Math.max(0, t.score + delta) } : t));
      return { ...state, teams, scores: scoresFromTeams(teams) };
    }

    case "NEXT_QUESTION": {
      const nextIndex = state.currentQuestionIndex + 1;
      if (nextIndex < state.questions.length) {
        const isStrategy = state.currentRound === "strategy";
        return { ...state, currentQuestionIndex: nextIndex, timer: { total: state.timerSecondsConfig, remaining: state.timerSecondsConfig, running: !isStrategy, locked: false }, gameStatus: isStrategy ? "ready" : "playing", strategySelection: null, answeringTeamId: null };
      }
      if (state.gameMode !== "tournament") {
        const entry = buildHistoryEntry(state);
        return { ...state, gameStatus: "finished", activeScreen: "results", results: buildResultsSnapshot(state.teams, state.questions.length, state.roundStatus), history: [entry, ...state.history].slice(0, 50) };
      }
      const newRoundStatus = { ...state.roundStatus, [state.currentRound]: "done" };
      const orderIdx = ROUND_ORDER.indexOf(state.currentRound);
      const remaining = ROUND_ORDER.slice(orderIdx + 1).filter((r) => state.selectedRounds.includes(r));
      if (remaining.length === 0) {
        const entry = buildHistoryEntry(state, "tournament");
        return { ...state, roundStatus: newRoundStatus, gameStatus: "finished", activeScreen: "results", currentRound: null, results: buildResultsSnapshot(state.teams, state.playedLog.length, newRoundStatus), history: [entry, ...state.history].slice(0, 50) };
      }
      const nextRound = remaining[0];
      return { ...state, roundStatus: newRoundStatus, currentRound: nextRound, activeScreen: screenForRound(nextRound, state.finalists), gameStatus: "ready", duel: null, strategySelection: null, answeringTeamId: null };
    }

    case "RESTART_TO_SETUP": {
      const mode = action.payload;
      const screenByMode = { quickplay: "quickplay-setup", practice: "practice-setup", tournament: "tournament-setup", board: "board-setup" };
      return {
        ...state, gameMode: mode, activeScreen: screenByMode[mode] || "quickplay-setup",
        teams: state.teams.map((t) => ({ ...t, score: 0, correct: 0, wrong: 0, streak: 0, bestStreak: 0 })),
        questions: [], currentQuestionIndex: 0, currentRound: null, roundStatus: {}, finalists: [],
        duel: null, strategySelection: null, answeringTeamId: null, playedLog: [], usedQuestionIds: [], selectedAnswers: {},
        topicFilter: null, difficultyFilter: null,
        boardQuestions: [], boardTiles: [], boardTurnOrder: [], boardTurnIndex: 0, boardActiveTileId: null,
        gameStatus: "setup", results: null, timer: { total: 30, remaining: 30, running: false, locked: false },
      };
    }

    case "RESET_GAME": return { ...initialGameState, questionBank: state.questionBank, classes: state.classes, activeClassId: state.activeClassId, teams: state.teams.map((t) => ({ ...t, score: 0, correct: 0, wrong: 0, streak: 0, bestStreak: 0 })), soundOn: state.soundOn, settings: state.settings, history: state.history, storageReady: state.storageReady };

    default: return state;
  }
}

/* ============================================================
   STYLE GLOBAL
   ============================================================ */
const GlobalStyle = () => (
  <style>{`
    .kc-root { --navy-deep:#060E24; --navy:#0B1C3D; --navy-card:#10254D; --navy-card-hi:#17305F; --blue:#2F6FED; --blue-soft:#4C86FF; --gold:#F2B705; --gold-soft:#FCD34D; --white:#F7F9FC; --slate:#93A5C9; --green:#22C55E; --red:#F0435B;
      font-family:'Sora','Inter',system-ui,sans-serif; background: radial-gradient(circle at 15% 0%, #10265A 0%, var(--navy) 42%, var(--navy-deep) 100%); color:var(--white); min-height:100vh; }
    .kc-root * { box-sizing:border-box; }
    .kc-font-body { font-family:'Inter',system-ui,sans-serif; }
    .kc-display { font-family:'Sora',system-ui,sans-serif; letter-spacing:-0.01em; }
    .kc-card { background: linear-gradient(160deg, var(--navy-card) 0%, var(--navy-card-hi) 100%); border:1px solid rgba(255,255,255,0.08); border-radius:18px; }
    .kc-card-cut { clip-path: polygon(0 12px, 12px 0, 100% 0, 100% calc(100% - 12px), calc(100% - 12px) 100%, 0 100%); }
    .kc-btn-primary { background: linear-gradient(135deg, var(--blue) 0%, #1A4FC4 100%); color:var(--white); border:none; box-shadow:0 8px 24px -8px rgba(47,111,237,0.7); }
    .kc-btn-primary:hover { filter:brightness(1.1); }
    .kc-btn-gold { background: linear-gradient(135deg, var(--gold) 0%, #C98F03 100%); color:#221800; border:none; box-shadow:0 8px 24px -8px rgba(242,183,5,0.6); }
    .kc-btn-gold:hover { filter:brightness(1.08); }
    .kc-btn-danger { background: linear-gradient(135deg, #F0435B 0%, #B4213A 100%); color:var(--white); border:none; }
    .kc-btn-ghost { background: rgba(255,255,255,0.05); border:1px solid rgba(255,255,255,0.14); color:var(--white); }
    .kc-btn-ghost:hover { background: rgba(255,255,255,0.1); }
    .kc-chip { border:1px solid rgba(255,255,255,0.16); background: rgba(255,255,255,0.04); }
    .kc-chip-active { border-color:var(--blue); background: rgba(47,111,237,0.18); box-shadow:0 0 0 2px rgba(47,111,237,0.35) inset; }
    .kc-navbar { background: rgba(6,14,36,0.85); backdrop-filter: blur(10px); border-bottom:1px solid rgba(255,255,255,0.08); }
    .kc-timer-ring { transition: stroke-dashoffset 1s linear, stroke 0.3s ease; }
    .kc-pulse { animation: kcPulse 1.4s ease-in-out infinite; }
    @keyframes kcPulse { 0%,100%{transform:scale(1);} 50%{transform:scale(1.04);} }
    .kc-pop { animation: kcPop 0.35s cubic-bezier(.2,1.4,.4,1); }
    @keyframes kcPop { 0%{transform:scale(0.85);opacity:0;} 100%{transform:scale(1);opacity:1;} }
    .kc-scroll::-webkit-scrollbar { height:6px; width:6px; }
    .kc-scroll::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.2); border-radius:4px; }
    .kc-urgent { border-color:var(--red) !important; box-shadow:0 0 0 2px rgba(240,67,91,0.35) inset !important; }
    @media (prefers-reduced-motion: reduce) { .kc-pulse, .kc-pop, .kc-pop-bounce { animation: none !important; } }
    .kc-popup-layer { position:fixed; inset:0; z-index:70; display:flex; align-items:flex-start; justify-content:center; padding-top:14vh; pointer-events:none; }
    .kc-popup-card { pointer-events:auto; cursor:pointer; background:linear-gradient(160deg, var(--navy-card) 0%, var(--navy-card-hi) 100%); border:3px solid var(--gold); border-radius:24px; padding:28px 40px; box-shadow:0 20px 60px -12px rgba(0,0,0,0.6); display:flex; flex-direction:column; align-items:center; max-width:90vw; }
    .kc-popup-emoji { font-size:64px; line-height:1; margin-bottom:8px; }
    @keyframes kcPopBounce {
      0% { transform: scale(0.3) translateY(-40px); opacity: 0; }
      55% { transform: scale(1.12) translateY(4px); opacity: 1; }
      75% { transform: scale(0.96) translateY(-2px); }
      100% { transform: scale(1) translateY(0); opacity: 1; }
    }
    .kc-pop-bounce { animation: kcPopBounce 0.55s cubic-bezier(.34,1.56,.64,1); }
    button:focus-visible, a:focus-visible, input:focus-visible, select:focus-visible { outline: 2px solid var(--gold); outline-offset: 2px; }
    .kc-input { background: rgba(255,255,255,0.05); border:1px solid rgba(255,255,255,0.15); border-radius:10px; color:var(--white); padding:9px 12px; font-size:13px; width:100%; outline:none; }
    .kc-input:focus { border-color: var(--blue); }
    .kc-input::placeholder { color: rgba(147,165,201,0.6); }
    .kc-modal-backdrop { position:fixed; inset:0; background:rgba(3,8,22,0.7); backdrop-filter:blur(4px); z-index:50; display:flex; align-items:flex-start; justify-content:center; overflow-y:auto; padding:24px 16px; }
    .kc-option-btn { cursor:pointer; text-align:left; width:100%; transition: transform .12s ease, filter .12s ease, box-shadow .12s ease; }
    .kc-option-btn:hover { filter:brightness(1.15); transform:translateY(-1px); }
    .kc-option-btn:active { transform:scale(0.98); }
    .kc-option-selected { border-color:var(--blue) !important; background:rgba(47,111,237,0.18) !important; box-shadow:0 0 0 2px rgba(47,111,237,0.4) inset; }
    .kc-spin { animation: kcSpin 1s linear infinite; }
    @keyframes kcSpin { to { transform: rotate(360deg); } }
  `}</style>
);

/* ============================================================
   KOMPONEN UMUM
   ============================================================ */
function NavBar({ state, dispatch, setScreen }) {
  return (
    <div className="kc-navbar sticky top-0 z-30 w-full">
      <div className="max-w-6xl mx-auto flex items-center justify-between gap-3 px-4 py-3">
        <button onClick={() => setScreen("home")} className="flex items-center gap-2 shrink-0">
          <div className="w-9 h-9 rounded-lg flex items-center justify-center kc-btn-gold"><Trophy size={18} /></div>
          <div className="text-left leading-tight hidden sm:block">
            <div className="kc-display font-bold text-sm text-white">KELAS CHAMPIONS</div>
            <div className="text-[10px] tracking-wide" style={{ color: "var(--slate)" }}>SMP · Academic Arena</div>
          </div>
        </button>
        <div className="flex items-center gap-1 overflow-x-auto kc-scroll">
          {NAV_ITEMS.map((item) => {
            const Icon = item.icon;
            const active =
              state.activeScreen === item.key ||
              (item.mode && state.gameMode === item.mode && ["round-transition", "finalist-select", "arena", "duel-arena", "strategy-arena", "board-arena", "results"].includes(state.activeScreen)) ||
              (item.teacherTab && state.activeScreen === "teacher-panel" && state.teacherTab === item.teacherTab);
            const onClick = () => {
              if (item.key === "quickplay-setup") dispatch({ type: "RESTART_TO_SETUP", payload: "quickplay" });
              else if (item.key === "tournament-setup") dispatch({ type: "RESTART_TO_SETUP", payload: "tournament" });
              else if (item.key === "practice-setup") dispatch({ type: "RESTART_TO_SETUP", payload: "practice" });
              else if (item.key === "board-setup") dispatch({ type: "RESTART_TO_SETUP", payload: "board" });
              else if (item.teacherTab) { dispatch({ type: "SET_TEACHER_TAB", payload: item.teacherTab }); setScreen("teacher-panel"); }
              else setScreen(item.key);
            };
            return (
              <button key={item.key} onClick={onClick} className={`flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-medium whitespace-nowrap transition-colors kc-font-body ${active ? "kc-chip-active" : "kc-chip"}`} title={item.ready ? item.label : `${item.label} — ${item.phase}`}>
                <Icon size={14} /><span>{item.label}</span>
                {!item.ready && <span className="w-1.5 h-1.5 rounded-full" style={{ background: "var(--gold)" }} />}
              </button>
            );
          })}
        </div>
        <button onClick={() => dispatch({ type: "TOGGLE_SOUND" })} className="w-9 h-9 rounded-lg flex items-center justify-center kc-btn-ghost shrink-0" title={state.soundOn ? "Matikan suara" : "Nyalakan suara"}>
          {state.soundOn ? <Volume2 size={16} /> : <VolumeX size={16} />}
        </button>
      </div>
    </div>
  );
}

function Placeholder({ title, phase, description, icon: Icon, setScreen }) {
  return (
    <div className="max-w-2xl mx-auto px-4 py-20 text-center">
      <div className="w-20 h-20 rounded-2xl kc-card flex items-center justify-center mx-auto mb-6"><Icon size={32} style={{ color: "var(--gold)" }} /></div>
      <h2 className="kc-display text-2xl font-bold mb-2">{title}</h2>
      <div className="inline-block px-3 py-1 rounded-full text-xs font-semibold mb-4" style={{ background: "rgba(242,183,5,0.15)", color: "var(--gold)" }}>Hadir di {phase}</div>
      <p className="kc-font-body text-sm mb-8" style={{ color: "var(--slate)" }}>{description}</p>
      <button onClick={() => setScreen("home")} className="kc-btn-primary px-6 py-3 rounded-xl font-semibold text-sm kc-font-body">Kembali ke Beranda</button>
    </div>
  );
}

function TimerRing({ timeLeft, total, running, size: sizeProp }) {
  const size = sizeProp || 168, stroke = 12, r = (size - stroke) / 2, circ = 2 * Math.PI * r;
  const pct = total > 0 ? Math.max(0, timeLeft / total) : 0;
  const offset = circ * (1 - pct);
  const color = timeLeft <= 5 ? "var(--red)" : timeLeft <= total * 0.4 ? "var(--gold)" : "var(--blue-soft)";
  return (
    <div className={`relative ${timeLeft <= 5 && running ? "kc-pulse" : ""}`} style={{ width: size, height: size }}>
      <svg width={size} height={size}>
        <circle cx={size / 2} cy={size / 2} r={r} stroke="rgba(255,255,255,0.1)" strokeWidth={stroke} fill="none" />
        <circle cx={size / 2} cy={size / 2} r={r} stroke={color} strokeWidth={stroke} fill="none" strokeDasharray={circ} strokeDashoffset={offset} strokeLinecap="round" className="kc-timer-ring" transform={`rotate(-90 ${size / 2} ${size / 2})`} />
      </svg>
      <div className="absolute inset-0 flex flex-col items-center justify-center">
        <div className="kc-display font-extrabold" style={{ color, fontSize: size * 0.34, lineHeight: 1 }}>{timeLeft}</div>
        <div className="text-xs font-bold tracking-widest mt-1" style={{ color: "var(--slate)" }}>DETIK</div>
      </div>
    </div>
  );
}

function RoundStatusBadge({ status }) {
  const map = { pending: { label: "Menunggu", color: "var(--slate)" }, active: { label: "Sedang Berlangsung", color: "var(--gold)" }, done: { label: "Selesai", color: "var(--green)" }, skipped: { label: "Dilewati", color: "var(--slate)" } };
  const s = map[status] || map.pending;
  return <span className="px-2.5 py-1 rounded-full text-[11px] font-bold kc-font-body" style={{ background: s.color + "22", color: s.color }}>{s.label}</span>;
}

/* Papan Skor ringkas — dipakai di sidebar Arena/Duel/Strategy */
function ScoreboardList({ teams, eliminatedIds = [] }) {
  const ranked = updateRanking(teams);
  return (
    <div className="space-y-2">
      {ranked.map((g, i) => {
        const Icon = GROUP_ICONS[g.iconIdx];
        const eliminated = eliminatedIds.includes(g.id);
        return (
          <div key={g.id} className={`flex items-center gap-2 ${eliminated ? "opacity-40" : ""}`}>
            <span className="w-5 text-sm font-extrabold text-center" style={{ color: "var(--slate)" }}>{i + 1}</span>
            <div className="w-7 h-7 rounded-md flex items-center justify-center shrink-0" style={{ background: g.color + "33" }}><Icon size={14} style={{ color: g.color }} /></div>
            <span className="text-sm font-semibold kc-font-body flex-1 truncate">{g.name}</span>
            <span className="text-[10px] font-bold" style={{ color: "var(--green)" }}>{g.correct}B</span>
            <span className="text-[10px] font-bold" style={{ color: "var(--red)" }}>{g.wrong}S</span>
            {eliminated && <span className="text-[9px]" style={{ color: "var(--slate)" }}>Tidak lolos</span>}
            <span className="kc-display text-base font-extrabold w-14 text-right">{g.score}</span>
          </div>
        );
      })}
    </div>
  );
}

/* Papan Skor layar penuh — untuk dipajang ke proyektor/IFP di antara sesi */
function ScoreboardOverlay({ teams, eliminatedIds = [], onClose }) {
  const ranked = updateRanking(teams);
  return (
    <div className="fixed inset-0 z-[60] kc-modal-backdrop flex items-center justify-center p-4 md:p-10" onClick={onClose}>
      <div className="kc-card rounded-3xl w-full max-w-4xl max-h-[90vh] overflow-y-auto kc-scroll p-6 md:p-10" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between mb-6 md:mb-8">
          <div className="flex items-center gap-3"><Trophy size={28} style={{ color: "var(--gold)" }} /><h2 className="kc-display font-extrabold text-2xl md:text-4xl">Papan Skor</h2></div>
          <button onClick={onClose} className="kc-btn-ghost w-11 h-11 rounded-xl flex items-center justify-center shrink-0"><X size={20} /></button>
        </div>
        <div className="space-y-3 md:space-y-4">
          {ranked.map((g, i) => {
            const Icon = GROUP_ICONS[g.iconIdx];
            const eliminated = eliminatedIds.includes(g.id);
            const answered = g.correct + g.wrong;
            return (
              <div key={g.id} className={`flex items-center gap-4 md:gap-6 rounded-2xl px-5 py-4 md:px-7 md:py-5 kc-chip ${eliminated ? "opacity-40" : ""}`}>
                <span className="kc-display font-extrabold text-2xl md:text-3xl w-10 text-center" style={{ color: i === 0 ? "var(--gold)" : "var(--slate)" }}>{i + 1}</span>
                <div className="w-14 h-14 md:w-16 md:h-16 rounded-2xl flex items-center justify-center shrink-0" style={{ background: g.color + "33" }}><Icon size={28} style={{ color: g.color }} /></div>
                <div className="flex-1 min-w-0">
                  <div className="kc-display font-bold text-xl md:text-2xl truncate">{g.name}</div>
                  <div className="text-xs md:text-sm kc-font-body mt-0.5" style={{ color: "var(--slate)" }}>
                    <span style={{ color: "var(--green)" }}>{g.correct} benar</span> · <span style={{ color: "var(--red)" }}>{g.wrong} salah</span> · {calcPercentage(g.correct, answered)}%
                    {eliminated && <span className="ml-2" style={{ color: "var(--slate)" }}>· Tidak lolos</span>}
                  </div>
                </div>
                <div className="kc-display font-extrabold text-3xl md:text-5xl" style={{ color: i === 0 ? "var(--gold)" : "var(--white)" }}>{g.score}</div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

// Penyesuaian poin manual — ditampilkan setelah jawaban keluar (reviewing) di semua arena
// (Quick Play/Practice, Head to Head, Strategy Arena, Papan Kuis) agar guru bisa memberi bonus
// (mis. jawaban kreatif di luar kunci) atau mengurangi poin (mis. kerja sama tidak sportif)
// tanpa mengubah statistik benar/salah/streak kelompok.
const SCORE_ADJUST_DELTAS = [-50, -10, 10, 50];
function ManualScoreAdjust({ teams, dispatch }) {
  return (
    <div className="kc-card rounded-xl p-4 mt-4">
      <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>SESUAIKAN POIN MANUAL</div>
      <div className="space-y-2">
        {teams.map((t) => {
          const Icon = GROUP_ICONS[t.iconIdx];
          return (
            <div key={t.id} className="flex items-center gap-2 flex-wrap kc-chip rounded-lg px-3 py-2">
              <div className="w-8 h-8 rounded-lg flex items-center justify-center shrink-0" style={{ background: t.color + "33" }}><Icon size={15} style={{ color: t.color }} /></div>
              <span className="text-xs md:text-sm font-semibold kc-font-body flex-1 min-w-[90px] truncate">{t.name}</span>
              <span className="kc-display font-extrabold text-sm md:text-base w-14 text-center">{t.score}</span>
              <div className="flex gap-1">
                {SCORE_ADJUST_DELTAS.map((d) => (
                  <button
                    key={d}
                    onClick={() => dispatch({ type: "ADJUST_SCORE", payload: { teamId: t.id, delta: d } })}
                    className="kc-btn-ghost px-2.5 py-1.5 rounded-md text-[11px] md:text-xs font-bold kc-font-body"
                    style={{ color: d > 0 ? "var(--green)" : "var(--red)" }}
                  >
                    {d > 0 ? `+${d}` : d}
                  </button>
                ))}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// Pemilih kelompok penjawab — muncul saat sebuah pilihan jawaban disentuh padahal belum ada kelompok yang
// ditunjuk sebagai penjawab. Dengan begini pilihan jawaban SELALU bisa disentuh (termasuk soal pertama),
// tanpa guru harus menekan "Tunjuk sebagai Penjawab" lebih dulu.
function AnswererPickerModal({ teams, optionLabel, onPick, onCancel }) {
  return (
    <div className="fixed inset-0 z-[60] kc-modal-backdrop flex items-center justify-center p-4" onClick={onCancel}>
      <div className="kc-card rounded-3xl w-full max-w-lg p-6 md:p-8 kc-pop" onClick={(e) => e.stopPropagation()}>
        <div className="text-center mb-5">
          <div className="text-xs font-bold tracking-wide mb-1" style={{ color: "var(--gold-soft)" }}>PILIHAN JAWABAN {optionLabel}</div>
          <h3 className="kc-display font-extrabold text-xl md:text-2xl">Kelompok mana yang menjawab?</h3>
        </div>
        <div className="grid grid-cols-2 gap-3 mb-4">
          {teams.map((t) => {
            const Icon = GROUP_ICONS[t.iconIdx];
            return (
              <button key={t.id} onClick={() => onPick(t.id)} className="kc-chip rounded-xl p-4 flex items-center gap-3 kc-font-body text-left">
                <div className="w-10 h-10 rounded-lg flex items-center justify-center shrink-0" style={{ background: t.color + "33" }}><Icon size={18} style={{ color: t.color }} /></div>
                <span className="text-sm md:text-base font-bold truncate">{t.name}</span>
              </button>
            );
          })}
        </div>
        <button onClick={onCancel} className="kc-btn-ghost w-full py-2.5 rounded-xl text-xs font-semibold kc-font-body">Batal</button>
      </div>
    </div>
  );
}
function answerOptionLabel(q, idx) {
  return q.type === "true_false" ? (idx === 0 ? "B" : "S") : String.fromCharCode(65 + idx);
}

function QuestionBody({ q, revealed, large = false, selectable = false, selectedIndex = null, onSelect = null }) {
  const isChoice = CHOICE_TYPES.includes(q.type);
  const optionTextCls = large ? "text-lg md:text-xl font-bold" : "text-sm font-semibold";
  const optionPadCls = large ? "px-5 py-5 md:px-6 md:py-6" : "px-4 py-4";
  const badgeSizeCls = large ? "w-10 h-10 md:w-11 md:h-11 text-base" : "w-7 h-7 text-xs";
  return (
    <>
      {q.imageUrl && (
        <div className="rounded-xl overflow-hidden mb-4 kc-chip">
          <img src={q.imageUrl} alt="Ilustrasi soal" className="w-full max-h-64 object-contain" onError={(e) => { e.currentTarget.style.display = "none"; }} />
        </div>
      )}
      {isChoice ? (
        <div className={`grid gap-3 md:gap-4 ${q.type === "true_false" ? "grid-cols-2" : "sm:grid-cols-2"}`}>
          {q.options.map((opt, idx) => {
            const showCorrect = revealed && idx === q.answer;
            const showWrongPick = revealed && selectedIndex === idx && idx !== q.answer;
            const isInteractive = selectable && !revealed;
            const Tag = isInteractive ? "button" : "div";
            let extraStyle = {};
            if (showCorrect) extraStyle = { background: "rgba(34,197,94,0.18)", border: "2px solid var(--green)" };
            else if (showWrongPick) extraStyle = { background: "rgba(240,67,91,0.16)", border: "2px solid var(--red)" };
            return (
              <Tag
                key={idx}
                type={isInteractive ? "button" : undefined}
                onClick={isInteractive ? () => onSelect(idx) : undefined}
                className={`rounded-xl md:rounded-2xl ${optionPadCls} flex items-center gap-3 md:gap-4 kc-font-body ${optionTextCls} ${showCorrect || showWrongPick ? "" : "kc-chip"} ${isInteractive ? "kc-option-btn" : ""}`}
                style={extraStyle}
              >
                <div className={`rounded-lg md:rounded-xl flex items-center justify-center font-extrabold shrink-0 ${badgeSizeCls}`} style={{ background: showCorrect ? "var(--green)" : showWrongPick ? "var(--red)" : "rgba(255,255,255,0.08)", color: showCorrect || showWrongPick ? "#fff" : "var(--slate)" }}>
                  {q.type === "true_false" ? (idx === 0 ? "B" : "S") : String.fromCharCode(65 + idx)}
                </div>
                <span className="leading-snug">{opt}</span>
                {showCorrect && <Check size={large ? 22 : 16} className="ml-auto shrink-0" style={{ color: "var(--green)" }} />}
                {showWrongPick && <X size={large ? 22 : 16} className="ml-auto shrink-0" style={{ color: "var(--red)" }} />}
              </Tag>
            );
          })}
        </div>
      ) : revealed ? (
        <div className={`rounded-xl px-5 py-5 kc-font-body ${large ? "text-lg font-bold" : "text-sm font-semibold"} kc-chip`} style={{ background: "rgba(34,197,94,0.18)", border: "2px solid var(--green)" }}>
          <div className="text-xs font-bold mb-1 tracking-wide" style={{ color: "var(--green)" }}>KUNCI JAWABAN</div>{q.answer}
        </div>
      ) : (
        <div className="rounded-xl px-4 py-4 kc-chip text-xs kc-font-body text-center" style={{ color: "var(--slate)" }}>
          Tipe: {QUESTION_TYPES.find((t) => t.key === q.type)?.label || q.type}. Guru membacakan/menampilkan tugas ke kelas, lalu menandai kelompok yang menjawab benar.
        </div>
      )}
    </>
  );
}

/* ============================================================
   POPUP UMPAN BALIK LUCU — muncul otomatis saat jawaban benar/salah
   ============================================================ */
const CORRECT_POPUP_MESSAGES = [
  { emoji: "🎉", text: "MANTAP JIWA!" },
  { emoji: "🚀", text: "MELESAT SEPERTI ROKET!" },
  { emoji: "🔥", text: "JAWABAN MEMBARA!" },
  { emoji: "⭐", text: "BINTANG KELAS!" },
  { emoji: "🏆", text: "JUARA SEJATI!" },
  { emoji: "✨", text: "OTAK CEMERLANG!" },
  { emoji: "🎯", text: "TEPAT SASARAN!" },
  { emoji: "🦾", text: "KUAT BANGET NIH TIM!" },
];
const WRONG_POPUP_MESSAGES = [
  { emoji: "😅", text: "YAH, MELESET DIKIT!" },
  { emoji: "🙈", text: "BELUM BERUNTUNG!" },
  { emoji: "🍋", text: "ASEM! KURANG TEPAT!" },
  { emoji: "🐢", text: "SANTAI, MASIH ADA SOAL LAIN!" },
  { emoji: "💥", text: "DUAR! MELESET!" },
  { emoji: "😬", text: "ADUH, HAMPIR SAJA!" },
  { emoji: "🎈", text: "KEMPES DIKIT POINNYA!" },
];
const TIMEUP_POPUP_MESSAGES = [
  { emoji: "⏰", text: "WAKTU HABIS!" },
  { emoji: "🐌", text: "KELAMAAN MIKIR!" },
  { emoji: "🥶", text: "BEKU KEHABISAN WAKTU!" },
];
function pickRandom(arr) { return arr[Math.floor(Math.random() * arr.length)]; }

function FeedbackPopup({ popup, onClose }) {
  const msg = useMemo(() => {
    if (!popup) return null;
    if (popup.kind === "correct") return pickRandom(CORRECT_POPUP_MESSAGES);
    if (popup.kind === "wrong") return pickRandom(WRONG_POPUP_MESSAGES);
    return pickRandom(TIMEUP_POPUP_MESSAGES);
  }, [popup?.nonce]);
  if (!popup || !msg) return null;
  const isCorrect = popup.kind === "correct";
  const isWrong = popup.kind === "wrong";
  const accent = isCorrect ? "var(--green)" : isWrong ? "var(--red)" : "var(--gold)";
  return (
    <div className="kc-popup-layer" onClick={onClose}>
      <div className="kc-popup-card kc-pop-bounce" style={{ borderColor: accent }}>
        <div className="kc-popup-emoji">{msg.emoji}</div>
        <div className="kc-display font-extrabold text-xl md:text-2xl text-center" style={{ color: accent }}>{msg.text}</div>
        {popup.teamName && <div className="text-sm md:text-base font-bold kc-font-body text-center mt-1" style={{ color: "var(--white)" }}>{popup.teamName}</div>}
        {popup.gain !== 0 && (
          <div className="kc-display font-extrabold text-2xl md:text-3xl text-center mt-1" style={{ color: accent }}>
            {popup.gain > 0 ? `+${popup.gain}` : popup.gain} poin
          </div>
        )}
      </div>
    </div>
  );
}

/* ============================================================
   BERANDA
   ============================================================ */
function HomeScreen({ state, dispatch, setScreen }) {
  const subjects = useMemo(() => getAvailableSubjects(state.questionBank), [state.questionBank]);
  const features = [
    { icon: Zap, title: "Quick Play", desc: "Mulai permainan dalam hitungan detik — pilih mapel, kelas, dan mainkan." },
    { icon: LayoutGrid, title: "Papan Kuis", desc: "Perwakilan kelompok maju bergiliran memilih kotak soal — tiap kotak soal & poinnya berbeda." },
    { icon: Swords, title: "Tournament Mode", desc: "5 babak kompetisi: Qualification, Head to Head, Strategy Arena, Semifinal, Grand Final." },
    { icon: GraduationCap, title: "Teacher Studio", desc: "Kelola bank soal, kelompok, pengaturan game, dan riwayat hasil dari satu panel." },
  ];
  const pausedClasses = state.classes.filter((c) => c.savedSession);
  return (
    <div className="max-w-6xl mx-auto px-4 pb-24">
      <div className="flex flex-wrap items-center justify-between gap-3 mt-6 md:mt-10 kc-card rounded-2xl px-4 py-3">
        <div className="flex items-center gap-2 flex-wrap">
          <GraduationCap size={16} style={{ color: "var(--gold)" }} />
          <span className="text-xs kc-font-body" style={{ color: "var(--slate)" }}>Kelas aktif:</span>
          <select className="kc-input py-1.5 text-xs w-auto" value={state.activeClassId || ""} onChange={(e) => dispatch({ type: "SET_ACTIVE_CLASS", payload: { id: e.target.value } })}>
            {state.classes.map((c) => (<option key={c.id} value={c.id}>{c.name}{c.savedSession ? " · ada kuis tersimpan" : ""}</option>))}
          </select>
        </div>
        <button onClick={() => { dispatch({ type: "SET_TEACHER_TAB", payload: "kelas" }); setScreen("teacher-panel"); }} className="kc-btn-ghost px-3 py-1.5 rounded-lg text-[11px] font-semibold kc-font-body flex items-center gap-1.5"><BarChart3 size={12} /> Kelola Kelas</button>
      </div>

      {pausedClasses.length > 0 && (
        <div className="mt-4 rounded-2xl p-4 flex items-center justify-between flex-wrap gap-3" style={{ background: "rgba(242,183,5,0.1)", border: "1px solid rgba(242,183,5,0.35)" }}>
          <div className="flex items-center gap-2">
            <Clock size={16} style={{ color: "var(--gold)" }} />
            <span className="text-xs md:text-sm font-semibold kc-font-body">
              {pausedClasses.length === 1 ? `Ada kuis belum selesai di kelas "${pausedClasses[0].name}".` : `Ada ${pausedClasses.length} kuis belum selesai di beberapa kelas.`}
            </span>
          </div>
          <div className="flex flex-wrap gap-2">
            {pausedClasses.map((c) => (
              <button key={c.id} onClick={() => dispatch({ type: "RESUME_SESSION", payload: { classId: c.id } })} className="kc-btn-gold px-4 py-2 rounded-lg text-xs font-bold kc-font-body flex items-center gap-1.5"><Play size={13} /> Lanjutkan {c.name}</button>
            ))}
          </div>
        </div>
      )}

      <div className="relative kc-card kc-card-cut mt-6 md:mt-10 px-6 py-10 md:px-14 md:py-16 overflow-hidden">
        <div className="absolute -top-10 -right-10 w-64 h-64 rounded-full opacity-20 blur-3xl" style={{ background: "var(--blue)" }} />
        <div className="absolute -bottom-16 -left-10 w-72 h-72 rounded-full opacity-10 blur-3xl" style={{ background: "var(--gold)" }} />
        <div className="relative flex flex-col md:flex-row items-start md:items-center gap-8">
          <div className="flex-1">
            <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full kc-chip text-xs font-semibold mb-5" style={{ color: "var(--gold-soft)" }}><Sparkles size={13} /> SMP YPKKP BANDUNG · Kurikulum Merdeka</div>
            <h1 className="kc-display font-extrabold text-4xl md:text-6xl leading-[1.05] mb-4">KELAS<br /><span style={{ color: "var(--gold)" }}>CHAMPIONS</span> SMP</h1>
            <p className="kc-font-body text-base md:text-lg mb-1" style={{ color: "var(--slate)" }}>The Ultimate Academic Arena</p>
            <p className="kc-display font-semibold text-lg md:text-xl mb-7">Berpikir Cerdas. Bergerak Cepat. Menjadi Juara.</p>
            <div className="flex flex-wrap gap-3">
              <button onClick={() => dispatch({ type: "RESTART_TO_SETUP", payload: "quickplay" })} className="kc-btn-gold px-6 py-3.5 rounded-xl font-bold text-sm kc-font-body flex items-center gap-2"><Play size={16} /> Mulai Quick Play</button>
              <button onClick={() => dispatch({ type: "RESTART_TO_SETUP", payload: "board" })} className="kc-btn-primary px-6 py-3.5 rounded-xl font-bold text-sm kc-font-body flex items-center gap-2"><LayoutGrid size={16} /> Papan Kuis</button>
              <button onClick={() => dispatch({ type: "RESTART_TO_SETUP", payload: "tournament" })} className="kc-btn-primary px-6 py-3.5 rounded-xl font-bold text-sm kc-font-body flex items-center gap-2"><Swords size={16} /> Mulai Tournament Mode</button>
              <button onClick={() => dispatch({ type: "RESTART_TO_SETUP", payload: "practice" })} className="kc-btn-ghost px-6 py-3.5 rounded-xl font-semibold text-sm kc-font-body flex items-center gap-2"><Target size={16} /> Mode Latihan</button>
              <button onClick={() => { dispatch({ type: "SET_TEACHER_TAB", payload: "dashboard" }); setScreen("teacher-panel"); }} className="kc-btn-ghost px-6 py-3.5 rounded-xl font-semibold text-sm kc-font-body flex items-center gap-2"><GraduationCap size={16} /> Teacher Studio</button>
            </div>
          </div>
          <div className="w-full md:w-64 shrink-0">
            <div className="kc-card rounded-2xl p-5 text-center kc-pop">
              <Trophy size={40} className="mx-auto mb-2" style={{ color: "var(--gold)" }} />
              <div className="text-xs font-semibold tracking-wide" style={{ color: "var(--slate)" }}>STATUS PENGEMBANGAN</div>
              <div className="kc-display font-bold text-lg mt-1">Fase 2 Lanjutan</div>
              <div className="text-xs mt-1" style={{ color: "var(--slate)" }}>Teacher Studio & Bank Soal Dinamis aktif</div>
            </div>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mt-8">
        {features.map((f, i) => { const Icon = f.icon; return (
          <div key={i} className="kc-card rounded-2xl p-5">
            <div className="w-10 h-10 rounded-xl flex items-center justify-center mb-3" style={{ background: "rgba(47,111,237,0.18)" }}><Icon size={18} style={{ color: "var(--blue-soft)" }} /></div>
            <div className="kc-display font-semibold text-sm mb-1">{f.title}</div>
            <div className="text-xs kc-font-body" style={{ color: "var(--slate)" }}>{f.desc}</div>
          </div>
        ); })}
      </div>

      <div className="mt-10">
        <h3 className="kc-display font-bold text-lg mb-4">Mata Pelajaran Tersedia</h3>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          {subjects.map((s) => (
            <div key={s.key} className="kc-card rounded-xl p-4 flex items-center gap-3">
              <div className="text-2xl">{s.icon}</div>
              <div><div className="text-sm font-semibold kc-font-body">{s.label}</div><div className="text-[11px]" style={{ color: "var(--slate)" }}>{state.questionBank.filter((q) => q.subject === s.key).length} soal siap main</div></div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   QUICK PLAY SETUP
   ============================================================ */
function QuickPlaySetup({ state, dispatch, setScreen, practiceMode = false }) {
  const subjects = useMemo(() => getAvailableSubjects(state.questionBank), [state.questionBank]);
  const [subject, setSubject] = useState(null);
  const [grade, setGrade] = useState(null);
  const [topic, setTopic] = useState("");
  const [difficulty, setDifficulty] = useState("all");
  const [questionCount, setQuestionCount] = useState(null);
  const [timerSeconds, setTimerSeconds] = useState(state.settings.defaultTimerSeconds);

  const availableGrades = useMemo(() => (subject ? [...new Set(state.questionBank.filter((q) => q.subject === subject).map((q) => q.grade))].sort() : []), [subject, state.questionBank]);
  const gradeScopedBank = useMemo(() => (subject && grade ? state.questionBank.filter((q) => q.subject === subject && q.grade === grade) : []), [subject, grade, state.questionBank]);
  const availableTopics = useMemo(() => [...new Set(gradeScopedBank.map((q) => q.topic))].sort(), [gradeScopedBank]);
  const scopedByTopic = useMemo(() => (topic ? gradeScopedBank.filter((q) => q.topic === topic) : gradeScopedBank), [gradeScopedBank, topic]);
  const availableCount = useMemo(() => (difficulty !== "all" ? scopedByTopic.filter((q) => q.difficulty === difficulty).length : scopedByTopic.length), [scopedByTopic, difficulty]);
  const questionCountOptions = useMemo(() => { const opts = [5, 10].filter((n) => n <= availableCount); return opts.length ? opts : (availableCount > 0 ? [availableCount] : []); }, [availableCount]);
  const canStart = subject && grade && questionCount && timerSeconds && state.teams.every((g) => g.name.trim().length > 0);

  const title = practiceMode ? "Practice Mode" : "Quick Play";
  const subtitle = practiceMode
    ? "Mode latihan tanpa tekanan skor — semua kelompok tetap bermain dan pembahasan selalu ditampilkan. Cocok untuk asesmen formatif."
    : "Atur permainan dalam beberapa langkah singkat.";

  return (
    <div className="max-w-4xl mx-auto px-4 pb-24 pt-8">
      <button onClick={() => setScreen("home")} className="flex items-center gap-1 text-sm mb-4 kc-font-body" style={{ color: "var(--slate)" }}><ChevronLeft size={16} /> Kembali</button>
      <h1 className="kc-display font-extrabold text-3xl mb-1">{title}</h1>
      <p className="text-sm mb-8 kc-font-body" style={{ color: "var(--slate)" }}>{subtitle}</p>

      <div className="kc-card rounded-2xl p-5 mb-4">
        <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>1 · PILIH MATA PELAJARAN</div>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          {subjects.map((s) => (
            <button key={s.key} onClick={() => { setSubject(s.key); setGrade(null); setTopic(""); setDifficulty("all"); setQuestionCount(null); }} className={`rounded-xl p-4 text-center kc-font-body ${subject === s.key ? "kc-chip-active" : "kc-chip"}`}>
              <div className="text-2xl mb-1">{s.icon}</div><div className="text-sm font-semibold">{s.label}</div>
            </button>
          ))}
        </div>
        {subjects.length === 0 && <p className="text-xs" style={{ color: "var(--slate)" }}>Belum ada soal di Bank Soal. Tambahkan lewat Teacher Studio.</p>}
      </div>

      {subject && (
        <div className="kc-card rounded-2xl p-5 mb-4 kc-pop">
          <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>2 · PILIH KELAS</div>
          <div className="flex gap-3 flex-wrap">{availableGrades.map((g) => (<button key={g} onClick={() => { setGrade(g); setTopic(""); setDifficulty("all"); setQuestionCount(null); }} className={`px-6 py-3 rounded-xl font-semibold text-sm kc-font-body ${grade === g ? "kc-chip-active" : "kc-chip"}`}>Kelas {g}</button>))}</div>
        </div>
      )}

      {grade && (
        <div className="kc-card rounded-2xl p-5 mb-4 kc-pop">
          <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>3 · PILIH MATERI</div>
          <div className="flex gap-2 flex-wrap">
            <button onClick={() => { setTopic(""); setQuestionCount(null); }} className={`px-5 py-2.5 rounded-xl font-semibold text-xs kc-font-body ${topic === "" ? "kc-chip-active" : "kc-chip"}`}>Semua Materi</button>
            {availableTopics.map((t) => (<button key={t} onClick={() => { setTopic(t); setQuestionCount(null); }} className={`px-5 py-2.5 rounded-xl font-semibold text-xs kc-font-body ${topic === t ? "kc-chip-active" : "kc-chip"}`}>{t}</button>))}
          </div>
        </div>
      )}

      {grade && (
        <div className="kc-card rounded-2xl p-5 mb-4 kc-pop">
          <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>4 · TINGKAT KESULITAN</div>
          <div className="flex gap-2 flex-wrap">
            {[{ key: "all", label: "Semua Tingkat" }, { key: "easy", label: "Mudah" }, { key: "medium", label: "Sedang" }, { key: "hard", label: "Sulit" }].map((d) => (
              <button key={d.key} onClick={() => { setDifficulty(d.key); setQuestionCount(null); }} className={`px-5 py-2.5 rounded-xl font-semibold text-xs kc-font-body ${difficulty === d.key ? "kc-chip-active" : "kc-chip"}`}>{d.label}</button>
            ))}
          </div>
        </div>
      )}

      {grade && (
        <div className="kc-card rounded-2xl p-5 mb-4 kc-pop">
          <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>5 · JUMLAH SOAL</div>
          {questionCountOptions.length > 0 ? (
            <div className="flex gap-3 flex-wrap">{questionCountOptions.map((n) => (<button key={n} onClick={() => setQuestionCount(n)} className={`px-6 py-3 rounded-xl font-semibold text-sm kc-font-body ${questionCount === n ? "kc-chip-active" : "kc-chip"}`}>{n} Soal</button>))}</div>
          ) : (
            <p className="text-xs" style={{ color: "var(--red)" }}>Tidak ada soal untuk kombinasi materi/kesulitan ini. Coba pilih "Semua Materi" atau "Semua Tingkat".</p>
          )}
        </div>
      )}

      <div className="kc-card rounded-2xl p-5 mb-4">
        <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>6 · WAKTU PER SOAL</div>
        <div className="flex gap-3 flex-wrap">{TIMER_OPTIONS.map((t) => (<button key={t} onClick={() => setTimerSeconds(t)} className={`px-6 py-3 rounded-xl font-semibold text-sm kc-font-body flex items-center gap-2 ${timerSeconds === t ? "kc-chip-active" : "kc-chip"}`}><Clock size={14} /> {t} detik</button>))}</div>
      </div>

      <div className="kc-card rounded-2xl p-5 mb-6">
        <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>7 · NAMA KELOMPOK (4 Kelompok)</div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          {state.teams.map((g) => { const Icon = GROUP_ICONS[g.iconIdx]; return (
            <div key={g.id} className="flex items-center gap-3 rounded-xl px-3 py-2.5 kc-chip">
              <div className="w-9 h-9 rounded-lg flex items-center justify-center shrink-0" style={{ background: g.color + "33" }}><Icon size={16} style={{ color: g.color }} /></div>
              <input value={g.name} maxLength={16} onChange={(e) => dispatch({ type: "UPDATE_TEAM_NAME", payload: { id: g.id, name: e.target.value } })} className="bg-transparent outline-none text-sm font-semibold kc-font-body w-full text-white" placeholder="Nama kelompok" />
            </div>
          ); })}
        </div>
      </div>

      <button disabled={!canStart} onClick={() => dispatch({ type: "START_QUICKPLAY", payload: { subject, grade, questionCount, timerSeconds, topic: topic || null, difficulty, mode: practiceMode ? "practice" : "quickplay" } })} className={`w-full py-4 rounded-xl font-bold text-base kc-font-body flex items-center justify-center gap-2 ${canStart ? "kc-btn-gold" : "kc-btn-ghost opacity-40 cursor-not-allowed"}`}>
        <Play size={18} /> Mulai Permainan!
      </button>
    </div>
  );
}

/* ============================================================
   PAPAN KUIS SETUP (Board Quiz)
   ============================================================ */
function BoardSetup({ state, dispatch, setScreen }) {
  const subjects = useMemo(() => getAvailableSubjects(state.questionBank), [state.questionBank]);
  const [subject, setSubject] = useState(null);
  const [grade, setGrade] = useState(null);
  const [topic, setTopic] = useState("");
  const [difficulty, setDifficulty] = useState("all");
  const [boardSize, setBoardSize] = useState(null);
  const [timerSeconds, setTimerSeconds] = useState(state.settings.defaultTimerSeconds);

  const availableGrades = useMemo(() => (subject ? [...new Set(state.questionBank.filter((q) => q.subject === subject).map((q) => q.grade))].sort() : []), [subject, state.questionBank]);
  const gradeScopedBank = useMemo(() => (subject && grade ? state.questionBank.filter((q) => q.subject === subject && q.grade === grade) : []), [subject, grade, state.questionBank]);
  const availableTopics = useMemo(() => [...new Set(gradeScopedBank.map((q) => q.topic))].sort(), [gradeScopedBank]);
  const scopedByTopic = useMemo(() => (topic ? gradeScopedBank.filter((q) => q.topic === topic) : gradeScopedBank), [gradeScopedBank, topic]);
  const availableCount = useMemo(() => (difficulty !== "all" ? scopedByTopic.filter((q) => q.difficulty === difficulty).length : scopedByTopic.length), [scopedByTopic, difficulty]);
  const boardSizePresetOptions = useMemo(() => BOARD_SIZE_PRESETS.filter((n) => n <= availableCount), [availableCount]);
  const canStart = subject && grade && boardSize && boardSize >= BOARD_MIN_SIZE && boardSize <= availableCount && timerSeconds && state.teams.every((g) => g.name.trim().length > 0);
  const setBoardSizeClamped = (n) => {
    if (isNaN(n)) { setBoardSize(null); return; }
    setBoardSize(Math.max(BOARD_MIN_SIZE, Math.min(availableCount, n)));
  };

  return (
    <div className="max-w-4xl mx-auto px-4 pb-24 pt-8">
      <button onClick={() => setScreen("home")} className="flex items-center gap-1 text-sm mb-4 kc-font-body" style={{ color: "var(--slate)" }}><ChevronLeft size={16} /> Kembali</button>
      <h1 className="kc-display font-extrabold text-3xl mb-1">Papan Kuis</h1>
      <p className="text-sm mb-8 kc-font-body" style={{ color: "var(--slate)" }}>Setiap perwakilan kelompok maju bergiliran memilih kotak soal di papan — setiap kotak berisi soal berbeda dengan poin berbeda.</p>

      <div className="kc-card rounded-2xl p-5 mb-4">
        <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>1 · PILIH MATA PELAJARAN</div>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          {subjects.map((s) => (
            <button key={s.key} onClick={() => { setSubject(s.key); setGrade(null); setTopic(""); setDifficulty("all"); setBoardSize(null); }} className={`rounded-xl p-4 text-center kc-font-body ${subject === s.key ? "kc-chip-active" : "kc-chip"}`}>
              <div className="text-2xl mb-1">{s.icon}</div><div className="text-sm font-semibold">{s.label}</div>
            </button>
          ))}
        </div>
        {subjects.length === 0 && <p className="text-xs" style={{ color: "var(--slate)" }}>Belum ada soal di Bank Soal. Tambahkan lewat Teacher Studio.</p>}
      </div>

      {subject && (
        <div className="kc-card rounded-2xl p-5 mb-4 kc-pop">
          <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>2 · PILIH KELAS</div>
          <div className="flex gap-3 flex-wrap">{availableGrades.map((g) => (<button key={g} onClick={() => { setGrade(g); setTopic(""); setDifficulty("all"); setBoardSize(null); }} className={`px-6 py-3 rounded-xl font-semibold text-sm kc-font-body ${grade === g ? "kc-chip-active" : "kc-chip"}`}>Kelas {g}</button>))}</div>
        </div>
      )}

      {grade && (
        <div className="kc-card rounded-2xl p-5 mb-4 kc-pop">
          <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>3 · PILIH MATERI</div>
          <div className="flex gap-2 flex-wrap">
            <button onClick={() => { setTopic(""); setBoardSize(null); }} className={`px-5 py-2.5 rounded-xl font-semibold text-xs kc-font-body ${topic === "" ? "kc-chip-active" : "kc-chip"}`}>Semua Materi</button>
            {availableTopics.map((t) => (<button key={t} onClick={() => { setTopic(t); setBoardSize(null); }} className={`px-5 py-2.5 rounded-xl font-semibold text-xs kc-font-body ${topic === t ? "kc-chip-active" : "kc-chip"}`}>{t}</button>))}
          </div>
        </div>
      )}

      {grade && (
        <div className="kc-card rounded-2xl p-5 mb-4 kc-pop">
          <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>4 · TINGKAT KESULITAN</div>
          <div className="flex gap-2 flex-wrap">
            {[{ key: "all", label: "Semua Tingkat" }, { key: "easy", label: "Mudah" }, { key: "medium", label: "Sedang" }, { key: "hard", label: "Sulit" }].map((d) => (
              <button key={d.key} onClick={() => { setDifficulty(d.key); setBoardSize(null); }} className={`px-5 py-2.5 rounded-xl font-semibold text-xs kc-font-body ${difficulty === d.key ? "kc-chip-active" : "kc-chip"}`}>{d.label}</button>
            ))}
          </div>
        </div>
      )}

      {grade && (
        <div className="kc-card rounded-2xl p-5 mb-4 kc-pop">
          <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>5 · JUMLAH KOTAK SOAL DI PAPAN</div>
          <p className="text-xs kc-font-body mb-3" style={{ color: "var(--slate)" }}>Guru menentukan sendiri ukuran papan — bebas berapa saja, mengikuti jumlah soal yang tersedia untuk kombinasi ini.</p>
          {availableCount >= BOARD_MIN_SIZE ? (
            <>
              <div className="flex items-center gap-3 mb-3 flex-wrap">
                <button onClick={() => setBoardSizeClamped((boardSize || BOARD_MIN_SIZE + 1) - 1)} disabled={!boardSize || boardSize <= BOARD_MIN_SIZE} className="kc-btn-ghost w-10 h-10 rounded-xl flex items-center justify-center font-bold text-lg disabled:opacity-30">−</button>
                <input
                  type="number"
                  min={BOARD_MIN_SIZE}
                  max={availableCount}
                  value={boardSize ?? ""}
                  placeholder="-"
                  onChange={(e) => setBoardSizeClamped(parseInt(e.target.value, 10))}
                  className="kc-input w-20 text-center kc-display font-extrabold text-lg"
                />
                <button onClick={() => setBoardSizeClamped((boardSize || BOARD_MIN_SIZE - 1) + 1)} disabled={!!boardSize && boardSize >= availableCount} className="kc-btn-ghost w-10 h-10 rounded-xl flex items-center justify-center font-bold text-lg disabled:opacity-30">+</button>
                <span className="text-xs kc-font-body" style={{ color: "var(--slate)" }}>kotak (min {BOARD_MIN_SIZE}, maks {availableCount} sesuai soal tersedia)</span>
              </div>
              <div className="flex gap-2 flex-wrap">
                {boardSizePresetOptions.map((n) => (<button key={n} onClick={() => setBoardSize(n)} className={`px-4 py-2 rounded-lg text-xs font-semibold kc-font-body ${boardSize === n ? "kc-chip-active" : "kc-chip"}`}>{n} Kotak</button>))}
                <button onClick={() => setBoardSize(availableCount)} className={`px-4 py-2 rounded-lg text-xs font-semibold kc-font-body ${boardSize === availableCount ? "kc-chip-active" : "kc-chip"}`}>Pakai Semua ({availableCount})</button>
              </div>
            </>
          ) : (
            <p className="text-xs" style={{ color: "var(--red)" }}>Soal yang tersedia untuk kombinasi ini belum cukup untuk membuat papan (minimal {BOARD_MIN_SIZE} soal, baru ada {availableCount}). Coba "Semua Materi" atau "Semua Tingkat".</p>
          )}
          {availableCount > 0 && <p className="text-[11px] mt-2" style={{ color: "var(--slate)" }}>{availableCount} soal tersedia untuk kombinasi ini.</p>}
        </div>
      )}

      <div className="kc-card rounded-2xl p-5 mb-4">
        <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>6 · WAKTU PER SOAL</div>
        <div className="flex gap-3 flex-wrap">{TIMER_OPTIONS.map((t) => (<button key={t} onClick={() => setTimerSeconds(t)} className={`px-6 py-3 rounded-xl font-semibold text-sm kc-font-body flex items-center gap-2 ${timerSeconds === t ? "kc-chip-active" : "kc-chip"}`}><Clock size={14} /> {t} detik</button>))}</div>
      </div>

      <div className="kc-card rounded-2xl p-5 mb-6">
        <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>7 · NAMA KELOMPOK (4 Kelompok, bergiliran maju)</div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          {state.teams.map((g) => { const Icon = GROUP_ICONS[g.iconIdx]; return (
            <div key={g.id} className="flex items-center gap-3 rounded-xl px-3 py-2.5 kc-chip">
              <div className="w-9 h-9 rounded-lg flex items-center justify-center shrink-0" style={{ background: g.color + "33" }}><Icon size={16} style={{ color: g.color }} /></div>
              <input value={g.name} maxLength={16} onChange={(e) => dispatch({ type: "UPDATE_TEAM_NAME", payload: { id: g.id, name: e.target.value } })} className="bg-transparent outline-none text-sm font-semibold kc-font-body w-full text-white" placeholder="Nama kelompok" />
            </div>
          ); })}
        </div>
      </div>

      <button disabled={!canStart} onClick={() => dispatch({ type: "START_BOARD", payload: { subject, grade, boardSize, timerSeconds, topic: topic || null, difficulty } })} className={`w-full py-4 rounded-xl font-bold text-base kc-font-body flex items-center justify-center gap-2 ${canStart ? "kc-btn-gold" : "kc-btn-ghost opacity-40 cursor-not-allowed"}`}>
        <LayoutGrid size={18} /> Buka Papan Kuis!
      </button>
    </div>
  );
}

/* ============================================================
   TOURNAMENT SETUP
   ============================================================ */
function TournamentSetup({ state, dispatch, setScreen }) {
  const subjects = useMemo(() => getAvailableSubjects(state.questionBank), [state.questionBank]);
  const [subject, setSubject] = useState(null);
  const [grade, setGrade] = useState(null);
  const [topic, setTopic] = useState("");
  const [difficulty, setDifficulty] = useState("all");
  const [selectedRounds, setSelectedRounds] = useState([...ROUND_ORDER]);
  const availableGrades = useMemo(() => (subject ? [...new Set(state.questionBank.filter((q) => q.subject === subject).map((q) => q.grade))].sort() : []), [subject, state.questionBank]);
  const availableTopics = useMemo(() => (subject && grade ? [...new Set(state.questionBank.filter((q) => q.subject === subject && q.grade === grade).map((q) => q.topic))].sort() : []), [subject, grade, state.questionBank]);
  const toggleRound = (key) => setSelectedRounds((prev) => (prev.includes(key) ? prev.filter((r) => r !== key) : [...prev, key]));
  const canStart = subject && grade && selectedRounds.length > 0 && state.teams.every((g) => g.name.trim().length > 0);

  return (
    <div className="max-w-4xl mx-auto px-4 pb-24 pt-8">
      <button onClick={() => setScreen("home")} className="flex items-center gap-1 text-sm mb-4 kc-font-body" style={{ color: "var(--slate)" }}><ChevronLeft size={16} /> Kembali</button>
      <h1 className="kc-display font-extrabold text-3xl mb-1">Tournament Mode</h1>
      <p className="text-sm mb-8 kc-font-body" style={{ color: "var(--slate)" }}>Susun turnamen 5 babak untuk kompetisi kelas yang lebih menantang.</p>

      <div className="kc-card rounded-2xl p-5 mb-4">
        <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>1 · PILIH MATA PELAJARAN</div>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          {subjects.map((s) => (<button key={s.key} onClick={() => { setSubject(s.key); setGrade(null); setTopic(""); setDifficulty("all"); }} className={`rounded-xl p-4 text-center kc-font-body ${subject === s.key ? "kc-chip-active" : "kc-chip"}`}><div className="text-2xl mb-1">{s.icon}</div><div className="text-sm font-semibold">{s.label}</div></button>))}
        </div>
      </div>

      {subject && (
        <div className="kc-card rounded-2xl p-5 mb-4 kc-pop">
          <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>2 · PILIH KELAS</div>
          <div className="flex gap-3 flex-wrap">{availableGrades.map((g) => (<button key={g} onClick={() => { setGrade(g); setTopic(""); setDifficulty("all"); }} className={`px-6 py-3 rounded-xl font-semibold text-sm kc-font-body ${grade === g ? "kc-chip-active" : "kc-chip"}`}>Kelas {g}</button>))}</div>
        </div>
      )}

      {grade && (
        <div className="kc-card rounded-2xl p-5 mb-4 kc-pop">
          <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>3 · MATERI & KESULITAN (berlaku di seluruh babak)</div>
          <div className="flex gap-2 flex-wrap mb-2">
            <button onClick={() => setTopic("")} className={`px-4 py-2 rounded-lg font-semibold text-xs kc-font-body ${topic === "" ? "kc-chip-active" : "kc-chip"}`}>Semua Materi</button>
            {availableTopics.map((t) => (<button key={t} onClick={() => setTopic(t)} className={`px-4 py-2 rounded-lg font-semibold text-xs kc-font-body ${topic === t ? "kc-chip-active" : "kc-chip"}`}>{t}</button>))}
          </div>
          <div className="flex gap-2 flex-wrap">
            {[{ key: "all", label: "Semua Tingkat" }, { key: "easy", label: "Mudah" }, { key: "medium", label: "Sedang" }, { key: "hard", label: "Sulit" }].map((d) => (
              <button key={d.key} onClick={() => setDifficulty(d.key)} className={`px-4 py-2 rounded-lg font-semibold text-xs kc-font-body ${difficulty === d.key ? "kc-chip-active" : "kc-chip"}`}>{d.label}</button>
            ))}
          </div>
        </div>
      )}

      <div className="kc-card rounded-2xl p-5 mb-4">
        <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>4 · PILIH BABAK YANG DIGUNAKAN</div>
        <div className="space-y-2">
          {ROUND_ORDER.map((key) => { const r = ROUND_DEFS[key]; const Icon = r.icon; const active = selectedRounds.includes(key); return (
            <button key={key} onClick={() => toggleRound(key)} className={`w-full flex items-center gap-3 rounded-xl px-4 py-3 text-left kc-font-body ${active ? "kc-chip-active" : "kc-chip"}`}>
              {active ? <CheckCircle2 size={18} style={{ color: "var(--blue-soft)" }} /> : <Circle size={18} style={{ color: "var(--slate)" }} />}
              <Icon size={16} style={{ color: "var(--gold)" }} />
              <div className="flex-1"><div className="text-sm font-semibold">{r.name}</div><div className="text-[11px]" style={{ color: "var(--slate)" }}>{r.questionCount} soal · {r.timerSeconds} detik/soal</div></div>
            </button>
          ); })}
        </div>
        {selectedRounds.length === 0 && <p className="text-xs mt-2" style={{ color: "var(--red)" }}>Pilih minimal 1 babak untuk memulai turnamen.</p>}
      </div>

      <div className="kc-card rounded-2xl p-5 mb-6">
        <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>5 · NAMA KELOMPOK (4 Kelompok)</div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          {state.teams.map((g) => { const Icon = GROUP_ICONS[g.iconIdx]; return (
            <div key={g.id} className="flex items-center gap-3 rounded-xl px-3 py-2.5 kc-chip">
              <div className="w-9 h-9 rounded-lg flex items-center justify-center shrink-0" style={{ background: g.color + "33" }}><Icon size={16} style={{ color: g.color }} /></div>
              <input value={g.name} maxLength={16} onChange={(e) => dispatch({ type: "UPDATE_TEAM_NAME", payload: { id: g.id, name: e.target.value } })} className="bg-transparent outline-none text-sm font-semibold kc-font-body w-full text-white" />
            </div>
          ); })}
        </div>
      </div>

      <button disabled={!canStart} onClick={() => dispatch({ type: "START_TOURNAMENT", payload: { subject, grade, selectedRounds, topic: topic || null, difficulty } })} className={`w-full py-4 rounded-xl font-bold text-base kc-font-body flex items-center justify-center gap-2 ${canStart ? "kc-btn-gold" : "kc-btn-ghost opacity-40 cursor-not-allowed"}`}>
        <Swords size={18} /> Mulai Turnamen!
      </button>
    </div>
  );
}

/* ============================================================
   ROUND TRANSITION
   ============================================================ */
function RoundTransitionScreen({ state, dispatch }) {
  const def = ROUND_DEFS[state.currentRound];
  const Icon = def.icon;
  const isHeadToHead = state.currentRound === "headtohead";
  const isSemifinal = state.currentRound === "semifinal";
  const isFinal = state.currentRound === "final";
  const [teamA, setTeamA] = useState(null);
  const [teamB, setTeamB] = useState(null);
  const eligibleTeams = isFinal || isSemifinal ? state.teams.filter((t) => state.finalists.includes(t.id)) : state.teams;
  const canStart = !isHeadToHead || (teamA != null && teamB != null && teamA !== teamB);
  const handleStart = () => dispatch({ type: "START_ROUND", payload: isHeadToHead ? { duelTeams: { teamAId: teamA, teamBId: teamB } } : {} });

  return (
    <div className="max-w-3xl mx-auto px-4 pb-24 pt-10">
      <div className="text-center mb-6">
        <div className="text-xs font-bold tracking-widest mb-2" style={{ color: "var(--slate)" }}>BABAK {ROUND_ORDER.indexOf(state.currentRound) + 1} / {ROUND_ORDER.length}</div>
        <div className="w-16 h-16 rounded-2xl kc-card flex items-center justify-center mx-auto mb-4"><Icon size={28} style={{ color: "var(--gold)" }} /></div>
        <h1 className="kc-display font-extrabold text-3xl mb-1">{def.name}</h1>
        <p className="text-sm kc-font-body" style={{ color: "var(--slate)" }}>{def.shortName}</p>
      </div>

      <div className="kc-card rounded-2xl p-6 mb-4">
        <p className="text-sm kc-font-body mb-4">{def.description}</p>
        <div className="grid grid-cols-3 gap-3 mb-4">
          <div className="kc-chip rounded-xl p-3 text-center"><div className="text-lg font-bold kc-display">{def.questionCount}</div><div className="text-[10px]" style={{ color: "var(--slate)" }}>SOAL</div></div>
          <div className="kc-chip rounded-xl p-3 text-center"><div className="text-lg font-bold kc-display">{def.timerSeconds}s</div><div className="text-[10px]" style={{ color: "var(--slate)" }}>TIMER</div></div>
          <div className="kc-chip rounded-xl p-3 text-center"><RoundStatusBadge status={state.roundStatus[state.currentRound] || "pending"} /></div>
        </div>
        <div className="text-xs font-bold tracking-wide mb-2" style={{ color: "var(--gold-soft)" }}>ATURAN BABAK</div>
        <ul className="space-y-1.5 mb-2">{def.rules.map((r, i) => (<li key={i} className="text-xs kc-font-body flex items-start gap-2" style={{ color: "var(--slate)" }}><ChevronRight size={12} className="mt-0.5 shrink-0" style={{ color: "var(--blue-soft)" }} /> {r}</li>))}</ul>
        <div className="text-xs kc-font-body" style={{ color: "var(--slate)" }}>Sistem poin: {def.pointsMode}</div>
      </div>

      {(isSemifinal || isFinal) && (
        <div className="kc-card rounded-2xl p-5 mb-4">
          <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>FINALIS TERPILIH</div>
          <div className="flex flex-wrap gap-2">{eligibleTeams.map((t) => { const TI = GROUP_ICONS[t.iconIdx]; return (<div key={t.id} className="kc-chip rounded-lg px-3 py-2 flex items-center gap-2"><TI size={14} style={{ color: t.color }} /><span className="text-xs font-semibold kc-font-body">{t.name}</span><span className="text-xs" style={{ color: "var(--slate)" }}>{t.score} pts</span></div>); })}</div>
          <button onClick={() => dispatch({ type: "SET_SCREEN", payload: "finalist-select" })} className="text-xs mt-3 kc-font-body underline" style={{ color: "var(--blue-soft)" }}>Ubah finalis</button>
        </div>
      )}

      {isHeadToHead && (
        <div className="kc-card rounded-2xl p-5 mb-4">
          <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>PILIH DUA KELOMPOK YANG BERDUEL</div>
          <div className="grid grid-cols-2 gap-4">
            {[{ label: "Kelompok A", value: teamA, set: setTeamA, exclude: teamB }, { label: "Kelompok B", value: teamB, set: setTeamB, exclude: teamA }].map((slot, si) => (
              <div key={si}>
                <div className="text-[11px] mb-2 font-semibold" style={{ color: "var(--slate)" }}>{slot.label}</div>
                <div className="space-y-2">{state.teams.filter((t) => t.id !== slot.exclude).map((t) => { const TI = GROUP_ICONS[t.iconIdx]; return (<button key={t.id} onClick={() => slot.set(t.id)} className={`w-full flex items-center gap-2 rounded-lg px-3 py-2 kc-font-body ${slot.value === t.id ? "kc-chip-active" : "kc-chip"}`}><TI size={14} style={{ color: t.color }} /><span className="text-xs font-semibold">{t.name}</span></button>); })}</div>
              </div>
            ))}
          </div>
        </div>
      )}

      <button onClick={handleStart} disabled={!canStart} className={`w-full py-4 rounded-xl font-bold text-base kc-font-body flex items-center justify-center gap-2 ${canStart ? "kc-btn-gold" : "kc-btn-ghost opacity-40 cursor-not-allowed"}`}><Play size={18} /> Mulai Babak Ini</button>
    </div>
  );
}

/* ============================================================
   FINALIST SELECT
   ============================================================ */
function FinalistSelectScreen({ state, dispatch }) {
  const ranked = updateRanking(state.teams);
  const [selected, setSelected] = useState(state.finalists.length ? state.finalists : ranked.slice(0, 2).map((t) => t.id));
  const toggle = (id) => setSelected((s) => (s.includes(id) ? s.filter((x) => x !== id) : [...s, id]));
  const confirm = () => { dispatch({ type: "SET_FINALISTS", payload: selected }); dispatch({ type: "SET_SCREEN", payload: "round-transition" }); };
  return (
    <div className="max-w-2xl mx-auto px-4 pb-24 pt-10">
      <div className="text-center mb-6">
        <div className="w-16 h-16 rounded-2xl kc-card flex items-center justify-center mx-auto mb-4"><Award size={28} style={{ color: "var(--gold)" }} /></div>
        <h1 className="kc-display font-extrabold text-3xl mb-1">Tentukan Finalis</h1>
        <p className="text-sm kc-font-body" style={{ color: "var(--slate)" }}>Pilih kelompok yang lanjut ke babak {ROUND_DEFS[state.currentRound].name}. Minimal 2 kelompok.</p>
      </div>
      <div className="kc-card rounded-2xl p-5 mb-6">
        <div className="space-y-2">{ranked.map((t, i) => { const Icon = GROUP_ICONS[t.iconIdx]; const active = selected.includes(t.id); return (
          <button key={t.id} onClick={() => toggle(t.id)} className={`w-full flex items-center gap-3 rounded-xl px-4 py-3 kc-font-body ${active ? "kc-chip-active" : "kc-chip"}`}>
            {active ? <CheckCircle2 size={18} style={{ color: "var(--blue-soft)" }} /> : <Circle size={18} style={{ color: "var(--slate)" }} />}
            <span className="text-xs font-bold" style={{ color: "var(--slate)" }}>#{i + 1}</span>
            <div className="w-7 h-7 rounded-lg flex items-center justify-center" style={{ background: t.color + "33" }}><Icon size={14} style={{ color: t.color }} /></div>
            <span className="text-sm font-semibold flex-1 text-left">{t.name}</span><span className="kc-display font-bold text-sm">{t.score} pts</span>
          </button>
        ); })}</div>
      </div>
      <button onClick={confirm} disabled={selected.length < 2} className={`w-full py-4 rounded-xl font-bold text-base kc-font-body flex items-center justify-center gap-2 ${selected.length >= 2 ? "kc-btn-gold" : "kc-btn-ghost opacity-40 cursor-not-allowed"}`}><ArrowRight size={18} /> Konfirmasi Finalis ({selected.length})</button>
    </div>
  );
}

/* ============================================================
   GENERIC ROUND ARENA
   ============================================================ */
function ArenaScreen({ state, dispatch }) {
  const total = state.questions.length;
  const q = state.questions[state.currentQuestionIndex];
  const { timer, gameStatus } = state;
  const revealed = timer.locked;
  const def = state.currentRound ? ROUND_DEFS[state.currentRound] : null;
  const restrictedRound = state.currentRound === "semifinal" || state.currentRound === "final";
  const activeTeams = restrictedRound && state.finalists.length > 0 ? state.teams.filter((t) => state.finalists.includes(t.id)) : state.teams;
  const eliminatedIds = restrictedRound && state.finalists.length > 0 ? state.teams.filter((t) => !state.finalists.includes(t.id)).map((t) => t.id) : [];
  const [showScoreboard, setShowScoreboard] = useState(false);
  const [pickIdx, setPickIdx] = useState(null); // pilihan jawaban yang disentuh sebelum ada kelompok penjawab
  if (!q) return null;
  const isChoice = CHOICE_TYPES.includes(q.type);
  const answeringTeam = state.answeringTeamId != null ? state.teams.find((t) => t.id === state.answeringTeamId) : null;
  const difficultyColor = { easy: "var(--green)", medium: "var(--gold)", hard: "var(--red)" }[q.difficulty] || "var(--blue)";
  const selectedIndexForReveal = state.selectedAnswers[q.id]?.selectedIndex ?? null;
  const handleMark = (teamId, result) => { if (!revealed) dispatch({ type: "ANSWER_LOCK", payload: { teamId, result, basePoints: q.points, difficulty: q.difficulty, penaltyEnabled: state.settings.selfSelectPenaltyEnabled, penaltyPoints: q.points } }); };
  const handleSkip = () => { if (!revealed) dispatch({ type: "ANSWER_LOCK", payload: { teamId: null, result: "skip", basePoints: 0 } }); };
  const gradeSelection = (teamId, idx) => {
    const correct = idx === q.answer;
    dispatch({ type: "ANSWER_LOCK", payload: { teamId, result: correct ? "correct" : "wrong", basePoints: q.points, difficulty: q.difficulty, penaltyEnabled: state.settings.selfSelectPenaltyEnabled, penaltyPoints: q.points, selectedIndex: idx } });
  };
  const handleSelfSelect = (idx) => {
    if (revealed) return;
    if (state.answeringTeamId == null) { setPickIdx(idx); return; } // belum ada penjawab → tanya kelompok mana
    gradeSelection(state.answeringTeamId, idx);
  };
  const toggleAnswerer = (teamId) => { if (!revealed) dispatch({ type: "SET_ANSWERING_TEAM", payload: state.answeringTeamId === teamId ? null : teamId }); };
  const goNext = () => dispatch({ type: "NEXT_QUESTION" });

  return (
    <div className="max-w-6xl mx-auto px-4 pb-16 pt-6">
      <div className="flex items-center justify-between flex-wrap gap-3 mb-5">
        <div className="flex items-center gap-2 flex-wrap">
          <span className="kc-chip px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold kc-font-body">{q.subject} · Kelas {q.grade}</span>
          <span className="kc-chip px-3 py-1.5 rounded-lg text-xs md:text-sm font-semibold kc-font-body">{q.topic}</span>
          {def && <span className="px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold kc-font-body" style={{ background: "rgba(242,183,5,0.15)", color: "var(--gold-soft)" }}>{def.name}</span>}
          <span className="px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold kc-font-body" style={{ background: difficultyColor + "22", color: difficultyColor }}>{q.difficulty === "easy" ? "Mudah" : q.difficulty === "medium" ? "Sedang" : "Sulit"} · {q.points} pts</span>
        </div>
        <div className="flex items-center gap-2">
          <button onClick={() => setShowScoreboard(true)} className="kc-btn-ghost px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold kc-font-body flex items-center gap-1.5"><Trophy size={14} style={{ color: "var(--gold)" }} /> Papan Skor</button>
          <button onClick={() => dispatch({ type: "PAUSE_SESSION" })} className="kc-btn-ghost px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold kc-font-body flex items-center gap-1.5"><Save size={14} /> Simpan &amp; Keluar</button>
          <div className="kc-display font-extrabold text-sm md:text-base" style={{ color: "var(--slate)" }}>Soal {state.currentQuestionIndex + 1} / {total}</div>
        </div>
      </div>
      <div className="w-full h-1.5 rounded-full bg-white/10 mb-6 overflow-hidden"><div className="h-full rounded-full" style={{ width: `${((state.currentQuestionIndex + (revealed ? 1 : 0)) / total) * 100}%`, background: "linear-gradient(90deg, var(--blue), var(--gold))", transition: "width .4s ease" }} /></div>

      <div className="grid grid-cols-1 lg:grid-cols-[1fr_280px] gap-5">
        <div>
          <div className={`kc-card rounded-2xl p-6 md:p-8 mb-5 flex flex-col items-center text-center ${gameStatus === "answering" ? "kc-urgent" : ""}`}>
            <TimerRing timeLeft={timer.remaining} total={timer.total} running={timer.running && !revealed} />
            <h2 className="kc-display font-bold text-2xl md:text-3xl lg:text-4xl mt-6 leading-snug">{q.question}</h2>
          </div>

          {isChoice && answeringTeam && !revealed && (
            <div className="rounded-xl px-4 py-3 mb-4 flex items-center gap-2 kc-pop" style={{ background: "rgba(47,111,237,0.14)", border: "1px solid rgba(47,111,237,0.4)" }}>
              {(() => { const Icon = GROUP_ICONS[answeringTeam.iconIdx]; return <Icon size={16} style={{ color: answeringTeam.color }} />; })()}
              <span className="text-sm font-bold kc-font-body">Giliran menjawab: {answeringTeam.name}</span>
              <span className="text-xs kc-font-body ml-auto" style={{ color: "var(--slate)" }}>Perwakilan tinggal sentuh salah satu pilihan di bawah →</span>
            </div>
          )}

          {isChoice && !answeringTeam && !revealed && (<div className="rounded-xl px-4 py-3 mb-4 text-xs md:text-sm font-semibold kc-font-body" style={{ background: "rgba(47,111,237,0.10)", border: "1px solid rgba(47,111,237,0.3)" }}>Sentuh salah satu pilihan jawaban, lalu pilih kelompok yang menjawab.</div>)}
          <QuestionBody q={q} revealed={revealed} large={true} selectable={isChoice} selectedIndex={selectedIndexForReveal} onSelect={handleSelfSelect} />
          {pickIdx != null && !revealed && (<AnswererPickerModal teams={activeTeams} optionLabel={answerOptionLabel(q, pickIdx)} onCancel={() => setPickIdx(null)} onPick={(teamId) => { const i = pickIdx; setPickIdx(null); gradeSelection(teamId, i); }} />)}

          {revealed && (
            <div className="kc-card rounded-2xl p-5 mt-5 kc-pop">
              <div className="flex items-center gap-2 mb-2"><ListChecks size={16} style={{ color: "var(--gold)" }} /><span className="text-xs font-bold tracking-wide" style={{ color: "var(--gold-soft)" }}>PEMBAHASAN</span></div>
              <p className="text-sm md:text-base kc-font-body">{q.explanation}</p>
              <ManualScoreAdjust teams={activeTeams} dispatch={dispatch} />
              <button onClick={goNext} className="kc-btn-primary mt-5 px-6 py-3.5 rounded-xl font-bold text-sm md:text-base kc-font-body flex items-center gap-2">{state.currentQuestionIndex + 1 < total ? "Soal Berikutnya" : "Lanjutkan"} <ArrowRight size={18} /></button>
            </div>
          )}
        </div>

        <div className="space-y-4">
          <div className="kc-card rounded-2xl p-4">
            <div className="flex items-center justify-between mb-3">
              <span className="text-xs font-bold tracking-wide" style={{ color: "var(--gold-soft)" }}>KONTROL GURU</span>
              <div className="flex items-center gap-1">
                <button onClick={() => dispatch({ type: timer.running ? "TIMER_PAUSE" : "TIMER_RESUME" })} disabled={revealed} className="w-9 h-9 rounded-lg flex items-center justify-center kc-btn-ghost disabled:opacity-30">{timer.running ? <PauseCircle size={18} /> : <PlayCircle size={18} />}</button>
                <button onClick={() => dispatch({ type: "TIMER_RESET" })} disabled={revealed} className="w-9 h-9 rounded-lg flex items-center justify-center kc-btn-ghost disabled:opacity-30"><RefreshCw size={16} /></button>
              </div>
            </div>
            {isChoice ? (
              <p className="text-[11px] mb-3 kc-font-body" style={{ color: "var(--slate)" }}>Tunjuk perwakilan yang maju agar mereka bisa langsung menyentuh pilihan jawaban (otomatis dinilai) — atau nilai manual dengan Benar/Salah di bawah.</p>
            ) : (
              <p className="text-[11px] mb-3 kc-font-body" style={{ color: "var(--slate)" }}>Klik kelompok yang menjawab benar/salah secara lisan.</p>
            )}
            <div className="space-y-2">
              {activeTeams.map((g) => { const Icon = GROUP_ICONS[g.iconIdx]; const isAnswerer = state.answeringTeamId === g.id; return (
                <div key={g.id} className={`rounded-xl p-2.5 ${isAnswerer ? "kc-chip-active" : "kc-chip"}`}>
                  {isChoice ? (
                    <button disabled={revealed} onClick={() => toggleAnswerer(g.id)} className="w-full flex items-center gap-2 mb-2 disabled:opacity-40 kc-font-body text-left">
                      <div className="w-7 h-7 rounded-md flex items-center justify-center shrink-0" style={{ background: g.color + "33" }}><Icon size={14} style={{ color: g.color }} /></div>
                      <span className="text-sm font-semibold truncate">{g.name}</span>
                      {g.streak >= 2 && <span className="text-[10px] flex items-center gap-0.5" style={{ color: "var(--gold)" }}><Flame size={11} />{g.streak}</span>}
                      <span className="text-[10px] font-bold ml-auto shrink-0" style={{ color: isAnswerer ? "var(--blue-soft)" : "var(--slate)" }}>{isAnswerer ? "✓ Menjawab" : "Tunjuk"}</span>
                    </button>
                  ) : (
                    <div className="flex items-center gap-2 mb-2">
                      <div className="w-7 h-7 rounded-md flex items-center justify-center shrink-0" style={{ background: g.color + "33" }}><Icon size={14} style={{ color: g.color }} /></div>
                      <span className="text-sm font-semibold kc-font-body truncate">{g.name}</span>
                      {g.streak >= 2 && <span className="text-[10px] ml-auto flex items-center gap-0.5" style={{ color: "var(--gold)" }}><Flame size={11} />{g.streak}</span>}
                    </div>
                  )}
                  <div className="flex gap-1.5">
                    <button disabled={revealed} onClick={() => handleMark(g.id, "correct")} className="flex-1 py-2.5 rounded-lg text-xs font-bold flex items-center justify-center gap-1 disabled:opacity-30 kc-font-body" style={{ background: "rgba(34,197,94,0.16)", color: "var(--green)" }}><Check size={14} /> Benar</button>
                    <button disabled={revealed} onClick={() => handleMark(g.id, "wrong")} className="flex-1 py-2.5 rounded-lg text-xs font-bold flex items-center justify-center gap-1 disabled:opacity-30 kc-font-body" style={{ background: "rgba(240,67,91,0.16)", color: "var(--red)" }}><X size={14} /> Salah</button>
                  </div>
                </div>
              ); })}
            </div>
            <button disabled={revealed} onClick={handleSkip} className="w-full mt-3 py-2.5 rounded-lg text-xs font-semibold kc-btn-ghost disabled:opacity-30 kc-font-body">Tidak Ada yang Menjawab</button>
          </div>

          <div className="kc-card rounded-2xl p-4">
            <div className="flex items-center gap-2 mb-3"><Trophy size={14} style={{ color: "var(--gold)" }} /><span className="text-xs font-bold tracking-wide" style={{ color: "var(--gold-soft)" }}>PAPAN SKOR</span></div>
            <ScoreboardList teams={state.teams} eliminatedIds={eliminatedIds} />
          </div>
        </div>
      </div>
      {showScoreboard && <ScoreboardOverlay teams={state.teams} eliminatedIds={eliminatedIds} onClose={() => setShowScoreboard(false)} />}
    </div>
  );
}

/* ============================================================
   DUEL ARENA
   ============================================================ */
function DuelArena({ state, dispatch }) {
  const total = state.questions.length;
  const q = state.questions[state.currentQuestionIndex];
  const { timer, gameStatus } = state;
  const revealed = timer.locked;
  const [showScoreboard, setShowScoreboard] = useState(false);
  const [pickIdx, setPickIdx] = useState(null);
  if (!q || !state.duel) return null;
  const teamA = state.teams.find((t) => t.id === state.duel.teamAId);
  const teamB = state.teams.find((t) => t.id === state.duel.teamBId);
  const winsA = state.duel.wins[teamA.id] || 0;
  const winsB = state.duel.wins[teamB.id] || 0;
  const isLastQuestion = state.currentQuestionIndex === total - 1;
  const isChoice = CHOICE_TYPES.includes(q.type);
  const answeringTeam = state.answeringTeamId != null ? state.teams.find((t) => t.id === state.answeringTeamId) : null;
  const selectedIndexForReveal = state.selectedAnswers[q.id]?.selectedIndex ?? null;
  const handleMark = (teamId, result) => { if (!revealed) dispatch({ type: "ANSWER_LOCK", payload: { teamId, result, basePoints: q.points, difficulty: q.difficulty, penaltyEnabled: state.settings.selfSelectPenaltyEnabled, penaltyPoints: q.points } }); };
  const handleSkip = () => { if (!revealed) dispatch({ type: "ANSWER_LOCK", payload: { teamId: null, result: "skip", basePoints: 0 } }); };
  const toggleAnswerer = (teamId) => { if (!revealed) dispatch({ type: "SET_ANSWERING_TEAM", payload: state.answeringTeamId === teamId ? null : teamId }); };
  const gradeSelection = (teamId, idx) => {
    const correct = idx === q.answer;
    dispatch({ type: "ANSWER_LOCK", payload: { teamId, result: correct ? "correct" : "wrong", basePoints: q.points, difficulty: q.difficulty, penaltyEnabled: state.settings.selfSelectPenaltyEnabled, penaltyPoints: q.points, selectedIndex: idx } });
  };
  const handleSelfSelect = (idx) => {
    if (revealed) return;
    if (state.answeringTeamId == null) { setPickIdx(idx); return; }
    gradeSelection(state.answeringTeamId, idx);
  };

  const DuelCard = ({ team, wins }) => { const Icon = GROUP_ICONS[team.iconIdx]; const isAnswerer = state.answeringTeamId === team.id; return (
    <div className={`kc-card rounded-2xl p-5 flex-1 ${isAnswerer ? "kc-chip-active" : ""}`}>
      <div className="flex items-center gap-3 mb-4">
        <div className="w-12 h-12 rounded-xl flex items-center justify-center" style={{ background: team.color + "33" }}><Icon size={22} style={{ color: team.color }} /></div>
        <div><div className="kc-display font-bold text-lg md:text-xl">{team.name}</div><div className="text-xs kc-font-body" style={{ color: "var(--slate)" }}>{team.score} poin total</div></div>
        <div className="ml-auto text-center"><div className="kc-display font-extrabold text-3xl" style={{ color: "var(--gold)" }}>{wins}</div><div className="text-[10px] font-bold tracking-wide" style={{ color: "var(--slate)" }}>MENANG</div></div>
      </div>
      {isChoice && (
        <button disabled={revealed} onClick={() => toggleAnswerer(team.id)} className="w-full mb-2 py-2 rounded-lg text-xs font-bold kc-font-body disabled:opacity-30" style={{ background: isAnswerer ? "rgba(47,111,237,0.25)" : "rgba(255,255,255,0.06)", color: isAnswerer ? "var(--blue-soft)" : "var(--slate)" }}>
          {isAnswerer ? "✓ Ditunjuk menjawab — sentuh pilihan di bawah" : "Tunjuk sebagai Penjawab"}
        </button>
      )}
      <div className="flex gap-2">
        <button disabled={revealed} onClick={() => handleMark(team.id, "correct")} className="flex-1 py-3 rounded-xl text-sm font-bold flex items-center justify-center gap-1.5 disabled:opacity-30 kc-font-body" style={{ background: "rgba(34,197,94,0.16)", color: "var(--green)" }}><Check size={15} /> Benar</button>
        <button disabled={revealed} onClick={() => handleMark(team.id, "wrong")} className="flex-1 py-3 rounded-xl text-sm font-bold flex items-center justify-center gap-1.5 disabled:opacity-30 kc-font-body" style={{ background: "rgba(240,67,91,0.16)", color: "var(--red)" }}><X size={15} /> Salah</button>
      </div>
    </div>
  ); };

  return (
    <div className="max-w-4xl mx-auto px-4 pb-16 pt-6">
      <div className="flex items-center justify-between mb-5 flex-wrap gap-2">
        <span className="kc-chip px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold kc-font-body flex items-center gap-1.5"><Swords size={13} /> Head to Head</span>
        <div className="flex items-center gap-2">
          <button onClick={() => setShowScoreboard(true)} className="kc-btn-ghost px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold kc-font-body flex items-center gap-1.5"><Trophy size={14} style={{ color: "var(--gold)" }} /> Papan Skor</button>
          <button onClick={() => dispatch({ type: "PAUSE_SESSION" })} className="kc-btn-ghost px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold kc-font-body flex items-center gap-1.5"><Save size={14} /> Simpan &amp; Keluar</button>
          <span className="kc-display font-extrabold text-sm md:text-base" style={{ color: "var(--slate)" }}>Soal {state.currentQuestionIndex + 1} / {total}</span>
        </div>
      </div>
      <div className="flex gap-4 mb-5">
        <DuelCard team={teamA} wins={winsA} /><div className="flex items-center kc-display font-extrabold text-2xl" style={{ color: "var(--slate)" }}>VS</div><DuelCard team={teamB} wins={winsB} />
      </div>
      {showScoreboard && <ScoreboardOverlay teams={state.teams} onClose={() => setShowScoreboard(false)} />}
      <div className={`kc-card rounded-2xl p-6 md:p-8 mb-5 flex flex-col items-center text-center ${gameStatus === "answering" ? "kc-urgent" : ""}`}>
        <TimerRing timeLeft={timer.remaining} total={timer.total} running={timer.running && !revealed} />
        <h2 className="kc-display font-bold text-2xl md:text-3xl lg:text-4xl mt-6 leading-snug">{q.question}</h2>
      </div>
      {isChoice && answeringTeam && !revealed && (
        <div className="rounded-xl px-4 py-3 mb-4 flex items-center gap-2" style={{ background: "rgba(47,111,237,0.14)", border: "1px solid rgba(47,111,237,0.4)" }}>
          <span className="text-sm font-bold kc-font-body">Giliran menjawab: {answeringTeam.name} — sentuh salah satu pilihan di bawah</span>
        </div>
      )}
      <div className="mb-5">{isChoice && !answeringTeam && !revealed && (<div className="rounded-xl px-4 py-3 mb-4 text-xs md:text-sm font-semibold kc-font-body" style={{ background: "rgba(47,111,237,0.10)", border: "1px solid rgba(47,111,237,0.3)" }}>Sentuh salah satu pilihan jawaban, lalu pilih kelompok yang menjawab.</div>)}<QuestionBody q={q} revealed={revealed} large={true} selectable={isChoice} selectedIndex={selectedIndexForReveal} onSelect={handleSelfSelect} />{pickIdx != null && !revealed && (<AnswererPickerModal teams={[teamA, teamB]} optionLabel={answerOptionLabel(q, pickIdx)} onCancel={() => setPickIdx(null)} onPick={(teamId) => { const i = pickIdx; setPickIdx(null); gradeSelection(teamId, i); }} />)}</div>
      <button disabled={revealed} onClick={handleSkip} className="w-full mb-4 py-2.5 rounded-xl text-xs font-semibold kc-btn-ghost disabled:opacity-30 kc-font-body">Tidak Ada yang Menjawab</button>
      {revealed && (
        <div className="kc-card rounded-2xl p-5 kc-pop">
          <div className="flex items-center gap-2 mb-2"><ListChecks size={16} style={{ color: "var(--gold)" }} /><span className="text-xs font-bold tracking-wide" style={{ color: "var(--gold-soft)" }}>PEMBAHASAN</span></div>
          <p className="text-sm kc-font-body mb-4">{q.explanation}</p>
          <ManualScoreAdjust teams={[teamA, teamB]} dispatch={dispatch} />
          {isLastQuestion && (
            <div className="rounded-xl p-4 mb-4 text-center" style={{ background: "rgba(242,183,5,0.12)", border: "1px solid rgba(242,183,5,0.4)" }}>
              <Trophy size={22} className="mx-auto mb-1" style={{ color: "var(--gold)" }} />
              <div className="kc-display font-bold text-sm">{winsA === winsB ? "Duel Berakhir Seri" : `Pemenang Duel: ${winsA > winsB ? teamA.name : teamB.name}`}</div>
            </div>
          )}
          <button onClick={() => dispatch({ type: "NEXT_QUESTION" })} className="kc-btn-primary px-6 py-3 rounded-xl font-bold text-sm kc-font-body flex items-center gap-2">{isLastQuestion ? "Lanjut ke Babak Berikutnya" : "Soal Berikutnya"} <ArrowRight size={16} /></button>
        </div>
      )}
    </div>
  );
}

/* ============================================================
   STRATEGY ARENA
   ============================================================ */
function StrategyArenaScreen({ state, dispatch }) {
  const total = state.questions.length;
  const q = state.questions[state.currentQuestionIndex];
  const { timer, gameStatus, strategySelection } = state;
  const revealed = timer.locked;
  const [pendingDifficulty, setPendingDifficulty] = useState(null);
  const [showScoreboard, setShowScoreboard] = useState(false);
  useEffect(() => { setPendingDifficulty(null); }, [state.currentQuestionIndex]);
  if (!q) return null;
  const step = !strategySelection ? "pick-team" : !strategySelection.confirmed ? "pick-difficulty" : "playing";
  const activeTeam = strategySelection ? state.teams.find((t) => t.id === strategySelection.teamId) : null;
  const level = strategySelection?.difficulty ? STRATEGY_LEVELS.find((l) => l.key === strategySelection.difficulty) : null;
  const isChoice = CHOICE_TYPES.includes(q.type);
  const selectedIndexForReveal = state.selectedAnswers[q.id]?.selectedIndex ?? null;
  const handleMark = (result) => { if (revealed || !level || !activeTeam) return; dispatch({ type: "ANSWER_LOCK", payload: { teamId: activeTeam.id, result, basePoints: level.points, difficulty: level.key, penaltyEnabled: state.settings.strategyPenaltyEnabled } }); };
  const handleSelfSelect = (idx) => {
    if (revealed || !level || !activeTeam) return;
    const correct = idx === q.answer;
    dispatch({ type: "ANSWER_LOCK", payload: { teamId: activeTeam.id, result: correct ? "correct" : "wrong", basePoints: level.points, difficulty: level.key, penaltyEnabled: state.settings.strategyPenaltyEnabled, selectedIndex: idx } });
  };

  return (
    <div className="max-w-3xl mx-auto px-4 pb-16 pt-6">
      <div className="flex items-center justify-between mb-5 flex-wrap gap-2">
        <span className="kc-chip px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold kc-font-body flex items-center gap-1.5"><Target size={13} /> Strategy Arena</span>
        <div className="flex items-center gap-2">
          <button onClick={() => setShowScoreboard(true)} className="kc-btn-ghost px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold kc-font-body flex items-center gap-1.5"><Trophy size={14} style={{ color: "var(--gold)" }} /> Papan Skor</button>
          <button onClick={() => dispatch({ type: "PAUSE_SESSION" })} className="kc-btn-ghost px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold kc-font-body flex items-center gap-1.5"><Save size={14} /> Simpan &amp; Keluar</button>
          <span className="kc-display font-extrabold text-sm md:text-base" style={{ color: "var(--slate)" }}>Soal {state.currentQuestionIndex + 1} / {total}</span>
        </div>
      </div>
      {showScoreboard && <ScoreboardOverlay teams={state.teams} onClose={() => setShowScoreboard(false)} />}

      {step === "pick-team" && (
        <div className="kc-card rounded-2xl p-6 kc-pop">
          <div className="text-center mb-5"><Target size={32} className="mx-auto mb-2" style={{ color: "var(--gold)" }} /><h2 className="kc-display font-bold text-xl">Kelompok Mana yang Tampil?</h2><p className="text-xs kc-font-body mt-1" style={{ color: "var(--slate)" }}>Pilih kelompok yang akan mengambil tantangan pada giliran ini.</p></div>
          <div className="grid grid-cols-2 gap-3">{state.teams.map((t) => { const Icon = GROUP_ICONS[t.iconIdx]; return (<button key={t.id} onClick={() => dispatch({ type: "STRATEGY_SELECT_TEAM", payload: t.id })} className="kc-chip rounded-xl p-4 flex items-center gap-3 kc-font-body"><div className="w-9 h-9 rounded-lg flex items-center justify-center" style={{ background: t.color + "33" }}><Icon size={16} style={{ color: t.color }} /></div><div className="text-left"><div className="text-sm font-semibold">{t.name}</div><div className="text-[11px]" style={{ color: "var(--slate)" }}>{t.score} pts</div></div></button>); })}</div>
        </div>
      )}

      {step === "pick-difficulty" && activeTeam && (
        <div className="kc-card rounded-2xl p-6 kc-pop">
          <div className="text-center mb-5"><div className="text-xs kc-font-body mb-1" style={{ color: "var(--slate)" }}>Giliran</div><h2 className="kc-display font-bold text-xl">{activeTeam.name}</h2><p className="text-xs kc-font-body mt-1" style={{ color: "var(--slate)" }}>Pilih tingkat tantangan — makin tinggi, makin besar risiko dan hadiahnya.</p></div>
          <div className="grid grid-cols-2 gap-3 mb-4">{STRATEGY_LEVELS.map((lvl) => { const LIcon = lvl.icon; return (<button key={lvl.key} onClick={() => setPendingDifficulty(lvl.key)} className={`rounded-xl p-4 text-left kc-font-body ${pendingDifficulty === lvl.key ? "kc-chip-active" : "kc-chip"}`}><LIcon size={18} style={{ color: lvl.color }} className="mb-2" /><div className="text-sm font-bold" style={{ color: lvl.color }}>{lvl.label}</div><div className="text-xs font-semibold mt-0.5">{lvl.points} poin</div><div className="text-[10px] mt-1" style={{ color: "var(--slate)" }}>{lvl.riskLabel}</div></button>); })}</div>
          {pendingDifficulty && (
            <div className="rounded-xl p-4 mb-4 kc-pop" style={{ background: "rgba(255,255,255,0.04)", border: "1px solid rgba(255,255,255,0.12)" }}>
              <div className="flex items-center gap-2 mb-2"><AlertTriangle size={14} style={{ color: "var(--gold)" }} /><span className="text-xs font-bold">Konfirmasi Pilihan</span></div>
              <p className="text-xs kc-font-body mb-3" style={{ color: "var(--slate)" }}>{activeTeam.name} memilih tingkat <strong>{STRATEGY_LEVELS.find((l) => l.key === pendingDifficulty).label}</strong> ({STRATEGY_LEVELS.find((l) => l.key === pendingDifficulty).points} poin). {STRATEGY_LEVELS.find((l) => l.key === pendingDifficulty).riskLabel}. Lanjutkan?</p>
              <div className="flex gap-2"><button onClick={() => { dispatch({ type: "STRATEGY_SELECT_DIFFICULTY", payload: pendingDifficulty }); dispatch({ type: "STRATEGY_CONFIRM" }); }} className="kc-btn-gold flex-1 py-2.5 rounded-lg text-sm font-bold kc-font-body">Ya, Lanjutkan</button><button onClick={() => setPendingDifficulty(null)} className="kc-btn-ghost flex-1 py-2.5 rounded-lg text-sm font-semibold kc-font-body">Batal</button></div>
            </div>
          )}
          <button onClick={() => dispatch({ type: "STRATEGY_RESET_SELECTION" })} className="text-xs kc-font-body underline" style={{ color: "var(--slate)" }}>Ganti kelompok</button>
        </div>
      )}

      {step === "playing" && level && activeTeam && (
        <>
          <div className="kc-card rounded-2xl p-4 mb-4 flex items-center justify-between">
            <div className="flex items-center gap-2">{(() => { const LIcon = level.icon; return <LIcon size={16} style={{ color: level.color }} />; })()}<span className="text-sm font-bold kc-font-body" style={{ color: level.color }}>{activeTeam.name} · {level.label} · {level.points} pts</span></div>
          </div>
          <div className={`kc-card rounded-2xl p-6 md:p-8 mb-5 flex flex-col items-center text-center ${gameStatus === "answering" ? "kc-urgent" : ""}`}>
            <TimerRing timeLeft={timer.remaining} total={timer.total} running={timer.running && !revealed} />
            <h2 className="kc-display font-bold text-2xl md:text-3xl lg:text-4xl mt-6 leading-snug">{q.question}</h2>
          </div>
          {isChoice && !revealed && (
            <div className="rounded-xl px-4 py-3 mb-4 flex items-center gap-2" style={{ background: "rgba(47,111,237,0.14)", border: "1px solid rgba(47,111,237,0.4)" }}>
              <span className="text-sm font-bold kc-font-body">{activeTeam.name}, sentuh langsung pilihan jawabanmu di bawah — otomatis dinilai.</span>
            </div>
          )}
          <div className="mb-5"><QuestionBody q={q} revealed={revealed} large={true} selectable={isChoice} selectedIndex={selectedIndexForReveal} onSelect={handleSelfSelect} /></div>
          {!revealed && (
            <div className="flex gap-3 mb-5">
              <button onClick={() => handleMark("correct")} className="flex-1 py-3 rounded-xl text-sm font-bold flex items-center justify-center gap-1.5 kc-font-body" style={{ background: "rgba(34,197,94,0.16)", color: "var(--green)" }}><Check size={15} /> {isChoice ? "Nilai Manual: Benar" : "Jawaban Benar"}</button>
              <button onClick={() => handleMark("wrong")} className="flex-1 py-3 rounded-xl text-sm font-bold flex items-center justify-center gap-1.5 kc-font-body" style={{ background: "rgba(240,67,91,0.16)", color: "var(--red)" }}><X size={15} /> {isChoice ? "Nilai Manual: Salah" : "Jawaban Salah"}</button>
            </div>
          )}
          {revealed && (
            <div className="kc-card rounded-2xl p-5 kc-pop">
              <div className="flex items-center gap-2 mb-2"><ListChecks size={16} style={{ color: "var(--gold)" }} /><span className="text-xs font-bold tracking-wide" style={{ color: "var(--gold-soft)" }}>PEMBAHASAN</span></div>
              <p className="text-sm kc-font-body mb-4">{q.explanation}</p>
              <ManualScoreAdjust teams={state.teams} dispatch={dispatch} />
              <button onClick={() => dispatch({ type: "NEXT_QUESTION" })} className="kc-btn-primary px-6 py-3 rounded-xl font-bold text-sm kc-font-body flex items-center gap-2">{state.currentQuestionIndex + 1 < total ? "Soal Berikutnya" : "Lanjutkan"} <ArrowRight size={16} /></button>
            </div>
          )}
        </>
      )}
    </div>
  );
}

/* ============================================================
   HASIL PERMAINAN
   ============================================================ */
function BoardArena({ state, dispatch }) {
  const [showScoreboard, setShowScoreboard] = useState(false);
  const [pickIdx, setPickIdx] = useState(null);
  const tile = state.boardActiveTileId != null ? state.boardTiles.find((t) => t.id === state.boardActiveTileId) : null;
  const q = tile ? state.questions[0] : null;
  const { timer, gameStatus } = state;
  const revealed = timer.locked;
  const currentTurnTeamId = state.boardTurnOrder[state.boardTurnIndex % Math.max(1, state.boardTurnOrder.length)];
  const currentTurnTeam = state.teams.find((t) => t.id === currentTurnTeamId);
  const remainingCount = state.boardTiles.filter((t) => !t.used).length;
  const isChoice = q ? CHOICE_TYPES.includes(q.type) : false;
  const answeringTeam = state.answeringTeamId != null ? state.teams.find((t) => t.id === state.answeringTeamId) : null;
  const selectedIndexForReveal = q ? state.selectedAnswers[q.id]?.selectedIndex ?? null : null;

  const handleMark = (teamId, result) => { if (!q || revealed) return; dispatch({ type: "ANSWER_LOCK", payload: { teamId, result, basePoints: tile.points, difficulty: q.difficulty, penaltyEnabled: state.settings.selfSelectPenaltyEnabled, penaltyPoints: tile.points } }); };
  const handleSkip = () => { if (!q || revealed) return; dispatch({ type: "ANSWER_LOCK", payload: { teamId: null, result: "skip", basePoints: 0 } }); };
  const gradeSelection = (teamId, idx) => {
    const correct = idx === q.answer;
    dispatch({ type: "ANSWER_LOCK", payload: { teamId, result: correct ? "correct" : "wrong", basePoints: tile.points, difficulty: q.difficulty, penaltyEnabled: state.settings.selfSelectPenaltyEnabled, penaltyPoints: tile.points, selectedIndex: idx } });
  };
  const handleSelfSelect = (idx) => {
    if (!q || revealed) return;
    if (state.answeringTeamId == null) { setPickIdx(idx); return; }
    gradeSelection(state.answeringTeamId, idx);
  };
  const toggleAnswerer = (teamId) => { if (!revealed) dispatch({ type: "SET_ANSWERING_TEAM", payload: state.answeringTeamId === teamId ? null : teamId }); };

  // TAMPILAN PAPAN (belum ada kotak yang dibuka)
  if (!tile) {
    return (
      <div className="max-w-5xl mx-auto px-4 pb-16 pt-6">
        <div className="flex items-center justify-between flex-wrap gap-3 mb-5">
          <span className="kc-chip px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold kc-font-body flex items-center gap-1.5"><LayoutGrid size={14} /> Papan Kuis · {state.subject} · Kelas {state.grade}</span>
          <div className="flex items-center gap-2">
            <button onClick={() => setShowScoreboard(true)} className="kc-btn-ghost px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold kc-font-body flex items-center gap-1.5"><Trophy size={14} style={{ color: "var(--gold)" }} /> Papan Skor</button>
            <button onClick={() => dispatch({ type: "PAUSE_SESSION" })} className="kc-btn-ghost px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold kc-font-body flex items-center gap-1.5"><Save size={14} /> Simpan &amp; Keluar</button>
          </div>
        </div>

        {currentTurnTeam && (
          <div className="kc-card rounded-2xl p-5 mb-5 flex items-center gap-4 kc-pop">
            {(() => { const Icon = GROUP_ICONS[currentTurnTeam.iconIdx]; return <div className="w-14 h-14 rounded-2xl flex items-center justify-center shrink-0" style={{ background: currentTurnTeam.color + "33" }}><Icon size={26} style={{ color: currentTurnTeam.color }} /></div>; })()}
            <div className="flex-1">
              <div className="text-xs font-bold tracking-wide" style={{ color: "var(--gold-soft)" }}>GILIRAN MAJU MEMILIH KOTAK</div>
              <div className="kc-display font-extrabold text-2xl md:text-3xl">{currentTurnTeam.name}</div>
            </div>
            <div className="text-right">
              <div className="kc-display font-bold text-xl">{currentTurnTeam.score}</div>
              <div className="text-[10px]" style={{ color: "var(--slate)" }}>POIN</div>
            </div>
          </div>
        )}

        <div className="grid grid-cols-3 sm:grid-cols-4 md:grid-cols-5 gap-3 mb-6">
          {state.boardTiles.map((t) => {
            const answeredTeam = t.answeredBy != null ? state.teams.find((team) => team.id === t.answeredBy) : null;
            if (t.used) {
              return (
                <div key={t.id} className="kc-card rounded-2xl aspect-square flex flex-col items-center justify-center opacity-50 p-2 text-center">
                  {t.result === "correct" ? <Check size={20} style={{ color: "var(--green)" }} /> : t.result === "wrong" ? <X size={20} style={{ color: "var(--red)" }} /> : <Circle size={16} style={{ color: "var(--slate)" }} />}
                  <span className="text-[10px] font-semibold kc-font-body mt-1 truncate max-w-full">{answeredTeam ? answeredTeam.name : "—"}</span>
                </div>
              );
            }
            return (
              <button key={t.id} onClick={() => dispatch({ type: "PICK_TILE", payload: t.id })} className="kc-card rounded-2xl aspect-square flex items-center justify-center kc-option-btn" style={{ border: "1px solid rgba(242,183,5,0.35)" }}>
                <span className="kc-display font-extrabold text-2xl md:text-3xl" style={{ color: "var(--gold)" }}>{t.points}</span>
              </button>
            );
          })}
        </div>
        <p className="text-xs kc-font-body text-center mb-4" style={{ color: "var(--slate)" }}>{remainingCount} kotak tersisa. Perwakilan {currentTurnTeam?.name} maju dan menyentuh salah satu kotak untuk membuka soal.</p>

        <div className="kc-card rounded-2xl p-4">
          <div className="flex items-center gap-2 mb-3"><Trophy size={14} style={{ color: "var(--gold)" }} /><span className="text-xs font-bold tracking-wide" style={{ color: "var(--gold-soft)" }}>PAPAN SKOR</span></div>
          <ScoreboardList teams={state.teams} />
        </div>
        {showScoreboard && <ScoreboardOverlay teams={state.teams} onClose={() => setShowScoreboard(false)} />}
      </div>
    );
  }

  // TAMPILAN SOAL (kotak sedang dibuka)
  return (
    <div className="max-w-6xl mx-auto px-4 pb-16 pt-6">
      <div className="flex items-center justify-between flex-wrap gap-3 mb-5">
        <div className="flex items-center gap-2 flex-wrap">
          <span className="kc-chip px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold kc-font-body flex items-center gap-1.5"><LayoutGrid size={14} /> Papan Kuis</span>
          <span className="px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold kc-font-body" style={{ background: "rgba(242,183,5,0.15)", color: "var(--gold-soft)" }}>{tile.points} poin</span>
        </div>
        <div className="flex items-center gap-2">
          <button onClick={() => setShowScoreboard(true)} className="kc-btn-ghost px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold kc-font-body flex items-center gap-1.5"><Trophy size={14} style={{ color: "var(--gold)" }} /> Papan Skor</button>
          <button onClick={() => dispatch({ type: "PAUSE_SESSION" })} className="kc-btn-ghost px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold kc-font-body flex items-center gap-1.5"><Save size={14} /> Simpan &amp; Keluar</button>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-[1fr_280px] gap-5">
        <div>
          <div className={`kc-card rounded-2xl p-6 md:p-8 mb-5 flex flex-col items-center text-center ${gameStatus === "answering" ? "kc-urgent" : ""}`}>
            <TimerRing timeLeft={timer.remaining} total={timer.total} running={timer.running && !revealed} />
            <h2 className="kc-display font-bold text-2xl md:text-3xl lg:text-4xl mt-6 leading-snug">{q.question}</h2>
          </div>

          {isChoice && answeringTeam && !revealed && (
            <div className="rounded-xl px-4 py-3 mb-4 flex items-center gap-2 kc-pop" style={{ background: "rgba(47,111,237,0.14)", border: "1px solid rgba(47,111,237,0.4)" }}>
              {(() => { const Icon = GROUP_ICONS[answeringTeam.iconIdx]; return <Icon size={16} style={{ color: answeringTeam.color }} />; })()}
              <span className="text-sm font-bold kc-font-body">{answeringTeam.name}, sentuh langsung pilihan jawabanmu →</span>
            </div>
          )}

          {isChoice && !answeringTeam && !revealed && (<div className="rounded-xl px-4 py-3 mb-4 text-xs md:text-sm font-semibold kc-font-body" style={{ background: "rgba(47,111,237,0.10)", border: "1px solid rgba(47,111,237,0.3)" }}>Sentuh salah satu pilihan jawaban, lalu pilih kelompok yang menjawab.</div>)}
          <QuestionBody q={q} revealed={revealed} large={true} selectable={isChoice} selectedIndex={selectedIndexForReveal} onSelect={handleSelfSelect} />
          {pickIdx != null && !revealed && (<AnswererPickerModal teams={state.teams} optionLabel={answerOptionLabel(q, pickIdx)} onCancel={() => setPickIdx(null)} onPick={(teamId) => { const i = pickIdx; setPickIdx(null); gradeSelection(teamId, i); }} />)}

          {revealed && (
            <div className="kc-card rounded-2xl p-5 mt-5 kc-pop">
              <div className="flex items-center gap-2 mb-2"><ListChecks size={16} style={{ color: "var(--gold)" }} /><span className="text-xs font-bold tracking-wide" style={{ color: "var(--gold-soft)" }}>PEMBAHASAN</span></div>
              <p className="text-sm md:text-base kc-font-body">{q.explanation}</p>
              <ManualScoreAdjust teams={state.teams} dispatch={dispatch} />
              <button onClick={() => dispatch({ type: "BOARD_NEXT" })} className="kc-btn-primary mt-5 px-6 py-3.5 rounded-xl font-bold text-sm md:text-base kc-font-body flex items-center gap-2">{remainingCount > 1 ? "Kembali ke Papan" : "Lihat Hasil Akhir"} <ArrowRight size={18} /></button>
            </div>
          )}
        </div>

        <div className="space-y-4">
          <div className="kc-card rounded-2xl p-4">
            <div className="flex items-center justify-between mb-3">
              <span className="text-xs font-bold tracking-wide" style={{ color: "var(--gold-soft)" }}>KONTROL GURU</span>
              <div className="flex items-center gap-1">
                <button onClick={() => dispatch({ type: timer.running ? "TIMER_PAUSE" : "TIMER_RESUME" })} disabled={revealed} className="w-9 h-9 rounded-lg flex items-center justify-center kc-btn-ghost disabled:opacity-30">{timer.running ? <PauseCircle size={18} /> : <PlayCircle size={18} />}</button>
                <button onClick={() => dispatch({ type: "TIMER_RESET" })} disabled={revealed} className="w-9 h-9 rounded-lg flex items-center justify-center kc-btn-ghost disabled:opacity-30"><RefreshCw size={16} /></button>
              </div>
            </div>
            {isChoice ? (
              <p className="text-[11px] mb-3 kc-font-body" style={{ color: "var(--slate)" }}>Tunjuk perwakilan yang maju agar mereka bisa langsung menyentuh pilihan jawaban (otomatis dinilai) — atau nilai manual dengan Benar/Salah.</p>
            ) : (
              <p className="text-[11px] mb-3 kc-font-body" style={{ color: "var(--slate)" }}>Klik kelompok yang menjawab benar/salah secara lisan.</p>
            )}
            <div className="space-y-2">
              {state.teams.map((g) => { const Icon = GROUP_ICONS[g.iconIdx]; const isAnswerer = state.answeringTeamId === g.id; return (
                <div key={g.id} className={`rounded-xl p-2.5 ${isAnswerer ? "kc-chip-active" : "kc-chip"}`}>
                  {isChoice ? (
                    <button disabled={revealed} onClick={() => toggleAnswerer(g.id)} className="w-full flex items-center gap-2 mb-2 disabled:opacity-40 kc-font-body text-left">
                      <div className="w-7 h-7 rounded-md flex items-center justify-center shrink-0" style={{ background: g.color + "33" }}><Icon size={14} style={{ color: g.color }} /></div>
                      <span className="text-sm font-semibold truncate">{g.name}</span>
                      <span className="text-[10px] font-bold ml-auto shrink-0" style={{ color: isAnswerer ? "var(--blue-soft)" : "var(--slate)" }}>{isAnswerer ? "✓ Menjawab" : "Tunjuk"}</span>
                    </button>
                  ) : (
                    <div className="flex items-center gap-2 mb-2">
                      <div className="w-7 h-7 rounded-md flex items-center justify-center shrink-0" style={{ background: g.color + "33" }}><Icon size={14} style={{ color: g.color }} /></div>
                      <span className="text-sm font-semibold kc-font-body truncate">{g.name}</span>
                    </div>
                  )}
                  <div className="flex gap-1.5">
                    <button disabled={revealed} onClick={() => handleMark(g.id, "correct")} className="flex-1 py-2.5 rounded-lg text-xs font-bold flex items-center justify-center gap-1 disabled:opacity-30 kc-font-body" style={{ background: "rgba(34,197,94,0.16)", color: "var(--green)" }}><Check size={14} /> Benar</button>
                    <button disabled={revealed} onClick={() => handleMark(g.id, "wrong")} className="flex-1 py-2.5 rounded-lg text-xs font-bold flex items-center justify-center gap-1 disabled:opacity-30 kc-font-body" style={{ background: "rgba(240,67,91,0.16)", color: "var(--red)" }}><X size={14} /> Salah</button>
                  </div>
                </div>
              ); })}
            </div>
            <button disabled={revealed} onClick={handleSkip} className="w-full mt-3 py-2.5 rounded-lg text-xs font-semibold kc-btn-ghost disabled:opacity-30 kc-font-body">Tidak Ada yang Menjawab</button>
          </div>

          <div className="kc-card rounded-2xl p-4">
            <div className="flex items-center gap-2 mb-3"><Trophy size={14} style={{ color: "var(--gold)" }} /><span className="text-xs font-bold tracking-wide" style={{ color: "var(--gold-soft)" }}>PAPAN SKOR</span></div>
            <ScoreboardList teams={state.teams} />
          </div>
        </div>
      </div>
      {showScoreboard && <ScoreboardOverlay teams={state.teams} onClose={() => setShowScoreboard(false)} />}
    </div>
  );
}

/* ============================================================
   HASIL PERMAINAN
   ============================================================ */
function ResultsScreen({ state, dispatch, setScreen }) {
  const [showReview, setShowReview] = useState(false);
  const ranked = useMemo(() => updateRanking(state.teams), [state.teams]);
  const podiumOrder = [1, 0, 2].filter((i) => ranked[i]);
  const totalAnswered = state.playedLog.length;

  return (
    <div className="max-w-5xl mx-auto px-4 pb-20 pt-8">
      <div className="text-center mb-8">
        <Trophy size={44} className="mx-auto mb-3" style={{ color: "var(--gold)" }} />
        <h1 className="kc-display font-extrabold text-3xl md:text-4xl mb-1">Hasil Permainan</h1>
        <p className="text-sm kc-font-body" style={{ color: "var(--slate)" }}>{state.gameMode === "tournament" ? "Tournament Mode" : state.gameMode === "practice" ? "Practice Mode" : state.gameMode === "board" ? "Papan Kuis" : "Quick Play"} · {state.subject} · Kelas {state.grade} · {totalAnswered} Soal Dimainkan</p>
      </div>

      {state.gameMode === "tournament" && (
        <div className="kc-card rounded-2xl p-5 mb-6">
          <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>RINGKASAN BABAK</div>
          <div className="grid grid-cols-2 sm:grid-cols-5 gap-2">{ROUND_ORDER.map((key) => (<div key={key} className="kc-chip rounded-lg p-3 text-center"><div className="text-xs font-semibold kc-font-body mb-1">{ROUND_DEFS[key].shortName}</div><RoundStatusBadge status={state.roundStatus[key] || "skipped"} /></div>))}</div>
        </div>
      )}

      <div className="flex items-end justify-center gap-3 md:gap-6 mb-10">
        {podiumOrder.map((idx) => { const g = ranked[idx]; if (!g) return null; const rank = idx + 1; const heights = { 1: "h-36", 2: "h-44", 3: "h-28" }; const medalColor = { 1: "#C0C8D6", 2: "var(--gold)", 3: "#CD7F32" }; const Icon = GROUP_ICONS[g.iconIdx]; return (
          <div key={g.id} className="flex flex-col items-center w-28 md:w-36">
            <Icon size={22} className="mb-1" style={{ color: g.color }} />
            <div className="text-xs font-bold kc-font-body truncate max-w-full mb-1">{g.name}</div>
            <div className="kc-display font-extrabold text-lg mb-2" style={{ color: medalColor[rank] }}>{g.score}</div>
            <div className={`w-full ${heights[rank]} rounded-t-xl kc-card flex items-start justify-center pt-3`} style={{ borderTop: `3px solid ${medalColor[rank]}` }}><Medal size={20} style={{ color: medalColor[rank] }} /></div>
            <div className="w-full py-1 text-center text-xs font-bold kc-font-body" style={{ background: medalColor[rank] + "22", color: medalColor[rank] }}>Juara {rank}</div>
          </div>
        ); })}
      </div>

      <div className="kc-card rounded-2xl p-5 mb-6 overflow-x-auto kc-scroll">
        <table className="w-full text-sm kc-font-body min-w-[560px]">
          <thead><tr className="text-left" style={{ color: "var(--slate)" }}><th className="pb-3 font-semibold text-xs">Peringkat</th><th className="pb-3 font-semibold text-xs">Kelompok</th><th className="pb-3 font-semibold text-xs text-center">Benar</th><th className="pb-3 font-semibold text-xs text-center">Salah</th><th className="pb-3 font-semibold text-xs text-center">Streak Terbaik</th><th className="pb-3 font-semibold text-xs text-center">Persentase</th><th className="pb-3 font-semibold text-xs text-right">Skor</th></tr></thead>
          <tbody>{ranked.map((g, i) => { const Icon = GROUP_ICONS[g.iconIdx]; const answered = g.correct + g.wrong; const pct = calcPercentage(g.correct, answered); return (
            <tr key={g.id} className="border-t" style={{ borderColor: "rgba(255,255,255,0.08)" }}>
              <td className="py-3 font-bold">{i + 1}</td>
              <td className="py-3"><div className="flex items-center gap-2"><div className="w-7 h-7 rounded-lg flex items-center justify-center" style={{ background: g.color + "33" }}><Icon size={14} style={{ color: g.color }} /></div><span className="font-semibold">{g.name}</span>{state.finalists.includes(g.id) && state.gameMode === "tournament" && <span className="text-[10px] px-1.5 py-0.5 rounded" style={{ background: "rgba(242,183,5,0.15)", color: "var(--gold)" }}>Finalis</span>}</div></td>
              <td className="py-3 text-center" style={{ color: "var(--green)" }}>{g.correct}</td>
              <td className="py-3 text-center" style={{ color: "var(--red)" }}>{g.wrong}</td>
              <td className="py-3 text-center">{g.bestStreak}</td>
              <td className="py-3 text-center">{pct}%</td>
              <td className="py-3 text-right kc-display font-bold text-base">{g.score}</td>
            </tr>
          ); })}</tbody>
        </table>
      </div>

      <div className="flex flex-wrap gap-3 justify-center mb-8">
        <button onClick={() => dispatch({ type: "RESTART_TO_SETUP", payload: state.gameMode || "quickplay" })} className="kc-btn-gold px-6 py-3 rounded-xl font-bold text-sm kc-font-body flex items-center gap-2"><RotateCcw size={16} /> Main Lagi</button>
        <button onClick={() => setShowReview((s) => !s)} className="kc-btn-primary px-6 py-3 rounded-xl font-bold text-sm kc-font-body flex items-center gap-2"><ListChecks size={16} /> {showReview ? "Sembunyikan" : "Lihat"} Pembahasan</button>
        <button onClick={() => exportHistoryCsv([{ id: "current", timestamp: Date.now(), mode: state.gameMode, subject: state.subject, grade: state.grade, ranking: ranked.map((t) => ({ name: t.name, score: t.score, correct: t.correct, wrong: t.wrong })) }])} className="kc-btn-ghost px-6 py-3 rounded-xl font-semibold text-sm kc-font-body flex items-center gap-2"><Download size={16} /> Ekspor Hasil (CSV)</button>
        <button onClick={() => setScreen("home")} className="kc-btn-ghost px-6 py-3 rounded-xl font-semibold text-sm kc-font-body flex items-center gap-2"><HomeIcon size={16} /> Kembali ke Beranda</button>
        <button onClick={() => dispatch({ type: "RESET_GAME" })} className="kc-btn-ghost px-6 py-3 rounded-xl font-semibold text-sm kc-font-body flex items-center gap-2" style={{ color: "var(--red)" }}><X size={16} /> Reset Game</button>
      </div>

      {showReview && (
        <div className="space-y-3 kc-pop">
          {state.playedLog.map((entry, i) => { const q = entry.question; const team = entry.teamId != null ? state.teams.find((t) => t.id === entry.teamId) : null; return (
            <div key={i} className="kc-card rounded-xl p-4">
              <div className="flex items-center gap-2 mb-1 flex-wrap">
                <span className="text-xs font-bold" style={{ color: "var(--gold-soft)" }}>Soal {i + 1} · {q.points} pts</span>
                {entry.round && <span className="text-[10px] px-1.5 py-0.5 rounded" style={{ background: "rgba(255,255,255,0.08)", color: "var(--slate)" }}>{ROUND_DEFS[entry.round]?.shortName}</span>}
                <span className="text-[10px] px-1.5 py-0.5 rounded" style={{ background: entry.result === "correct" ? "rgba(34,197,94,0.15)" : entry.result === "wrong" ? "rgba(240,67,91,0.15)" : "rgba(255,255,255,0.08)", color: entry.result === "correct" ? "var(--green)" : entry.result === "wrong" ? "var(--red)" : "var(--slate)" }}>{team ? team.name : "Tidak terjawab"}</span>
              </div>
              <div className="text-sm font-semibold mb-2 kc-font-body">{q.question}</div>
              <div className="mb-2"><QuestionBody q={q} revealed={true} /></div>
              <div className="text-xs kc-font-body" style={{ color: "var(--slate)" }}>{q.explanation}</div>
            </div>
          ); })}
        </div>
      )}
    </div>
  );
}

/* ============================================================
   PANDUAN & PENGATURAN GLOBAL
   ============================================================ */
function GuideScreen({ setScreen }) {
  const quickSteps = [
    { title: "Pilih Quick Play, Papan Kuis, Practice, atau Tournament Mode", desc: "Quick Play: permainan cepat satu sesi. Papan Kuis: perwakilan kelompok bergiliran memilih kotak soal bernilai berbeda. Tournament Mode: kompetisi 5 babak berjenjang." },
    { title: "Atur Mapel, Kelas & Kelompok", desc: "Pilih mata pelajaran, kelas, dan beri nama 4 kelompok. Untuk Papan Kuis, guru juga menentukan jumlah kotak soal. Untuk turnamen, pilih juga babak yang digunakan." },
    { title: "Mainkan di Arena", desc: "Bacakan soal ke seluruh kelas dari layar proyektor/IFP. Timer berjalan otomatis, guru menandai jawaban." },
    { title: "Mode Jawab Otomatis (opsional)", desc: "Untuk soal pilihan ganda/benar-salah: klik 'Tunjuk' pada kelompok yang perwakilannya maju, lalu perwakilan itu langsung menyentuh pilihan jawaban di layar — sistem otomatis menilai benar/salah dan mengubah skor, tanpa guru perlu menekan Benar/Salah manual." },
    { title: "Ikuti Alur Babak", desc: "Qualification → Head to Head → Strategy Arena → Semifinal → Grand Final." },
    { title: "Kelola dari Teacher Studio", desc: "Tambah/edit soal di Bank Soal, atur nama & warna kelompok, atur default game, dan lihat riwayat hasil." },
  ];
  return (
    <div className="max-w-3xl mx-auto px-4 pb-20 pt-8">
      <button onClick={() => setScreen("home")} className="flex items-center gap-1 text-sm mb-4 kc-font-body" style={{ color: "var(--slate)" }}><ChevronLeft size={16} /> Kembali</button>
      <h1 className="kc-display font-extrabold text-3xl mb-2">Panduan Bermain</h1>
      <p className="text-sm mb-8 kc-font-body" style={{ color: "var(--slate)" }}>Kelas Champions SMP dioperasikan guru dari satu layar (laptop + proyektor/IFP) — siswa tidak memerlukan perangkat.</p>
      <div className="space-y-3">{quickSteps.map((s, i) => (<div key={i} className="kc-card rounded-xl p-4 flex gap-4"><div className="w-9 h-9 rounded-lg flex items-center justify-center shrink-0 kc-display font-bold" style={{ background: "rgba(242,183,5,0.15)", color: "var(--gold)" }}>{i + 1}</div><div><div className="font-semibold text-sm kc-font-body mb-0.5">{s.title}</div><div className="text-xs kc-font-body" style={{ color: "var(--slate)" }}>{s.desc}</div></div></div>))}</div>
      <div className="kc-card rounded-xl p-4 mt-6">
        <div className="text-xs font-bold mb-2" style={{ color: "var(--gold-soft)" }}>SISTEM POIN</div>
        <ul className="text-xs kc-font-body space-y-1" style={{ color: "var(--slate)" }}>
          <li>• Mudah = 100 poin · Sedang = 200 poin · Sulit = 300 poin</li>
          <li>• Strategy Arena: Easy 100 · Medium 200 · Hard 300 · Extreme 500 poin</li>
          <li>• Bonus streak +50 poin (3x berturut-turut), +100 poin (5x berturut-turut)</li>
          <li>• Penalti Strategy Arena: aktif/nonaktif di Teacher Studio → Pengaturan Game</li>
          <li>• Jawaban salah (baik dinilai manual maupun dijawab langsung perwakilan) mengurangi penuh poin soal tersebut — bisa dimatikan di Teacher Studio → Pengaturan Game</li>
        </ul>
      </div>
      <div className="kc-card rounded-xl p-4 mt-4">
        <div className="text-xs font-bold mb-2 flex items-center gap-1.5" style={{ color: "var(--gold-soft)" }}><Info size={13} /> TENTANG CP, TP & KURIKULUM</div>
        <p className="text-xs kc-font-body" style={{ color: "var(--slate)" }}>Aplikasi ini mendukung 12 mata pelajaran SMP (IPS, Informatika, IPA, Matematika, Bahasa Indonesia, Bahasa Inggris, Pendidikan Pancasila, Pendidikan Agama, PJOK, Seni Budaya, Prakarya, Bahasa Sunda) plus mapel custom, untuk kelas 7–9. Field Capaian Pembelajaran (CP) dan Tujuan Pembelajaran (TP) pada setiap soal <strong>kosong secara default dan bukan data resmi</strong> — silakan isi sesuai dokumen kurikulum sekolah Anda yang berlaku melalui Teacher Studio → Bank Soal.</p>
      </div>
    </div>
  );
}

function SettingsScreen({ setScreen, state, dispatch }) {
  return (
    <div className="max-w-2xl mx-auto px-4 pb-20 pt-8">
      <button onClick={() => setScreen("home")} className="flex items-center gap-1 text-sm mb-4 kc-font-body" style={{ color: "var(--slate)" }}><ChevronLeft size={16} /> Kembali</button>
      <h1 className="kc-display font-extrabold text-3xl mb-2">Pengaturan</h1>
      <p className="text-sm mb-6 kc-font-body" style={{ color: "var(--slate)" }}>Pengaturan default game (timer, penalti) kini ada di Teacher Studio → Pengaturan Game.</p>
      <div className="kc-card rounded-2xl p-5 flex items-center justify-between mb-4">
        <div><div className="font-semibold text-sm kc-font-body">Efek Suara</div><div className="text-xs kc-font-body" style={{ color: "var(--slate)" }}>Bunyi saat waktu berjalan, jawaban benar/salah, dan waktu habis.</div></div>
        <button onClick={() => dispatch({ type: "TOGGLE_SOUND" })} className="w-14 h-8 rounded-full relative transition-colors" style={{ background: state.soundOn ? "var(--blue)" : "rgba(255,255,255,0.15)" }}><span className="absolute top-1 w-6 h-6 rounded-full bg-white transition-all" style={{ left: state.soundOn ? 30 : 4 }} /></button>
      </div>
      <button onClick={() => { dispatch({ type: "SET_TEACHER_TAB", payload: "pengaturan" }); setScreen("teacher-panel"); }} className="kc-btn-primary w-full py-3 rounded-xl font-semibold text-sm kc-font-body flex items-center justify-center gap-2"><GraduationCap size={16} /> Buka Teacher Studio</button>
    </div>
  );
}

/* ============================================================
   TEACHER STUDIO (PANEL GURU)
   ============================================================ */
const TEACHER_TABS = [
  { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { key: "kelas", label: "Kelas", icon: BarChart3 },
  { key: "bank", label: "Bank Soal", icon: BookOpen },
  { key: "kelompok", label: "Kelompok", icon: Users },
  { key: "pengaturan-game", label: "Pengaturan Game", icon: SlidersHorizontal },
  { key: "riwayat", label: "Riwayat", icon: History },
  { key: "pengaturan", label: "Pengaturan", icon: Settings },
];

function TeacherStudio({ state, dispatch, setScreen }) {
  const tab = state.teacherTab;
  return (
    <div className="max-w-6xl mx-auto px-4 pb-24 pt-8">
      <button onClick={() => setScreen("home")} className="flex items-center gap-1 text-sm mb-4 kc-font-body" style={{ color: "var(--slate)" }}><ChevronLeft size={16} /> Kembali</button>
      <div className="flex items-center gap-3 mb-6">
        <div className="w-11 h-11 rounded-xl kc-card flex items-center justify-center"><GraduationCap size={20} style={{ color: "var(--gold)" }} /></div>
        <div><h1 className="kc-display font-extrabold text-2xl">Teacher Studio</h1><p className="text-xs kc-font-body" style={{ color: "var(--slate)" }}>Panel guru untuk mengelola seluruh aspek permainan</p></div>
      </div>

      <div className="flex gap-1 mb-6 overflow-x-auto kc-scroll">
        {TEACHER_TABS.map((t) => { const Icon = t.icon; return (
          <button key={t.key} onClick={() => dispatch({ type: "SET_TEACHER_TAB", payload: t.key })} className={`flex items-center gap-1.5 px-4 py-2.5 rounded-lg text-xs font-semibold whitespace-nowrap kc-font-body ${tab === t.key ? "kc-chip-active" : "kc-chip"}`}>
            <Icon size={14} /> {t.label}
          </button>
        ); })}
      </div>

      {tab === "dashboard" && <DashboardTab state={state} dispatch={dispatch} />}
      {tab === "kelas" && <KelasTab state={state} dispatch={dispatch} />}
      {tab === "bank" && <BankSoalTab state={state} dispatch={dispatch} />}
      {tab === "kelompok" && <KelompokTab state={state} dispatch={dispatch} />}
      {tab === "pengaturan-game" && <PengaturanGameTab state={state} dispatch={dispatch} />}
      {tab === "riwayat" && <RiwayatTab state={state} dispatch={dispatch} />}
      {tab === "pengaturan" && <PengaturanGuruTab state={state} dispatch={dispatch} />}
    </div>
  );
}

function DashboardTab({ state, dispatch }) {
  const subjects = useMemo(() => getAvailableSubjects(state.questionBank), [state.questionBank]);
  return (
    <div>
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
        <div className="kc-card rounded-2xl p-5"><BookOpen size={18} style={{ color: "var(--gold)" }} className="mb-2" /><div className="kc-display font-extrabold text-2xl">{state.questionBank.length}</div><div className="text-xs kc-font-body" style={{ color: "var(--slate)" }}>Total Soal</div></div>
        <div className="kc-card rounded-2xl p-5"><Users size={18} style={{ color: "var(--blue-soft)" }} className="mb-2" /><div className="kc-display font-extrabold text-2xl">{state.teams.length}</div><div className="text-xs kc-font-body" style={{ color: "var(--slate)" }}>Kelompok Aktif</div></div>
        <div className="kc-card rounded-2xl p-5"><History size={18} style={{ color: "var(--green)" }} className="mb-2" /><div className="kc-display font-extrabold text-2xl">{state.history.length}</div><div className="text-xs kc-font-body" style={{ color: "var(--slate)" }}>Riwayat Permainan</div></div>
        <div className="kc-card rounded-2xl p-5"><DatabaseZap size={18} style={{ color: "var(--gold)" }} className="mb-2" /><div className="kc-display font-extrabold text-2xl">{subjects.length}</div><div className="text-xs kc-font-body" style={{ color: "var(--slate)" }}>Mata Pelajaran</div></div>
      </div>
      <div className="kc-card rounded-2xl p-5 mb-6">
        <div className="text-xs font-bold tracking-wide mb-3" style={{ color: "var(--gold-soft)" }}>SOAL PER MATA PELAJARAN</div>
        <div className="space-y-2">{subjects.map((s) => { const count = state.questionBank.filter((q) => q.subject === s.key).length; return (
          <div key={s.key} className="flex items-center gap-3">
            <span className="text-lg">{s.icon}</span><span className="text-sm font-semibold kc-font-body flex-1">{s.label}</span>
            <div className="flex-1 h-2 rounded-full bg-white/10 overflow-hidden max-w-[200px]"><div className="h-full rounded-full" style={{ width: `${Math.min(100, (count / Math.max(...subjects.map(sj=>state.questionBank.filter(q=>q.subject===sj.key).length),1)) * 100)}%`, background: s.color }} /></div>
            <span className="text-xs kc-display font-bold w-8 text-right">{count}</span>
          </div>
        ); })}</div>
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        <button onClick={() => dispatch({ type: "SET_TEACHER_TAB", payload: "bank" })} className="kc-card rounded-2xl p-5 text-left kc-font-body"><Plus size={18} style={{ color: "var(--gold)" }} className="mb-2" /><div className="text-sm font-semibold">Kelola Bank Soal</div><div className="text-xs" style={{ color: "var(--slate)" }}>Tambah, edit, hapus, dan filter soal.</div></button>
        <button onClick={() => dispatch({ type: "SET_TEACHER_TAB", payload: "riwayat" })} className="kc-card rounded-2xl p-5 text-left kc-font-body"><History size={18} style={{ color: "var(--blue-soft)" }} className="mb-2" /><div className="text-sm font-semibold">Lihat Riwayat</div><div className="text-xs" style={{ color: "var(--slate)" }}>Semua hasil permainan tersimpan otomatis.</div></button>
      </div>
    </div>
  );
}

function FormSoalModal({ initial, subjects, onClose, onSave }) {
  const [form, setForm] = useState(initial || makeEmptyQuestion());
  const [errors, setErrors] = useState([]);
  const isChoice = CHOICE_TYPES.includes(form.type);

  const handleTypeChange = (type) => {
    if (type === "multiple_choice") setForm((f) => ({ ...f, type, options: f.options && f.options.length === 4 ? f.options : ["", "", "", ""], answer: 0 }));
    else if (type === "true_false") setForm((f) => ({ ...f, type, options: ["Benar", "Salah"], answer: 0 }));
    else setForm((f) => ({ ...f, type, options: [], answer: typeof f.answer === "string" ? f.answer : "" }));
  };

  const validate = () => {
    const errs = [];
    if (!form.subject.trim()) errs.push("Mata pelajaran wajib diisi.");
    if (!form.grade) errs.push("Kelas wajib dipilih.");
    if (!form.topic.trim()) errs.push("Materi wajib diisi.");
    if (!form.question.trim()) errs.push("Pertanyaan wajib diisi.");
    if (form.type === "multiple_choice") {
      const filled = form.options.filter((o) => o.trim());
      if (filled.length < 2) errs.push("Pilihan ganda butuh minimal 2 opsi terisi.");
      if (form.answer == null || form.answer >= form.options.length || !form.options[form.answer]?.trim()) errs.push("Pilih kunci jawaban dari opsi yang terisi.");
    } else if (form.type !== "true_false") {
      if (!form.answer || !form.answer.toString().trim()) errs.push("Kunci jawaban wajib diisi.");
    }
    if (!form.points || form.points <= 0) errs.push("Poin harus lebih dari 0.");
    if (!form.timeLimit || form.timeLimit <= 0) errs.push("Waktu harus lebih dari 0.");
    return errs;
  };

  const handleSave = () => {
    const errs = validate();
    if (errs.length) { setErrors(errs); return; }
    const payload = { ...form, id: form.id || genQuestionId(), grade: parseInt(form.grade, 10), points: parseInt(form.points, 10), timeLimit: parseInt(form.timeLimit, 10), imageUrl: form.imageUrl?.trim() || null };
    onSave(payload);
  };

  return (
    <div className="kc-modal-backdrop" onClick={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="kc-card rounded-2xl w-full max-w-2xl mt-6 mb-6 p-6 kc-pop">
        <div className="flex items-center justify-between mb-5">
          <h3 className="kc-display font-bold text-xl">{form.id ? "Edit Soal" : "Tambah Soal Baru"}</h3>
          <button onClick={onClose} className="w-8 h-8 rounded-lg kc-btn-ghost flex items-center justify-center"><X size={16} /></button>
        </div>

        {errors.length > 0 && (
          <div className="rounded-xl p-3 mb-4" style={{ background: "rgba(240,67,91,0.12)", border: "1px solid rgba(240,67,91,0.4)" }}>
            {errors.map((e, i) => (<div key={i} className="text-xs kc-font-body" style={{ color: "var(--red)" }}>• {e}</div>))}
          </div>
        )}

        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3">
          <div>
            <label className="text-[11px] font-semibold block mb-1" style={{ color: "var(--slate)" }}>Mata Pelajaran</label>
            <input list="kc-subject-list" className="kc-input" value={form.subject} onChange={(e) => setForm((f) => ({ ...f, subject: e.target.value }))} placeholder="mis. IPS, Informatika, PJOK..." />
            <datalist id="kc-subject-list">{subjects.map((s) => (<option key={s.key} value={s.key} />))}</datalist>
          </div>
          <div>
            <label className="text-[11px] font-semibold block mb-1" style={{ color: "var(--slate)" }}>Kelas</label>
            <select className="kc-input" value={form.grade} onChange={(e) => setForm((f) => ({ ...f, grade: parseInt(e.target.value, 10) }))}>
              <option value={7}>Kelas 7</option><option value={8}>Kelas 8</option><option value={9}>Kelas 9</option>
            </select>
          </div>
        </div>

        <div className="mb-3">
          <label className="text-[11px] font-semibold block mb-1" style={{ color: "var(--slate)" }}>Materi</label>
          <input className="kc-input" value={form.topic} onChange={(e) => setForm((f) => ({ ...f, topic: e.target.value }))} placeholder="Nama materi/bab" />
        </div>

        <div className="rounded-xl p-3 mb-3 flex items-start gap-2" style={{ background: "rgba(76,134,255,0.08)", border: "1px solid rgba(76,134,255,0.25)" }}>
          <Info size={14} className="mt-0.5 shrink-0" style={{ color: "var(--blue-soft)" }} />
          <p className="text-[11px] kc-font-body" style={{ color: "var(--slate)" }}>CP dan TP di bawah ini <strong>bukan data resmi bawaan</strong> — kosong secara default. Isi/sesuaikan sendiri dengan dokumen Capaian Pembelajaran dan Tujuan Pembelajaran dari kurikulum sekolah Anda yang berlaku.</p>
        </div>

        <div className="mb-3">
          <label className="text-[11px] font-semibold block mb-1" style={{ color: "var(--slate)" }}>Capaian Pembelajaran (CP)</label>
          <textarea className="kc-input" rows={2} value={form.cp || ""} onChange={(e) => setForm((f) => ({ ...f, cp: e.target.value }))} placeholder="Salin/tulis CP sesuai dokumen kurikulum sekolah Anda (opsional)" />
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3">
          <div><label className="text-[11px] font-semibold block mb-1" style={{ color: "var(--slate)" }}>Tujuan Pembelajaran (TP)</label><input className="kc-input" value={form.learningObjective} onChange={(e) => setForm((f) => ({ ...f, learningObjective: e.target.value }))} placeholder="Opsional" /></div>
          <div><label className="text-[11px] font-semibold block mb-1" style={{ color: "var(--slate)" }}>Indikator</label><input className="kc-input" value={form.indicator} onChange={(e) => setForm((f) => ({ ...f, indicator: e.target.value }))} placeholder="Opsional" /></div>
        </div>

        <div className="mb-3">
          <label className="text-[11px] font-semibold block mb-1" style={{ color: "var(--slate)" }}>Pertanyaan</label>
          <textarea className="kc-input" rows={3} value={form.question} onChange={(e) => setForm((f) => ({ ...f, question: e.target.value }))} placeholder="Tulis soal di sini..." />
        </div>

        <div className="mb-3">
          <label className="text-[11px] font-semibold block mb-1" style={{ color: "var(--slate)" }}>Tipe Soal</label>
          <select className="kc-input" value={form.type} onChange={(e) => handleTypeChange(e.target.value)}>
            {QUESTION_TYPES.map((t) => (<option key={t.key} value={t.key}>{t.label}</option>))}
          </select>
        </div>

        {form.type === "multiple_choice" && (
          <div className="mb-3">
            <label className="text-[11px] font-semibold block mb-1" style={{ color: "var(--slate)" }}>Pilihan Jawaban & Kunci (klik radio untuk kunci)</label>
            <div className="space-y-2">
              {[0, 1, 2, 3].map((idx) => (
                <div key={idx} className="flex items-center gap-2">
                  <input type="radio" name="answer-radio" checked={form.answer === idx} onChange={() => setForm((f) => ({ ...f, answer: idx }))} />
                  <input className="kc-input" value={form.options[idx] || ""} onChange={(e) => setForm((f) => { const opts = [...f.options]; opts[idx] = e.target.value; return { ...f, options: opts }; })} placeholder={`Opsi ${String.fromCharCode(65 + idx)}`} />
                </div>
              ))}
            </div>
          </div>
        )}
        {form.type === "true_false" && (
          <div className="mb-3">
            <label className="text-[11px] font-semibold block mb-1" style={{ color: "var(--slate)" }}>Kunci Jawaban</label>
            <div className="flex gap-3">
              <label className="flex items-center gap-2 text-sm kc-font-body"><input type="radio" name="tf-radio" checked={form.answer === 0} onChange={() => setForm((f) => ({ ...f, answer: 0 }))} /> Benar</label>
              <label className="flex items-center gap-2 text-sm kc-font-body"><input type="radio" name="tf-radio" checked={form.answer === 1} onChange={() => setForm((f) => ({ ...f, answer: 1 }))} /> Salah</label>
            </div>
          </div>
        )}
        {!isChoice && (
          <div className="mb-3">
            <label className="text-[11px] font-semibold block mb-1" style={{ color: "var(--slate)" }}>Kunci Jawaban (teks)</label>
            <textarea className="kc-input" rows={2} value={form.answer} onChange={(e) => setForm((f) => ({ ...f, answer: e.target.value }))} placeholder="Tulis kunci jawaban / jawaban benar" />
          </div>
        )}

        <div className="mb-3">
          <label className="text-[11px] font-semibold block mb-1" style={{ color: "var(--slate)" }}>Pembahasan</label>
          <textarea className="kc-input" rows={2} value={form.explanation} onChange={(e) => setForm((f) => ({ ...f, explanation: e.target.value }))} />
        </div>

        <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 mb-3">
          <div><label className="text-[11px] font-semibold block mb-1" style={{ color: "var(--slate)" }}>Poin</label><input type="number" className="kc-input" value={form.points} onChange={(e) => setForm((f) => ({ ...f, points: e.target.value }))} /></div>
          <div><label className="text-[11px] font-semibold block mb-1" style={{ color: "var(--slate)" }}>Waktu (detik)</label><input type="number" className="kc-input" value={form.timeLimit} onChange={(e) => setForm((f) => ({ ...f, timeLimit: e.target.value }))} /></div>
          <div><label className="text-[11px] font-semibold block mb-1" style={{ color: "var(--slate)" }}>Kesulitan</label><select className="kc-input" value={form.difficulty} onChange={(e) => setForm((f) => ({ ...f, difficulty: e.target.value }))}>{DIFFICULTY_OPTIONS.map((d) => (<option key={d.key} value={d.key}>{d.label}</option>))}</select></div>
        </div>

        <div className="mb-5">
          <label className="text-[11px] font-semibold block mb-1" style={{ color: "var(--slate)" }}>URL Gambar (opsional)</label>
          <input className="kc-input" value={form.imageUrl || ""} onChange={(e) => setForm((f) => ({ ...f, imageUrl: e.target.value }))} placeholder="https://..." />
        </div>

        <div className="flex gap-2">
          <button onClick={handleSave} className="kc-btn-gold flex-1 py-3 rounded-xl font-bold text-sm kc-font-body flex items-center justify-center gap-2"><Save size={16} /> Simpan Soal</button>
          <button onClick={onClose} className="kc-btn-ghost px-6 py-3 rounded-xl font-semibold text-sm kc-font-body">Batal</button>
        </div>
      </div>
    </div>
  );
}

function BankSoalTab({ state, dispatch }) {
  const [search, setSearch] = useState("");
  const [fSubject, setFSubject] = useState("");
  const [fGrade, setFGrade] = useState("");
  const [fTopic, setFTopic] = useState("");
  const [fType, setFType] = useState("");
  const [fDifficulty, setFDifficulty] = useState("");
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [pendingDeleteId, setPendingDeleteId] = useState(null);
  const [importErrors, setImportErrors] = useState([]);
  const fileInputRef = useRef(null);
  const [sheetUrl, setSheetUrl] = useState(() => { try { return localStorage.getItem("kc-sheet-url") || ""; } catch (e) { return ""; } });
  const [sheetBusy, setSheetBusy] = useState(false);
  const [sheetMsg, setSheetMsg] = useState(null); // { ok: bool, text: string }
  const syncFromSheet = async () => {
    const url = sheetUrl.trim();
    if (!/^https:\/\//i.test(url)) { setSheetMsg({ ok: false, text: "Isi dulu alamat (URL) Web App Apps Script atau link CSV Google Sheets yang diawali https://" }); return; }
    try { localStorage.setItem("kc-sheet-url", url); } catch (e) { /* abaikan */ }
    setSheetBusy(true); setSheetMsg(null);
    try {
      const res = await fetch(url, { redirect: "follow" });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const { questions, errors } = parseSheetPayload(await res.text());
      if (questions.length) {
        dispatch({ type: "SYNC_SHEET_QUESTIONS", payload: questions });
        setSheetMsg({ ok: true, text: `${questions.length} soal berhasil disinkronkan dari Google Sheets.${errors.length ? ` (${errors.length} baris dilewati, lihat catatan di bawah)` : ""}` });
      } else setSheetMsg({ ok: false, text: "Tidak ada soal valid yang ditemukan di Google Sheets." });
      setImportErrors(errors);
    } catch (e) {
      setSheetMsg({ ok: false, text: navigator.onLine === false ? "Tidak ada koneksi internet. Soal yang sudah pernah disinkronkan tetap tersimpan dan bisa dipakai offline." : `Gagal mengambil data (${e?.message || "kesalahan jaringan"}). Cek alamat URL dan pengaturan akses (harus \"Siapa saja\").` });
    }
    setSheetBusy(false);
  };

  const subjects = useMemo(() => getAvailableSubjects(state.questionBank), [state.questionBank]);
  // Daftar saran mapel di form soal: gabungan 12 mapel baku SMP + mapel custom yang sudah pernah dipakai guru.
  const subjectSuggestions = useMemo(() => {
    const names = [...new Set([...STANDARD_SUBJECT_NAMES, ...state.questionBank.map((q) => q.subject)])].sort();
    return names.map(getSubjectMeta);
  }, [state.questionBank]);
  const topics = useMemo(() => [...new Set(state.questionBank.filter((q) => !fSubject || q.subject === fSubject).map((q) => q.topic))].sort(), [state.questionBank, fSubject]);

  const filtered = useMemo(() => {
    return state.questionBank.filter((q) => {
      if (search && !q.question.toLowerCase().includes(search.toLowerCase()) && !q.topic.toLowerCase().includes(search.toLowerCase())) return false;
      if (fSubject && q.subject !== fSubject) return false;
      if (fGrade && q.grade !== parseInt(fGrade, 10)) return false;
      if (fTopic && q.topic !== fTopic) return false;
      if (fType && q.type !== fType) return false;
      if (fDifficulty && q.difficulty !== fDifficulty) return false;
      return true;
    });
  }, [state.questionBank, search, fSubject, fGrade, fTopic, fType, fDifficulty]);

  const handleSave = (q) => {
    if (state.questionBank.some((x) => x.id === q.id)) dispatch({ type: "UPDATE_QUESTION", payload: q });
    else dispatch({ type: "ADD_QUESTION", payload: q });
    setShowForm(false); setEditing(null);
  };

  const handleImportFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const name = file.name.toLowerCase();
    const isExcel = name.endsWith(".xlsx") || name.endsWith(".xls");
    const isCsv = name.endsWith(".csv");
    if (!isExcel && !isCsv) {
      setImportErrors(["Format file tidak dikenali. Gunakan file .csv, .xlsx, atau .xls."]);
      e.target.value = "";
      return;
    }
    const reader = new FileReader();
    reader.onload = (ev) => {
      const { questions, errors } = isExcel ? parseImportedQuestionsXlsx(ev.target.result) : parseImportedQuestionsCsv(ev.target.result);
      if (questions.length) dispatch({ type: "IMPORT_QUESTIONS", payload: questions });
      setImportErrors(errors.length ? errors : questions.length ? [] : ["Tidak ada soal valid ditemukan dalam file."]);
    };
    reader.onerror = () => setImportErrors(["Gagal membaca file. Coba lagi atau gunakan file lain."]);
    if (isExcel) reader.readAsArrayBuffer(file); else reader.readAsText(file);
    e.target.value = "";
  };

  return (
    <div>
      <div className="kc-card rounded-2xl p-5 mb-4">
        <div className="flex items-center gap-2 mb-4 flex-wrap">
          <div className="relative flex-1 min-w-[200px]">
            <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: "var(--slate)" }} />
            <input className="kc-input pl-9" placeholder="Cari soal atau materi..." value={search} onChange={(e) => setSearch(e.target.value)} />
          </div>
          <button onClick={() => { setEditing(null); setShowForm(true); }} className="kc-btn-gold px-4 py-2.5 rounded-lg text-xs font-bold kc-font-body flex items-center gap-1.5"><Plus size={14} /> Tambah Soal</button>
          <button onClick={() => fileInputRef.current?.click()} className="kc-btn-ghost px-4 py-2.5 rounded-lg text-xs font-semibold kc-font-body flex items-center gap-1.5"><Upload size={14} /> Import Soal (Excel/CSV)</button>
          <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={handleImportFile} />
          <button onClick={() => downloadQuestionTemplateXlsx()} className="kc-btn-ghost px-4 py-2.5 rounded-lg text-xs font-semibold kc-font-body flex items-center gap-1.5"><Download size={14} /> Template Excel</button>
          <button onClick={() => exportQuestionsXlsx(filtered)} className="kc-btn-ghost px-4 py-2.5 rounded-lg text-xs font-semibold kc-font-body flex items-center gap-1.5"><Download size={14} /> Ekspor Excel</button>
          <button onClick={() => exportQuestionsCsv(filtered)} className="kc-btn-ghost px-4 py-2.5 rounded-lg text-xs font-semibold kc-font-body flex items-center gap-1.5"><Download size={14} /> Ekspor CSV</button>
        </div>

        <div className="rounded-xl p-3 mb-4" style={{ background: "rgba(47,111,237,0.08)", border: "1px solid rgba(47,111,237,0.3)" }}>
          <div className="text-xs font-bold mb-2 kc-font-body">Sinkron dari Google Sheets (Apps Script)</div>
          <div className="flex gap-2 flex-wrap">
            <input className="kc-input flex-1 min-w-[220px]" placeholder="https://script.google.com/macros/s/.../exec  (atau link CSV hasil Publikasikan ke web)" value={sheetUrl} onChange={(e) => setSheetUrl(e.target.value)} />
            <button onClick={syncFromSheet} disabled={sheetBusy} className="kc-btn-gold px-4 py-2.5 rounded-lg text-xs font-bold kc-font-body flex items-center gap-1.5">{sheetBusy ? <Loader2 size={14} className="animate-spin" /> : <RefreshCw size={14} />} {sheetBusy ? "Mengambil..." : "Ambil Soal dari Sheets"}</button>
          </div>
          {sheetMsg && <div className="text-[11px] mt-2 kc-font-body font-semibold" style={{ color: sheetMsg.ok ? "var(--green)" : "var(--red)" }}>{sheetMsg.text}</div>}
          <div className="text-[11px] mt-1 kc-font-body" style={{ color: "var(--slate)" }}>Soal dari Sheets bisa disinkron ulang kapan saja: soal yang sama diperbarui, bukan digandakan. Setelah tersinkron, soal tersimpan di perangkat dan tetap bisa dipakai offline.</div>
        </div>

        {importErrors.length > 0 && (
          <div className="rounded-xl p-3 mb-4" style={{ background: "rgba(242,183,5,0.1)", border: "1px solid rgba(242,183,5,0.35)" }}>
            <div className="text-xs font-bold mb-1" style={{ color: "var(--gold-soft)" }}>Catatan Import:</div>
            {importErrors.map((e, i) => (<div key={i} className="text-[11px] kc-font-body" style={{ color: "var(--slate)" }}>• {e}</div>))}
          </div>
        )}

        <div className="grid grid-cols-2 sm:grid-cols-5 gap-2">
          <select className="kc-input" value={fSubject} onChange={(e) => { setFSubject(e.target.value); setFTopic(""); }}><option value="">Semua Mapel</option>{subjects.map((s) => (<option key={s.key} value={s.key}>{s.label}</option>))}</select>
          <select className="kc-input" value={fGrade} onChange={(e) => setFGrade(e.target.value)}><option value="">Semua Kelas</option><option value="7">Kelas 7</option><option value="8">Kelas 8</option><option value="9">Kelas 9</option></select>
          <select className="kc-input" value={fTopic} onChange={(e) => setFTopic(e.target.value)}><option value="">Semua Materi</option>{topics.map((t) => (<option key={t} value={t}>{t}</option>))}</select>
          <select className="kc-input" value={fType} onChange={(e) => setFType(e.target.value)}><option value="">Semua Tipe</option>{QUESTION_TYPES.map((t) => (<option key={t.key} value={t.key}>{t.label}</option>))}</select>
          <select className="kc-input" value={fDifficulty} onChange={(e) => setFDifficulty(e.target.value)}><option value="">Semua Kesulitan</option>{DIFFICULTY_OPTIONS.map((d) => (<option key={d.key} value={d.key}>{d.label}</option>))}</select>
        </div>
        <div className="text-[11px] mt-3" style={{ color: "var(--slate)" }}>
          Menampilkan {filtered.length} dari {state.questionBank.length} soal. Import mendukung file .xlsx, .xls, dan .csv dengan kolom: subject, grade, cp, topic, learningObjective, indicator, question, type, option1-4, answer, explanation, points, timeLimit, difficulty, imageUrl. Belum punya filenya? Unduh "Template Excel" di atas untuk contoh format siap isi.
        </div>
      </div>

      <div className="space-y-2">
        {filtered.map((q) => { const sMeta = getSubjectMeta(q.subject); const diffColor = { easy: "var(--green)", medium: "var(--gold)", hard: "var(--red)" }[q.difficulty]; return (
          <div key={q.id} className="kc-card rounded-xl p-4 flex items-start gap-3">
            <div className="flex-1 min-w-0">
              <div className="flex items-center gap-1.5 flex-wrap mb-1.5">
                <span className="text-[10px] px-2 py-0.5 rounded kc-font-body font-semibold" style={{ background: sMeta.color + "22", color: sMeta.color }}>{sMeta.icon} {q.subject}</span>
                <span className="text-[10px] px-2 py-0.5 rounded kc-chip kc-font-body">Kelas {q.grade}</span>
                <span className="text-[10px] px-2 py-0.5 rounded kc-chip kc-font-body">{q.topic}</span>
                <span className="text-[10px] px-2 py-0.5 rounded kc-chip kc-font-body">{QUESTION_TYPES.find((t) => t.key === q.type)?.label}</span>
                <span className="text-[10px] px-2 py-0.5 rounded font-semibold kc-font-body" style={{ background: diffColor + "22", color: diffColor }}>{DIFFICULTY_OPTIONS.find((d) => d.key === q.difficulty)?.label} · {q.points}pts · {q.timeLimit}s</span>
              </div>
              <div className="text-sm kc-font-body truncate">{q.question}</div>
            </div>
            <div className="flex gap-1.5 shrink-0">
              <button onClick={() => { setEditing(q); setShowForm(true); }} className="w-8 h-8 rounded-lg kc-btn-ghost flex items-center justify-center"><Pencil size={14} /></button>
              <button
                onClick={() => { if (pendingDeleteId === q.id) { dispatch({ type: "DELETE_QUESTION", payload: q.id }); setPendingDeleteId(null); } else setPendingDeleteId(q.id); }}
                onBlur={() => setPendingDeleteId(null)}
                className="w-8 h-8 rounded-lg flex items-center justify-center"
                style={{ background: pendingDeleteId === q.id ? "var(--red)" : "rgba(240,67,91,0.12)", color: pendingDeleteId === q.id ? "#fff" : "var(--red)" }}
                title={pendingDeleteId === q.id ? "Klik lagi untuk konfirmasi hapus" : "Hapus soal"}
              ><Trash2 size={14} /></button>
            </div>
          </div>
        ); })}
        {filtered.length === 0 && <div className="text-center py-10 text-sm kc-font-body" style={{ color: "var(--slate)" }}>Tidak ada soal yang cocok dengan filter.</div>}
      </div>

      <div className="mt-6">
        <button
          onClick={() => { if (pendingDeleteId === "reset-demo") { dispatch({ type: "RESET_QUESTION_BANK_DEMO" }); setPendingDeleteId(null); } else setPendingDeleteId("reset-demo"); }}
          className="kc-btn-ghost px-4 py-2.5 rounded-lg text-xs font-semibold kc-font-body flex items-center gap-1.5"
          style={{ color: pendingDeleteId === "reset-demo" ? "var(--red)" : undefined }}
        >
          <RotateCcw size={14} /> {pendingDeleteId === "reset-demo" ? "Klik lagi untuk konfirmasi reset ke 30 soal demo" : "Reset Bank Soal ke Demo Awal"}
        </button>
      </div>

      {showForm && <FormSoalModal initial={editing} subjects={subjectSuggestions} onClose={() => { setShowForm(false); setEditing(null); }} onSave={handleSave} />}
    </div>
  );
}

function KelasTab({ state, dispatch }) {
  const [newName, setNewName] = useState("");
  const [renamingId, setRenamingId] = useState(null);
  const [renameValue, setRenameValue] = useState("");
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);

  const handleAdd = () => {
    const name = newName.trim();
    if (!name) return;
    dispatch({ type: "ADD_CLASS", payload: { name } });
    setNewName("");
  };
  const startRename = (c) => { setRenamingId(c.id); setRenameValue(c.name); };
  const commitRename = () => {
    if (renameValue.trim()) dispatch({ type: "RENAME_CLASS", payload: { id: renamingId, name: renameValue } });
    setRenamingId(null);
  };

  return (
    <div>
      <div className="kc-card rounded-2xl p-5 mb-4">
        <div className="text-sm font-semibold kc-font-body mb-1">Kelola Kelas</div>
        <div className="text-xs kc-font-body mb-4" style={{ color: "var(--slate)" }}>
          Setiap kelas punya daftar kelompok sendiri. Saat sedang bermain, tekan <b>"Simpan &amp; Keluar"</b> untuk menjeda kuis — nanti bisa dilanjutkan lagi persis dari soal yang sama, kapan saja, dari kelas yang sama.
        </div>
        <div className="flex gap-2 mb-5">
          <input className="kc-input flex-1" placeholder="Nama kelas baru, contoh: 7A, 8B Matematika..." value={newName} onChange={(e) => setNewName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && handleAdd()} />
          <button onClick={handleAdd} disabled={!newName.trim()} className="kc-btn-gold px-4 py-2.5 rounded-lg text-xs font-bold kc-font-body flex items-center gap-1.5 disabled:opacity-40"><Plus size={14} /> Tambah Kelas</button>
        </div>
        <div className="space-y-3">
          {state.classes.map((c) => {
            const isActive = c.id === state.activeClassId;
            const teamCount = isActive ? state.teams.length : c.teams.length;
            const saved = c.savedSession;
            return (
              <div key={c.id} className={`kc-chip rounded-xl p-4 ${isActive ? "kc-chip-active" : ""}`}>
                <div className="flex items-center justify-between flex-wrap gap-2">
                  <div className="flex items-center gap-2 flex-1 min-w-[160px]">
                    {renamingId === c.id ? (
                      <input autoFocus className="kc-input flex-1" value={renameValue} onChange={(e) => setRenameValue(e.target.value)} onBlur={commitRename} onKeyDown={(e) => e.key === "Enter" && commitRename()} />
                    ) : (
                      <>
                        <span className="text-sm font-bold kc-font-body">{c.name}</span>
                        {isActive && <span className="text-[10px] px-2 py-0.5 rounded font-bold" style={{ background: "rgba(242,183,5,0.18)", color: "var(--gold)" }}>AKTIF</span>}
                        <button onClick={() => startRename(c)} className="w-6 h-6 rounded-md kc-btn-ghost flex items-center justify-center shrink-0"><Pencil size={11} /></button>
                      </>
                    )}
                  </div>
                  <div className="flex items-center gap-2 shrink-0">
                    <span className="text-[11px] kc-font-body" style={{ color: "var(--slate)" }}>{teamCount} kelompok</span>
                    {!isActive && <button onClick={() => dispatch({ type: "SET_ACTIVE_CLASS", payload: { id: c.id } })} className="kc-btn-ghost px-3 py-1.5 rounded-lg text-[11px] font-bold kc-font-body">Jadikan Aktif</button>}
                    {state.classes.length > 1 && (
                      confirmDeleteId === c.id ? (
                        <>
                          <button onClick={() => { dispatch({ type: "DELETE_CLASS", payload: { id: c.id } }); setConfirmDeleteId(null); }} className="px-3 py-1.5 rounded-lg text-[11px] font-bold kc-font-body" style={{ background: "rgba(240,67,91,0.18)", color: "var(--red)" }}>Yakin Hapus?</button>
                          <button onClick={() => setConfirmDeleteId(null)} className="kc-btn-ghost px-2 py-1.5 rounded-lg text-[11px] font-bold kc-font-body">Batal</button>
                        </>
                      ) : (
                        <button onClick={() => setConfirmDeleteId(c.id)} className="w-7 h-7 rounded-lg kc-btn-ghost flex items-center justify-center shrink-0"><Trash2 size={12} /></button>
                      )
                    )}
                  </div>
                </div>
                {saved && (
                  <div className="mt-3 rounded-lg p-3 flex items-center justify-between flex-wrap gap-2" style={{ background: "rgba(47,111,237,0.12)", border: "1px solid rgba(47,111,237,0.35)" }}>
                    <div className="text-[11px] kc-font-body flex items-center gap-1" style={{ color: "var(--blue-soft)" }}>
                      <Clock size={11} /> Kuis tersimpan · {GAME_MODE_LABELS[saved.gameMode] || saved.gameMode} · {new Date(saved.savedAt).toLocaleString("id-ID")}
                    </div>
                    <div className="flex items-center gap-2">
                      <button onClick={() => dispatch({ type: "RESUME_SESSION", payload: { classId: c.id } })} className="kc-btn-gold px-3 py-1.5 rounded-lg text-[11px] font-bold kc-font-body flex items-center gap-1"><Play size={11} /> Lanjutkan</button>
                      <button onClick={() => dispatch({ type: "DISCARD_SESSION", payload: { classId: c.id } })} className="kc-btn-ghost px-3 py-1.5 rounded-lg text-[11px] font-bold kc-font-body">Buang</button>
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

function KelompokTab({ state, dispatch }) {
  const count = state.teams.length;
  const setCount = (n) => dispatch({ type: "SET_TEAM_COUNT", payload: n });
  const canRemove = count > MIN_TEAM_COUNT;
  const canAdd = count < MAX_TEAM_COUNT;
  const activeClassName = state.classes.find((c) => c.id === state.activeClassId)?.name || "Kelas Saya";
  return (
    <div>
      <div className="rounded-xl px-4 py-3 mb-4 flex items-center gap-2" style={{ background: "rgba(242,183,5,0.1)", border: "1px solid rgba(242,183,5,0.3)" }}>
        <GraduationCap size={14} style={{ color: "var(--gold)" }} />
        <span className="text-xs font-semibold kc-font-body">Kelompok untuk kelas: {activeClassName}</span>
        <span className="text-[11px] kc-font-body ml-auto" style={{ color: "var(--slate)" }}>Kelola kelas lain di tab "Kelas"</span>
      </div>
      <div className="kc-card rounded-2xl p-5 mb-4">
        <div className="flex items-center justify-between flex-wrap gap-3 mb-1">
          <div>
            <div className="text-sm font-semibold kc-font-body mb-0.5">Jumlah Kelompok</div>
            <div className="text-xs kc-font-body" style={{ color: "var(--slate)" }}>Atur jumlah kelompok sesuai kebutuhan mapel/kelas Anda — mendukung 2 hingga 10 kelompok.</div>
          </div>
          <div className="flex items-center gap-3">
            <button onClick={() => canRemove && setCount(count - 1)} disabled={!canRemove} className="kc-btn-ghost w-10 h-10 rounded-xl flex items-center justify-center font-bold text-lg disabled:opacity-30">−</button>
            <div className="kc-display font-extrabold text-2xl w-10 text-center">{count}</div>
            <button onClick={() => canAdd && setCount(count + 1)} disabled={!canAdd} className="kc-btn-ghost w-10 h-10 rounded-xl flex items-center justify-center font-bold text-lg disabled:opacity-30">+</button>
          </div>
        </div>
        <div className="flex flex-wrap gap-2 mt-3">
          {[2, 4, 6, 8, 10].map((n) => (
            <button key={n} onClick={() => setCount(n)} className={`px-3 py-1.5 rounded-lg text-xs font-bold kc-font-body ${count === n ? "kc-chip-active" : "kc-chip"}`}>{n} kelompok</button>
          ))}
        </div>
        {!canAdd && <div className="text-[11px] mt-2 kc-font-body" style={{ color: "var(--slate)" }}>Maksimal 10 kelompok tercapai.</div>}
        {!canRemove && <div className="text-[11px] mt-2 kc-font-body" style={{ color: "var(--slate)" }}>Minimal 2 kelompok diperlukan untuk bermain.</div>}
      </div>
      <div className="kc-card rounded-2xl p-5 mb-4">
        <div className="text-xs kc-font-body mb-4" style={{ color: "var(--slate)" }}>Nama dan warna kelompok tersimpan otomatis dan akan digunakan di setiap sesi baru. Skor kembali ke 0 setiap kali permainan baru dimulai. Mengurangi jumlah kelompok akan menghapus kelompok paling akhir dalam daftar.</div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          {state.teams.map((t) => { const Icon = GROUP_ICONS[t.iconIdx]; return (
            <div key={t.id} className="kc-chip rounded-xl p-4">
              <div className="flex items-center gap-3 mb-3">
                <div className="w-10 h-10 rounded-lg flex items-center justify-center" style={{ background: t.color + "33" }}><Icon size={18} style={{ color: t.color }} /></div>
                <input className="kc-input flex-1" value={t.name} maxLength={16} onChange={(e) => dispatch({ type: "UPDATE_TEAM_NAME", payload: { id: t.id, name: e.target.value } })} />
              </div>
              <div className="flex gap-2 flex-wrap">{KELOMPOK_PALETTE.map((c) => (<button key={c} onClick={() => dispatch({ type: "UPDATE_TEAM_COLOR", payload: { id: t.id, color: c } })} className="w-6 h-6 rounded-full" style={{ background: c, boxShadow: t.color === c ? "0 0 0 2px var(--navy), 0 0 0 4px " + c : "none" }} />))}</div>
              <div className="text-xs mt-3 kc-font-body" style={{ color: "var(--slate)" }}>Statistik terakhir: {t.score} pts · {t.correct} benar · {t.wrong} salah</div>
            </div>
          ); })}
        </div>
      </div>
      <button onClick={() => dispatch({ type: "RESET_TEAMS_CONFIG" })} className="kc-btn-ghost px-4 py-2.5 rounded-lg text-xs font-semibold kc-font-body flex items-center gap-1.5"><RotateCcw size={14} /> Kembalikan Nama & Warna Default</button>
    </div>
  );
}

function PengaturanGameTab({ state, dispatch }) {
  return (
    <div className="kc-card rounded-2xl p-5 max-w-xl">
      <div className="mb-5">
        <div className="text-xs font-bold tracking-wide mb-2" style={{ color: "var(--gold-soft)" }}>TIMER DEFAULT QUICK PLAY</div>
        <div className="flex gap-2 flex-wrap">{TIMER_OPTIONS.map((t) => (<button key={t} onClick={() => dispatch({ type: "SET_SETTINGS", payload: { defaultTimerSeconds: t } })} className={`px-4 py-2 rounded-lg text-xs font-semibold kc-font-body ${state.settings.defaultTimerSeconds === t ? "kc-chip-active" : "kc-chip"}`}>{t} detik</button>))}</div>
      </div>
      <div className="mb-5">
        <div className="text-xs font-bold tracking-wide mb-2" style={{ color: "var(--gold-soft)" }}>JUMLAH SOAL DEFAULT QUICK PLAY</div>
        <div className="flex gap-2 flex-wrap">{[5, 10].map((n) => (<button key={n} onClick={() => dispatch({ type: "SET_SETTINGS", payload: { quickplayDefaultCount: n } })} className={`px-4 py-2 rounded-lg text-xs font-semibold kc-font-body ${state.settings.quickplayDefaultCount === n ? "kc-chip-active" : "kc-chip"}`}>{n} Soal</button>))}</div>
      </div>
      <div className="mb-5">
        <div className="text-xs font-bold tracking-wide mb-2" style={{ color: "var(--gold-soft)" }}>UKURAN PAPAN KUIS DEFAULT</div>
        <p className="text-xs kc-font-body mb-2" style={{ color: "var(--slate)" }}>Guru tetap bisa mengubah jumlah kotak setiap kali membuka Papan Kuis — ini hanya nilai bawaan.</p>
        <div className="flex gap-2 flex-wrap">{BOARD_SIZE_PRESETS.map((n) => (<button key={n} onClick={() => dispatch({ type: "SET_SETTINGS", payload: { boardDefaultSize: n } })} className={`px-4 py-2 rounded-lg text-xs font-semibold kc-font-body ${state.settings.boardDefaultSize === n ? "kc-chip-active" : "kc-chip"}`}>{n} Kotak</button>))}</div>
      </div>
      <div className="flex items-center justify-between mb-5">
        <div><div className="text-sm font-semibold kc-font-body">Penalti Strategy Arena</div><div className="text-xs kc-font-body" style={{ color: "var(--slate)" }}>Kurangi poin saat jawaban salah pada tingkat Hard/Extreme.</div></div>
        <button onClick={() => dispatch({ type: "SET_SETTINGS", payload: { strategyPenaltyEnabled: !state.settings.strategyPenaltyEnabled } })} className="w-14 h-8 rounded-full relative transition-colors shrink-0" style={{ background: state.settings.strategyPenaltyEnabled ? "var(--blue)" : "rgba(255,255,255,0.15)" }}><span className="absolute top-1 w-6 h-6 rounded-full bg-white transition-all" style={{ left: state.settings.strategyPenaltyEnabled ? 30 : 4 }} /></button>
      </div>
      <div className="flex items-center justify-between">
        <div><div className="text-sm font-semibold kc-font-body">Penalti Jawaban Salah</div><div className="text-xs kc-font-body" style={{ color: "var(--slate)" }}>Berlaku di Quick Play, Practice, Papan Kuis, Qualification, Head to Head, Semifinal & Final — baik saat guru menilai manual (tombol Salah) maupun saat perwakilan menyentuh jawaban sendiri: kurangi penuh poin soal jika jawaban salah. Nonaktifkan agar permainan lebih santai tanpa risiko poin minus.</div></div>
        <button onClick={() => dispatch({ type: "SET_SETTINGS", payload: { selfSelectPenaltyEnabled: !state.settings.selfSelectPenaltyEnabled } })} className="w-14 h-8 rounded-full relative transition-colors shrink-0" style={{ background: state.settings.selfSelectPenaltyEnabled ? "var(--blue)" : "rgba(255,255,255,0.15)" }}><span className="absolute top-1 w-6 h-6 rounded-full bg-white transition-all" style={{ left: state.settings.selfSelectPenaltyEnabled ? 30 : 4 }} /></button>
      </div>
    </div>
  );
}

function RiwayatTab({ state, dispatch }) {
  const [expandedId, setExpandedId] = useState(null);
  const [confirmClear, setConfirmClear] = useState(false);
  return (
    <div>
      <div className="flex items-center gap-2 mb-4">
        <button onClick={() => exportHistoryCsv(state.history)} disabled={!state.history.length} className="kc-btn-ghost px-4 py-2.5 rounded-lg text-xs font-semibold kc-font-body flex items-center gap-1.5 disabled:opacity-40"><Download size={14} /> Ekspor Semua Riwayat (CSV)</button>
        <button onClick={() => { if (confirmClear) { dispatch({ type: "CLEAR_HISTORY" }); setConfirmClear(false); } else setConfirmClear(true); }} disabled={!state.history.length} className="kc-btn-ghost px-4 py-2.5 rounded-lg text-xs font-semibold kc-font-body flex items-center gap-1.5 disabled:opacity-40" style={{ color: confirmClear ? "var(--red)" : undefined }}><Trash2 size={14} /> {confirmClear ? "Klik lagi untuk konfirmasi" : "Hapus Riwayat"}</button>
      </div>
      <div className="space-y-2">
        {state.history.map((h) => { const top = h.ranking[0]; return (
          <div key={h.id} className="kc-card rounded-xl p-4">
            <button className="w-full flex items-center gap-3 text-left" onClick={() => setExpandedId(expandedId === h.id ? null : h.id)}>
              <Trophy size={16} style={{ color: "var(--gold)" }} />
              <div className="flex-1 min-w-0">
                <div className="text-sm font-semibold kc-font-body">{h.mode === "tournament" ? "Tournament" : h.mode === "practice" ? "Practice Mode" : h.mode === "board" ? "Papan Kuis" : "Quick Play"} · {h.subject} · Kelas {h.grade}</div>
                <div className="text-[11px] kc-font-body" style={{ color: "var(--slate)" }}>{new Date(h.timestamp).toLocaleString("id-ID")} · Juara: {top?.name} ({top?.score} pts)</div>
              </div>
              {expandedId === h.id ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
            </button>
            {expandedId === h.id && (
              <div className="mt-3 space-y-1.5">{h.ranking.map((r, i) => (<div key={i} className="flex items-center gap-2 text-xs kc-font-body"><span className="w-4 font-bold" style={{ color: "var(--slate)" }}>{i + 1}</span><span className="flex-1">{r.name}</span><span style={{ color: "var(--green)" }}>{r.correct} benar</span><span style={{ color: "var(--red)" }}>{r.wrong} salah</span><span className="font-bold">{r.score} pts</span></div>))}</div>
            )}
          </div>
        ); })}
        {state.history.length === 0 && <div className="text-center py-10 text-sm kc-font-body" style={{ color: "var(--slate)" }}>Belum ada riwayat permainan.</div>}
      </div>
    </div>
  );
}

function PengaturanGuruTab({ state, dispatch }) {
  const [confirmReset, setConfirmReset] = useState(false);
  return (
    <div className="max-w-xl">
      <div className="kc-card rounded-2xl p-5 mb-4 flex items-center justify-between">
        <div><div className="font-semibold text-sm kc-font-body">Efek Suara</div><div className="text-xs kc-font-body" style={{ color: "var(--slate)" }}>Bunyi timer, jawaban benar/salah, dan waktu habis.</div></div>
        <button onClick={() => dispatch({ type: "TOGGLE_SOUND" })} className="w-14 h-8 rounded-full relative transition-colors" style={{ background: state.soundOn ? "var(--blue)" : "rgba(255,255,255,0.15)" }}><span className="absolute top-1 w-6 h-6 rounded-full bg-white transition-all" style={{ left: state.soundOn ? 30 : 4 }} /></button>
      </div>
      <div className="kc-card rounded-2xl p-5 mb-4">
        <div className="flex items-start gap-2 mb-3"><Info size={15} className="mt-0.5 shrink-0" style={{ color: "var(--blue-soft)" }} /><p className="text-xs kc-font-body" style={{ color: "var(--slate)" }}>Semua data (bank soal, kelompok, pengaturan, riwayat) disimpan otomatis menggunakan penyimpanan persisten milik Claude — pengganti localStorage yang tidak didukung di lingkungan artifact ini. Data bersifat pribadi untuk Anda dan tetap ada meski browser ditutup atau di-refresh.</p></div>
        <div className="text-xs kc-font-body flex items-center gap-2" style={{ color: state.storageReady ? "var(--green)" : "var(--gold)" }}>
          {state.storageReady ? <CheckCircle2 size={13} /> : <Loader2 size={13} className="kc-spin" />} {state.storageReady ? "Tersimpan & tersinkron" : "Memuat data tersimpan..."}
        </div>
      </div>
      <div className="kc-card rounded-2xl p-5">
        <div className="text-sm font-semibold kc-font-body mb-1">Reset Semua Data Demo</div>
        <div className="text-xs kc-font-body mb-3" style={{ color: "var(--slate)" }}>Mengembalikan bank soal ke 30 soal demo awal, nama/warna kelompok ke default, dan menghapus seluruh riwayat. Tindakan ini tidak dapat dibatalkan.</div>
        <button
          onClick={() => { if (confirmReset) { dispatch({ type: "RESET_QUESTION_BANK_DEMO" }); dispatch({ type: "RESET_TEAMS_CONFIG" }); dispatch({ type: "CLEAR_HISTORY" }); setConfirmReset(false); } else setConfirmReset(true); }}
          className="kc-btn-danger px-5 py-2.5 rounded-lg text-xs font-bold kc-font-body flex items-center gap-1.5"
        ><RotateCcw size={14} /> {confirmReset ? "Klik lagi untuk konfirmasi Reset Semua Data" : "Reset Semua Data Demo"}</button>
      </div>
    </div>
  );
}

/* ============================================================
   MAIN APP
   ============================================================ */
export default function KelasChampionsSMP() {
  const [state, dispatch] = useReducer(gameReducer, initialGameState);
  const setScreen = useCallback((s) => dispatch({ type: "SET_SCREEN", payload: s }), []);
  const hasHydratedRef = useRef(false);

  // Hidrasi data dari penyimpanan persisten saat pertama kali dimuat
  useEffect(() => {
    let cancelled = false;
    (async () => {
      const [questionBank, teamsConfig, settings, history, classesData] = await Promise.all([
        loadStorageSlice(STORAGE_KEYS.bank, null),
        loadStorageSlice(STORAGE_KEYS.teams, null),
        loadStorageSlice(STORAGE_KEYS.settings, null),
        loadStorageSlice(STORAGE_KEYS.history, null),
        loadStorageSlice(STORAGE_KEYS.classes, null),
      ]);
      if (cancelled) return;
      dispatch({ type: "HYDRATE", payload: { questionBank, teamsConfig, settings, history, classesData } });
      hasHydratedRef.current = true;
    })();
    return () => { cancelled = true; };
  }, []);

  // Simpan otomatis setiap ada perubahan (setelah hidrasi awal selesai)
  useEffect(() => { if (hasHydratedRef.current) saveStorageSlice(STORAGE_KEYS.bank, state.questionBank); }, [state.questionBank]);
  useEffect(() => { if (hasHydratedRef.current) saveStorageSlice(STORAGE_KEYS.settings, state.settings); }, [state.settings]);
  useEffect(() => { if (hasHydratedRef.current) saveStorageSlice(STORAGE_KEYS.history, state.history); }, [state.history]);
  // Kelas (beserta roster tiap kelas & sesi kuis yang disimpan untuk dilanjutkan) tersimpan sebagai satu kesatuan.
  // Roster kelas yang sedang aktif selalu dibaca langsung dari state.teams (bukan classes[i].teams, yang baru
  // disinkronkan pada titik pindah/pause) supaya setiap suntingan nama/warna langsung ikut tersimpan.
  const classesSaveKey = state.classes.map((c) => `${c.id}:${c.name}:${c.savedSession ? c.savedSession.savedAt : ""}`).join("|") + "::" + state.teams.map((t) => t.name + t.color).join(",");
  useEffect(() => {
    if (!hasHydratedRef.current) return;
    const patched = state.classes.map((c) => (c.id === state.activeClassId ? { ...c, teams: state.teams } : c));
    saveStorageSlice(STORAGE_KEYS.classes, { classes: patched, activeClassId: state.activeClassId });
  }, [classesSaveKey]); // eslint-disable-line react-hooks/exhaustive-deps

  // TIMER ENGINE — satu interval aktif, berhenti saat locked / finished
  useEffect(() => {
    if (!state.timer.running || state.timer.locked || state.gameStatus === "finished") return;
    const id = setInterval(() => dispatch({ type: "TIMER_TICK" }), 1000);
    return () => clearInterval(id);
  }, [state.timer.running, state.timer.locked, state.gameStatus]);

  const prevRemainingRef = useRef(state.timer.remaining);
  useEffect(() => {
    if (state.soundOn && state.timer.running && !state.timer.locked && state.timer.remaining <= 4 && state.timer.remaining < prevRemainingRef.current) playBeep("tick");
    prevRemainingRef.current = state.timer.remaining;
  }, [state.timer.remaining, state.soundOn, state.timer.running, state.timer.locked]);

  const prevStatusRef = useRef(state.gameStatus);
  const popupTimeoutRef = useRef(null);
  const [popup, setPopup] = useState(null); // { kind, teamName, teamColor, gain }
  useEffect(() => {
    if (prevStatusRef.current !== "reviewing" && state.gameStatus === "reviewing") {
      const qId = state.questions[state.currentQuestionIndex]?.id;
      const ans = state.selectedAnswers[qId];
      if (state.soundOn) {
        if (ans?.result === "correct") playBeep("correct");
        else if (ans?.result === "wrong") playBeep("wrong");
        else playBeep("timeup");
      }
      if (popupTimeoutRef.current) clearTimeout(popupTimeoutRef.current);
      if (ans && ans.teamId != null && (ans.result === "correct" || ans.result === "wrong")) {
        const team = state.teams.find((t) => t.id === ans.teamId);
        setPopup({ kind: ans.result, teamName: team?.name || "Kelompok", teamColor: team?.color || "var(--gold)", gain: team?.lastGain ?? 0, nonce: Date.now() });
        popupTimeoutRef.current = setTimeout(() => setPopup(null), 2400);
      } else if (ans && ans.result === "timeup") {
        setPopup({ kind: "timeup", teamName: null, teamColor: "var(--gold)", gain: 0, nonce: Date.now() });
        popupTimeoutRef.current = setTimeout(() => setPopup(null), 1800);
      }
    }
    prevStatusRef.current = state.gameStatus;
  }, [state.gameStatus]);
  useEffect(() => () => { if (popupTimeoutRef.current) clearTimeout(popupTimeoutRef.current); }, []);

  let body;
  switch (state.activeScreen) {
    case "home": body = <HomeScreen state={state} dispatch={dispatch} setScreen={setScreen} />; break;
    case "quickplay-setup": body = <QuickPlaySetup state={state} dispatch={dispatch} setScreen={setScreen} />; break;
    case "practice-setup": body = <QuickPlaySetup state={state} dispatch={dispatch} setScreen={setScreen} practiceMode={true} />; break;
    case "board-setup": body = <BoardSetup state={state} dispatch={dispatch} setScreen={setScreen} />; break;
    case "tournament-setup": body = <TournamentSetup state={state} dispatch={dispatch} setScreen={setScreen} />; break;
    case "round-transition": body = <RoundTransitionScreen state={state} dispatch={dispatch} />; break;
    case "finalist-select": body = <FinalistSelectScreen state={state} dispatch={dispatch} />; break;
    case "arena": body = <ArenaScreen state={state} dispatch={dispatch} />; break;
    case "duel-arena": body = <DuelArena state={state} dispatch={dispatch} />; break;
    case "strategy-arena": body = <StrategyArenaScreen state={state} dispatch={dispatch} />; break;
    case "board-arena": body = <BoardArena state={state} dispatch={dispatch} />; break;
    case "results": body = <ResultsScreen state={state} dispatch={dispatch} setScreen={setScreen} />; break;
    case "guide": body = <GuideScreen setScreen={setScreen} />; break;
    case "settings": body = <SettingsScreen setScreen={setScreen} state={state} dispatch={dispatch} />; break;
    case "teacher-panel": body = <TeacherStudio state={state} dispatch={dispatch} setScreen={setScreen} />; break;
    default: body = <HomeScreen state={state} dispatch={dispatch} setScreen={setScreen} />;
  }

  return (
    <div className="kc-root">
      <GlobalStyle />
      <NavBar state={state} dispatch={dispatch} setScreen={setScreen} />
      {body}
      <FeedbackPopup popup={popup} onClose={() => setPopup(null)} />
    </div>
  );
}
