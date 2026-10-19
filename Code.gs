/**
 * Kelas Champions SMP — Input soal lewat Google Sheets.
 * Cara pasang: lihat APPS_SCRIPT.md
 */
var SHEET_NAME = "Soal";
var HEADERS = ["id","subject","grade","cp","topic","learningObjective","indicator","question","type",
  "option1","option2","option3","option4","answer","explanation","points","timeLimit","difficulty","imageUrl"];
var TYPES = ["multiple_choice","true_false","short_answer"];
var LEVELS = ["easy","medium","hard"];

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Kelas Champions")
    .addItem("1. Buat / reset template soal", "buatTemplate")
    .addItem("2. Periksa soal (cari baris yang salah)", "periksaSoal")
    .addToUi();
}

/** Membuat sheet "Soal" lengkap dengan judul kolom, pilihan dropdown, dan 1 contoh baris. */
function buatTemplate() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (sh && sh.getLastRow() > 1) {
    var ui = SpreadsheetApp.getUi();
    if (ui.alert("Sheet 'Soal' sudah berisi data. Tetap pasang ulang judul kolom & dropdown? (isi soal TIDAK dihapus)", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  }
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold").setBackground("#2f6fed").setFontColor("#ffffff");
  sh.setFrozenRows(1);
  if (sh.getLastRow() < 2) {
    sh.getRange(2, 1, 1, HEADERS.length).setValues([[
      "", "IPA", 7, "", "Tata Surya", "", "", "Planet terdekat dengan Matahari adalah?", "multiple_choice",
      "Merkurius", "Venus", "Bumi", "Mars", 0, "Merkurius adalah planet terdekat dengan Matahari.", 100, 30, "easy", ""]]);
  }
  var n = 1000;
  setDropdown_(sh, "type", TYPES, n);
  setDropdown_(sh, "difficulty", LEVELS, n);
  setDropdown_(sh, "grade", [7, 8, 9], n);
  sh.getRange(1, HEADERS.indexOf("answer") + 1).setNote(
    "Pilihan ganda: nomor jawaban benar, mulai dari 0 (0=option1, 1=option2, 2=option3, 3=option4).\n" +
    "Benar/Salah: 0 = Benar, 1 = Salah.\nIsian singkat: tulis jawabannya.");
  sh.getRange(1, 1).setNote("Boleh dikosongkan. Isi kode unik (mis. ipa-001) jika ingin soal tetap dikenali walau kalimatnya diubah.");
  sh.autoResizeColumns(1, HEADERS.length);
  SpreadsheetApp.getUi().alert("Template siap. Isi soal mulai baris 2.");
}

function setDropdown_(sh, header, values, rows) {
  var col = HEADERS.indexOf(header) + 1;
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(values.map(String), true).setAllowInvalid(false).build();
  sh.getRange(2, col, rows, 1).setDataValidation(rule);
}

/** Membaca semua soal dari sheet menjadi array objek (kunci = judul kolom). */
function bacaSoal_(sheetName) {
  var sh = SpreadsheetApp.getActive().getSheetByName(sheetName || SHEET_NAME);
  if (!sh) throw new Error("Sheet '" + (sheetName || SHEET_NAME) + "' tidak ditemukan. Jalankan menu Kelas Champions > Buat template soal.");
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var head = values[0].map(function (h) { return String(h).trim(); });
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var obj = {}, empty = true;
    for (var c = 0; c < head.length; c++) {
      if (!head[c]) continue;
      var v = values[i][c];
      if (v !== "" && v !== null) empty = false;
      obj[head[c]] = v;
    }
    if (!empty) { obj._row = i + 1; rows.push(obj); }
  }
  return rows;
}

/** Menandai baris yang bermasalah dengan warna merah dan menampilkan ringkasan. */
function periksaSoal() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  var rows = bacaSoal_();
  sh.getRange(2, 1, Math.max(1, sh.getMaxRows() - 1), HEADERS.length).setBackground(null);
  var problems = [];
  rows.forEach(function (r) {
    var msg = [];
    var type = String(r.type || "multiple_choice");
    if (!String(r.question || "").trim()) msg.push("kolom question kosong");
    if (TYPES.indexOf(type) < 0) msg.push("type tidak dikenal");
    if (type === "multiple_choice") {
      var opts = [r.option1, r.option2, r.option3, r.option4].filter(function (o) { return String(o).trim() !== ""; });
      if (opts.length < 2) msg.push("pilihan ganda butuh minimal 2 opsi");
      var a = parseInt(r.answer, 10);
      if (isNaN(a) || a < 0 || a >= opts.length) msg.push("answer harus 0.." + (opts.length - 1));
    } else if (type === "true_false") {
      if (String(r.answer) !== "0" && String(r.answer) !== "1") msg.push("answer harus 0 (Benar) atau 1 (Salah)");
    } else if (String(r.answer).trim() === "") msg.push("jawaban isian kosong");
    if (msg.length) { problems.push("Baris " + r._row + ": " + msg.join(", ")); sh.getRange(r._row, 1, 1, HEADERS.length).setBackground("#fde0e0"); }
  });
  SpreadsheetApp.getUi().alert(problems.length
    ? rows.length + " soal diperiksa, " + problems.length + " bermasalah (diberi warna merah):\n\n" + problems.slice(0, 15).join("\n")
    : "Semua " + rows.length + " soal valid. Siap disinkronkan ke aplikasi!");
}

/** Dipanggil aplikasi Kelas Champions lewat alamat Web App. Hasilnya JSON berisi semua soal. */
function doGet(e) {
  var out;
  try {
    var rows = bacaSoal_(e && e.parameter && e.parameter.sheet);
    rows.forEach(function (r) { delete r._row; });
    out = { rows: rows, updated: new Date().toISOString() };
  } catch (err) {
    out = { error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
