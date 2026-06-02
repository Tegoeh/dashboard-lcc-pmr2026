/**
 * Google Apps Script - Otomasi Tabulasi LCC PMR 2026
 * 
 * Skrip ini dipasang di Google Sheets untuk mengotomatiskan penulisan rumus
 * dan penghitungan nilai LCC PMR tingkat MULA, MADYA, dan WIRA secara real-time.
 * 
 * Fitur:
 * 1. Membuat Menu Kustom "🏆 Tabulasi LCC PMR" di Google Sheets.
 * 2. Mengisi rumus otomatis untuk babak Penyisihan, Semifinal, dan Final di sheet PMR Mula, Madya, dan Wira.
 *    (Kode diletakkan di Column B, Nama Sekolah di Column C).
 * 3. Mengambil KODE (Kolom D) dan NAMA kontingen (Kolom E) secara otomatis di sheet LIVE BEL MULA, MADYA, WIRA.
 *    (Secara cerdas membersihkan Data Validation yang salah terpasang di Kolom D & E agar bebas dari error).
 * 4. Peringatan Otomatis Nilai Sama (Tie Warning Alert) menggunakan Conditional Formatting kustom.
 * 5. Membuat Rekapitulasi Juara & Penghargaan (Sistem Sepertiga Bracket Juara 1, 2, 3) secara otomatis.
 * 
 * Cara Memasang:
 * 1. Buka Google Spreadsheet Utama Anda.
 * 2. Klik Extensions (Ekstensi) -> Apps Script.
 * 3. Hapus kode bawaan, lalu paste (tempel) kode di bawah ini.
 * 4. Klik ikon Save (Disket) atau tekan Ctrl + S.
 * 5. Refresh halaman Google Spreadsheet Anda. Menu "🏆 Tabulasi LCC PMR" akan muncul di toolbar atas!
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🏆 Tabulasi LCC PMR')
      .addItem('🔄 1. Inisialisasi Rumus PMR MULA', 'initMula')
      .addItem('🔄 2. Inisialisasi Rumus PMR MADYA', 'initMadya')
      .addItem('🔄 3. Inisialisasi Rumus PMR WIRA', 'initWira')
      .addItem('🏆 4. Inisialisasi Rekap Juara & Penghargaan', 'initRekap')
      .addSeparator()
      .addItem('⚡ Inisialisasi Semua Rumus', 'initAll')
      .addSeparator()
      .addItem('🔍 Diagnosis & Cek Error Semifinal', 'debugTabulasi')
      .addToUi();
}

/**
 * Helper untuk menerapkan aturan Format Bersyarat Peringatan Nilai Sama (Tie Warning)
 */
function applyTieWarning(sheet, rangeString, formula) {
  var range = sheet.getRange(rangeString);
  var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground("#FEE2E2")
      .setFontColor("#991B1B")
      .setBold(true)
      .setRanges([range])
      .build();
  
  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

/**
 * Inisialisasi rumus pada sheet PMR MULA dan LIVE BEL MULA
 */
function initMula() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPmr = ss.getSheetByName("PMR MULA");
  var sheetLive = ss.getSheetByName("LIVE BEL MULA");
  
  if (!sheetPmr || !sheetLive) {
    SpreadsheetApp.getUi().alert("Sheet 'PMR MULA' atau 'LIVE BEL MULA' tidak ditemukan!");
    return;
  }
  
  // Hapus format bersyarat lama untuk mencegah penumpukan rule
  sheetPmr.setConditionalFormatRules([]);
  sheetLive.setConditionalFormatRules([]);
  
  // Pembersihan sisa-sisa rumus yang salah kolom sebelumnya
  sheetPmr.getRange("B24:H33").clearContent();
  sheetPmr.getRange("B37:F41").clearContent();
  
  // ==========================================
  // A. FORMULA DI SHEET PMR MULA
  // ==========================================
  
  // 1. Babak Penyisihan (Rows 6-19)
  for (var r = 6; r <= 19; r++) {
    sheetPmr.getRange("E" + r).setFormula('=IF(D' + r + '>0, RANK(D' + r + ', $D$6:$D$19), "")');
    sheetPmr.getRange("F" + r).setFormula('=IF(E' + r + '="", "", IF(E' + r + '<=10, "LOLOS SEMIFINAL", ""))');
  }
  
  // 2. Babak Semifinal (Rows 24-33)
  for (var r = 24; r <= 33; r++) {
    sheetPmr.getRange("B" + r).setFormula('=IFERROR(INDEX(QUERY($B$6:$F$19, "select B where F = \'LOLOS SEMIFINAL\' order by E", 0), A' + r + '), "")');
    sheetPmr.getRange("C" + r).setFormula('=IFERROR(VLOOKUP(B' + r + ', $B$6:$C$19, 2, FALSE), "")');
    sheetPmr.getRange("D" + r).setFormula('=IF(COUNTIF(\'LIVE BEL MULA\'!$D$8:$D$12, B' + r + ')>0, 1, IF(COUNTIF(\'LIVE BEL MULA\'!$C$18:$C$22, B' + r + ')>0, 2, ""))');
    sheetPmr.getRange("E" + r).setFormula('=IF(B' + r + '="", "", SUM(IFERROR(FILTER(\'LIVE BEL MULA\'!$U$8:$U$12, \'LIVE BEL MULA\'!$D$8:$D$12 = B' + r + '), 0)) + SUM(IFERROR(FILTER(\'LIVE BEL MULA\'!$U$18:$U$22, \'LIVE BEL MULA\'!$C$18:$C$22 = B' + r + '), 0)))');
    sheetPmr.getRange("F" + r).setFormula('=IF(ISNUMBER(E' + r + '), RANK(E' + r + ', $E$24:$E$33), "")');
    sheetPmr.getRange("G" + r).setFormula('=IF(F' + r + '="", "", IF(F' + r + '<=5, "LOLOS FINAL", ""))');
  }
  
  // 3. Babak Final (Rows 37-41)
  for (var r = 37; r <= 41; r++) {
    sheetPmr.getRange("B" + r).setFormula('=IF(SUM(\'LIVE BEL MULA\'!$Y$28:$Y$32)=0, \'LIVE BEL MULA\'!D' + (r-9) + ', IFERROR(INDEX(SORT(\'LIVE BEL MULA\'!$D$28:$D$32, \'LIVE BEL MULA\'!$Y$28:$Y$32, FALSE), A' + r + '), ""))');
    sheetPmr.getRange("C" + r).setFormula('=IFERROR(VLOOKUP(B' + r + ', $B$6:$C$19, 2, FALSE), "")');
    sheetPmr.getRange("D" + r).setFormula('=SUMIF(\'LIVE BEL MULA\'!$D$28:$D$32, B' + r + ', \'LIVE BEL MULA\'!$Y$28:$Y$32)');
    sheetPmr.getRange("E" + r).setFormula('=IF(ISNUMBER(D' + r + '), RANK(D' + r + ', $D$37:$D$41), "")');
    sheetPmr.getRange("F" + r).setFormula('=IF(E' + r + '="", "", IF(E' + r + '=1, "JUARA 1", IF(E' + r + '=2, "JUARA 2", IF(E' + r + '=3, "JUARA 3", IF(E' + r + '=4, "JUARA HARAPAN 1", "JUARA HARAPAN 2")))))');
  }
  
  // ==========================================
  // B. FORMULA DI SHEET LIVE BEL MULA
  // ==========================================
  
  // Proaktif bersihkan Data Validation & Konten di Kolom D & E (Kelompok 1 & Final) dan C & D (Kelompok 2) agar bebas dari error
  sheetLive.getRange("D8:E12").clearDataValidations().clearContent();
  sheetLive.getRange("C18:D22").clearDataValidations().clearContent();
  sheetLive.getRange("D28:E32").clearDataValidations().clearContent();
  
  // Bersihkan juga validasi data di kolom Rank & Status agar penulisan rumus tidak ditolak oleh Google Sheets
  sheetLive.getRange("W8:W12").clearDataValidations();
  sheetLive.getRange("Z8:Z12").clearDataValidations();
  sheetLive.getRange("W18:W22").clearDataValidations();
  sheetLive.getRange("X18:X22").clearDataValidations();
  sheetLive.getRange("Z28:Z32").clearDataValidations();
  
  // Set header kolom undian semifinal (Bisa diisi manual oleh panitia untuk mengacak grup)
  sheetPmr.getRange("G5").setValue("No. Undi SF");
  sheetPmr.getRange("G5").setFontWeight("bold").setHorizontalAlignment("center");

  // 1. Semifinal Kelompok 1 (Rows 8-12) - Mengambil berdasarkan No. Undi SF (Fallback ke Rank Penyisihan)
  for (var r = 8; r <= 12; r++) {
    var rank = r - 7; // No. Undi 1 s.d 5
    sheetLive.getRange("D" + r).setFormula('=IFERROR(IFERROR(INDEX(\'PMR MULA\'!$B$6:$B$19, MATCH(' + rank + ', \'PMR MULA\'!$G$6:$G$19, 0)), INDEX(\'PMR MULA\'!$B$6:$B$19, MATCH(' + rank + ', \'PMR MULA\'!$E$6:$E$19, 0))), "")');
    sheetLive.getRange("E" + r).setFormula('=IFERROR(VLOOKUP(D' + r + ', \'PMR MULA\'!$B$6:$C$19, 2, FALSE), "")');
    sheetLive.getRange("W" + r).setFormula('=IF(ISNUMBER(U' + r + '), RANK(U' + r + ', $U$8:$U$12), "")');
    sheetLive.getRange("Z" + r).setFormula('=IF(W' + r + '<=3, "Lolos Final", "-")');
  }
  
  // 2. Semifinal Kelompok 2 (Rows 18-22) - Mengambil berdasarkan No. Undi SF (Fallback ke Rank Penyisihan)
  for (var r = 18; r <= 22; r++) {
    var rank = r - 12; // No. Undi 6 s.d 10
    sheetLive.getRange("C" + r).setFormula('=IFERROR(IFERROR(INDEX(\'PMR MULA\'!$B$6:$B$19, MATCH(' + rank + ', \'PMR MULA\'!$G$6:$G$19, 0)), INDEX(\'PMR MULA\'!$B$6:$B$19, MATCH(' + rank + ', \'PMR MULA\'!$E$6:$E$19, 0))), "")');
    sheetLive.getRange("D" + r).setFormula('=IFERROR(VLOOKUP(C' + r + ', \'PMR MULA\'!$B$6:$C$19, 2, FALSE), "")');
    sheetLive.getRange("W" + r).setFormula('=IF(ISNUMBER(U' + r + '), RANK(U' + r + ', $U$18:$U$22), "")');
    sheetLive.getRange("X" + r).setFormula('=IF(W' + r + '<=3, "Lolos Final", "-")');
  }
  
  // 3. Final Kelompok (Rows 28-32) - Mengambil 5 tim yang berstatus "LOLOS FINAL" di Semifinal PMR MULA.
  for (var r = 28; r <= 32; r++) {
    var itemIndex = r - 27; // Index 1 s.d 5
    sheetLive.getRange("D" + r).setFormula('=IFERROR(INDEX(QUERY(\'PMR MULA\'!$B$24:$G$33, "select B where G = \'LOLOS FINAL\'", 0), ' + itemIndex + '), "")');
    sheetLive.getRange("E" + r).setFormula('=IFERROR(VLOOKUP(D' + r + ', \'PMR MULA\'!$B$6:$C$19, 2, FALSE), "")');
    sheetLive.getRange("Z" + r).setFormula('=IF(ISNUMBER(Y' + r + '), RANK(Y' + r + ', $Y$28:$Y$32), "")');
  }
  
  // ==========================================
  // C. ALARM DETEKSI NILAI KEMBAR (TIE WARNING)
  // ==========================================
  applyTieWarning(sheetPmr, "E24:E33", "=AND(E24>0, COUNTIF($E$24:$E$33, E24)>1)");
  applyTieWarning(sheetPmr, "D37:D41", "=AND(D37<>-5250, D37<>0, COUNTIF($D$37:$D$41, D37)>1)");
  
  applyTieWarning(sheetLive, "U8:U12", "=AND(U8>0, COUNTIF($U$8:$U$12, U8)>1)");
  applyTieWarning(sheetLive, "U18:U22", "=AND(U18>0, COUNTIF($U$18:$U$22, U18)>1)");
  applyTieWarning(sheetLive, "Y28:Y32", "=AND(Y28<>-5250, Y28<>0, COUNTIF($Y$28:$Y$32, Y28)>1)");
  
  ss.toast("Rumus PMR MULA & LIVE BEL MULA berhasil diinisialisasi!", "Sukses", 3);
}

/**
 * Inisialisasi rumus pada sheet PMR MADYA dan LIVE BEL MADYA
 */
function initMadya() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPmr = ss.getSheetByName("PMR MADYA");
  var sheetLive = ss.getSheetByName("LIVE BEL MADYA");
  
  if (!sheetPmr || !sheetLive) {
    SpreadsheetApp.getUi().alert("Sheet 'PMR MADYA' atau 'LIVE BEL MADYA' tidak ditemukan!");
    return;
  }
  
  // Hapus format bersyarat lama untuk mencegah penumpukan rule
  sheetPmr.setConditionalFormatRules([]);
  sheetLive.setConditionalFormatRules([]);
  
  // Pembersihan sisa-sisa rumus yang salah kolom sebelumnya
  sheetPmr.getRange("B28:H37").clearContent();
  sheetPmr.getRange("B41:F45").clearContent();
  
  // ==========================================
  // A. FORMULA DI SHEET PMR MADYA
  // ==========================================
  
  // 1. Babak Penyisihan (Rows 6-23)
  for (var r = 6; r <= 23; r++) {
    sheetPmr.getRange("F" + r).setFormula('=IF(E' + r + '>0, RANK(E' + r + ', $E$6:$E$23), "")');
    sheetPmr.getRange("G" + r).setFormula('=IF(F' + r + '="", "", IF(F' + r + '<=10, "LOLOS SEMIFINAL", ""))');
  }
  
  // 2. Babak Semifinal (Rows 28-37)
  for (var r = 28; r <= 37; r++) {
    sheetPmr.getRange("B" + r).setFormula('=IFERROR(INDEX(QUERY($B$6:$G$23, "select B where G = \'LOLOS SEMIFINAL\' order by F", 0), A' + r + '), "")');
    sheetPmr.getRange("C" + r).setFormula('=IFERROR(VLOOKUP(B' + r + ', $B$6:$C$23, 2, FALSE), "")');
    sheetPmr.getRange("D" + r).setFormula('=IF(COUNTIF(\'LIVE BEL MADYA\'!$D$8:$D$12, B' + r + ')>0, 1, IF(COUNTIF(\'LIVE BEL MADYA\'!$C$18:$C$22, B' + r + ')>0, 2, ""))');
    sheetPmr.getRange("E" + r).setFormula('=IF(B' + r + '="", "", SUM(IFERROR(FILTER(\'LIVE BEL MADYA\'!$U$8:$U$12, \'LIVE BEL MADYA\'!$D$8:$D$12 = B' + r + '), 0)) + SUM(IFERROR(FILTER(\'LIVE BEL MADYA\'!$U$18:$U$22, \'LIVE BEL MADYA\'!$C$18:$C$22 = B' + r + '), 0)))');
    sheetPmr.getRange("F" + r).setFormula('=IF(ISNUMBER(E' + r + '), RANK(E' + r + ', $E$28:$E$37), "")');
    sheetPmr.getRange("G" + r).setFormula('=IF(F' + r + '="", "", IF(F' + r + '<=5, "LOLOS FINAL", ""))');
  }
  
  // 3. Babak Final (Rows 41-45)
  for (var r = 41; r <= 45; r++) {
    sheetPmr.getRange("B" + r).setFormula('=IF(SUM(\'LIVE BEL MADYA\'!$Y$28:$Y$32)=0, \'LIVE BEL MADYA\'!D' + (r-13) + ', IFERROR(INDEX(SORT(\'LIVE BEL MADYA\'!$D$28:$D$32, \'LIVE BEL MADYA\'!$Y$28:$Y$32, FALSE), A' + r + '), ""))');
    sheetPmr.getRange("C" + r).setFormula('=IFERROR(VLOOKUP(B' + r + ', $B$6:$C$23, 2, FALSE), "")');
    sheetPmr.getRange("D" + r).setFormula('=SUMIF(\'LIVE BEL MADYA\'!$D$28:$D$32, B' + r + ', \'LIVE BEL MADYA\'!$Y$28:$Y$32)');
    sheetPmr.getRange("E" + r).setFormula('=IF(ISNUMBER(D' + r + '), RANK(D' + r + ', $D$41:$D$45), "")');
    sheetPmr.getRange("F" + r).setFormula('=IF(E' + r + '="", "", IF(E' + r + '=1, "JUARA 1", IF(E' + r + '=2, "JUARA 2", IF(E' + r + '=3, "JUARA 3", IF(E' + r + '=4, "JUARA HARAPAN 1", "JUARA HARAPAN 2")))))');
  }
  
  // ==========================================
  // B. FORMULA DI SHEET LIVE BEL MADYA
  // ==========================================
  
  // Proaktif bersihkan Data Validation & Konten di Kolom D & E (Kelompok 1 & Final) dan C & D (Kelompok 2) agar bebas dari error
  sheetLive.getRange("D8:E12").clearDataValidations().clearContent();
  sheetLive.getRange("C18:D22").clearDataValidations().clearContent();
  sheetLive.getRange("D28:E32").clearDataValidations().clearContent();
  
  // Bersihkan juga validasi data di kolom Rank & Status agar penulisan rumus tidak ditolak oleh Google Sheets
  sheetLive.getRange("W8:W12").clearDataValidations();
  sheetLive.getRange("Z8:Z12").clearDataValidations();
  sheetLive.getRange("W18:W22").clearDataValidations();
  sheetLive.getRange("X18:X22").clearDataValidations();
  sheetLive.getRange("Z28:Z32").clearDataValidations();
  
  // Set header kolom undian semifinal (Bisa diisi manual oleh panitia untuk mengacak grup)
  sheetPmr.getRange("H5").setValue("No. Undi SF");
  sheetPmr.getRange("H5").setFontWeight("bold").setHorizontalAlignment("center");

  // 1. Semifinal Kelompok 1 (Rows 8-12) - Mengambil berdasarkan No. Undi SF (Fallback ke Rank Penyisihan)
  for (var r = 8; r <= 12; r++) {
    var rank = r - 7; // No. Undi 1 s.d 5
    sheetLive.getRange("D" + r).setFormula('=IFERROR(IFERROR(INDEX(\'PMR MADYA\'!$B$6:$B$23, MATCH(' + rank + ', \'PMR MADYA\'!$H$6:$H$23, 0)), INDEX(\'PMR MADYA\'!$B$6:$B$23, MATCH(' + rank + ', \'PMR MADYA\'!$F$6:$F$23, 0))), "")');
    sheetLive.getRange("E" + r).setFormula('=IFERROR(VLOOKUP(D' + r + ', \'PMR MADYA\'!$B$6:$C$23, 2, FALSE), "")');
    sheetLive.getRange("W" + r).setFormula('=IF(ISNUMBER(U' + r + '), RANK(U' + r + ', $U$8:$U$12), "")');
    sheetLive.getRange("Z" + r).setFormula('=IF(W' + r + '<=3, "Lolos Final", "-")');
  }
  
  // 2. Semifinal Kelompok 2 (Rows 18-22) - Mengambil berdasarkan No. Undi SF (Fallback ke Rank Penyisihan)
  for (var r = 18; r <= 22; r++) {
    var rank = r - 12; // No. Undi 6 s.d 10
    sheetLive.getRange("C" + r).setFormula('=IFERROR(IFERROR(INDEX(\'PMR MADYA\'!$B$6:$B$23, MATCH(' + rank + ', \'PMR MADYA\'!$H$6:$H$23, 0)), INDEX(\'PMR MADYA\'!$B$6:$B$23, MATCH(' + rank + ', \'PMR MADYA\'!$F$6:$F$23, 0))), "")');
    sheetLive.getRange("D" + r).setFormula('=IFERROR(VLOOKUP(C' + r + ', \'PMR MADYA\'!$B$6:$C$23, 2, FALSE), "")');
    sheetLive.getRange("W" + r).setFormula('=IF(ISNUMBER(U' + r + '), RANK(U' + r + ', $U$18:$U$22), "")');
    sheetLive.getRange("X" + r).setFormula('=IF(W' + r + '<=3, "Lolos Final", "-")');
  }
  
  // 3. Final Kelompok (Rows 28-32) - Mengambil 5 tim yang berstatus "LOLOS FINAL" di Semifinal PMR MADYA.
  for (var r = 28; r <= 32; r++) {
    var itemIndex = r - 27; // Index 1 s.d 5
    sheetLive.getRange("D" + r).setFormula('=IFERROR(INDEX(QUERY(\'PMR MADYA\'!$B$28:$G$37, "select B where G = \'LOLOS FINAL\'", 0), ' + itemIndex + '), "")');
    sheetLive.getRange("E" + r).setFormula('=IFERROR(VLOOKUP(D' + r + ', \'PMR MADYA\'!$B$6:$C$23, 2, FALSE), "")');
    sheetLive.getRange("Z" + r).setFormula('=IF(ISNUMBER(Y' + r + '), RANK(Y' + r + ', $Y$28:$Y$32), "")');
  }
  
  // ==========================================
  // C. ALARM DETEKSI NILAI KEMBAR (TIE WARNING)
  // ==========================================
  applyTieWarning(sheetPmr, "E28:E37", "=AND(E28>0, COUNTIF($E$28:$E$37, E28)>1)");
  applyTieWarning(sheetPmr, "D41:D45", "=AND(D41<>-5250, D41<>0, COUNTIF($D$41:$D$45, D41)>1)");
  
  applyTieWarning(sheetLive, "U8:U12", "=AND(U8>0, COUNTIF($U$8:$U$12, U8)>1)");
  applyTieWarning(sheetLive, "U18:U22", "=AND(U18>0, COUNTIF($U$18:$U$22, U18)>1)");
  applyTieWarning(sheetLive, "Y28:Y32", "=AND(Y28<>-5250, Y28<>0, COUNTIF($Y$28:$Y$32, Y28)>1)");
  
  ss.toast("Rumus PMR MADYA & LIVE BEL MADYA berhasil diinisialisasi!", "Sukses", 3);
}

/**
 * Inisialisasi rumus pada sheet PMR WIRA dan LIVE BEL WIRA
 */
function initWira() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPmr = ss.getSheetByName("PMR WIRA");
  var sheetLive = ss.getSheetByName("LIVE BEL WIRA");
  
  if (!sheetPmr || !sheetLive) {
    SpreadsheetApp.getUi().alert("Sheet 'PMR WIRA' atau 'LIVE BEL WIRA' tidak ditemukan!");
    return;
  }
  
  // Hapus format bersyarat lama untuk mencegah penumpukan rule
  sheetPmr.setConditionalFormatRules([]);
  sheetLive.setConditionalFormatRules([]);
  
  // Pembersihan sisa-sisa rumus yang salah kolom sebelumnya
  sheetPmr.getRange("B24:H33").clearContent();
  sheetPmr.getRange("B37:F41").clearContent();
  
  // ==========================================
  // A. FORMULA DI SHEET PMR WIRA
  // ==========================================
  
  // 1. Babak Penyisihan (Rows 6-19)
  for (var r = 6; r <= 19; r++) {
    sheetPmr.getRange("F" + r).setFormula('=IF(E' + r + '>0, IF(RANK(E' + r + ', $E$6:$E$19)<=10, "LOLOS SEMIFINAL", ""), "")');
  }
  
  // 2. Babak Semifinal (Rows 24-33)
  for (var r = 24; r <= 33; r++) {
    sheetPmr.getRange("B" + r).setFormula('=IFERROR(INDEX(QUERY($B$6:$F$19, "select B where F = \'LOLOS SEMIFINAL\' order by E desc", 0), A' + r + '), "")');
    sheetPmr.getRange("C" + r).setFormula('=IFERROR(VLOOKUP(B' + r + ', $B$6:$C$19, 2, FALSE), "")');
    sheetPmr.getRange("D" + r).setFormula('=IF(COUNTIF(\'LIVE BEL WIRA\'!$D$8:$D$12, B' + r + ')>0, 1, IF(COUNTIF(\'LIVE BEL WIRA\'!$C$18:$C$22, B' + r + ')>0, 2, ""))');
    sheetPmr.getRange("E" + r).setFormula('=IF(B' + r + '="", "", SUM(IFERROR(FILTER(\'LIVE BEL WIRA\'!$U$8:$U$12, \'LIVE BEL WIRA\'!$D$8:$D$12 = B' + r + '), 0)) + SUM(IFERROR(FILTER(\'LIVE BEL WIRA\'!$U$18:$U$22, \'LIVE BEL WIRA\'!$C$18:$C$22 = B' + r + '), 0)))');
    sheetPmr.getRange("F" + r).setFormula('=IF(ISNUMBER(E' + r + '), RANK(E' + r + ', $E$24:$E$33), "")');
    sheetPmr.getRange("G" + r).setFormula('=IF(F' + r + '="", "", IF(F' + r + '<=5, "LOLOS FINAL", ""))');
  }
  
  // 3. Babak Final (Rows 37-41)
  for (var r = 37; r <= 41; r++) {
    sheetPmr.getRange("B" + r).setFormula('=IF(SUM(\'LIVE BEL WIRA\'!$Y$28:$Y$32)=0, \'LIVE BEL WIRA\'!D' + (r-9) + ', IFERROR(INDEX(SORT(\'LIVE BEL WIRA\'!$D$28:$D$32, \'LIVE BEL WIRA\'!$Y$28:$Y$32, FALSE), A' + r + '), ""))');
    sheetPmr.getRange("C" + r).setFormula('=IFERROR(VLOOKUP(B' + r + ', $B$6:$C$19, 2, FALSE), "")');
    sheetPmr.getRange("D" + r).setFormula('=SUMIF(\'LIVE BEL WIRA\'!$D$28:$D$32, B' + r + ', \'LIVE BEL WIRA\'!$Y$28:$Y$32)');
    sheetPmr.getRange("E" + r).setFormula('=IF(ISNUMBER(D' + r + '), RANK(D' + r + ', $D$37:$D$41), "")');
    sheetPmr.getRange("F" + r).setFormula('=IF(E' + r + '="", "", IF(E' + r + '=1, "JUARA 1", IF(E' + r + '=2, "JUARA 2", IF(E' + r + '=3, "JUARA 3", IF(E' + r + '=4, "JUARA HARAPAN 1", "JUARA HARAPAN 2")))))');
  }
  
  // ==========================================
  // B. FORMULA DI SHEET LIVE BEL WIRA
  // ==========================================
  
  // Proaktif bersihkan Data Validation & Konten di Kolom D & E (Kelompok 1 & Final) dan C & D (Kelompok 2) agar bebas dari error
  sheetLive.getRange("D8:E12").clearDataValidations().clearContent();
  sheetLive.getRange("C18:D22").clearDataValidations().clearContent();
  sheetLive.getRange("D28:E32").clearDataValidations().clearContent();
  
  // Bersihkan juga validasi data di kolom Rank & Status agar penulisan rumus tidak ditolak oleh Google Sheets
  sheetLive.getRange("W8:W12").clearDataValidations();
  sheetLive.getRange("Z8:Z12").clearDataValidations();
  sheetLive.getRange("W18:W22").clearDataValidations();
  sheetLive.getRange("X18:X22").clearDataValidations();
  sheetLive.getRange("Z28:Z32").clearDataValidations();
  
  // Set header kolom undian semifinal (Bisa diisi manual oleh panitia untuk mengacak grup)
  sheetPmr.getRange("G5").setValue("No. Undi SF");
  sheetPmr.getRange("G5").setFontWeight("bold").setHorizontalAlignment("center");

  // 1. Semifinal Kelompok 1 (Rows 8-12) - Mengambil berdasarkan No. Undi SF (Fallback ke Rank Penyisihan)
  for (var r = 8; r <= 12; r++) {
    var rank = r - 7; // No. Undi 1 s.d 5
    sheetLive.getRange("D" + r).setFormula('=IFERROR(IFERROR(INDEX(\'PMR WIRA\'!$B$6:$B$19, MATCH(' + rank + ', \'PMR WIRA\'!$G$6:$G$19, 0)), INDEX(\'PMR WIRA\'!$B$6:$B$19, MATCH(LARGE(\'PMR WIRA\'!$E$6:$E$19, ' + rank + '), \'PMR WIRA\'!$E$6:$E$19, 0))), "")');
    sheetLive.getRange("E" + r).setFormula('=IFERROR(VLOOKUP(D' + r + ', \'PMR WIRA\'!$B$6:$C$19, 2, FALSE), "")');
    sheetLive.getRange("W" + r).setFormula('=IF(ISNUMBER(U' + r + '), RANK(U' + r + ', $U$8:$U$12), "")');
    sheetLive.getRange("Z" + r).setFormula('=IF(W' + r + '<=3, "Lolos Final", "-")');
  }
  
  // 2. Semifinal Kelompok 2 (Rows 18-22) - Mengambil berdasarkan No. Undi SF (Fallback ke Rank Penyisihan)
  for (var r = 18; r <= 22; r++) {
    var rank = r - 12; // No. Undi 6 s.d 10
    sheetLive.getRange("C" + r).setFormula('=IFERROR(IFERROR(INDEX(\'PMR WIRA\'!$B$6:$B$19, MATCH(' + rank + ', \'PMR WIRA\'!$G$6:$G$19, 0)), INDEX(\'PMR WIRA\'!$B$6:$B$19, MATCH(LARGE(\'PMR WIRA\'!$E$6:$E$19, ' + rank + '), \'PMR WIRA\'!$E$6:$E$19, 0))), "")');
    sheetLive.getRange("D" + r).setFormula('=IFERROR(VLOOKUP(C' + r + ', \'PMR WIRA\'!$B$6:$C$19, 2, FALSE), "")');
    sheetLive.getRange("W" + r).setFormula('=IF(ISNUMBER(U' + r + '), RANK(U' + r + ', $U$18:$U$22), "")');
    sheetLive.getRange("X" + r).setFormula('=IF(W' + r + '<=3, "Lolos Final", "-")');
  }
  
  // 3. Final Kelompok (Rows 28-32) - Mengambil 5 tim yang berstatus "LOLOS FINAL" di Semifinal PMR WIRA
  for (var r = 28; r <= 32; r++) {
    var itemIndex = r - 27; // Index 1 s.d 5
    sheetLive.getRange("D" + r).setFormula('=IFERROR(INDEX(QUERY(\'PMR WIRA\'!$B$24:$G$33, "select B where G = \'LOLOS FINAL\'", 0), ' + itemIndex + '), "")');
    sheetLive.getRange("E" + r).setFormula('=IFERROR(VLOOKUP(D' + r + ', \'PMR WIRA\'!$B$6:$C$19, 2, FALSE), "")');
    sheetLive.getRange("Z" + r).setFormula('=IF(ISNUMBER(Y' + r + '), RANK(Y' + r + ', $Y$28:$Y$32), "")');
  }
  
  // ==========================================
  // C. ALARM DETEKSI NILAI KEMBAR (TIE WARNING)
  // ==========================================
  applyTieWarning(sheetPmr, "E24:E33", "=AND(E24>0, COUNTIF($E$24:$E$33, E24)>1)");
  applyTieWarning(sheetPmr, "D37:D41", "=AND(D37<>-5250, D37<>0, COUNTIF($D$37:$D$41, D37)>1)");
  
  applyTieWarning(sheetLive, "U8:U12", "=AND(U8>0, COUNTIF($U$8:$U$12, U8)>1)");
  applyTieWarning(sheetLive, "U18:U22", "=AND(U18>0, COUNTIF($U$18:$U$22, U18)>1)");
  applyTieWarning(sheetLive, "Y28:Y32", "=AND(Y28<>-5250, Y28<>0, COUNTIF($Y$28:$Y$32, Y28)>1)");
  
  ss.toast("Rumus PMR WIRA & LIVE BEL WIRA berhasil diinisialisasi!", "Sukses", 3);
}

/**
 * Inisialisasi rumus pada sheet Rekap Juara & Penghargaan
 */
function initRekap() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rekap Juara & Penghargaan");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Sheet 'Rekap Juara & Penghargaan' tidak ditemukan!");
    return;
  }
  
  // ==========================================================
  // 1. TINGKAT MULA (SD)
  // ==========================================================
  // Podium Mula (Rows 7-11: Juara 1 to 5, Row 12: Juara Harapan 3 yang diambil dari Rank 6 Semifinal)
  sheet.getRange("C7").setFormula("='PMR MULA'!B37");
  sheet.getRange("C8").setFormula("='PMR MULA'!B38");
  sheet.getRange("C9").setFormula("='PMR MULA'!B39");
  sheet.getRange("C10").setFormula("='PMR MULA'!B40");
  sheet.getRange("C11").setFormula("='PMR MULA'!B41");
  sheet.getRange("C12").setFormula("=IFERROR(INDEX('PMR MULA'!$B$24:$B$33, MATCH(6, 'PMR MULA'!$F$24:$F$33, 0)), \"\")");
  
  for (var r = 7; r <= 12; r++) {
    sheet.getRange("D" + r).setFormula('=IFERROR(VLOOKUP(C' + r + ', \'PMR MULA\'!$B$6:$C$19, 2, FALSE), "")');
    sheet.getRange("E" + r).setFormula('=IFERROR(VLOOKUP(C' + r + ', \'PMR MULA\'!$B$37:$D$41, 3, FALSE), IFERROR(VLOOKUP(C' + r + ', \'PMR MULA\'!$B$24:$E$33, 4, FALSE), 0))');
  }
  
  // Apresiasi Mula (Rows 16-29) - Sistem Bracket Juara 1, 2, 3
  for (var r = 16; r <= 29; r++) {
    sheet.getRange("C" + r).setFormula('=IFERROR(VLOOKUP(B' + r + ', \'PMR MULA\'!$B$6:$C$19, 2, FALSE), "")');
    sheet.getRange("D" + r).setFormula('=IFERROR(VLOOKUP(B' + r + ', \'PMR MULA\'!$B$37:$D$41, 3, FALSE), IFERROR(VLOOKUP(B' + r + ', \'PMR MULA\'!$B$24:$E$33, 4, FALSE), VLOOKUP(B' + r + ', \'PMR MULA\'!$B$6:$D$19, 3, FALSE)))');
    sheet.getRange("E" + r).setFormula('=IF(D' + r + '<=0, "", IF(RANK(D' + r + ', $D$16:$D$29) <= CEILING(COUNT($D$16:$D$29)/3), "JUARA 1", IF(RANK(D' + r + ', $D$16:$D$29) <= CEILING(2*COUNT($D$16:$D$29)/3), "JUARA 2", "JUARA 3")))');
  }
  
  // ==========================================================
  // 2. TINGKAT MADYA (SMP)
  // ==========================================================
  // Podium Madya (Rows 36-40: Juara 1 to 5, Row 41: Juara Harapan 3 dari Rank 6 Semifinal)
  sheet.getRange("C36").setFormula("='PMR MADYA'!B41");
  sheet.getRange("C37").setFormula("='PMR MADYA'!B42");
  sheet.getRange("C38").setFormula("='PMR MADYA'!B43");
  sheet.getRange("C39").setFormula("='PMR MADYA'!B44");
  sheet.getRange("C40").setFormula("='PMR MADYA'!B45");
  sheet.getRange("C41").setFormula("=IFERROR(INDEX('PMR MADYA'!$B$28:$B$37, MATCH(6, 'PMR MADYA'!$F$28:$F$37, 0)), \"\")");
  
  for (var r = 36; r <= 41; r++) {
    sheet.getRange("D" + r).setFormula('=IFERROR(VLOOKUP(C' + r + ', \'PMR MADYA\'!$B$6:$C$23, 2, FALSE), "")');
    sheet.getRange("E" + r).setFormula('=IFERROR(VLOOKUP(C' + r + ', \'PMR MADYA\'!$B$41:$D$45, 3, FALSE), IFERROR(VLOOKUP(C' + r + ', \'PMR MADYA\'!$B$28:$E$37, 4, FALSE), 0))');
  }
  
  // Apresiasi Madya (Rows 45-62)
  for (var r = 45; r <= 62; r++) {
    sheet.getRange("C" + r).setFormula('=IFERROR(VLOOKUP(B' + r + ', \'PMR MADYA\'!$B$6:$C$23, 2, FALSE), "")');
    sheet.getRange("D" + r).setFormula('=IFERROR(VLOOKUP(B' + r + ', \'PMR MADYA\'!$B$41:$D$45, 3, FALSE), IFERROR(VLOOKUP(B' + r + ', \'PMR MADYA\'!$B$28:$E$37, 4, FALSE), VLOOKUP(B' + r + ', \'PMR MADYA\'!$B$6:$E$23, 4, FALSE)))');
    sheet.getRange("E" + r).setFormula('=IF(D' + r + '<=0, "", IF(RANK(D' + r + ', $D$45:$D$62) <= CEILING(COUNT($D$45:$D$62)/3), "JUARA 1", IF(RANK(D' + r + ', $D$45:$D$62) <= CEILING(2*COUNT($D$45:$D$62)/3), "JUARA 2", "JUARA 3")))');
  }
  
  // ==========================================================
  // 3. TINGKAT WIRA (SMA)
  // ==========================================================
  // Podium Wira (Rows 69-73: Juara 1 to 5, Row 74: Juara Harapan 3 dari Rank 6 Semifinal)
  sheet.getRange("C69").setFormula("='PMR WIRA'!B37");
  sheet.getRange("C70").setFormula("='PMR WIRA'!B38");
  sheet.getRange("C71").setFormula("='PMR WIRA'!B39");
  sheet.getRange("C72").setFormula("='PMR WIRA'!B40");
  sheet.getRange("C73").setFormula("='PMR WIRA'!B41");
  sheet.getRange("C74").setFormula("=IFERROR(INDEX('PMR WIRA'!$B$24:$B$33, MATCH(6, 'PMR WIRA'!$F$24:$F$33, 0)), \"\")");
  
  for (var r = 69; r <= 74; r++) {
    sheet.getRange("D" + r).setFormula('=IFERROR(VLOOKUP(C' + r + ', \'PMR WIRA\'!$B$6:$C$19, 2, FALSE), "")');
    sheet.getRange("E" + r).setFormula('=IFERROR(VLOOKUP(C' + r + ', \'PMR WIRA\'!$B$37:$D$41, 3, FALSE), IFERROR(VLOOKUP(C' + r + ', \'PMR WIRA\'!$B$24:$E$33, 4, FALSE), 0))');
  }
  
  // Apresiasi Wira (Rows 78-91)
  for (var r = 78; r <= 91; r++) {
    sheet.getRange("C" + r).setFormula('=IFERROR(VLOOKUP(B' + r + ', \'PMR WIRA\'!$B$6:$C$19, 2, FALSE), "")');
    sheet.getRange("D" + r).setFormula('=IFERROR(VLOOKUP(B' + r + ', \'PMR WIRA\'!$B$37:$D$41, 3, FALSE), IFERROR(VLOOKUP(B' + r + ', \'PMR WIRA\'!$B$24:$E$33, 4, FALSE), VLOOKUP(B' + r + ', \'PMR WIRA\'!$B$6:$F$19, 4, FALSE)))');
    sheet.getRange("E" + r).setFormula('=IF(D' + r + '<=0, "", IF(RANK(D' + r + ', $D$78:$D$91) <= CEILING(COUNT($D$78:$D$91)/3), "JUARA 1", IF(RANK(D' + r + ', $D$78:$D$91) <= CEILING(2*COUNT($D$78:$D$91)/3), "JUARA 2", "JUARA 3")))');
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Rumus Rekap Juara & Penghargaan berhasil diinisialisasi!", "Sukses", 3);
}

/**
 * Inisialisasi seluruh rumus di semua Sheet sekaligus
 */
function initAll() {
  initMula();
  initMadya();
  initWira();
  initRekap();
  SpreadsheetApp.getUi().alert("Selamat! Semua rumus kustom berhasil diinisialisasi di seluruh Sheet!");
}

/**
 * Fungsi diagnosis kustom untuk memeriksa persis letak error dan rumus di sheet
 */
function debugTabulasi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("DEBUG_LOG") || ss.insertSheet("DEBUG_LOG");
  logSheet.clear();
  
  var logData = [];
  
  // Helper untuk melakukan padding baris data agar selalu memiliki tepat 8 kolom
  function pushLog(row) {
    while (row.length < 8) {
      row.push("");
    }
    logData.push(row);
  }
  
  pushLog(["=== DEBUG TABULASI LCC PMR ==="]);
  pushLog(["Tanggal/Waktu: " + new Date().toString()]);
  pushLog([]);
  
  // 1. Periksa Sheet PMR MULA Semifinal (Rows 24-33)
  var sheetPmr = ss.getSheetByName("PMR MULA");
  if (sheetPmr) {
    pushLog(["--- DATA PMR MULA (Semifinal Rows 24-33) ---"]);
    pushLog(["Row", "Col A (No)", "Col B (Kode)", "Col C (Nama)", "Col D (Kelompok)", "Col E (Total)", "Col F (Rank)", "Col G (Status)"]);
    for (var r = 24; r <= 33; r++) {
      var rowVals = [];
      rowVals.push("Row " + r);
      rowVals.push(sheetPmr.getRange("A" + r).getValue() + " (" + sheetPmr.getRange("A" + r).getFormula() + ")");
      rowVals.push(sheetPmr.getRange("B" + r).getValue() + " (" + sheetPmr.getRange("B" + r).getFormula() + ")");
      rowVals.push(sheetPmr.getRange("C" + r).getValue() + " (" + sheetPmr.getRange("C" + r).getFormula() + ")");
      rowVals.push(sheetPmr.getRange("D" + r).getValue() + " (" + sheetPmr.getRange("D" + r).getFormula() + ")");
      rowVals.push(sheetPmr.getRange("E" + r).getValue() + " (" + sheetPmr.getRange("E" + r).getFormula() + ")");
      rowVals.push(sheetPmr.getRange("F" + r).getValue() + " (" + sheetPmr.getRange("F" + r).getFormula() + ")");
      rowVals.push(sheetPmr.getRange("G" + r).getValue() + " (" + sheetPmr.getRange("G" + r).getFormula() + ")");
      pushLog(rowVals);
    }
    
    pushLog([]);
    pushLog(["--- DATA PMR MULA (Final Rows 37-41) ---"]);
    pushLog(["Row", "Col A (No)", "Col B (Kode)", "Col C (Nama)", "Col D (Total)", "", "", ""]);
    for (var r = 37; r <= 41; r++) {
      var rowVals = [];
      rowVals.push("Row " + r);
      rowVals.push(sheetPmr.getRange("A" + r).getValue() + " (" + sheetPmr.getRange("A" + r).getFormula() + ")");
      rowVals.push(sheetPmr.getRange("B" + r).getValue() + " (" + sheetPmr.getRange("B" + r).getFormula() + ")");
      rowVals.push(sheetPmr.getRange("C" + r).getValue() + " (" + sheetPmr.getRange("C" + r).getFormula() + ")");
      rowVals.push(sheetPmr.getRange("D" + r).getValue() + " (" + sheetPmr.getRange("D" + r).getFormula() + ")");
      pushLog(rowVals);
    }
  } else {
    pushLog(["Sheet PMR MULA tidak ditemukan!"]);
  }
  pushLog([]);
  
  // 2. Periksa Sheet LIVE BEL MULA Group 1 (Rows 8-12), Group 2 (Rows 18-22) & Final (Rows 28-32)
  var sheetLive = ss.getSheetByName("LIVE BEL MULA");
  if (sheetLive) {
    pushLog(["--- DATA LIVE BEL MULA (Group 1 Rows 8-12) ---"]);
    pushLog(["Row", "Col D (Kode)", "Col E (Nama)", "Col U (Total)", "Col W (Rank SF)", "Col Z (Status)"]);
    for (var r = 8; r <= 12; r++) {
      var rowVals = [];
      rowVals.push("Row " + r);
      rowVals.push(sheetLive.getRange("D" + r).getValue() + " (" + sheetLive.getRange("D" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("E" + r).getValue() + " (" + sheetLive.getRange("E" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("U" + r).getValue() + " (" + sheetLive.getRange("U" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("W" + r).getValue() + " (" + sheetLive.getRange("W" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("Z" + r).getValue() + " (" + sheetLive.getRange("Z" + r).getFormula() + ")");
      pushLog(rowVals);
    }
    pushLog([]);
    
    pushLog(["--- DATA LIVE BEL MULA (Group 2 Rows 18-22) ---"]);
    pushLog(["Row", "Col C (Kode)", "Col D (Nama)", "Col E", "Col U (Total)", "Col W (Rank SF)", "Col X (Status)"]);
    for (var r = 18; r <= 22; r++) {
      var rowVals = [];
      rowVals.push("Row " + r);
      rowVals.push(sheetLive.getRange("C" + r).getValue() + " (" + sheetLive.getRange("C" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("D" + r).getValue() + " (" + sheetLive.getRange("D" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("E" + r).getValue() + " (" + sheetLive.getRange("E" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("U" + r).getValue() + " (" + sheetLive.getRange("U" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("W" + r).getValue() + " (" + sheetLive.getRange("W" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("X" + r).getValue() + " (" + sheetLive.getRange("X" + r).getFormula() + ")");
      pushLog(rowVals);
    }
    pushLog([]);
    pushLog(["--- DATA LIVE BEL MULA (Final Rows 28-32) ---"]);
    pushLog(["Row", "D (Kode)", "E (Nama)", "Col U", "Col V", "Col W", "Col X", "Col Y", "Col Z", "Col AA"]);
    for (var r = 28; r <= 32; r++) {
      var rowVals = [];
      rowVals.push("Row " + r);
      rowVals.push(sheetLive.getRange("D" + r).getValue() + " (" + sheetLive.getRange("D" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("E" + r).getValue() + " (" + sheetLive.getRange("E" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("U" + r).getValue() + " (" + sheetLive.getRange("U" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("V" + r).getValue() + " (" + sheetLive.getRange("V" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("W" + r).getValue() + " (" + sheetLive.getRange("W" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("X" + r).getValue() + " (" + sheetLive.getRange("X" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("Y" + r).getValue() + " (" + sheetLive.getRange("Y" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("Z" + r).getValue() + " (" + sheetLive.getRange("Z" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("AA" + r).getValue() + " (" + sheetLive.getRange("AA" + r).getFormula() + ")");
      pushLog(rowVals);
    }
  } else {
    pushLog(["Sheet LIVE BEL MULA tidak ditemukan!"]);
  }
  
  logSheet.getRange(1, 1, logData.length, 10).setValues(logData);
  logSheet.autoResizeColumns(1, 10);
  ss.setActiveSheet(logSheet);
  
  SpreadsheetApp.getUi().alert("Diagnosis selesai! Silakan periksa sheet 'DEBUG_LOG' yang baru dibuat untuk melihat semua rumus dan nilai secara detail, lalu infokan hasilnya ke saya.");
}

/**
 * API Web Service - Mengekspos data Final secara real-time dalam format JSON
 * Digunakan oleh website dashboard offline/online untuk memantau skor.
 */
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {
    mula: getFinalData(ss.getSheetByName("PMR MULA"), 37, 41),
    madya: getFinalData(ss.getSheetByName("PMR MADYA"), 41, 45),
    wira: getFinalData(ss.getSheetByName("PMR WIRA"), 37, 41),
    timestamp: new Date().toISOString()
  };
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Helper untuk menyusun data finalis dari sheet rekap PMR babak Final
 */
function getFinalData(sheet, startRow, endRow) {
  if (!sheet) return [];
  var data = [];
  for (var r = startRow; r <= endRow; r++) {
    var kode = sheet.getRange("B" + r).getValue();
    var nama = sheet.getRange("C" + r).getValue();
    var total = sheet.getRange("D" + r).getValue();
    var rank = sheet.getRange("E" + r).getValue();
    var status = sheet.getRange("F" + r).getValue();
    
    if (kode) {
      data.push({
        kode: kode,
        nama: nama,
        total: typeof total === 'number' ? total : 0,
        rank: rank,
        status: status
      });
    }
  }
  
  // Urutkan berdasarkan total skor tertinggi secara default
  data.sort(function(a, b) {
    return b.total - a.total;
  });
  
  return data;
}
