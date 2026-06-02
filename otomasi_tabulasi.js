/**
 * SISTEM TABULASI OTOMATIS LCC PMR 2026
 * Dibuat khusus untuk otomatisasi, sinkronisasi, validasi,
 * dan penghitungan nilai LCC PMR tingkat MULA, MADYA, dan WIRA secara real-time.
 * 
 * Fitur Utama:
 * 1. Mencegah error duplikasi / tim ganda di babak Semifinal dan Final.
 * 2. Mengisi rumus otomatis untuk babak Penyisihan, Semifinal, dan Final di sheet PMR Mula, Madya, dan Wira.
 * 3. Mengambil KODE (Kolom D) dan NAMA kontingen (Kolom E) secara otomatis di sheet LIVE BEL MULA, MADYA, WIRA.
 * 4. Menyinkronkan total skor panggung secara dinamis (=SUM(H:V) untuk SF, =SUM(G:Z) untuk Final).
 * 5. Fitur Tie Warning (Deteksi Nilai Seri/Kembar) dengan highlight merah bersyarat.
 * 6. Penjamin Dimensi Sheet Dinamis (ensureSheetDimensions) agar bebas dari error koordinat out of bounds.
 * 7. API Web Service doGet() untuk integrasi Real-time Dashboard Offline/Online.
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🏆 Tabulasi LCC PMR')
      .addItem('⚡ Inisialisasi Semua Rumus', 'initAll')
      .addSeparator()
      .addItem('🔄 1. Inisialisasi Rumus PMR MULA', 'initMula')
      .addItem('🔄 2. Inisialisasi Rumus PMR MADYA', 'initMadya')
      .addItem('🔄 3. Inisialisasi Rumus PMR WIRA', 'initWira')
      .addItem('🔄 4. Inisialisasi Rumus REKAP PODIUM', 'initRekap')
      .addSeparator()
      .addItem('🩺 5. Jalankan Diagnosis Tabulasi', 'debugTabulasi')
      .addToUi();
}

/**
 * Memastikan sheet memiliki dimensi baris dan kolom yang cukup sebelum menulis range.
 * Jika dimensi kurang, baris/kolom baru akan otomatis disisipkan di akhir.
 */
function ensureSheetDimensions(sheet, requiredRows, requiredCols) {
  var maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();
  
  if (maxRows < requiredRows) {
    sheet.insertRowsAfter(maxRows, requiredRows - maxRows);
  }
  if (maxCols < requiredCols) {
    sheet.insertColumnsAfter(maxCols, requiredCols - maxCols);
  }
}

/**
 * Menerapkan Aturan Format Bersyarat untuk Deteksi Nilai Seri (Tie Warning)
 */
function applyTieWarning(sheet, rangeString, formula) {
  var range = sheet.getRange(rangeString);
  var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground("#F4C7C3") // Warna merah pastel elegan
      .setFontColor("#C53929")
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
  
  // Pastikan dimensi sheet cukup untuk mencegah range out of bounds
  ensureSheetDimensions(sheetPmr, 45, 10);
  ensureSheetDimensions(sheetLive, 35, 30); // Butuh minimal 30 kolom (s.d kolom AD)
  
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
    sheetPmr.getRange("B" + r).setFormula('=IFERROR(INDEX(SORT(FILTER($B$6:$B$19, $F$6:$F$19 = "LOLOS SEMIFINAL"), FILTER($E$6:$E$19, $F$6:$F$19 = "LOLOS SEMIFINAL"), TRUE), A' + r + '), "")');
    sheetPmr.getRange("C" + r).setFormula('=IFERROR(VLOOKUP(B' + r + ', $B$6:$C$19, 2, FALSE), "")');
    sheetPmr.getRange("D" + r).setFormula('=IF(COUNTIF(\'LIVE BEL MULA\'!$D$8:$D$12, B' + r + ')>0, 1, IF(COUNTIF(\'LIVE BEL MULA\'!$D$18:$D$22, B' + r + ')>0, 2, ""))');
    sheetPmr.getRange("E" + r).setFormula('=IF(B' + r + '="", "", SUM(IFERROR(FILTER(\'LIVE BEL MULA\'!$W$8:$W$12, \'LIVE BEL MULA\'!$D$8:$D$12 = B' + r + '), 0)) + SUM(IFERROR(FILTER(\'LIVE BEL MULA\'!$W$18:$W$22, \'LIVE BEL MULA\'!$D$18:$D$22 = B' + r + '), 0)))');
    sheetPmr.getRange("F" + r).setFormula('=IF(ISNUMBER(E' + r + '), RANK(E' + r + ', $E$24:$E$33), "")');
    sheetPmr.getRange("G" + r).setFormula('=IF(F' + r + '="", "", IF(F' + r + '<=5, "LOLOS FINAL", ""))');
  }
  
  // 3. Babak Final (Rows 37-41)
  for (var r = 37; r <= 41; r++) {
    sheetPmr.getRange("B" + r).setFormula('=IF(SUM(\'LIVE BEL MULA\'!$AA$28:$AA$32)=0, \'LIVE BEL MULA\'!D28, IFERROR(INDEX(SORT(\'LIVE BEL MULA\'!$D$28:$D$32, \'LIVE BEL MULA\'!$AA$28:$AA$32, FALSE), A' + r + '), ""))');
    sheetPmr.getRange("C" + r).setFormula('=IFERROR(VLOOKUP(B' + r + ', $B$6:$C$19, 2, FALSE), "")');
    sheetPmr.getRange("D" + r).setFormula('=SUMIF(\'LIVE BEL MULA\'!$D$28:$D$32, B' + r + ', \'LIVE BEL MULA\'!$AA$28:$AA$32)');
    sheetPmr.getRange("E" + r).setFormula('=IF(ISNUMBER(D' + r + '), RANK(D' + r + ', $D$37:$D$41), "")');
    sheetPmr.getRange("F" + r).setFormula('=IF(E' + r + '="", "", IF(E' + r + '=1, "JUARA 1", IF(E' + r + '=2, "JUARA 2", IF(E' + r + '=3, "JUARA 3", IF(E' + r + '=4, "JUARA HARAPAN 1", "JUARA HARAPAN 2")))))');
  }
  
  // ==========================================
  // B. FORMULA DI SHEET LIVE BEL MULA
  // ==========================================
  
  // Proaktif bersihkan Data Validation & Konten di Kolom D & E (Kelompok 1, Kelompok 2 & Final) agar bebas dari error
  sheetLive.getRange("D8:E12").clearDataValidations().clearContent();
  sheetLive.getRange("D18:E22").clearDataValidations().clearContent();
  sheetLive.getRange("D28:E32").clearDataValidations().clearContent();
  
  // Bersihkan juga validasi data di kolom Rank & Status agar penulisan rumus tidak ditolak oleh Google Sheets
  sheetLive.getRange("Y8:Y12").clearDataValidations();
  sheetLive.getRange("Z8:Z12").clearDataValidations();
  sheetLive.getRange("Y18:Y22").clearDataValidations();
  sheetLive.getRange("Z18:Z22").clearDataValidations();
  sheetLive.getRange("AB28:AB32").clearDataValidations();
  sheetLive.getRange("AC28:AC32").clearDataValidations();
  
  // Set header kolom undian semifinal & final (Bisa diisi manual oleh panitia untuk mengacak grup)
  sheetPmr.getRange("G5").setValue("No. Undi SF");
  sheetPmr.getRange("G5").setFontWeight("bold").setHorizontalAlignment("center");
  sheetPmr.getRange("H5").setValue("No. Undi Final");
  sheetPmr.getRange("H5").setFontWeight("bold").setHorizontalAlignment("center");
 
  // 1. Semifinal Kelompok 1 (Rows 8-12) - Mengambil berdasarkan No. Undi SF (Fallback ke Peringkat Penyisihan)
  for (var r = 8; r <= 12; r++) {
    var rank = r - 7; // No. Undi SF 1 s.d 5
    sheetLive.getRange("D" + r).setFormula('=IFERROR(INDEX(FILTER(\'PMR MULA\'!$B$6:$B$19, (\'PMR MULA\'!$F$6:$F$19 = "LOLOS SEMIFINAL") * (\'PMR MULA\'!$G$6:$G$19 = ' + rank + ')), 1), IFERROR(INDEX(SORT(FILTER(\'PMR MULA\'!$B$6:$B$19, \'PMR MULA\'!$F$6:$F$19 = "LOLOS SEMIFINAL"), FILTER(\'PMR MULA\'!$E$6:$E$19, \'PMR MULA\'!$F$6:$F$19 = "LOLOS SEMIFINAL"), TRUE), ' + rank + '), ""))');
    sheetLive.getRange("E" + r).setFormula('=IFERROR(VLOOKUP(D' + r + ', \'PMR MULA\'!$B$6:$C$19, 2, FALSE), "")');
    sheetLive.getRange("W" + r).setFormula('=SUM(H' + r + ':V' + r + ')');
    sheetLive.getRange("Y" + r).setFormula('=IF(ISNUMBER(W' + r + '), RANK(W' + r + ', $W$8:$W$12), "")');
    sheetLive.getRange("Z" + r).setFormula('=IF(Y' + r + '<=3, "Lolos Final", "-")');
  }
  
  // 2. Semifinal Kelompok 2 (Rows 18-22) - Mengambil berdasarkan No. Undi SF (Fallback ke Peringkat Penyisihan)
  for (var r = 18; r <= 22; r++) {
    var rank = r - 12; // No. Undi SF 6 s.d 10 (r=18 -> 6, r=22 -> 10)
    sheetLive.getRange("D" + r).setFormula('=IFERROR(INDEX(FILTER(\'PMR MULA\'!$B$6:$B$19, (\'PMR MULA\'!$F$6:$F$19 = "LOLOS SEMIFINAL") * (\'PMR MULA\'!$G$6:$G$19 = ' + rank + ')), 1), IFERROR(INDEX(SORT(FILTER(\'PMR MULA\'!$B$6:$B$19, \'PMR MULA\'!$F$6:$F$19 = "LOLOS SEMIFINAL"), FILTER(\'PMR MULA\'!$E$6:$E$19, \'PMR MULA\'!$F$6:$F$19 = "LOLOS SEMIFINAL"), TRUE), ' + rank + '), ""))');
    sheetLive.getRange("E" + r).setFormula('=IFERROR(VLOOKUP(D' + r + ', \'PMR MULA\'!$B$6:$C$19, 2, FALSE), "")');
    sheetLive.getRange("W" + r).setFormula('=SUM(H' + r + ':V' + r + ')');
    sheetLive.getRange("Y" + r).setFormula('=IF(ISNUMBER(W' + r + '), RANK(W' + r + ', $W$18:$W$22), "")');
    sheetLive.getRange("Z" + r).setFormula('=IF(Y' + r + '<=3, "Lolos Final", "-")');
  }
  
  // 3. Final Kelompok (Rows 28-32) - Mengambil 5 tim yang berstatus "LOLOS FINAL" di Semifinal PMR MULA.
  for (var r = 28; r <= 32; r++) {
    var itemIndex = r - 27; // Index 1 s.d 5
    sheetLive.getRange("D" + r).setFormula('=IFERROR(INDEX(FILTER(\'PMR MULA\'!$B$24:$B$33, (\'PMR MULA\'!$G$24:$G$33 = "LOLOS FINAL") * (IFERROR(VLOOKUP(\'PMR MULA\'!$B$24:$B$33, \'PMR MULA\'!$B$6:$H$19, 7, FALSE), 0) = ' + itemIndex + ')), 1), IFERROR(INDEX(SORT(FILTER(\'PMR MULA\'!$B$24:$B$33, \'PMR MULA\'!$G$24:$G$33 = "LOLOS FINAL"), FILTER(\'PMR MULA\'!$F$24:$F$33, \'PMR MULA\'!$G$24:$G$33 = "LOLOS FINAL"), TRUE), ' + itemIndex + '), ""))');
    sheetLive.getRange("E" + r).setFormula('=IFERROR(VLOOKUP(D' + r + ', \'PMR MULA\'!$B$6:$C$19, 2, FALSE), "")');
    sheetLive.getRange("AA" + r).setFormula('=SUM(G' + r + ':Z' + r + ')');
    sheetLive.getRange("AB" + r).setFormula('=IF(ISNUMBER(AA' + r + '), RANK(AA' + r + ', $AA$28:$AA$32), "")');
    sheetLive.getRange("AC" + r).setFormula('=IF(AB' + r + '="", "", IF(AB' + r + '=1, "JUARA 1", IF(AB' + r + '=2, "JUARA 2", IF(AB' + r + '=3, "JUARA 3", IF(AB' + r + '=4, "JUARA HARAPAN 1", "JUARA HARAPAN 2")))))');
  }
  
  // ==========================================
  // C. ALARM DETEKSI NILAI KEMBAR (TIE WARNING)
  // ==========================================
  applyTieWarning(sheetPmr, "E24:E33", "=AND(E24>0, COUNTIF($E$24:$E$33, E24)>1)");
  applyTieWarning(sheetPmr, "D37:D41", "=AND(D37<>-5250, D37<>0, COUNTIF($D$37:$D$41, D37)>1)");
  
  applyTieWarning(sheetLive, "W8:W12", "=AND(W8>0, COUNTIF($W$8:$W$12, W8)>1)");
  applyTieWarning(sheetLive, "W18:W22", "=AND(W18>0, COUNTIF($W$18:$W$22, W18)>1)");
  applyTieWarning(sheetLive, "AA28:AA32", "=AND(AA28<>-5250, AA28<>0, COUNTIF($AA$28:$AA$32, AA28)>1)");
  
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
  
  // Pastikan dimensi sheet cukup untuk mencegah range out of bounds
  ensureSheetDimensions(sheetPmr, 50, 10);
  ensureSheetDimensions(sheetLive, 35, 30); // Butuh minimal 30 kolom (s.d kolom AD)
  
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
    sheetPmr.getRange("B" + r).setFormula('=IFERROR(INDEX(SORT(FILTER($B$6:$B$23, $G$6:$G$23 = "LOLOS SEMIFINAL"), FILTER($F$6:$F$23, $G$6:$G$23 = "LOLOS SEMIFINAL"), TRUE), A' + r + '), "")');
    sheetPmr.getRange("C" + r).setFormula('=IFERROR(VLOOKUP(B' + r + ', $B$6:$C$23, 2, FALSE), "")');
    sheetPmr.getRange("D" + r).setFormula('=IF(COUNTIF(\'LIVE BEL MADYA\'!$C$8:$C$12, B' + r + ')>0, 1, IF(COUNTIF(\'LIVE BEL MADYA\'!$C$18:$C$22, B' + r + ')>0, 2, ""))');
    sheetPmr.getRange("E" + r).setFormula('=IF(B' + r + '="", "", SUM(IFERROR(FILTER(\'LIVE BEL MADYA\'!$U$8:$U$12, \'LIVE BEL MADYA\'!$C$8:$C$12 = B' + r + '), 0)) + SUM(IFERROR(FILTER(\'LIVE BEL MADYA\'!$U$18:$U$22, \'LIVE BEL MADYA\'!$C$18:$C$22 = B' + r + '), 0)))');
    sheetPmr.getRange("F" + r).setFormula('=IF(ISNUMBER(E' + r + '), RANK(E' + r + ', $E$28:$E$37), "")');
    sheetPmr.getRange("G" + r).setFormula('=IF(F' + r + '="", "", IF(F' + r + '<=5, "LOLOS FINAL", ""))');
  }
  
  // 3. Babak Final (Rows 41-45)
  for (var r = 41; r <= 45; r++) {
    sheetPmr.getRange("B" + r).setFormula('=IF(SUM(\'LIVE BEL MADYA\'!$Y$28:$Y$32)=0, \'LIVE BEL MADYA\'!C28, IFERROR(INDEX(SORT(\'LIVE BEL MADYA\'!$C$28:$C$32, \'LIVE BEL MADYA\'!$Y$28:$Y$32, FALSE), A' + r + '), ""))');
    sheetPmr.getRange("C" + r).setFormula('=IFERROR(VLOOKUP(B' + r + ', $B$6:$C$23, 2, FALSE), "")');
    sheetPmr.getRange("D" + r).setFormula('=SUMIF(\'LIVE BEL MADYA\'!$C$28:$C$32, B' + r + ', \'LIVE BEL MADYA\'!$Y$28:$Y$32)');
    sheetPmr.getRange("E" + r).setFormula('=IF(ISNUMBER(D' + r + '), RANK(D' + r + ', $D$41:$D$45), "")');
    sheetPmr.getRange("F" + r).setFormula('=IF(E' + r + '="", "", IF(E' + r + '=1, "JUARA 1", IF(E' + r + '=2, "JUARA 2", IF(E' + r + '=3, "JUARA 3", IF(E' + r + '=4, "JUARA HARAPAN 1", "JUARA HARAPAN 2")))))');
  }
  
  // ==========================================
  // B. FORMULA DI SHEET LIVE BEL MADYA
  // ==========================================
  
  // Proaktif bersihkan Data Validation & Konten di Kolom C & D (Kelompok 1, Kelompok 2 & Final) agar bebas dari error
  sheetLive.getRange("C8:D12").clearDataValidations().clearContent();
  sheetLive.getRange("C18:D22").clearDataValidations().clearContent();
  sheetLive.getRange("C28:D32").clearDataValidations().clearContent();
  
  // Bersihkan juga validasi data di kolom Rank & Status agar penulisan rumus tidak ditolak oleh Google Sheets
  sheetLive.getRange("W8:W12").clearDataValidations();
  sheetLive.getRange("X8:X12").clearDataValidations();
  sheetLive.getRange("W18:W22").clearDataValidations();
  sheetLive.getRange("X18:X22").clearDataValidations();
  sheetLive.getRange("Z28:Z32").clearDataValidations();
  sheetLive.getRange("AA28:AA32").clearDataValidations();
  
  // Set header kolom undian semifinal & final (Bisa diisi manual oleh panitia untuk mengacak grup)
  sheetPmr.getRange("H5").setValue("No. Undi SF");
  sheetPmr.getRange("H5").setFontWeight("bold").setHorizontalAlignment("center");
  sheetPmr.getRange("I5").setValue("No. Undi Final");
  sheetPmr.getRange("I5").setFontWeight("bold").setHorizontalAlignment("center");
 
  // 1. Semifinal Kelompok 1 (Rows 8-12) - Mengambil berdasarkan No. Undi SF (Fallback ke Peringkat Penyisihan)
  for (var r = 8; r <= 12; r++) {
    var rank = r - 7; // No. Undi SF 1 s.d 5
    sheetLive.getRange("C" + r).setFormula('=IFERROR(INDEX(FILTER(\'PMR MADYA\'!$B$6:$B$23, (\'PMR MADYA\'!$G$6:$G$23 = "LOLOS SEMIFINAL") * (\'PMR MADYA\'!$H$6:$H$23 = ' + rank + ')), 1), IFERROR(INDEX(SORT(FILTER(\'PMR MADYA\'!$B$6:$B$23, \'PMR MADYA\'!$G$6:$G$23 = "LOLOS SEMIFINAL"), FILTER(\'PMR MADYA\'!$F$6:$F$23, \'PMR MADYA\'!$G$6:$G$23 = "LOLOS SEMIFINAL"), TRUE), ' + rank + '), ""))');
    sheetLive.getRange("D" + r).setFormula('=IFERROR(VLOOKUP(C' + r + ', \'PMR MADYA\'!$B$6:$C$23, 2, FALSE), "")');
    sheetLive.getRange("U" + r).setFormula('=SUM(F' + r + ':T' + r + ')');
    sheetLive.getRange("V" + r).setValue(0); // Reset tie break to static 0
    sheetLive.getRange("W" + r).setFormula('=IF(ISNUMBER(U' + r + '), RANK(U' + r + ', $U$8:$U$12), "")');
    sheetLive.getRange("X" + r).setFormula('=IF(W' + r + '<=3, "Lolos Final", "-")');
    sheetLive.getRange("Z" + r).setFormula('=IF(W' + r + '<=3, "Lolos Final", "-")');
  }
  
  // 2. Semifinal Kelompok 2 (Rows 18-22) - Mengambil berdasarkan No. Undi SF (Fallback ke Peringkat Penyisihan)
  for (var r = 18; r <= 22; r++) {
    var rank = r - 12; // No. Undi SF 6 s.d 10 (r=18 -> 6, r=22 -> 10)
    sheetLive.getRange("C" + r).setFormula('=IFERROR(INDEX(FILTER(\'PMR MADYA\'!$B$6:$B$23, (\'PMR MADYA\'!$G$6:$G$23 = "LOLOS SEMIFINAL") * (\'PMR MADYA\'!$H$6:$H$23 = ' + rank + ')), 1), IFERROR(INDEX(SORT(FILTER(\'PMR MADYA\'!$B$6:$B$23, \'PMR MADYA\'!$G$6:$G$23 = "LOLOS SEMIFINAL"), FILTER(\'PMR MADYA\'!$F$6:$F$23, \'PMR MADYA\'!$G$6:$G$23 = "LOLOS SEMIFINAL"), TRUE), ' + rank + '), ""))');
    sheetLive.getRange("D" + r).setFormula('=IFERROR(VLOOKUP(C' + r + ', \'PMR MADYA\'!$B$6:$C$23, 2, FALSE), "")');
    sheetLive.getRange("U" + r).setFormula('=SUM(F' + r + ':T' + r + ')');
    sheetLive.getRange("V" + r).setValue(0); // Reset tie break to static 0
    sheetLive.getRange("W" + r).setFormula('=IF(ISNUMBER(U' + r + '), RANK(U' + r + ', $U$18:$U$22), "")');
    sheetLive.getRange("X" + r).setFormula('=IF(W' + r + '<=3, "Lolos Final", "-")');
    sheetLive.getRange("Z" + r).setFormula('=IF(W' + r + '<=3, "Lolos Final", "-")');
  }
  
  // 3. Final Kelompok (Rows 28-32) - Mengambil tim berstatus "LOLOS FINAL" di Semifinal PMR MADYA.
  for (var r = 28; r <= 32; r++) {
    var itemIndex = r - 27; // Index 1 s.d 5
    sheetLive.getRange("C" + r).setFormula('=IFERROR(IFERROR(INDEX(\'PMR MADYA\'!$B$6:$B$23, MATCH(' + itemIndex + ', \'PMR MADYA\'!$I$6:$I$23, 0)), INDEX(SORT(FILTER(\'PMR MADYA\'!$B$28:$B$37, \'PMR MADYA\'!$G$28:$G$37 = "LOLOS FINAL"), FILTER(\'PMR MADYA\'!$F$28:$F$37, \'PMR MADYA\'!$G$28:$G$37 = "LOLOS FINAL"), TRUE), ' + itemIndex + ')), "")');
    sheetLive.getRange("D" + r).setFormula('=IFERROR(VLOOKUP(C' + r + ', \'PMR MADYA\'!$B$6:$C$23, 2, FALSE), "")');
    sheetLive.getRange("Y" + r).setFormula('=IF(E' + r + '=1, 50, IF(E' + r + '=-1, -25, 0)) + IF(F' + r + '=1, 100, IF(F' + r + '=-1, -50, 0)) + IF(G' + r + '=1, 150, IF(G' + r + '=-1, -75, 0)) + IF(H' + r + '=1, 200, IF(H' + r + '=-1, -100, 0)) + IF(I' + r + '=1, 250, IF(I' + r + '=-1, -125, 0)) + IF(J' + r + '=1, 300, IF(J' + r + '=-1, -150, 0)) + IF(K' + r + '=1, 350, IF(K' + r + '=-1, -175, 0)) + IF(L' + r + '=1, 400, IF(L' + r + '=-1, -200, 0)) + IF(M' + r + '=1, 450, IF(M' + r + '=-1, -225, 0)) + IF(N' + r + '=1, 500, IF(N' + r + '=-1, -250, 0)) + IF(O' + r + '=1, 550, IF(O' + r + '=-1, -275, 0)) + IF(P' + r + '=1, 600, IF(P' + r + '=-1, -300, 0)) + IF(Q' + r + '=1, 650, IF(Q' + r + '=-1, -325, 0)) + IF(R' + r + '=1, 700, IF(R' + r + '=-1, -350, 0)) + IF(S' + r + '=1, 750, IF(S' + r + '=-1, -375, 0)) + IF(T' + r + '=1, 800, IF(T' + r + '=-1, -400, 0)) + IF(U' + r + '=1, 850, IF(U' + r + '=-1, -425, 0)) + IF(V' + r + '=1, 900, IF(V' + r + '=-1, -450, 0)) + IF(W' + r + '=1, 950, IF(W' + r + '=-1, -475, 0)) + IF(X' + r + '=1, 1000, IF(X' + r + '=-1, -500, 0))');
    sheetLive.getRange("Z" + r).setFormula('=IF(ISNUMBER(Y' + r + '), RANK(Y' + r + ', $Y$28:$Y$32), "")');
    sheetLive.getRange("AA" + r).setFormula('=IF(Z' + r + '="", "", IF(Z' + r + '=1, "JUARA 1", IF(Z' + r + '=2, "JUARA 2", IF(Z' + r + '=3, "JUARA 3", IF(Z' + r + '=4, "JUARA HARAPAN 1", "JUARA HARAPAN 2")))))');
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
  
  // Pastikan dimensi sheet cukup untuk mencegah range out of bounds
  ensureSheetDimensions(sheetPmr, 45, 10);
  ensureSheetDimensions(sheetLive, 35, 30); // Merujuk s.d kolom AD
  
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
    sheetPmr.getRange("F" + r).setFormula('=IF(E' + r + '>0, RANK(E' + r + ', $E$6:$E$19), "")');
    sheetPmr.getRange("G" + r).setFormula('=IF(F' + r + '="", "", IF(F' + r + '<=10, "LOLOS SEMIFINAL", ""))');
  }
  
  // 2. Babak Semifinal (Rows 24-33)
  for (var r = 24; r <= 33; r++) {
    sheetPmr.getRange("B" + r).setFormula('=IFERROR(INDEX(SORT(FILTER($B$6:$B$19, $G$6:$G$19 = "LOLOS SEMIFINAL"), FILTER($F$6:$F$19, $G$6:$G$19 = "LOLOS SEMIFINAL"), TRUE), A' + r + '), "")');
    sheetPmr.getRange("C" + r).setFormula('=IFERROR(VLOOKUP(B' + r + ', $B$6:$C$19, 2, FALSE), "")');
    sheetPmr.getRange("D" + r).setFormula('=IF(COUNTIF(\'LIVE BEL WIRA\'!$C$8:$C$12, B' + r + ')>0, 1, IF(COUNTIF(\'LIVE BEL WIRA\'!$C$18:$C$22, B' + r + ')>0, 2, ""))');
    sheetPmr.getRange("E" + r).setFormula('=IF(B' + r + '="", "", SUM(IFERROR(FILTER(\'LIVE BEL WIRA\'!$U$8:$U$12, \'LIVE BEL WIRA\'!$C$8:$C$12 = B' + r + '), 0)) + SUM(IFERROR(FILTER(\'LIVE BEL WIRA\'!$U$18:$U$22, \'LIVE BEL WIRA\'!$C$18:$C$22 = B' + r + '), 0)))');
    sheetPmr.getRange("F" + r).setFormula('=IF(ISNUMBER(E' + r + '), RANK(E' + r + ', $E$24:$E$33), "")');
    sheetPmr.getRange("G" + r).setFormula('=IF(F' + r + '="", "", IF(F' + r + '<=5, "LOLOS FINAL", ""))');
  }
  
  // 3. Babak Final (Rows 37-41)
  for (var r = 37; r <= 41; r++) {
    sheetPmr.getRange("B" + r).setFormula('=IF(SUM(\'LIVE BEL WIRA\'!$Y$28:$Y$32)=0, \'LIVE BEL WIRA\'!C28, IFERROR(INDEX(SORT(\'LIVE BEL WIRA\'!$C$28:$C$32, \'LIVE BEL WIRA\'!$Y$28:$Y$32, FALSE), A' + r + '), ""))');
    sheetPmr.getRange("C" + r).setFormula('=IFERROR(VLOOKUP(B' + r + ', $B$6:$C$19, 2, FALSE), "")');
    sheetPmr.getRange("D" + r).setFormula('=SUMIF(\'LIVE BEL WIRA\'!$C$28:$C$32, B' + r + ', \'LIVE BEL WIRA\'!$Y$28:$Y$32)');
    sheetPmr.getRange("E" + r).setFormula('=IF(ISNUMBER(D' + r + '), RANK(D' + r + ', $D$37:$D$41), "")');
    sheetPmr.getRange("F" + r).setFormula('=IF(E' + r + '="", "", IF(E' + r + '=1, "JUARA 1", IF(E' + r + '=2, "JUARA 2", IF(E' + r + '=3, "JUARA 3", IF(E' + r + '=4, "JUARA HARAPAN 1", "JUARA HARAPAN 2")))))');
  }
  
  // ==========================================
  // B. FORMULA DI SHEET LIVE BEL WIRA
  // ==========================================
  
  // Proaktif bersihkan Data Validation & Konten di Kolom C & D (Kelompok 1, Kelompok 2 & Final) agar bebas dari error
  sheetLive.getRange("C8:D12").clearDataValidations().clearContent();
  sheetLive.getRange("C18:D22").clearDataValidations().clearContent();
  sheetLive.getRange("C28:D32").clearDataValidations().clearContent();
  
  // Bersihkan juga validasi data di kolom Rank & Status agar penulisan rumus tidak ditolak oleh Google Sheets
  sheetLive.getRange("W8:W12").clearDataValidations();
  sheetLive.getRange("X8:X12").clearDataValidations();
  sheetLive.getRange("W18:W22").clearDataValidations();
  sheetLive.getRange("X18:X22").clearDataValidations();
  sheetLive.getRange("Z28:Z32").clearDataValidations();
  sheetLive.getRange("AA28:AA32").clearDataValidations();
  
  // Set header kolom undian semifinal & final (Bisa diisi manual oleh panitia untuk mengacak grup)
  sheetPmr.getRange("H5").setValue("No. Undi SF");
  sheetPmr.getRange("H5").setFontWeight("bold").setHorizontalAlignment("center");
  sheetPmr.getRange("I5").setValue("No. Undi Final");
  sheetPmr.getRange("I5").setFontWeight("bold").setHorizontalAlignment("center");

  // 1. Semifinal Kelompok 1 (Rows 8-12) - Mengambil berdasarkan No. Undi SF (Fallback ke Peringkat Penyisihan)
  for (var r = 8; r <= 12; r++) {
    var rank = r - 7; // No. Undi SF 1 s.d 5
    sheetLive.getRange("C" + r).setFormula('=IFERROR(INDEX(FILTER(\'PMR WIRA\'!$B$6:$B$19, (\'PMR WIRA\'!$G$6:$G$19 = "LOLOS SEMIFINAL") * (\'PMR WIRA\'!$H$6:$H$19 = ' + rank + ')), 1), IFERROR(INDEX(SORT(FILTER(\'PMR WIRA\'!$B$6:$B$19, \'PMR WIRA\'!$G$6:$G$19 = "LOLOS SEMIFINAL"), FILTER(\'PMR WIRA\'!$F$6:$F$19, \'PMR WIRA\'!$G$6:$G$19 = "LOLOS SEMIFINAL"), TRUE), ' + rank + '), ""))');
    sheetLive.getRange("D" + r).setFormula('=IFERROR(VLOOKUP(C' + r + ', \'PMR WIRA\'!$B$6:$C$19, 2, FALSE), "")');
    sheetLive.getRange("U" + r).setFormula('=SUM(F' + r + ':T' + r + ')');
    sheetLive.getRange("V" + r).setValue(0); // Reset tie break to static 0
    sheetLive.getRange("W" + r).setFormula('=IF(ISNUMBER(U' + r + '), RANK(U' + r + ', $U$8:$U$12), "")');
    sheetLive.getRange("X" + r).setFormula('=IF(W' + r + '<=3, "Lolos Final", "-")');
    sheetLive.getRange("Z" + r).setFormula('=IF(W' + r + '<=3, "Lolos Final", "-")');
  }
  
  // 2. Semifinal Kelompok 2 (Rows 18-22) - Mengambil berdasarkan No. Undi SF (Fallback ke Peringkat Penyisihan)
  for (var r = 18; r <= 22; r++) {
    var rank = r - 12; // No. Undi SF 6 s.d 10 (r=18 -> 6, r=22 -> 10)
    sheetLive.getRange("C" + r).setFormula('=IFERROR(INDEX(FILTER(\'PMR WIRA\'!$B$6:$B$19, (\'PMR WIRA\'!$G$6:$G$19 = "LOLOS SEMIFINAL") * (\'PMR WIRA\'!$H$6:$H$19 = ' + rank + ')), 1), IFERROR(INDEX(SORT(FILTER(\'PMR WIRA\'!$B$6:$B$19, \'PMR WIRA\'!$G$6:$G$19 = "LOLOS SEMIFINAL"), FILTER(\'PMR WIRA\'!$F$6:$F$19, \'PMR WIRA\'!$G$6:$G$19 = "LOLOS SEMIFINAL"), TRUE), ' + rank + '), ""))');
    sheetLive.getRange("D" + r).setFormula('=IFERROR(VLOOKUP(C' + r + ', \'PMR WIRA\'!$B$6:$C$19, 2, FALSE), "")');
    sheetLive.getRange("U" + r).setFormula('=SUM(F' + r + ':T' + r + ')');
    sheetLive.getRange("V" + r).setValue(0); // Reset tie break to static 0
    sheetLive.getRange("W" + r).setFormula('=IF(ISNUMBER(U' + r + '), RANK(U' + r + ', $U$18:$U$22), "")');
    sheetLive.getRange("X" + r).setFormula('=IF(W' + r + '<=3, "Lolos Final", "-")');
    sheetLive.getRange("Z" + r).setFormula('=IF(W' + r + '<=3, "Lolos Final", "-")');
  }
  
  // 3. Final Kelompok (Rows 28-32) - Mengambil 5 tim yang berstatus "LOLOS FINAL" di Semifinal PMR WIRA
  for (var r = 28; r <= 32; r++) {
    var itemIndex = r - 27; // Index 1 s.d 5
    sheetLive.getRange("C" + r).setFormula('=IFERROR(IFERROR(INDEX(\'PMR WIRA\'!$B$6:$B$19, MATCH(' + itemIndex + ', \'PMR WIRA\'!$I$6:$I$19, 0)), INDEX(SORT(FILTER(\'PMR WIRA\'!$B$24:$B$33, \'PMR WIRA\'!$G$24:$G$33 = "LOLOS FINAL"), FILTER(\'PMR WIRA\'!$F$24:$F$33, \'PMR WIRA\'!$G$24:$G$33 = "LOLOS FINAL"), TRUE), ' + itemIndex + ')), "")');
    sheetLive.getRange("D" + r).setFormula('=IFERROR(VLOOKUP(C' + r + ', \'PMR WIRA\'!$B$6:$C$19, 2, FALSE), "")');
    sheetLive.getRange("Y" + r).setFormula('=IF(E' + r + '=1, 50, IF(E' + r + '=-1, -25, 0)) + IF(F' + r + '=1, 100, IF(F' + r + '=-1, -50, 0)) + IF(G' + r + '=1, 150, IF(G' + r + '=-1, -75, 0)) + IF(H' + r + '=1, 200, IF(H' + r + '=-1, -100, 0)) + IF(I' + r + '=1, 250, IF(I' + r + '=-1, -125, 0)) + IF(J' + r + '=1, 300, IF(J' + r + '=-1, -150, 0)) + IF(K' + r + '=1, 350, IF(K' + r + '=-1, -175, 0)) + IF(L' + r + '=1, 400, IF(L' + r + '=-1, -200, 0)) + IF(M' + r + '=1, 450, IF(M' + r + '=-1, -225, 0)) + IF(N' + r + '=1, 500, IF(N' + r + '=-1, -250, 0)) + IF(O' + r + '=1, 550, IF(O' + r + '=-1, -275, 0)) + IF(P' + r + '=1, 600, IF(P' + r + '=-1, -300, 0)) + IF(Q' + r + '=1, 650, IF(Q' + r + '=-1, -325, 0)) + IF(R' + r + '=1, 700, IF(R' + r + '=-1, -350, 0)) + IF(S' + r + '=1, 750, IF(S' + r + '=-1, -375, 0)) + IF(T' + r + '=1, 800, IF(T' + r + '=-1, -400, 0)) + IF(U' + r + '=1, 850, IF(U' + r + '=-1, -425, 0)) + IF(V' + r + '=1, 900, IF(V' + r + '=-1, -450, 0)) + IF(W' + r + '=1, 950, IF(W' + r + '=-1, -475, 0)) + IF(X' + r + '=1, 1000, IF(X' + r + '=-1, -500, 0))');
    sheetLive.getRange("Z" + r).setFormula('=IF(ISNUMBER(Y' + r + '), RANK(Y' + r + ', $Y$28:$Y$32), "")');
    sheetLive.getRange("AA" + r).setFormula('=IF(Z' + r + '="", "", IF(Z' + r + '=1, "JUARA 1", IF(Z' + r + '=2, "JUARA 2", IF(Z' + r + '=3, "JUARA 3", IF(Z' + r + '=4, "JUARA HARAPAN 1", "JUARA HARAPAN 2")))))');
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
  
  // Helper untuk melakukan padding baris data agar selalu memiliki tepat 10 kolom
  function pushLog(row) {
    while (row.length < 10) {
      row.push("");
    }
    logData.push(row);
  }
  
  pushLog(["[DEBUG TABULASI LCC PMR]"]);
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
  
  // 2. Periksa Sheet LIVE BEL MULA Semifinal Kelompok 1 (Rows 8-12), Kelompok 2 (Rows 18-22) & Final (Rows 28-32)
  var sheetLive = ss.getSheetByName("LIVE BEL MULA");
  if (sheetLive) {
    pushLog(["--- DATA LIVE BEL MULA (Semifinal Kelompok 1 Rows 8-12) ---"]);
    pushLog(["Row", "Col D (Kode)", "Col E (Nama)", "Col W (Total)", "Col Y (Rank SF)", "Col Z (Status)"]);
    for (var r = 8; r <= 12; r++) {
      var rowVals = [];
      rowVals.push("Row " + r);
      rowVals.push(sheetLive.getRange("D" + r).getValue() + " (" + sheetLive.getRange("D" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("E" + r).getValue() + " (" + sheetLive.getRange("E" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("W" + r).getValue() + " (" + sheetLive.getRange("W" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("Y" + r).getValue() + " (" + sheetLive.getRange("Y" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("Z" + r).getValue() + " (" + sheetLive.getRange("Z" + r).getFormula() + ")");
      pushLog(rowVals);
    }
    pushLog([]);
    
    pushLog(["--- DATA LIVE BEL MULA (Semifinal Kelompok 2 Rows 18-22) ---"]);
    pushLog(["Row", "Col D (Kode)", "Col E (Nama)", "Col W (Total)", "Col Y (Rank SF)", "Col Z (Status)"]);
    for (var r = 18; r <= 22; r++) {
      var rowVals = [];
      rowVals.push("Row " + r);
      rowVals.push(sheetLive.getRange("D" + r).getValue() + " (" + sheetLive.getRange("D" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("E" + r).getValue() + " (" + sheetLive.getRange("E" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("W" + r).getValue() + " (" + sheetLive.getRange("W" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("Y" + r).getValue() + " (" + sheetLive.getRange("Y" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("Z" + r).getValue() + " (" + sheetLive.getRange("Z" + r).getFormula() + ")");
      pushLog(rowVals);
    }
    pushLog([]);
    pushLog(["--- DATA LIVE BEL MULA (Final Rows 28-32) ---"]);
    pushLog(["Row", "Col D (Kode)", "Col E (Nama)", "Col AA (Total)", "Col AB (Rank Final)", "Col AC (Keterangan)"]);
    for (var r = 28; r <= 32; r++) {
      var rowVals = [];
      rowVals.push("Row " + r);
      rowVals.push(sheetLive.getRange("D" + r).getValue() + " (" + sheetLive.getRange("D" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("E" + r).getValue() + " (" + sheetLive.getRange("E" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("AA" + r).getValue() + " (" + sheetLive.getRange("AA" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("AB" + r).getValue() + " (" + sheetLive.getRange("AB" + r).getFormula() + ")");
      rowVals.push(sheetLive.getRange("AC" + r).getValue() + " (" + sheetLive.getRange("AC" + r).getFormula() + ")");
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
 * Juga mendukung aksi edit No. Undian secara real-time dari website.
 */
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Deteksi Aksi Update Undian dari Website
  if (e && e.parameter && e.parameter.action === 'updateUndi') {
    var cat = e.parameter.cat;   // 'mula', 'madya', 'wira'
    var kode = e.parameter.kode; // e.g., 'MU12'
    var undi = e.parameter.undi; // e.g., '3' or '' to clear
    var isFinal = e.parameter.round === 'final' || e.parameter.isFinal === 'true';
    
    try {
      var success = updateUndiInSheet(ss, cat, kode, undi, isFinal);
      var responseObj = { success: success, message: success ? "Nomor undian berhasil diperbarui di Google Sheets!" : "Kode sekolah tidak ditemukan di babak penyisihan." };
      
      // Support JSONP callback jika dipanggil via JSONP script tag
      if (e.parameter.callback) {
        var callback = e.parameter.callback;
        return ContentService.createTextOutput(callback + "(" + JSON.stringify(responseObj) + ")")
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      
      return ContentService.createTextOutput(JSON.stringify(responseObj))
          .setMimeType(ContentService.MimeType.JSON)
          .setHeaders({ 'Access-Control-Allow-Origin': '*' });
    } catch (err) {
      var errObj = { success: false, error: err.toString() };
      if (e.parameter.callback) {
        return ContentService.createTextOutput(e.parameter.callback + "(" + JSON.stringify(errObj) + ")")
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return ContentService.createTextOutput(JSON.stringify(errObj))
        .setMimeType(ContentService.MimeType.JSON)
        .setHeaders({ 'Access-Control-Allow-Origin': '*' });
    }
  }

  // Aksi default: mengambil data live untuk dashboard
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
 * Helper untuk menulis nilai undian ke sheet penyisihan secara real-time
 */
function updateUndiInSheet(ss, cat, kode, undi, isFinal) {
  var sheetName = "";
  var startRow = 6;
  var endRow = 19;
  var undiColName = ""; // Column letter: 'G', 'H', or 'I'
  
  if (cat === "mula") {
    sheetName = "PMR MULA";
    endRow = 19;
    undiColName = isFinal ? "H" : "G";
  } else if (cat === "madya") {
    sheetName = "PMR MADYA";
    endRow = 23;
    undiColName = isFinal ? "I" : "H";
  } else if (cat === "wira") {
    sheetName = "PMR WIRA";
    endRow = 19;
    undiColName = isFinal ? "H" : "G";
  } else {
    return false;
  }
  
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return false;
  
  var rangeB = sheet.getRange("B" + startRow + ":B" + endRow);
  var valuesB = rangeB.getValues();
  
  for (var i = 0; i < valuesB.length; i++) {
    if (valuesB[i][0] === null || valuesB[i][0] === undefined || valuesB[i][0] === "") continue;
    var cellValue = valuesB[i][0].toString().trim();
    if (cellValue === kode.trim()) {
      var row = startRow + i;
      // Set nilai undian (konversi ke angka jika ada nilainya, or kosongkan)
      if (undi !== null && undi !== undefined && undi !== "") {
        sheet.getRange(undiColName + row).setValue(parseInt(undi));
      } else {
        sheet.getRange(undiColName + row).setValue("");
      }
      return true;
    }
  }
  
  return false;
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
