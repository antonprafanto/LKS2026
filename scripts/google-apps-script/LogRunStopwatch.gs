/**
 * LKS 2026 — Stopwatch Log Run Otonom untuk Google Sheets
 *
 * Cara pasang (sekali):
 * 1. Buka spreadsheet Google Sheets Anda
 * 2. Extensions → Apps Script
 * 3. Hapus isi default, paste seluruh file ini, Save
 * 4. Refresh spreadsheet → menu "⏱ Stopwatch Run" muncul di toolbar
 * 5. Saat run: Stopwatch Run → SIAP / START / SELESAI
 *
 * Sel waktu: B8 (SIAP), E8 (START), H8 (SELESAI)
 * Durasi final otomatis di K8/K9 (isi rumus di panduan)
 */

var LOG_SHEET = 'Log Run Otonom';
var CELL_SIAP = 'B8';
var CELL_START = 'E8';
var CELL_SELESAI = 'H8';
var CELL_LIVE = 'K7';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('\u23F1 Stopwatch Run')
    .addItem('\u25B6 SIAP', 'catatSIAP')
    .addItem('\u25B6 START', 'catatSTART')
    .addItem('\u25A0 SELESAI', 'catatSELESAI')
    .addItem('\u21BB Reset waktu', 'resetWaktuRun')
    .addSeparator()
    .addItem('Buka stopwatch live (sidebar)', 'showStopwatchSidebar')
    .addSeparator()
    .addItem('\u2699 Setup sheet Log Run (sekali)', 'setupLogRunSheet')
    .addToUi();
}

/**
 * Jalankan sekali setelah paste script — isi rumus durasi + perbaiki teks petunjuk.
 * Bisa juga di Script Editor: pilih setupLogRunSheet → Run.
 */
function setupLogRunSheet() {
  var sheet = getLogSheet_();

  sheet.getRange('K7').setFormula('=IF(E8="","",IF(H8<>"",H8-E8,NOW()-E8))');
  sheet.getRange('K8').setFormula('=IF(AND(E8<>"",H8<>""),H8-E8,"")');
  sheet.getRange('K9').setFormula('=IF(K8="","",ROUND(K8*24*60,1))');
  sheet.getRange('K7:K8').setNumberFormat('[h]:mm:ss');
  sheet.getRange('K9').setNumberFormat('0.0');

  sheet.getRange('J7').setValue('Durasi live:');
  sheet.getRange('J8').setValue('Durasi final:');
  sheet.getRange('J9').setValue('Menit:');

  try {
    sheet.getRange('J6:L6').merge();
  } catch (e) { /* sudah merge */ }
  sheet.getRange('J6').setValue(
    'STOPWATCH — menu \u23F1 Stopwatch Run (Google Sheets)'
  );

  try {
    sheet.getRange('J10:L10').merge();
  } catch (e) { /* sudah merge */ }
  sheet.getRange('J10').setValue(
    'Klik menu Stopwatch Run \u2192 SIAP / START / SELESAI. ' +
    'Manual: Ctrl+Shift+; di sel B8/E8/H8.'
  );

  // Skor otomatis dari Modul E (baris HASIL SINGKAT)
  var hasilRow = findRowByLabel_(sheet, 'Kubus berhasil ditempatkan benar:');
  if (hasilRow > 0) {
    sheet.getRange(hasilRow, 3).setFormula("=COUNTIF('Modul E'!H7:H15,1)");
    sheet.getRange(hasilRow + 1, 3).setFormula("='Modul E'!J16");
  }

  // Modul E tarik waktu START dari Log Run
  var modulE = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Modul E');
  if (modulE) {
    modulE.getRange('D4').setFormula("='Log Run Otonom'!E8");
    modulE.getRange('D4').setNumberFormat('HH:mm:ss');
  }

  SpreadsheetApp.getActive().toast(
    'Setup selesai — rumus K7/K8/K9 + skor Modul E aktif. Refresh halaman.',
    'LKS 2026',
    5
  );
}

function findRowByLabel_(sheet, label) {
  var finder = sheet.createTextFinder(label).matchEntireCell(false);
  var cell = finder.findNext();
  return cell ? cell.getRow() : -1;
}

function getLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet) {
    throw new Error('Sheet "' + LOG_SHEET + '" tidak ditemukan.');
  }
  return sheet;
}

function stampNow_(sheet, a1) {
  var cell = sheet.getRange(a1);
  cell.setValue(new Date());
  cell.setNumberFormat('HH:mm:ss');
}

function catatSIAP() {
  stampNow_(getLogSheet_(), CELL_SIAP);
  SpreadsheetApp.getActive().toast('Waktu SIAP tercatat di ' + CELL_SIAP, 'Stopwatch', 3);
}

function catatSTART() {
  var sheet = getLogSheet_();
  stampNow_(sheet, CELL_START);
  PropertiesService.getDocumentProperties().setProperty('lks_stopwatch_start', String(new Date().getTime()));
  sheet.getRange(CELL_LIVE).setValue(0).setNumberFormat('[h]:mm:ss');
  SpreadsheetApp.getActive().toast('START — stopwatch jalan. Buka sidebar untuk live timer.', 'Stopwatch', 4);
}

function catatSELESAI() {
  var sheet = getLogSheet_();
  stampNow_(sheet, CELL_SELESAI);
  PropertiesService.getDocumentProperties().deleteProperty('lks_stopwatch_start');
  updateLiveDuration_(sheet);
  SpreadsheetApp.getActive().toast('SELESAI — durasi final di K8/K9.', 'Stopwatch', 3);
}

function resetWaktuRun() {
  var sheet = getLogSheet_();
  sheet.getRange(CELL_SIAP).clearContent();
  sheet.getRange(CELL_START).clearContent();
  sheet.getRange(CELL_SELESAI).clearContent();
  sheet.getRange(CELL_LIVE).clearContent();
  PropertiesService.getDocumentProperties().deleteProperty('lks_stopwatch_start');
  SpreadsheetApp.getActive().toast('Waktu run direset.', 'Stopwatch', 3);
}

function updateLiveDuration_(sheet) {
  var start = sheet.getRange(CELL_START).getValue();
  var end = sheet.getRange(CELL_SELESAI).getValue();
  if (start && end) {
    var ms = end.getTime() - start.getTime();
    sheet.getRange(CELL_LIVE).setValue(ms / 86400000).setNumberFormat('[h]:mm:ss');
  }
}

function showStopwatchSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('StopwatchSidebar')
    .setTitle('Stopwatch Run')
    .setWidth(260);
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Dipanggil dari sidebar HTML setiap detik */
function getStopwatchState() {
  var sheet = getLogSheet_();
  var startVal = sheet.getRange(CELL_START).getValue();
  var endVal = sheet.getRange(CELL_SELESAI).getValue();
  var props = PropertiesService.getDocumentProperties();
  var t0 = props.getProperty('lks_stopwatch_start');

  if (!startVal || endVal) {
    return { running: false, display: '--:--:--' };
  }

  var startMs = startVal.getTime ? startVal.getTime() : Number(t0);
  if (!startMs) {
    return { running: false, display: '--:--:--' };
  }

  var elapsed = Math.max(0, Date.now() - startMs);
  return { running: true, display: formatMs_(elapsed) };
}

function formatMs_(ms) {
  var s = Math.floor(ms / 1000);
  var h = Math.floor(s / 3600);
  var m = Math.floor((s % 3600) / 60);
  var sec = s % 60;
  return pad_(h) + ':' + pad_(m) + ':' + pad_(sec);
}

function pad_(n) {
  return (n < 10 ? '0' : '') + n;
}
