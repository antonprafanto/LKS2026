/**
 * LKS 2026 — Stopwatch Log Run (SATU FILE — paste ke Code.gs saja)
 *
 * PENTING: buka dari spreadsheet → Extensions → Apps Script (bukan script.google.com)
 *
 * Setelah paste + Save:
 * 1. Pilih fungsi buatMenuStopwatch → Run → izinkan akses
 * 2. Kembali ke spreadsheet → menu "LKS Stopwatch" muncul di bar atas
 * 3. Pilih setupLogRunSheet → Run (isi rumus K7/K8/K9)
 */

var LOG_SHEET = 'Log Run Otonom';
var CELL_SIAP = 'B8';
var CELL_START = 'E8';
var CELL_SELESAI = 'H8';
var CELL_LIVE = 'K7';
var MENU_NAME = 'LKS Stopwatch';
/** WITA — Samarinda, Banjarmasin, Makassar (UTC+8). Ganti jika lomba di WIB/WIT. */
var TIMEZONE = 'Asia/Makassar';

function onOpen() {
  try {
    buatMenuStopwatch_();
  } catch (err) {
    Logger.log('onOpen gagal: ' + err);
  }
}

function onInstall() {
  onOpen();
}

/** Jalankan manual dari Script Editor jika menu belum muncul */
function buatMenuStopwatch() {
  buatMenuStopwatch_();
  SpreadsheetApp.getUi().alert(
    'Menu "' + MENU_NAME + '" sudah dibuat.\n\n' +
    'Kembali ke spreadsheet (tab browser). Menu ada di bar atas kanan.\n' +
    'Jika belum kelihatan, tutup tab spreadsheet lalu buka lagi dari Drive.'
  );
}

function buatMenuStopwatch_() {
  SpreadsheetApp.getUi()
    .createMenu(MENU_NAME)
    .addItem('SIAP', 'catatSIAP')
    .addItem('START', 'catatSTART')
    .addItem('SELESAI', 'catatSELESAI')
    .addItem('Reset waktu', 'resetWaktuRun')
    .addSeparator()
    .addItem('Stopwatch live (sidebar)', 'showStopwatchSidebar')
    .addSeparator()
    .addItem('Setup sheet (sekali)', 'setupLogRunSheet')
    .addToUi();
}

function setupLogRunSheet() {
  ensureTimezone_();

  var sheet = getLogSheet_();

  sheet.getRange('K7').setFormula('=IF(E8="","",IF(H8<>"",H8-E8,JAM_WITA()-E8))');
  sheet.getRange('K8').setFormula('=IF(AND(E8<>"",H8<>""),H8-E8,"")');
  sheet.getRange('K9').setFormula('=IF(K8="","",ROUND(K8*24*60,1))');
  sheet.getRange('K7:K8').setNumberFormat('[h]:mm:ss');
  sheet.getRange('K9').setNumberFormat('0.0');

  sheet.getRange('J7').setValue('Durasi run (dari 0):');
  sheet.getRange('J8').setValue('Durasi final:');
  sheet.getRange('J9').setValue('Menit:');

  safeMerge_(sheet, 'J6:L6');
  sheet.getRange('J6').setValue('STOPWATCH — durasi run di K7 (bukan jam SIAP di B8)');

  safeMerge_(sheet, 'J10:L10');
  sheet.getRange('J10').setValue(
    'B8/E8/H8 = jam lokal (' + TIMEZONE + '). Durasi run = K7/K8. Reset lalu klik ulang jika jam salah.'
  );

  sheet.getRange('A8').setValue('Jam SIAP (nyata):');
  sheet.getRange('D8').setValue('Jam START (nyata):');
  sheet.getRange('G8').setValue('Jam selesai (nyata):');

  var hasilRow = findRowByLabel_(sheet, 'Kubus berhasil ditempatkan benar:');
  if (hasilRow > 0) {
    sheet.getRange(hasilRow, 3).setFormula("=COUNTIF('Modul E'!H7:H15,1)");
    sheet.getRange(hasilRow + 1, 3).setFormula("='Modul E'!J16");
  }

  var modulE = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Modul E');
  if (modulE) {
    modulE.getRange('D4').setFormula("='Log Run Otonom'!E8");
    modulE.getRange('D4').setNumberFormat('HH:mm:ss');
  }

  buatMenuStopwatch_();

  SpreadsheetApp.getActive().toast(
    'Setup selesai. Timezone: ' + TIMEZONE + '. Klik Reset waktu lalu SIAP/START lagi.',
    'LKS 2026',
    8
  );
}

function getTimezone_() {
  return TIMEZONE;
}

/** Paksa timezone spreadsheet ke WITA (Samarinda) */
function ensureTimezone_() {
  SpreadsheetApp.getActiveSpreadsheet().setSpreadsheetTimeZone(TIMEZONE);
}

/** Jam WITA → angka 0–1 (format HH:mm:ss di sheet) — tampil 09:22 bukan 01:22 */
function witaTimeFraction_() {
  var now = new Date();
  var h = Number(Utilities.formatDate(now, TIMEZONE, 'H'));
  var m = Number(Utilities.formatDate(now, TIMEZONE, 'm'));
  var s = Number(Utilities.formatDate(now, TIMEZONE, 's'));
  return (h * 3600 + m * 60 + s) / 86400;
}

/**
 * Fungsi custom untuk rumus sel (K7).
 * @customfunction
 */
function JAM_WITA() {
  return witaTimeFraction_();
}

function durationBetween_(start, end) {
  if (start === '' || start === null || end === '' || end === null) {
    return null;
  }
  var dur = Number(end) - Number(start);
  if (dur < 0) {
    dur += 1;
  }
  return dur;
}

function safeMerge_(sheet, a1) {
  try {
    sheet.getRange(a1).merge();
  } catch (e) {
    /* sudah merge */
  }
}

function findRowByLabel_(sheet, label) {
  var cell = sheet.createTextFinder(label).matchEntireCell(false).findNext();
  return cell ? cell.getRow() : -1;
}

function getLogSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  if (!sheet) {
    throw new Error('Tab "' + LOG_SHEET + '" tidak ada. Pastikan nama sheet persis sama.');
  }
  return sheet;
}

function stampNow_(sheet, a1) {
  ensureTimezone_();
  sheet.getRange(a1).setValue(witaTimeFraction_());
  sheet.getRange(a1).setNumberFormat('HH:mm:ss');
}

function catatSIAP() {
  stampNow_(getLogSheet_(), CELL_SIAP);
  SpreadsheetApp.getActive().toast('SIAP tercatat di B8', MENU_NAME, 3);
}

function catatSTART() {
  var sheet = getLogSheet_();
  stampNow_(sheet, CELL_START);
  PropertiesService.getDocumentProperties().setProperty(
    'lks_stopwatch_start',
    String(new Date().getTime())
  );
  sheet.getRange(CELL_LIVE).setValue(0).setNumberFormat('[h]:mm:ss');
  SpreadsheetApp.getActive().toast('START — buka sidebar untuk timer live', MENU_NAME, 4);
}

function catatSELESAI() {
  var sheet = getLogSheet_();
  stampNow_(sheet, CELL_SELESAI);
  PropertiesService.getDocumentProperties().deleteProperty('lks_stopwatch_start');
  updateLiveDuration_(sheet);
  SpreadsheetApp.getActive().toast('SELESAI — lihat durasi di K8/K9', MENU_NAME, 3);
}

function resetWaktuRun() {
  var sheet = getLogSheet_();
  sheet.getRange(CELL_SIAP).clearContent();
  sheet.getRange(CELL_START).clearContent();
  sheet.getRange(CELL_SELESAI).clearContent();
  sheet.getRange(CELL_LIVE).clearContent();
  PropertiesService.getDocumentProperties().deleteProperty('lks_stopwatch_start');
  SpreadsheetApp.getActive().toast('Waktu direset', MENU_NAME, 3);
}

function updateLiveDuration_(sheet) {
  var dur = durationBetween_(
    sheet.getRange(CELL_START).getValue(),
    sheet.getRange(CELL_SELESAI).getValue()
  );
  if (dur !== null) {
    sheet.getRange(CELL_LIVE).setValue(dur).setNumberFormat('[h]:mm:ss');
  }
}

function showStopwatchSidebar() {
  var html = HtmlService.createHtmlOutput(getSidebarHtml_())
    .setTitle('Stopwatch Run')
    .setWidth(280);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSidebarHtml_() {
  return [
    '<!DOCTYPE html><html><head><base target="_top">',
    '<style>',
    'body{font-family:Arial,sans-serif;padding:12px;text-align:center}',
    '#clock{font-size:36px;font-weight:bold;color:#c00000;margin:16px 0}',
    'button{display:block;width:100%;margin:6px 0;padding:10px;font-size:14px;cursor:pointer}',
    '.hint{font-size:12px;color:#555}',
    '</style></head><body>',
    '<div class="hint">Timer live (update tiap detik)</div>',
    '<div id="clock">00:00:00</div>',
    '<button onclick="run(\'catatSIAP\')">SIAP</button>',
    '<button onclick="run(\'catatSTART\')">START</button>',
    '<button onclick="run(\'catatSELESAI\')">SELESAI</button>',
    '<button onclick="run(\'resetWaktuRun\')">Reset</button>',
    '<p class="hint">Waktu di B8, E8, H8</p>',
    '<script>',
    'function tick(){google.script.run.withSuccessHandler(function(s){',
    'document.getElementById("clock").textContent=s.display;}).getStopwatchState();}',
    'function run(fn){google.script.run.withSuccessHandler(tick)[fn]();}',
    'tick();setInterval(tick,1000);',
    '</script></body></html>',
  ].join('');
}

function getStopwatchState() {
  var sheet = getLogSheet_();
  var startVal = sheet.getRange(CELL_START).getValue();
  var endVal = sheet.getRange(CELL_SELESAI).getValue();
  var t0 = PropertiesService.getDocumentProperties().getProperty('lks_stopwatch_start');

  if (!startVal || endVal) {
    return { running: false, display: '--:--:--' };
  }

  var startMs = startVal.getTime ? startVal.getTime() : Number(t0);
  if (!startMs) {
    return { running: false, display: '--:--:--' };
  }

  return { running: true, display: formatMs_(Math.max(0, Date.now() - startMs)) };
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

/** Tes: Run dari Script Editor — harus muncul popup "Script OK" */
function testScriptOK() {
  getLogSheet_();
  SpreadsheetApp.getUi().alert('Script OK — sheet Log Run Otonom ditemukan.');
}
