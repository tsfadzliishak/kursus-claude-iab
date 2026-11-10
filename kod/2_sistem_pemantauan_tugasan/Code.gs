/**
 * Sistem Pemantauan Tugasan PKPB Siri 3/2026 (Latihan)
 * Pangkalan data: helaian "Tugasan" dalam Google Sheet ini.
 * Pasang: Extensions > Apps Script, tampal fail ini sebagai Code.gs,
 * cipta fail HTML bernama "Index", jalankan setupSheet() sekali, kemudian Deploy > Web app.
 */

const SHEET_TUGASAN = 'Tugasan';
const TAJUK = ['ID', 'Tugasan', 'Fasa', 'Pemilik', 'Tarikh Mula', 'Tarikh Akhir', 'Bergantung Kepada',
               'Status', '% Siap', 'Catatan', 'Dikemas Kini', 'Oleh'];
const STATUS = ['Belum mula', 'Sedang', 'Selesai', 'Tersekat'];
const KOL = { id: 1, status: 8, peratus: 9, catatan: 10, dikemaskini: 11, oleh: 12 };

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Pemantauan Tugasan PKPB3')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Jalankan dari editor: cipta helaian/tajuk (jika tiada) dan pasang senarai pilihan Status. */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_TUGASAN) || ss.insertSheet(SHEET_TUGASAN);
  if (sh.getLastRow() === 0) sh.appendRow(TAJUK);
  sh.getRange(1, 1, 1, TAJUK.length).setFontWeight('bold').setBackground('#d9e1f2');
  sh.setFrozenRows(1);

  const baris = Math.max(sh.getMaxRows() - 1, 1);
  const pilihanStatus = SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).build();
  sh.getRange(2, KOL.status, baris, 1).setDataValidation(pilihanStatus);
  const julatPeratus = SpreadsheetApp.newDataValidation().requireNumberBetween(0, 100).build();
  sh.getRange(2, KOL.peratus, baris, 1).setDataValidation(julatPeratus);
}

function getSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TUGASAN);
  if (!sh) throw new Error('Helaian "Tugasan" tiada. Jalankan setupSheet() dahulu.');
  return sh;
}

function tarikh_(v, corak) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), corak || 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}

/** READ: semua tugasan + masa pelayan. Tarikh dihantar sebagai teks (Date tidak boleh dihantar ke HTML). */
function getData() {
  const hariIni = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const nilai = getSheet_().getDataRange().getValues().slice(1);
  const tugasan = nilai
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      const status = String(r[7] || 'Belum mula');
      const akhir = tarikh_(r[5]);
      const peratus = status === 'Selesai' ? 100 : Math.max(0, Math.min(100, Number(r[8]) || 0));
      return {
        id: String(r[0]), tugasan: String(r[1]), fasa: String(r[2]), pemilik: String(r[3]),
        mula: tarikh_(r[4]), akhir: akhir, bergantung: String(r[6] || ''),
        status: status, peratus: peratus, catatan: String(r[9] || ''),
        dikemaskini: tarikh_(r[10], 'dd/MM/yyyy HH:mm'), oleh: String(r[11] || ''),
        lewat: status !== 'Selesai' && /^\d{4}-\d{2}-\d{2}$/.test(akhir) && akhir < hariIni
      };
    });
  return {
    tugasan: tugasan,
    hariIni: hariIni,
    masa: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
  };
}

/** UPDATE: kemas kini status, % siap dan catatan satu tugasan. */
function kemaskiniTugasan(k) {
  if (!k || !k.id) throw new Error('ID tugasan tiada.');
  if (STATUS.indexOf(k.status) === -1) throw new Error('Status tidak sah.');
  let peratus = Number(k.peratus);
  if (isNaN(peratus) || peratus < 0 || peratus > 100) throw new Error('% Siap mesti 0 hingga 100.');
  if (k.status === 'Selesai') peratus = 100;

  const emel = Session.getActiveUser().getEmail(); // kosong bagi sesetengah jenis akaun
  const oleh = emel || String(k.nama || '').trim();
  if (!oleh) throw new Error('Sila isi nama anda.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_();
    const ids = sh.getRange(2, KOL.id, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
    const i = ids.findIndex(function (r) { return String(r[0]) === String(k.id); });
    if (i === -1) throw new Error('Tugasan ' + k.id + ' tidak dijumpai.');
    const baris = i + 2;
    sh.getRange(baris, KOL.status, 1, 5).setValues([[k.status, peratus, String(k.catatan || '').trim(), new Date(), oleh]]);
    return true;
  } finally {
    lock.releaseLock();
  }
}
