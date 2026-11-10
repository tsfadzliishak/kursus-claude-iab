/**
 * Sistem Pengurusan Program IAB (Latihan)
 * Pangkalan data: Google Sheet ini (helaian "Peserta" dan "Penilaian").
 * Pasang: Extensions > Apps Script, tampal fail ini sebagai Code.gs,
 * cipta fail HTML bernama "Index", jalankan setupSheets() sekali, kemudian Deploy > Web app.
 */

const SHEET_PESERTA = 'Peserta';
const SHEET_PENILAIAN = 'Penilaian';

const SENARAI_SESI = [
  'S02 Kepimpinan Instruksional',
  'S04 Kurikulum dan PdP',
  'S05 Coaching dan Mentoring',
  'S07 Pengurusan Kewangan',
  'S08 Data dan Pelaporan',
  'S09 Lawatan Sekolah',
  'S10 Komuniti dan PIBG',
  'S11 Aset dan Perolehan',
  'S12 Pelan Strategik',
  'S13 Pembentangan Pelan Tindakan'
];

/** Paparkan aplikasi web. */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sistem Pengurusan Program IAB')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Jalankan SEKALI dari editor untuk mencipta helaian dan tajuk lajur. */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const definisi = {};
  definisi[SHEET_PESERTA] = ['ID', 'Nama', 'Jantina', 'Negeri', 'Jenis Sekolah', 'E-mel', 'Tarikh Daftar'];
  definisi[SHEET_PENILAIAN] = ['Masa', 'ID Peserta', 'Sesi', 'Skor', 'Komen'];

  Object.keys(definisi).forEach(function (nama) {
    const sh = ss.getSheetByName(nama) || ss.insertSheet(nama);
    if (sh.getLastRow() === 0) {
      const tajuk = definisi[nama];
      sh.appendRow(tajuk);
      sh.getRange(1, 1, 1, tajuk.length).setFontWeight('bold').setBackground('#d9e1f2');
      sh.setFrozenRows(1);
    }
  });
}

function getSheet_(nama) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nama);
  if (!sh) throw new Error('Helaian "' + nama + '" tiada. Jalankan setupSheets() dahulu.');
  return sh;
}

/** Senarai sesi untuk menu pilihan. */
function getSesi() {
  return SENARAI_SESI;
}

/** READ: semua peserta (tanpa lajur tarikh kerana Date tidak boleh dihantar ke HTML). */
function senaraiPeserta() {
  const nilai = getSheet_(SHEET_PESERTA).getDataRange().getValues().slice(1);
  return nilai
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return { id: r[0], nama: r[1], jantina: r[2], negeri: r[3], jenis: r[4], emel: r[5] };
    });
}

/** CREATE: daftar peserta baharu, pulangkan ID. */
function daftarPeserta(p) {
  if (!p || !p.nama || !p.emel) throw new Error('Nama dan e-mel wajib diisi.');
  const emel = String(p.emel).trim().toLowerCase();

  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // elak dua pendaftaran serentak mendapat ID yang sama
  try {
    const sh = getSheet_(SHEET_PESERTA);
    const ada = senaraiPeserta().some(function (x) {
      return String(x.emel).toLowerCase() === emel;
    });
    if (ada) throw new Error('E-mel ini sudah didaftarkan.');

    const id = 'P' + String(sh.getLastRow()).padStart(3, '0');
    sh.appendRow([id, String(p.nama).trim(), p.jantina, String(p.negeri).trim(), p.jenis, emel, new Date()]);
    return id;
  } finally {
    lock.releaseLock();
  }
}

/** CREATE: simpan satu penilaian sesi. */
function hantarPenilaian(n) {
  const skor = Number(n && n.skor);
  if (!Number.isInteger(skor) || skor < 1 || skor > 5) throw new Error('Skor mesti nombor bulat 1 hingga 5.');
  if (SENARAI_SESI.indexOf(n.sesi) === -1) throw new Error('Sesi tidak sah.');
  const wujud = senaraiPeserta().some(function (p) { return p.id === n.id; });
  if (!wujud) throw new Error('ID peserta tidak dijumpai.');

  getSheet_(SHEET_PENILAIAN).appendRow([new Date(), n.id, n.sesi, skor, String(n.komen || '').trim()]);
  return true;
}

/** READ + kira: min skor dan bilangan respons bagi setiap sesi, tertinggi dahulu. */
function ringkasanPenilaian() {
  const nilai = getSheet_(SHEET_PENILAIAN).getDataRange().getValues().slice(1);
  const kumpul = {};
  nilai.forEach(function (r) {
    const sesi = r[2];
    const skor = Number(r[3]);
    if (!sesi || !skor) return;
    if (!kumpul[sesi]) kumpul[sesi] = { jumlah: 0, n: 0 };
    kumpul[sesi].jumlah += skor;
    kumpul[sesi].n += 1;
  });
  return Object.keys(kumpul)
    .map(function (s) {
      return { sesi: s, min: Math.round((kumpul[s].jumlah / kumpul[s].n) * 100) / 100, n: kumpul[s].n };
    })
    .sort(function (a, b) { return b.min - a.min; });
}
