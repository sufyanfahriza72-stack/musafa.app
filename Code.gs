const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

const ADMIN_CREDENTIALS = ['MUSAFA2026', 'ADMIN_MUSAFA'];
const SECURITY_PINS = ['1234', '2026'];

const TTE_SPECIMEN_URL = 'https://lh3.googleusercontent.com/d/1mDzoyRWgLuyOutFn81FJd2XQ_JcAQtCo';
const LOGO_URL = 'https://lh3.googleusercontent.com/d/1v9XLw_NjWKP9zdtIgREkPYmgakhKMq_n';

const SHEETS = {
  AGENDA: 'Agenda',
  PELAJARAN: 'Pelajaran',
  UJIAN: 'Ujian',
  TUGAS: 'Tugas',
  NOTULENSI: 'Notulensi'
};

function doGet(e) {
  const params = (e && e.parameter) || {};

  if (params.manifest === '1' || params.manifest === 'true') {
    const manifestObj = {
      name: 'MUSAFA Mobile - Manajemen Agenda & Tugas',
      short_name: 'MUSAFA',
      start_url: './',
      scope: './',
      display: 'standalone',
      background_color: '#0f172a',
      theme_color: '#2563eb',
      icons: [
        {
          src: LOGO_URL,
          sizes: '192x192',
          type: 'image/png',
          purpose: 'any maskable'
        },
        {
          src: LOGO_URL,
          sizes: '512x512',
          type: 'image/png',
          purpose: 'any maskable'
        }
      ]
    };
    return ContentService
      .createTextOutput(JSON.stringify(manifestObj))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (params.action) {
    let args = [];
    if (params.args) {
      try {
        args = JSON.parse(params.args);
      } catch (err) {
        args = [];
      }
    }
    const result = dispatchAction(String(params.action), Array.isArray(args) ? args : []);
    if (params.callback) {
      const safeCb = String(params.callback).replace(/[^a-zA-Z0-9_$]/g, '');
      return ContentService
        .createTextOutput(`${safeCb}(${JSON.stringify(result)});`)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('MUSAFA Mobile - Sistem Managemen Agenda dan Tugas')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    body = {};
  }
  const action = String(body.action || '');
  const args = Array.isArray(body.args) ? body.args : [];
  const result = dispatchAction(action, args);
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function dispatchAction(action, args) {
  switch (action) {
    case 'verifyLogin':
      return verifyLogin(args[0]);
    case 'getAllData':
      return getAllData();
    case 'saveRecord':
      return saveRecord(args[0], args[1], args[2]);
    case 'deleteRecord':
      return deleteRecord(args[0], args[1], args[2]);
    case 'toggleUjianStatus':
      return toggleUjianStatus(args[0], args[1], args[2]);
    case 'saveNotulensi':
      return saveNotulensi(args[0], args[1]);
    case 'saveNotulensiArchive':
      return saveNotulensiArchive(args[0]);
    case 'deleteNotulensiArchive':
      return deleteNotulensiArchive(args[0], args[1]);
    case 'getAdminTteToken':
      return getAdminTteToken(args[0]);
    case 'verifyTteToken':
      return verifyTteToken(args[0]);
    default:
      return { success: false, message: 'Aksi tidak dikenal: ' + action };
  }
}

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
    }
  } else if (headers && headers.length > 0 && sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
  return sheet;
}

function sheetToObjects(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return rows
    .filter(r => r.some(cell => String(cell).trim() !== ''))
    .map(r => {
      const obj = {};
      headers.forEach((h, idx) => {
        if (h) obj[h] = r[idx];
      });
      return obj;
    });
}

function verifyLogin(cred) {
  const val = String(cred || '').trim();
  if (!val) return { success: false, message: 'Kredensial kosong.' };
  if (ADMIN_CREDENTIALS.indexOf(val) !== -1) {
    return { success: true, role: 'admin', message: 'Login admin berhasil.' };
  }
  return { success: false, message: 'Kredensial admin tidak valid.' };
}

function verifyPin(pin) {
  return SECURITY_PINS.indexOf(String(pin || '').trim()) !== -1;
}

function getAllData() {
  try {
    const agendaSheet = getOrCreateSheet(SHEETS.AGENDA, [
      'ID', 'Judul', 'Lokasi', 'JatuhTempo', 'Durasi', 'Prioritas', 'Status', 'DisposisiKepada', 'TamuKhusus', 'Notulensi'
    ]);
    const pelajaranSheet = getOrCreateSheet(SHEETS.PELAJARAN, [
      'ID', 'Hari', 'WaktuMulai', 'WaktuSelesai', 'NamaMapel', 'NamaGuru'
    ]);
    const ujianSheet = getOrCreateSheet(SHEETS.UJIAN, [
      'ID', 'Hari', 'Jam', 'NamaMapel', 'StatusAktif'
    ]);
    const tugasSheet = getOrCreateSheet(SHEETS.TUGAS, [
      'ID', 'Deadline', 'NamaMapel', 'NamaGuru', 'DetailTugas', 'Selesai'
    ]);
    const notulensiSheet = getOrCreateSheet(SHEETS.NOTULENSI, [
      'ID', 'AgendaID', 'Judul', 'Isi', 'TokenTTE', 'Penandatangan', 'Jabatan', 'Tanggal'
    ]);

    return {
      success: true,
      agenda: sheetToObjects(agendaSheet),
      pelajaran: sheetToObjects(pelajaranSheet),
      ujian: sheetToObjects(ujianSheet),
      tugas: sheetToObjects(tugasSheet),
      notulensi: sheetToObjects(notulensiSheet)
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function saveRecord(sheetName, payload, pin) {
  if (!verifyPin(pin)) {
    return { success: false, message: 'PIN Keamanan salah!' };
  }
  try {
    const headersMap = {
      Agenda: ['ID', 'Judul', 'Lokasi', 'JatuhTempo', 'Durasi', 'Prioritas', 'Status', 'DisposisiKepada', 'TamuKhusus', 'Notulensi'],
      Pelajaran: ['ID', 'Hari', 'WaktuMulai', 'WaktuSelesai', 'NamaMapel', 'NamaGuru'],
      Ujian: ['ID', 'Hari', 'Jam', 'NamaMapel', 'StatusAktif'],
      Tugas: ['ID', 'Deadline', 'NamaMapel', 'NamaGuru', 'DetailTugas', 'Selesai']
    };
    const headers = headersMap[sheetName];
    if (!headers) return { success: false, message: 'Sheet tidak valid.' };

    const sheet = getOrCreateSheet(sheetName, headers);
    const data = Object.assign({}, payload || {});
    if (!data.ID || String(data.ID).startsWith('TMP-')) {
      data.ID = sheetName.slice(0, 3).toUpperCase() + '-' + new Date().getTime();
    }

    const lastRow = sheet.getLastRow();
    const idColIndex = headers.indexOf('ID') + 1;
    let foundRow = -1;

    if (lastRow >= 2 && idColIndex > 0) {
      const ids = sheet.getRange(2, idColIndex, lastRow - 1, 1).getValues().flat();
      const idx = ids.findIndex(v => String(v).trim() === String(data.ID).trim());
      if (idx !== -1) foundRow = idx + 2;
    }

    const rowValues = headers.map(h => (data[h] !== undefined ? data[h] : ''));
    if (foundRow > 0) {
      sheet.getRange(foundRow, 1, 1, headers.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }

    return { success: true, id: data.ID, message: 'Data berhasil disimpan ke Spreadsheet.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function deleteRecord(sheetName, id, pin) {
  if (!verifyPin(pin)) {
    return { success: false, message: 'PIN Keamanan salah!' };
  }
  try {
    const sheet = getOrCreateSheet(sheetName, []);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Data tidak ditemukan.' };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const idCol = headers.indexOf('ID') + 1;
    const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues().flat();
    const rowIdx = ids.findIndex(v => String(v).trim() === String(id).trim());

    if (rowIdx === -1) return { success: false, message: 'ID tidak ditemukan.' };
    sheet.deleteRow(rowIdx + 2);
    return { success: true, message: 'Data berhasil dihapus.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function toggleUjianStatus(id, nextStatus, pin) {
  if (!verifyPin(pin)) {
    return { success: false, message: 'PIN Keamanan salah!' };
  }
  try {
    const sheet = getOrCreateSheet(SHEETS.UJIAN, ['ID', 'Hari', 'Jam', 'NamaMapel', 'StatusAktif']);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Data ujian kosong.' };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const idCol = headers.indexOf('ID') + 1;
    const stCol = headers.indexOf('StatusAktif') + 1;
    const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues().flat();
    const idx = ids.findIndex(v => String(v).trim() === String(id).trim());
    if (idx === -1) return { success: false, message: 'Jadwal ujian tidak ditemukan.' };

    sheet.getRange(idx + 2, stCol).setValue(nextStatus);
    return { success: true, message: 'Status ujian diubah menjadi ' + nextStatus };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function saveNotulensi(agendaId, isi) {
  try {
    const sheet = getOrCreateSheet(SHEETS.AGENDA, [
      'ID', 'Judul', 'Lokasi', 'JatuhTempo', 'Durasi', 'Prioritas', 'Status', 'DisposisiKepada', 'TamuKhusus', 'Notulensi'
    ]);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Agenda tidak ditemukan.' };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const idCol = headers.indexOf('ID') + 1;
    const notulCol = headers.indexOf('Notulensi') + 1;
    const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues().flat();
    const idx = ids.findIndex(v => String(v).trim() === String(agendaId).trim());
    if (idx !== -1 && notulCol > 0) {
      sheet.getRange(idx + 2, notulCol).setValue(isi);
    }
    return { success: true, message: 'Notulensi agenda tersimpan.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function saveNotulensiArchive(record) {
  try {
    const headers = ['ID', 'AgendaID', 'Judul', 'Isi', 'TokenTTE', 'Penandatangan', 'Jabatan', 'Tanggal'];
    const sheet = getOrCreateSheet(SHEETS.NOTULENSI, headers);
    const item = Object.assign({}, record || {});
    if (!item.ID) item.ID = 'NOT-' + new Date().getTime();

    const lastRow = sheet.getLastRow();
    const idCol = headers.indexOf('ID') + 1;
    let foundRow = -1;

    if (lastRow >= 2) {
      const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues().flat();
      const idx = ids.findIndex(v => String(v).trim() === String(item.ID).trim());
      if (idx !== -1) foundRow = idx + 2;
    }

    const values = headers.map(h => (item[h] !== undefined ? item[h] : ''));
    if (foundRow > 0) {
      sheet.getRange(foundRow, 1, 1, headers.length).setValues([values]);
    } else {
      sheet.appendRow(values);
    }

    return { success: true, id: item.ID, message: 'Arsip notulensi tersimpan.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function deleteNotulensiArchive(id, pin) {
  if (!verifyPin(pin)) {
    return { success: false, message: 'PIN Keamanan salah!' };
  }
  return deleteRecord(SHEETS.NOTULENSI, id, pin);
}

function computeServerToken(minuteOffset) {
  const slot = Math.floor(Date.now() / 60000) + (minuteOffset || 0);
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'MUSAFA-TTE-FT-' + slot);
  const hex = bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('').toUpperCase();
  return hex.slice(-6);
}

function getAdminTteToken(username) {
  const val = String(username || '').trim();
  if (ADMIN_CREDENTIALS.indexOf(val) === -1) {
    return { success: false, message: 'Akses ditolak.' };
  }
  return {
    success: true,
    token: computeServerToken(0),
    specimenUrl: TTE_SPECIMEN_URL
  };
}

function verifyTteToken(tokenInput) {
  const clean = String(tokenInput || '').trim().toUpperCase();
  if (!clean) return { success: false, message: 'Token kosong.' };
  if (clean === computeServerToken(0) || clean === computeServerToken(-1)) {
    return {
      success: true,
      ttdUrl: TTE_SPECIMEN_URL,
      message: 'Token TTE valid.'
    };
  }
  return { success: false, message: 'Token TTE tidak valid atau kedaluwarsa.' };
}