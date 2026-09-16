/******************************************************
 * SISTEM INFORMASI REALISASI KEGIATAN PEMBINAAN SD
 * Google Apps Script - Code.gs
 ******************************************************/

const SPREADSHEET_ID = '1Mb2HKjsJhkMizockGVczDhGmgujH1QcHU_Ub_2buwLE';
const SHEET_NAME = 'Data_Realisasi';

/*
 * GANTI PASSWORD ADMIN DI SINI.
 * Jangan taruh password ini di HTML.
 */
const ADMIN_PASSWORD = 'GantiPasswordAdminAnda123!';

/*
 * Masa berlaku session admin dalam detik.
 * CacheService maksimum 6 jam.
 */
const SESSION_TTL = 21600;


/* =====================================================
 * RESPONSE JSON
 * ===================================================== */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/* =====================================================
 * GET - PUBLIC READ ONLY
 * ===================================================== */
function doGet(e) {
  try {
    const action = e && e.parameter ? (e.parameter.action || 'getData') : 'getData';

    if (action === 'getData') {
      return getData_();
    }

    return jsonResponse({
      success: true,
      message: 'API aktif.',
      action: action
    });

  } catch (err) {
    return jsonResponse({
      success: false,
      message: err.message
    });
  }
}


/* =====================================================
 * POST - LOGIN + CRUD ADMIN
 * ===================================================== */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({
        success: false,
        message: 'POST body tidak ditemukan.'
      });
    }

    const req = JSON.parse(e.postData.contents);
    const action = String(req.action || '');

    // LOGIN: password hanya diperiksa di server.
    if (action === 'login') {
      return login_(String(req.password || ''));
    }

    // LOGOUT
    if (action === 'logout') {
      return logout_(String(req.token || ''));
    }

    // Semua operasi tulis wajib punya session admin valid.
    if (['insert', 'update', 'delete'].indexOf(action) !== -1) {
      const auth = requireAdmin_(String(req.token || ''));
      if (!auth.ok) {
        return jsonResponse({
          success: false,
          authenticated: false,
          message: auth.message
        });
      }
    }

    if (action === 'insert') return insert_(req);
    if (action === 'update') return update_(req);
    if (action === 'delete') return delete_(req);

    return jsonResponse({
      success: false,
      message: 'Action tidak dikenali: ' + action
    });

  } catch (err) {
    return jsonResponse({
      success: false,
      message: err.message
    });
  }
}


/* =====================================================
 * LOGIN
 * ===================================================== */
function login_(password) {
  if (!password || password !== ADMIN_PASSWORD) {
    return jsonResponse({
      success: false,
      authenticated: false,
      message: 'Password admin salah.'
    });
  }

  const token = Utilities.getUuid().replace(/-/g, '') +
                Utilities.getUuid().replace(/-/g, '');

  CacheService.getScriptCache().put(
    'admin_' + token,
    JSON.stringify({
      role: 'admin',
      createdAt: new Date().toISOString()
    }),
    SESSION_TTL
  );

  return jsonResponse({
    success: true,
    authenticated: true,
    role: 'admin',
    token: token,
    message: 'Login admin berhasil.'
  });
}


/* =====================================================
 * VALIDASI SESSION
 * ===================================================== */
function requireAdmin_(token) {
  if (!token) {
    return {
      ok: false,
      message: 'Sesi admin tidak ditemukan. Silakan login.'
    };
  }

  const value = CacheService.getScriptCache()
    .get('admin_' + token);

  if (!value) {
    return {
      ok: false,
      message: 'Sesi admin telah kedaluwarsa. Silakan login kembali.'
    };
  }

  try {
    const session = JSON.parse(value);

    if (session.role !== 'admin') {
      return {
        ok: false,
        message: 'Akses ditolak.'
      };
    }

    // Perpanjang session setiap ada aktivitas.
    CacheService.getScriptCache().put(
      'admin_' + token,
      value,
      SESSION_TTL
    );

    return { ok: true };

  } catch (err) {
    return {
      ok: false,
      message: 'Session tidak valid.'
    };
  }
}


/* =====================================================
 * LOGOUT
 * ===================================================== */
function logout_(token) {
  if (token) {
    CacheService.getScriptCache().remove('admin_' + token);
  }

  return jsonResponse({
    success: true,
    message: 'Logout berhasil.'
  });
}


/* =====================================================
 * BACA DATA
 * ===================================================== */
function getData_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();

  if (!values.length) {
    return jsonResponse({
      success: true,
      data: []
    });
  }

  const headers = values[0].map(String);

  validateHeaders_(headers);

  const data = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    if (row.join('').trim() === '') continue;

    const obj = {
      _row: i + 1
    };

    headers.forEach(function(header, index) {
      obj[header] = row[index];
    });

    data.push(obj);
  }

  return jsonResponse({
    success: true,
    data: data
  });
}


/* =====================================================
 * INSERT
 * ===================================================== */
function insert_(req) {
  validateInput_(req);

  const sheet = getSheet_();

  const pagu = number_(req.Pagu_Anggaran);
  const realisasiKeuangan = number_(req.Realisasi_Keuangan);
  const realisasiFisik = number_(req.Realisasi_Fisik);

  const sisa = pagu - realisasiKeuangan;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    sheet.appendRow([
      String(req.Kode_Rek).trim(),
      String(req.Sub_Kegiatan).trim(),
      String(req.Kategori).trim(),
      pagu,
      realisasiKeuangan,
      realisasiFisik,
      sisa
    ]);
  } finally {
    lock.releaseLock();
  }

  return jsonResponse({
    success: true,
    message: 'Data berhasil ditambahkan.'
  });
}


/* =====================================================
 * UPDATE
 * ===================================================== */
function update_(req) {
  validateInput_(req);

  const sheet = getSheet_();
  const rowNumber = parseInt(req._row, 10);

  if (!rowNumber || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
    throw new Error('Nomor baris tidak valid.');
  }

  const pagu = number_(req.Pagu_Anggaran);
  const realisasiKeuangan = number_(req.Realisasi_Keuangan);
  const realisasiFisik = number_(req.Realisasi_Fisik);

  const sisa = pagu - realisasiKeuangan;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    sheet.getRange(rowNumber, 1, 1, 7).setValues([[
      String(req.Kode_Rek).trim(),
      String(req.Sub_Kegiatan).trim(),
      String(req.Kategori).trim(),
      pagu,
      realisasiKeuangan,
      realisasiFisik,
      sisa
    ]]);
  } finally {
    lock.releaseLock();
  }

  return jsonResponse({
    success: true,
    message: 'Data berhasil diperbarui.'
  });
}


/* =====================================================
 * DELETE
 * ===================================================== */
function delete_(req) {
  const sheet = getSheet_();
  const rowNumber = parseInt(req._row, 10);

  if (!rowNumber || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
    throw new Error('Nomor baris tidak valid.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    sheet.deleteRow(rowNumber);
  } finally {
    lock.releaseLock();
  }

  return jsonResponse({
    success: true,
    message: 'Data berhasil dihapus.'
  });
}


/* =====================================================
 * VALIDASI INPUT
 * ===================================================== */
function validateInput_(req) {
  const kode = String(req.Kode_Rek || '').trim();
  const sub = String(req.Sub_Kegiatan || '').trim();
  const kategori = String(req.Kategori || '').trim();

  if (!kode) throw new Error('Kode_Rek wajib diisi.');
  if (!sub) throw new Error('Sub_Kegiatan wajib diisi.');

  if (kategori !== 'Sosbud' && kategori !== 'Sarpras') {
    throw new Error('Kategori harus Sosbud atau Sarpras.');
  }

  const pagu = number_(req.Pagu_Anggaran);
  const realisasi = number_(req.Realisasi_Keuangan);
  const fisik = number_(req.Realisasi_Fisik);

  if (pagu < 0) throw new Error('Pagu_Anggaran tidak boleh negatif.');
  if (realisasi < 0) throw new Error('Realisasi_Keuangan tidak boleh negatif.');
  if (realisasi > pagu) {
    throw new Error('Realisasi_Keuangan tidak boleh melebihi Pagu_Anggaran.');
  }
  if (fisik < 0 || fisik > 100) {
    throw new Error('Realisasi_Fisik harus antara 0 sampai 100.');
  }
}


/* =====================================================
 * UTILITAS
 * ===================================================== */
function number_(value) {
  const n = Number(value);
  if (!isFinite(n)) {
    throw new Error('Nilai angka tidak valid.');
  }
  return n;
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error('Sheet "' + SHEET_NAME + '" tidak ditemukan.');
  }

  return sheet;
}

function validateHeaders_(headers) {
  const required = [
    'Kode_Rek',
    'Sub_Kegiatan',
    'Kategori',
    'Pagu_Anggaran',
    'Realisasi_Keuangan',
    'Realisasi_Fisik',
    'Sisa_Anggaran'
  ];

  required.forEach(function(header) {
    if (headers.indexOf(header) === -1) {
      throw new Error('Header tidak ditemukan: ' + header);
    }
  });
}


/* =====================================================
 * TES API DARI EDITOR GAS
 * ===================================================== */
function testGetData() {
  Logger.log(getData_().getContent());
}
