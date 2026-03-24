// ============================================================
// WAREHOUSE FCL - Google Apps Script Backend
// Code.gs - Main Server-Side Logic
// ============================================================

// ---- KONFIGURASI ----
const CONFIG = {
  SPREADSHEET_ID: '', // Akan diisi otomatis saat setup
  SHEETS: {
    USERS: 'Users',
    KAS_GUDANG: 'KasGudang',
    TEAM_BUILDING: 'TeamBuilding',
    KARYAWAN: 'Karyawan',
    PEMBAYARAN_TB: 'PembayaranTB',
    SOP: 'SOP',
    ORGANISASI: 'Organisasi',
    SETTINGS: 'Settings'
  },
  VERSION: '1.0.0'
};

// ============================================================
// ENTRY POINT
// ============================================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Warehouse FCL')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ============================================================
// SETUP DATABASE
// ============================================================
function setupDatabase() {
  let ss;
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID');

  if (!ssId) {
    ss = SpreadsheetApp.create('Warehouse FCL - Database');
    ssId = ss.getId();
    props.setProperty('SPREADSHEET_ID', ssId);
    Logger.log('Spreadsheet baru dibuat: ' + ssId);
  } else {
    ss = SpreadsheetApp.openById(ssId);
  }

  // Buat semua sheet yang diperlukan
  setupSheet(ss, CONFIG.SHEETS.USERS, ['id', 'username', 'password', 'nama', 'role', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.KAS_GUDANG, ['id', 'tanggal', 'tipe', 'keterangan', 'nominal', 'buktiUrl', 'createdBy', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.TEAM_BUILDING, ['id', 'tanggal', 'keterangan', 'nominal', 'buktiUrl', 'createdBy', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.KARYAWAN, ['id', 'nama', 'jabatan', 'departemen', 'telepon', 'email', 'tanggalMasuk', 'status', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.PEMBAYARAN_TB, ['id', 'karyawanId', 'karyawanNama', 'periode', 'nominal', 'status', 'tanggalBayar', 'keterangan', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.SOP, ['id', 'judul', 'konten', 'kategori', 'createdBy', 'updatedAt']);
  setupSheet(ss, CONFIG.SHEETS.ORGANISASI, ['id', 'nama', 'jabatan', 'atasan', 'departemen', 'foto', 'urutan']);
  setupSheet(ss, CONFIG.SHEETS.SETTINGS, ['key', 'value']);

  // Tambah user admin default jika belum ada
  const usersSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
  if (usersSheet.getLastRow() <= 1) {
    usersSheet.appendRow([
      generateId(),
      'admin',
      hashPassword('admin123'),
      'Administrator',
      'admin',
      new Date().toISOString()
    ]);
    Logger.log('User admin default dibuat: username=admin, password=admin123');
  }

  return { success: true, spreadsheetId: ssId, url: ss.getUrl() };
}

function setupSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    Logger.log('Sheet dibuat: ' + sheetName);
  }
  return sheet;
}

// ============================================================
// AUTENTIKASI
// ============================================================
function login(username, password) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === username && data[i][2] === hashPassword(password)) {
        return {
          success: true,
          user: {
            id: data[i][0],
            username: data[i][1],
            nama: data[i][3],
            role: data[i][4]
          }
        };
      }
    }
    return { success: false, message: 'Username atau password salah' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ============================================================
// KAS GUDANG
// ============================================================
function getKasGudang() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KAS_GUDANG);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue; // skip baris kosong
      result.push({
        id: data[i][0],
        tanggal: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]),
        tipe: data[i][2],
        keterangan: data[i][3],
        nominal: parseFloat(data[i][4]) || 0,
        buktiUrl: data[i][5],
        createdBy: data[i][6],
        createdAt: data[i][7] instanceof Date ? data[i][7].toISOString() : String(data[i][7])
      });
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addKasGudang(tanggal, tipe, keterangan, nominal, buktiUrl, createdBy) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KAS_GUDANG);
    const id = generateId();
    sheet.appendRow([id, tanggal, tipe, keterangan, parseFloat(nominal), buktiUrl, createdBy, new Date().toISOString()]);
    return { success: true, id: id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteKasGudang(id) {
  return deleteRow(CONFIG.SHEETS.KAS_GUDANG, id);
}

function getSaldoGudang() {
  try {
    const result = getKasGudang();
    if (!result.success) return { success: false, message: result.message };
    let saldo = 0;
    result.data.forEach(d => {
      if (d.tipe === 'IN') saldo += d.nominal;
      else if (d.tipe === 'OUT') saldo -= d.nominal;
    });
    return { success: true, saldo: saldo };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// TEAM BUILDING
// ============================================================
function getTeamBuilding() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.TEAM_BUILDING);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({
        id: data[i][0],
        tanggal: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]),
        keterangan: data[i][2],
        nominal: parseFloat(data[i][3]) || 0,
        buktiUrl: data[i][4],
        createdBy: data[i][5],
        createdAt: data[i][6] instanceof Date ? data[i][6].toISOString() : String(data[i][6]),
        tipe: data[i][7] || 'Pengeluaran'
      });
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addTeamBuilding(tanggal, keterangan, nominal, buktiUrl, createdBy, tipe) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.TEAM_BUILDING);
    const id = generateId();
    sheet.appendRow([id, tanggal, keterangan, parseFloat(nominal), buktiUrl, createdBy, new Date().toISOString(), tipe || 'Pengeluaran']);
    return { success: true, id: id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteTeamBuilding(id) {
  return deleteRow(CONFIG.SHEETS.TEAM_BUILDING, id);
}

function getSaldoTeamBuilding() {
  try {
    const result = getTeamBuilding();
    if (!result.success) return { success: false, message: result.message };
    const total = result.data.reduce((s, d) => s + d.nominal, 0);
    return { success: true, saldo: total };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// KARYAWAN
// ============================================================
function getKaryawan() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KARYAWAN);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue; // skip baris kosong
      result.push({
        id: String(data[i][0]),
        nama: data[i][1],
        jabatan: data[i][2],
        departemen: data[i][3],
        telepon: data[i][4],
        email: data[i][5],
        tanggalMasuk: data[i][6] instanceof Date
          ? Utilities.formatDate(data[i][6], Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : String(data[i][6] || ''),
        status: data[i][7] || 'Aktif',
        createdAt: data[i][8] instanceof Date ? data[i][8].toISOString() : String(data[i][8] || '')
      });
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addKaryawan(nama, jabatan, departemen, telepon, email, tanggalMasuk, status) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KARYAWAN);
    const id = generateId();
    sheet.appendRow([id, nama, jabatan, departemen, telepon, email, tanggalMasuk, status || 'Aktif', new Date().toISOString()]);
    return { success: true, id: id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateKaryawan(id, nama, jabatan, departemen, telepon, email, tanggalMasuk, status) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KARYAWAN);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.getRange(i + 1, 2, 1, 7).setValues([[nama, jabatan, departemen, telepon, email, tanggalMasuk, status]]);
        return { success: true };
      }
    }
    return { success: false, message: 'Karyawan tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteKaryawan(id) {
  return deleteRow(CONFIG.SHEETS.KARYAWAN, id);
}

// ============================================================
// UPLOAD FILE KE GOOGLE DRIVE (Single + Chunked)
// ============================================================

// Single upload — untuk file kecil (sudah dikompres di browser)
function uploadFileToDrive(base64Data, fileName, mimeType, folderName) {
  try {
    const folder = getOrCreateBuktiFolder(folderName);
    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, mimeType || 'application/octet-stream', fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: 'https://drive.google.com/file/d/' + file.getId() + '/view' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Chunked upload — terima satu chunk base64 dan simpan di CacheService
function uploadChunk(chunkData, chunkIndex, uploadId) {
  try {
    const cache = CacheService.getScriptCache();
    // Buat uploadId baru jika chunk pertama
    const id = uploadId || Utilities.getUuid();
    // Simpan chunk (CacheService max 100KB per item, expiry 6 jam)
    const key = 'chunk_' + id + '_' + chunkIndex;
    // Simpan dalam 2 bagian jika perlu (base64 chunk 200KB = ~266KB chars, CacheService max 100KB)
    if (chunkData.length <= 90000) {
      cache.put(key, chunkData, 21600);
      cache.put('meta_' + id + '_count', String(chunkIndex + 1), 21600);
    } else {
      // Split chunk lagi
      const half = Math.ceil(chunkData.length / 2);
      cache.put(key + '_a', chunkData.substring(0, half), 21600);
      cache.put(key + '_b', chunkData.substring(half), 21600);
      cache.put(key + '_split', '1', 21600);
      cache.put('meta_' + id + '_count', String(chunkIndex + 1), 21600);
    }
    return { success: true, uploadId: id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Finalize — gabungkan semua chunk dan buat file di Drive
function finalizeChunkedUpload(uploadId, fileName, mimeType, folderName) {
  try {
    const cache = CacheService.getScriptCache();
    const countStr = cache.get('meta_' + uploadId + '_count');
    if (!countStr) return { success: false, message: 'Upload session tidak ditemukan atau kedaluwarsa' };

    const totalChunks = parseInt(countStr);
    let fullBase64 = '';

    for (let i = 0; i < totalChunks; i++) {
      const key = 'chunk_' + uploadId + '_' + i;
      const isSplit = cache.get(key + '_split');
      if (isSplit) {
        fullBase64 += (cache.get(key + '_a') || '') + (cache.get(key + '_b') || '');
      } else {
        fullBase64 += cache.get(key) || '';
      }
    }

    if (!fullBase64) return { success: false, message: 'Data chunk kosong' };

    const folder = getOrCreateBuktiFolder(folderName);
    const decoded = Utilities.base64Decode(fullBase64);
    const blob = Utilities.newBlob(decoded, mimeType || 'application/octet-stream', fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return { success: true, url: 'https://drive.google.com/file/d/' + file.getId() + '/view' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Helper: dapatkan atau buat folder bukti
function getOrCreateBuktiFolder(subFolderName) {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty('BUKTI_FOLDER_ID');
  let folder;

  if (folderId) {
    try { folder = DriveApp.getFolderById(folderId); }
    catch(e) { folderId = null; }
  }
  if (!folderId) {
    folder = DriveApp.createFolder('Warehouse FCL - Bukti Invoice');
    props.setProperty('BUKTI_FOLDER_ID', folder.getId());
  }

  // Subfolder per kategori
  const name = subFolderName || 'Umum';
  const subs = folder.getFoldersByName(name);
  return subs.hasNext() ? subs.next() : folder.createFolder(name);
}

// ============================================================
// EKSPOR SOP KE GOOGLE DOCS / TEXT
// ============================================================
function exportSOP() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SOP);
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) return { success: false, message: 'Tidak ada data SOP untuk diekspor' };

    // Buat Google Doc
    const doc = DocumentApp.create('SOP Gudang - Warehouse FCL - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'));
    const body = doc.getBody();

    // Header
    const title = body.appendParagraph('SOP GUDANG — WAREHOUSE FCL');
    title.setHeading(DocumentApp.ParagraphHeading.TITLE);
    title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    body.appendParagraph('Diekspor pada: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMMM yyyy HH:mm'))
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendHorizontalRule();

    // Kelompokkan per kategori
    const grouped = {};
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const kat = data[i][3] || 'Lainnya';
      if (!grouped[kat]) grouped[kat] = [];
      grouped[kat].push({ judul: data[i][1], konten: data[i][2] });
    }

    Object.keys(grouped).forEach(kat => {
      // Heading kategori
      body.appendParagraph(kat).setHeading(DocumentApp.ParagraphHeading.HEADING1);

      grouped[kat].forEach((sop, idx) => {
        // Sub judul
        body.appendParagraph((idx + 1) + '. ' + sop.judul)
          .setHeading(DocumentApp.ParagraphHeading.HEADING2);
        // Konten
        body.appendParagraph(sop.konten || '-');
        body.appendParagraph('');
      });
    });

    // Set sharing
    const docFile = DriveApp.getFileById(doc.getId());
    docFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const url = 'https://docs.google.com/document/d/' + doc.getId() + '/edit';
    return { success: true, url: url };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// PEMBAYARAN TEAM BUILDING
// ============================================================
function getPembayaranTB() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PEMBAYARAN_TB);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      result.push({
        id: data[i][0], karyawanId: data[i][1], karyawanNama: data[i][2],
        periode: data[i][3], nominal: data[i][4], status: data[i][5],
        tanggalBayar: data[i][6], keterangan: data[i][7], createdAt: data[i][8]
      });
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addPembayaranTB(karyawanId, karyawanNama, periode, nominal, status, tanggalBayar, keterangan) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PEMBAYARAN_TB);
    const id = generateId();
    sheet.appendRow([id, karyawanId, karyawanNama, periode, parseFloat(nominal), status, tanggalBayar, keterangan, new Date().toISOString()]);
    return { success: true, id: id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateStatusPembayaran(id, status, tanggalBayar) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PEMBAYARAN_TB);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.getRange(i + 1, 6, 1, 2).setValues([[status, tanggalBayar || new Date().toISOString()]]);
        return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deletePembayaranTB(id) {
  return deleteRow(CONFIG.SHEETS.PEMBAYARAN_TB, id);
}

// ============================================================
// SOP GUDANG
// ============================================================
function getSOP() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SOP);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      result.push({
        id: data[i][0], judul: data[i][1], konten: data[i][2],
        kategori: data[i][3], createdBy: data[i][4], updatedAt: data[i][5]
      });
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addSOP(judul, konten, kategori, createdBy) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SOP);
    const id = generateId();
    sheet.appendRow([id, judul, konten, kategori, createdBy, new Date().toISOString()]);
    return { success: true, id: id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateSOP(id, judul, konten, kategori) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SOP);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.getRange(i + 1, 2, 1, 4).setValues([[judul, konten, kategori, new Date().toISOString()]]);
        return { success: true };
      }
    }
    return { success: false, message: 'SOP tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteSOP(id) {
  return deleteRow(CONFIG.SHEETS.SOP, id);
}

// ============================================================
// STRUKTUR ORGANISASI
// ============================================================
function getOrganisasi() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ORGANISASI);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      result.push({
        id: data[i][0], nama: data[i][1], jabatan: data[i][2],
        atasan: data[i][3], departemen: data[i][4], foto: data[i][5], urutan: data[i][6]
      });
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addOrganisasi(nama, jabatan, atasan, departemen, foto, urutan) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ORGANISASI);
    const id = generateId();
    sheet.appendRow([id, nama, jabatan, atasan, departemen, foto, urutan || 0]);
    return { success: true, id: id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateOrganisasi(id, nama, jabatan, atasan, departemen, foto, urutan) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ORGANISASI);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.getRange(i + 1, 2, 1, 6).setValues([[nama, jabatan, atasan, departemen, foto, urutan]]);
        return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteOrganisasi(id) {
  return deleteRow(CONFIG.SHEETS.ORGANISASI, id);
}

// ============================================================
// DASHBOARD - RINGKASAN
// ============================================================
function getDashboardData() {
  try {
    const saldoGudang = getSaldoGudang();
    const saldoTB = getSaldoTeamBuilding();
    const kasGudang = getKasGudang();   // sudah di-fix, tanggal sudah string
    const teamBuilding = getTeamBuilding(); // sudah di-fix

    let history = [];
    if (kasGudang.success && kasGudang.data) {
      kasGudang.data.forEach(k => {
        history.push({
          tanggal: k.tanggal,
          tipe: k.tipe === 'IN' ? 'Kas Masuk' : 'Kas Keluar',
          keterangan: k.keterangan,
          nominal: k.nominal,
          kategori: 'Kas Gudang'
        });
      });
    }
    if (teamBuilding.success && teamBuilding.data) {
      teamBuilding.data.forEach(t => {
        history.push({
          tanggal: t.tanggal,
          tipe: 'Team Building',
          keterangan: t.keterangan,
          nominal: t.nominal,
          kategori: 'Team Building'
        });
      });
    }

    // Sort aman — tanggal sudah string 'yyyy-MM-dd'
    history.sort((a, b) => {
      const da = a.tanggal ? new Date(a.tanggal) : new Date(0);
      const db = b.tanggal ? new Date(b.tanggal) : new Date(0);
      return db - da;
    });
    history = history.slice(0, 20);

    const kasData = kasGudang.data || [];
    const totalKasIn  = kasData.filter(k => k.tipe === 'IN').reduce((s, k) => s + k.nominal, 0);
    const totalKasOut = kasData.filter(k => k.tipe === 'OUT').reduce((s, k) => s + k.nominal, 0);

    return {
      success: true,
      saldoGudang: saldoGudang.saldo || 0,
      saldoTB: saldoTB.saldo || 0,
      history: history,
      totalKasIn: totalKasIn,
      totalKasOut: totalKasOut,
      kasData: kasGudang.data || [],
      tbData: teamBuilding.data || []
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// USER MANAGEMENT
// ============================================================
function getUsers() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      result.push({ id: data[i][0], username: data[i][1], nama: data[i][3], role: data[i][4], createdAt: data[i][5] });
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addUser(username, password, nama, role) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === username) return { success: false, message: 'Username sudah ada' };
    }
    const id = generateId();
    sheet.appendRow([id, username, hashPassword(password), nama, role || 'user', new Date().toISOString()]);
    return { success: true, id: id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function changePassword(username, oldPassword, newPassword) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === username && data[i][2] === hashPassword(oldPassword)) {
        sheet.getRange(i + 1, 3).setValue(hashPassword(newPassword));
        return { success: true };
      }
    }
    return { success: false, message: 'Password lama salah' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// HELPERS
// ============================================================
function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID');
  if (!ssId) {
    const result = setupDatabase();
    ssId = result.spreadsheetId;
  }
  return SpreadsheetApp.openById(ssId);
}

function getSheet(sheetName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    setupDatabase();
    sheet = ss.getSheetByName(sheetName);
  }
  return sheet;
}

function deleteRow(sheetName, id) {
  try {
    const sheet = getSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function generateId() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 16);
}

function hashPassword(password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function getSpreadsheetUrl() {
  try {
    const ss = getSpreadsheet();
    return { success: true, url: ss.getUrl() };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
