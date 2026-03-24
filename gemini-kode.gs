// ============================================================
// GUDANG FCL - Google Apps Script Backend
// Code.gs - Main Server-Side Logic
// ============================================================

// ---- KONFIGURASI ----
const CONFIG = {
  SPREADSHEET_ID: '',
  SHEETS: {
    USERS: 'Users',
    KAS_GUDANG: 'KasGudang',
    TEAM_BUILDING: 'TeamBuilding',
    KARYAWAN: 'Karyawan',
    PEMBAYARAN_TB: 'PembayaranTB',
    SOP: 'SOP',
    ORGANISASI: 'Organisasi',
    SETTINGS: 'Settings',
    STOCK: 'Stock',
    SURAT_JALAN_MASUK: 'SuratJalanMasuk',
    SURAT_JALAN_MASUK_DETAIL: 'SuratJalanMasukDetail',
    SURAT_JALAN_KELUAR: 'SuratJalanKeluar',
    SURAT_JALAN_KELUAR_DETAIL: 'SuratJalanKeluarDetail',
    ORDER: 'Order',
    ORDER_DETAIL: 'OrderDetail'
  },
  VERSION: '1.1.0'
};

// ============================================================
// ENTRY POINT
// ============================================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Gudang FCL')
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
    ss = SpreadsheetApp.create('Gudang FCL - Database');
    ssId = ss.getId();
    props.setProperty('SPREADSHEET_ID', ssId);
    Logger.log('Spreadsheet baru dibuat: ' + ssId);
  } else {
    ss = SpreadsheetApp.openById(ssId);
  }

  // Buat semua sheet yang diperlukan
  setupSheet(ss, CONFIG.SHEETS.USERS, ['id', 'username', 'password', 'nama', 'role', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.KAS_GUDANG, ['id', 'tanggal', 'tipe', 'keterangan', 'nominal', 'buktiUrl', 'createdBy', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.TEAM_BUILDING, ['id', 'tanggal', 'keterangan', 'nominal', 'buktiUrl', 'createdBy', 'createdAt', 'tipe']);
  setupSheet(ss, CONFIG.SHEETS.KARYAWAN, ['id', 'nama', 'jabatan', 'departemen', 'telepon', 'email', 'tanggalMasuk', 'status', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.PEMBAYARAN_TB, ['id', 'karyawanId', 'karyawanNama', 'periode', 'nominal', 'status', 'tanggalBayar', 'keterangan', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.SOP, ['id', 'judul', 'konten', 'kategori', 'createdBy', 'updatedAt']);
  setupSheet(ss, CONFIG.SHEETS.ORGANISASI, ['id', 'nama', 'jabatan', 'atasan', 'departemen', 'foto', 'urutan']);
  setupSheet(ss, CONFIG.SHEETS.SETTINGS, ['key', 'value']);
  // Inventory
  setupSheet(ss, CONFIG.SHEETS.STOCK, ['id','sku','nama','barcode','batch','expDate','satuan','stok','stokMin','kategori','lokasi','createdAt','updatedAt']);
  setupSheet(ss, CONFIG.SHEETS.SURAT_JALAN_MASUK, ['id','noSJ','tanggal','supplier','keterangan','createdBy','createdAt']);
  setupSheet(ss, CONFIG.SHEETS.SURAT_JALAN_MASUK_DETAIL, ['id','sjId','noSJ','stockId','sku','nama','qty','satuan','batch','expDate']);
  setupSheet(ss, CONFIG.SHEETS.SURAT_JALAN_KELUAR, ['id','noSJ','tanggal','tujuan','keterangan','createdBy','createdAt']);
  setupSheet(ss, CONFIG.SHEETS.SURAT_JALAN_KELUAR_DETAIL, ['id','sjId','noSJ','stockId','sku','nama','qty','satuan','batch','expDate']);
  setupSheet(ss, CONFIG.SHEETS.ORDER, ['id','noOrder','tanggal','pelanggan','alamat','status','totalItem','keterangan','createdBy','createdAt','sentAt']);
  setupSheet(ss, CONFIG.SHEETS.ORDER_DETAIL, ['id','orderId','noOrder','stockId','sku','nama','qty','satuan']);
  
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
      if (!data[i][0]) continue;
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
    const total = result.data.reduce((s, d) => {
      if (d.tipe === 'Pemasukan') return s + d.nominal;
      return s - d.nominal;
    }, 0);
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
      if (!data[i][0]) continue;
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

function uploadChunk(chunkData, chunkIndex, uploadId) {
  try {
    const cache = CacheService.getScriptCache();
    const id = uploadId || Utilities.getUuid();
    const key = 'chunk_' + id + '_' + chunkIndex;
    if (chunkData.length <= 90000) {
      cache.put(key, chunkData, 21600);
      cache.put('meta_' + id + '_count', String(chunkIndex + 1), 21600);
    } else {
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

function getOrCreateBuktiFolder(subFolderName) {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty('BUKTI_FOLDER_ID');
  let folder;
  if (folderId) {
    try { folder = DriveApp.getFolderById(folderId); }
    catch(e) { folderId = null; }
  }
  if (!folderId) {
    folder = DriveApp.createFolder('Gudang FCL - Bukti Invoice');
    props.setProperty('BUKTI_FOLDER_ID', folder.getId());
  }

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
    const doc = DocumentApp.create('SOP Gudang - Gudang FCL - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'));
    const body = doc.getBody();
    const title = body.appendParagraph('SOP GUDANG — GUDANG FCL');
    title.setHeading(DocumentApp.ParagraphHeading.TITLE);
    title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('Diekspor pada: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMMM yyyy HH:mm'))
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendHorizontalRule();
    const grouped = {};
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const kat = data[i][3] || 'Lainnya';
      if (!grouped[kat]) grouped[kat] = [];
      grouped[kat].push({ judul: data[i][1], konten: data[i][2] });
    }

    Object.keys(grouped).forEach(kat => {
      body.appendParagraph(kat).setHeading(DocumentApp.ParagraphHeading.HEADING1);
      grouped[kat].forEach((sop, idx) => {
        body.appendParagraph((idx + 1) + '. ' + sop.judul)
          .setHeading(DocumentApp.ParagraphHeading.HEADING2);
        body.appendParagraph(sop.konten || '-');
        body.appendParagraph('');
      });
    });
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
      if (!data[i][0]) continue;
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
      if (!data[i][0]) continue;
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
      if (!data[i][0]) continue;
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
    const kasGudang = getKasGudang();
    const teamBuilding = getTeamBuilding();

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
      if (!data[i][0]) continue;
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

// ============================================================
// INVENTORY — STOCK BARANG
// ============================================================
function generateSKU(nama) {
  const prefix = nama.replace(/[^A-Za-z]/g,'').toUpperCase().substring(0,3) || 'SKU';
  const num = String(Date.now()).slice(-5);
  return prefix + '-' + num;
}

function getStock() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.STOCK);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({
        id: data[i][0], sku: data[i][1], nama: data[i][2],
        barcode: data[i][3], batch: data[i][4],
        expDate: data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][5]||''),
        satuan: data[i][6], stok: parseFloat(data[i][7])||0,
        stokMin: parseFloat(data[i][8])||0, kategori: data[i][9],
        lokasi: data[i][10],
        createdAt: data[i][11] instanceof Date ? data[i][11].toISOString() : String(data[i][11]||''),
        updatedAt: data[i][12] instanceof Date ? data[i][12].toISOString() : String(data[i][12]||'')
      });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addStock(skuInput, nama, barcode, batch, expDate, satuan, stok, stokMin, kategori, lokasi) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.STOCK);
    const id = generateId();
    const sku = (skuInput && skuInput.trim() !== '') ? skuInput.trim() : generateSKU(nama);
    const now = new Date().toISOString();
    sheet.appendRow([id, sku, nama, barcode, batch, expDate, satuan, parseFloat(stok)||0, parseFloat(stokMin)||0, kategori, lokasi, now, now]);
    return { success: true, id, sku };
  } catch(e) { return { success: false, message: e.message }; }
}

function updateStock(id, sku, nama, barcode, batch, expDate, satuan, stok, stokMin, kategori, lokasi) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.STOCK);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.getRange(i+1, 2, 1, 10).setValues([[sku, nama, barcode, batch, expDate, satuan, parseFloat(stok)||0, parseFloat(stokMin)||0, kategori, lokasi]]);
        sheet.getRange(i+1, 13).setValue(new Date().toISOString());
        return { success: true };
      }
    }
    return { success: false, message: 'Barang tidak ditemukan' };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteStock(id) { return deleteRow(CONFIG.SHEETS.STOCK, id); }

function updateStokQty(id, delta) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.STOCK);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        const newStok = (parseFloat(data[i][7])||0) + delta;
        if (newStok < 0) return { success: false, message: 'Stok tidak cukup! Sisa: ' + (parseFloat(data[i][7])||0) };
        sheet.getRange(i+1,8).setValue(newStok);
        sheet.getRange(i+1,13).setValue(new Date().toISOString());
        return { success: true, newStok };
      }
    }
    return { success: false, message: 'Barang tidak ditemukan' };
  } catch(e) { return { success: false, message: e.message }; }
}

// ============================================================
// SURAT JALAN MASUK
// ============================================================
function generateNoSJ(prefix) {
  const d = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  return prefix + '/' + d + '/' + String(Math.floor(Math.random()*9000)+1000);
}

function getSuratJalanMasuk() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SURAT_JALAN_MASUK);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({
        id: data[i][0], noSJ: data[i][1],
        tanggal: data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][2]||''),
        supplier: data[i][3], keterangan: data[i][4],
        createdBy: data[i][5],
        createdAt: data[i][6] instanceof Date ? data[i][6].toISOString() : String(data[i][6]||'')
      });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addSuratJalanMasuk(tanggal, supplier, keterangan, items, createdBy) {
  try {
    const noSJ = generateNoSJ('SJM');
    const id = generateId();
    const sheet = getSheet(CONFIG.SHEETS.SURAT_JALAN_MASUK);
    sheet.appendRow([id, noSJ, tanggal, supplier, keterangan, createdBy, new Date().toISOString()]);

    const detSheet = getSheet(CONFIG.SHEETS.SURAT_JALAN_MASUK_DETAIL);
    const parsedItems = typeof items === 'string' ? JSON.parse(items) : items;
    parsedItems.forEach(item => {
      detSheet.appendRow([generateId(), id, noSJ, item.stockId, item.sku, item.nama, parseFloat(item.qty)||0, item.satuan, item.batch||'', item.expDate||'']);
      updateStokQty(item.stockId, parseFloat(item.qty)||0);
    });
    return { success: true, id, noSJ };
  } catch(e) { return { success: false, message: e.message }; }
}

function getSJMasukDetail(sjId) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SURAT_JALAN_MASUK_DETAIL);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0] || data[i][1] !== sjId) continue;
      result.push({ id:data[i][0], sjId:data[i][1], noSJ:data[i][2], stockId:data[i][3], sku:data[i][4], nama:data[i][5], qty:parseFloat(data[i][6])||0, satuan:data[i][7], batch:data[i][8], expDate:data[i][9] });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function getSJDetailData(sjId, tipe) {
  try {
    const sheetName = tipe === 'masuk' ? CONFIG.SHEETS.SURAT_JALAN_MASUK_DETAIL : CONFIG.SHEETS.SURAT_JALAN_KELUAR_DETAIL;
    const sheet = getSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0] || data[i][1] !== sjId) continue;
      result.push({ sku:data[i][4], nama:data[i][5], qty:parseFloat(data[i][6])||0, satuan:data[i][7], batch:data[i][8], expDate:data[i][9] });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

// ============================================================
// SURAT JALAN KELUAR
// ============================================================
function getSuratJalanKeluar() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SURAT_JALAN_KELUAR);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({
        id: data[i][0], noSJ: data[i][1],
        tanggal: data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][2]||''),
        tujuan: data[i][3], keterangan: data[i][4],
        createdBy: data[i][5],
        createdAt: data[i][6] instanceof Date ? data[i][6].toISOString() : String(data[i][6]||'')
      });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addSuratJalanKeluar(tanggal, tujuan, keterangan, items, createdBy) {
  try {
    const noSJ = generateNoSJ('SJK');
    const id = generateId();
    const sheet = getSheet(CONFIG.SHEETS.SURAT_JALAN_KELUAR);
    sheet.appendRow([id, noSJ, tanggal, tujuan, keterangan, createdBy, new Date().toISOString()]);

    const detSheet = getSheet(CONFIG.SHEETS.SURAT_JALAN_KELUAR_DETAIL);
    const parsedItems = typeof items === 'string' ? JSON.parse(items) : items;
    for (const item of parsedItems) {
      const res = updateStokQty(item.stockId, -(parseFloat(item.qty)||0));
      if (!res.success) return { success: false, message: 'Barang "' + item.nama + '": ' + res.message };
      detSheet.appendRow([generateId(), id, noSJ, item.stockId, item.sku, item.nama, parseFloat(item.qty)||0, item.satuan, item.batch||'', item.expDate||'']);
    }
    return { success: true, id, noSJ };
  } catch(e) { return { success: false, message: e.message }; }
}

// ============================================================
// ORDER
// ============================================================
function getOrders() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ORDER);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({
        id:data[i][0], noOrder:data[i][1],
        tanggal: data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][2]||''),
        pelanggan:data[i][3], alamat:data[i][4], status:data[i][5],
        totalItem:parseFloat(data[i][6])||0, keterangan:data[i][7],
        createdBy:data[i][8],
        createdAt: data[i][9] instanceof Date ? data[i][9].toISOString() : String(data[i][9]||''),
        sentAt: data[i][10] instanceof Date ? data[i][10].toISOString() : String(data[i][10]||'')
      });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addOrder(tanggal, pelanggan, alamat, keterangan, items, createdBy) {
  try {
    const noOrder = generateNoSJ('WHFCL');
    const id = generateId();
    const parsedItems = typeof items === 'string' ? JSON.parse(items) : items;
    const totalItem = parsedItems.reduce((s,x) => s + (parseFloat(x.qty)||0), 0);
    const sheet = getSheet(CONFIG.SHEETS.ORDER);
    sheet.appendRow([id, noOrder, tanggal, pelanggan, alamat, 'Pending', totalItem, keterangan, createdBy, new Date().toISOString(), '']);

    const detSheet = getSheet(CONFIG.SHEETS.ORDER_DETAIL);
    parsedItems.forEach(item => {
      detSheet.appendRow([generateId(), id, noOrder, item.stockId, item.sku, item.nama, parseFloat(item.qty)||0, item.satuan]);
    });
    return { success: true, id, noOrder };
  } catch(e) { return { success: false, message: e.message }; }
}

function getOrderDetail(orderId) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ORDER_DETAIL);
    const data = sheet.getDataRange().getValues();
    const stockSheet = getSheet(CONFIG.SHEETS.STOCK);
    const stockData = stockSheet.getDataRange().getValues();
    const stockMap = {};
    for (let j = 1; j < stockData.length; j++) {
      if (stockData[j][0]) {
         stockMap[stockData[j][0]] = {
           batch: stockData[j][4] || '-',
           expDate: stockData[j][5] instanceof Date ? Utilities.formatDate(stockData[j][5], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(stockData[j][5]||'-')
         };
      }
    }

    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0] || data[i][1] !== orderId) continue;
      const sId = data[i][3];
      const st = stockMap[sId] || { batch: '-', expDate: '-' };
      
      result.push({ 
        id:data[i][0], orderId:data[i][1], noOrder:data[i][2], 
        stockId:sId, sku:data[i][4], nama:data[i][5], 
        qty:parseFloat(data[i][6])||0, satuan:data[i][7],
        batch: st.batch, expDate: st.expDate
      });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function kirimOrder(orderId) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ORDER);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== orderId) continue;
      if (data[i][5] === 'Terkirim') return { success: false, message: 'Order sudah terkirim' };
      
      const details = getOrderDetail(orderId);
      if (!details.success) return details;
      
      const rolled = [];
      for (const item of details.data) {
        const res = updateStokQty(item.stockId, -(parseFloat(item.qty)||0));
        if (!res.success) {
          rolled.forEach(r => updateStokQty(r.stockId, r.qty));
          return { success: false, message: 'Stok "' + item.nama + '": ' + res.message };
        }
        rolled.push({ stockId: item.stockId, qty: item.qty });
      }

      sheet.getRange(i+1, 6).setValue('Terkirim');
      sheet.getRange(i+1, 11).setValue(new Date().toISOString());
      return { success: true };
    }
    return { success: false, message: 'Order tidak ditemukan' };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteOrder(id) { return deleteRow(CONFIG.SHEETS.ORDER, id); }

// ============================================================
// ANALISIS PEMAKAIAN STOCK
// ============================================================
function getAnalisisStock() {
  try {
    const now = new Date();
    const oneWeekAgo = new Date(now - 7*24*60*60*1000);
    const oneMonthAgo = new Date(now.getFullYear(), now.getMonth()-1, now.getDate());
    
    const sjkDet = getSheet(CONFIG.SHEETS.SURAT_JALAN_KELUAR_DETAIL).getDataRange().getValues();
    const sjkHead = getSheet(CONFIG.SHEETS.SURAT_JALAN_KELUAR).getDataRange().getValues();
    const ordDet = getSheet(CONFIG.SHEETS.ORDER_DETAIL).getDataRange().getValues();
    const ordHead = getSheet(CONFIG.SHEETS.ORDER).getDataRange().getValues();

    const sjkDateMap = {};
    for (let i = 1; i < sjkHead.length; i++) {
      if (!sjkHead[i][0]) continue;
      sjkDateMap[sjkHead[i][0]] = sjkHead[i][2] instanceof Date ? sjkHead[i][2] : new Date(sjkHead[i][2]);
    }
    
    const ordDateMap = {};
    for (let i = 1; i < ordHead.length; i++) {
      if (!ordHead[i][0] || ordHead[i][5] !== 'Terkirim') continue;
      ordDateMap[ordHead[i][0]] = ordHead[i][9] instanceof Date ? ordHead[i][9] : new Date(ordHead[i][9]);
    }

    const weekly = {}, monthly = {};
    for (let i = 1; i < sjkDet.length; i++) {
      if (!sjkDet[i][0]) continue;
      const tgl = sjkDateMap[sjkDet[i][1]];
      if (!tgl) continue;
      const key = sjkDet[i][4] + '|' + sjkDet[i][5];
      const qty = parseFloat(sjkDet[i][6])||0;
      if (tgl >= oneWeekAgo) weekly[key] = (weekly[key]||0) + qty;
      if (tgl >= oneMonthAgo) monthly[key] = (monthly[key]||0) + qty;
    }

    for (let i = 1; i < ordDet.length; i++) {
      if (!ordDet[i][0]) continue;
      const tgl = ordDateMap[ordDet[i][1]];
      if (!tgl) continue;
      const key = ordDet[i][4] + '|' + ordDet[i][5];
      const qty = parseFloat(ordDet[i][6])||0;
      if (tgl >= oneWeekAgo) weekly[key] = (weekly[key]||0) + qty;
      if (tgl >= oneMonthAgo) monthly[key] = (monthly[key]||0) + qty;
    }

    const stockRes = getStock();
    const stockMap = {};
    (stockRes.data||[]).forEach(s => { stockMap[s.sku] = s; });
    const allKeys = new Set([...Object.keys(weekly), ...Object.keys(monthly)]);
    const rows = [];
    allKeys.forEach(key => {
      const [sku, nama] = key.split('|');
      const st = stockMap[sku] || {};
      rows.push({
        sku, nama,
        stokSaat: st.stok || 0,
        satuan: st.satuan || '',
        minggu: weekly[key] || 0,
        bulan: monthly[key] || 0,
        rataHarian: Math.round(((monthly[key]||0) / 30) * 100) / 100,
        statusStok: (st.stok||0) <= (st.stokMin||0) ? 'Kritis' : (st.stok||0) <= (st.stokMin||0)*2 ? 'Rendah' : 'Aman'
      });
    });
    rows.sort((a,b) => b.bulan - a.bulan);

    return { success: true, data: rows };
  } catch(e) { return { success: false, message: e.message }; }
}

// ============================================================
// FITUR BARU: IMPORT ORDER BULK
// ============================================================
function importOrdersBulk(ordersData, createdBy) {
  try {
    const parsedOrders = typeof ordersData === 'string' ? JSON.parse(ordersData) : ordersData;
    const sheetOrd = getSheet(CONFIG.SHEETS.ORDER);
    const detSheet = getSheet(CONFIG.SHEETS.ORDER_DETAIL);
    const stockSheet = getSheet(CONFIG.SHEETS.STOCK);
    const stockData = stockSheet.getDataRange().getValues();
    
    const stockMap = {};
    for(let i = 1; i < stockData.length; i++) {
       if(stockData[i][0]) stockMap[stockData[i][1]] = { id: stockData[i][0], nama: stockData[i][2], satuan: stockData[i][6] };
    }

    let importedCount = 0;
    parsedOrders.forEach(ord => {
       const noOrder = generateNoSJ('WHFCL');
       const id = generateId();
       let totalQty = 0;
       const validItems = [];
       
       ord.items.forEach(it => {
          const st = stockMap[it.sku];
          if(st) {
             validItems.push({ stockId: st.id, sku: it.sku, nama: st.nama, qty: parseFloat(it.qty)||0, satuan: st.satuan });
             totalQty += (parseFloat(it.qty)||0);
          }
       });
       
       if(validItems.length > 0) {
          sheetOrd.appendRow([id, noOrder, ord.tanggal, ord.pelanggan, ord.alamat, 'Pending', totalQty, ord.keterangan, createdBy, new Date().toISOString(), '']);
          validItems.forEach(item => {
             detSheet.appendRow([generateId(), id, noOrder, item.stockId, item.sku, item.nama, item.qty, item.satuan]);
          });
          importedCount++;
       }
    });
    return { success: true, count: importedCount };
  } catch(e) { 
    return { success: false, message: e.message }; 
  }
}
