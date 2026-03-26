// ============================================================
// GUDANG FCL - Google Apps Script Backend
// Code.gs - Main Server-Side Logic
// ============================================================

const CONFIG = {
  SPREADSHEET_ID: '',
  SHEETS: {
    USERS: 'Users',
    KAS_GUDANG: 'KasGudang',
    TEAM_BUILDING: 'TeamBuilding',
    EXPENSE: 'Expense',
    KARYAWAN: 'Karyawan',
    IJIN: 'Ijin',
    LEMBUR: 'Lembur',
    LAPORAN_KERJA: 'LaporanKerja',
    SOP: 'SOP',
    ORGANISASI: 'Organisasi',
    SETTINGS: 'Settings',
    STOCK: 'Stock',
    SURAT_JALAN_MASUK: 'SuratJalanMasuk',
    SURAT_JALAN_MASUK_DETAIL: 'SuratJalanMasukDetail',
    SURAT_JALAN_KELUAR: 'SuratJalanKeluar',
    SURAT_JALAN_KELUAR_DETAIL: 'SuratJalanKeluarDetail',
    ORDER: 'Order',
    ORDER_DETAIL: 'OrderDetail',
    RETUR: 'Retur',
    RETUR_DETAIL: 'ReturDetail',
    HANDOVER: 'Handover',
    KLAIM: 'Klaim',
    TUGAS_PROJECT: 'TugasProject'
  }
};

// ============================================================
// ENTRY POINT
// ============================================================
function doGet(e) {
  // Menggunakan 'Index' dengan I besar menyesuaikan default Google Apps Script
  var html = HtmlService.createHtmlOutputFromFile('Index'); 
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return html;
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
  } else {
    ss = SpreadsheetApp.openById(ssId);
  }

  setupSheet(ss, CONFIG.SHEETS.USERS, ['id', 'username', 'password', 'nama', 'role', 'createdAt', 'permissions']);
  setupSheet(ss, CONFIG.SHEETS.KAS_GUDANG, ['id', 'tanggal', 'tipe', 'keterangan', 'nominal', 'buktiUrl', 'createdBy', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.TEAM_BUILDING, ['id', 'tanggal', 'keterangan', 'nominal', 'buktiUrl', 'createdBy', 'createdAt', 'tipe']);
  setupSheet(ss, CONFIG.SHEETS.EXPENSE, ['id', 'tanggal', 'perusahaan', 'kategori', 'keterangan', 'nominal', 'bank', 'rekening', 'createdBy', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.KARYAWAN, ['id', 'nama', 'jabatan', 'departemenLama', 'telepon', 'email', 'tanggalMasuk', 'status', 'createdAt', 'tanggalSelesai']);
  setupSheet(ss, CONFIG.SHEETS.IJIN, ['id', 'tanggal', 'nama', 'jenis', 'keterangan', 'bukti', 'status', 'createdBy', 'createdAt', 'history']);
  setupSheet(ss, CONFIG.SHEETS.LEMBUR, ['id', 'tanggal', 'nama', 'divisi', 'jamMulai', 'jamSelesai', 'keterangan', 'status', 'createdBy', 'createdAt', 'history']);
  setupSheet(ss, CONFIG.SHEETS.LAPORAN_KERJA, ['id', 'tanggal', 'divisi', 'pic', 'totalOrang', 'perbantuan', 'pengurangan', 'jamLembur', 'totalJamKerja', 'kendala', 'totalStaff', 'totalAdmin', 'totalOrder', 'createdBy', 'createdAt', 'sisaOrder']);
  setupSheet(ss, CONFIG.SHEETS.SOP, ['id', 'judul', 'konten', 'kategori', 'createdBy', 'updatedAt']);
  setupSheet(ss, CONFIG.SHEETS.ORGANISASI, ['id', 'nama', 'jabatan', 'atasan', 'departemen', 'foto', 'urutan']);
  setupSheet(ss, CONFIG.SHEETS.STOCK, ['id','sku','nama','barcode','batch','expDate','satuan','stok','stokMin','kategori','lokasi','createdAt','updatedAt']);
  setupSheet(ss, CONFIG.SHEETS.SURAT_JALAN_MASUK, ['id','noSJ','tanggal','supplier','keterangan','createdBy','createdAt']);
  setupSheet(ss, CONFIG.SHEETS.SURAT_JALAN_MASUK_DETAIL, ['id','sjId','noSJ','stockId','sku','nama','qty','satuan','batch','expDate']);
  setupSheet(ss, CONFIG.SHEETS.SURAT_JALAN_KELUAR, ['id','noSJ','tanggal','tujuan','keterangan','createdBy','createdAt']);
  setupSheet(ss, CONFIG.SHEETS.SURAT_JALAN_KELUAR_DETAIL, ['id','sjId','noSJ','stockId','sku','nama','qty','satuan','batch','expDate']);
  setupSheet(ss, CONFIG.SHEETS.ORDER, ['id','noOrder','tanggal','pelanggan','alamat','status','totalItem','keterangan','createdBy','createdAt','sentAt']);
  setupSheet(ss, CONFIG.SHEETS.ORDER_DETAIL, ['id','orderId','noOrder','stockId','sku','nama','qty','satuan']);
  setupSheet(ss, CONFIG.SHEETS.RETUR, ['id','noRetur','tanggal','sumber','alasan','keterangan','createdBy','createdAt']);
  setupSheet(ss, CONFIG.SHEETS.RETUR_DETAIL, ['id','returId','noRetur','stockId','sku','nama','qty','satuan','batch','expDate']);
  setupSheet(ss, CONFIG.SHEETS.HANDOVER, ['id', 'tanggal', 'pic', 'resi', 'pengerjaan', 'keterangan', 'status', 'createdBy', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.KLAIM, ['id', 'tanggal', 'pic', 'resi', 'harga', 'keterangan', 'status', 'createdBy', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.TUGAS_PROJECT, ['id','judul','assignee','assigneeName','prioritas','tanggalMulai','deadline','targetHari','status','kategori','deskripsi','createdBy','createdAt','updatedAt','log']);

  const usersSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
  if (usersSheet.getLastRow() <= 1) {
    usersSheet.appendRow([generateId(), 'admin', hashPassword('admin123'), 'Administrator', 'admin', new Date().toISOString(), '[]']);
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
  }
  return sheet;
}

// ============================================================
// FUNGSI UTILITIES DASAR
// ============================================================
function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID');
  if (!ssId) { const result = setupDatabase(); ssId = result.spreadsheetId; }
  return SpreadsheetApp.openById(ssId);
}

function getSheet(sheetName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) { setupDatabase(); sheet = ss.getSheetByName(sheetName); }
  return sheet;
}

function deleteRow(sheetName, id) {
  try {
    const sheet = getSheet(sheetName); 
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) { 
        sheet.deleteRow(i + 1); 
        return { success: true }; 
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

function generateId() { 
  return Utilities.getUuid().replace(/-/g, '').substring(0, 16); 
}

function hashPassword(password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function getSpreadsheetUrl() { 
  try { return { success: true, url: getSpreadsheet().getUrl() }; } 
  catch (e) { return { success: false, message: e.message }; } 
}

// ============================================================
// AUTENTIKASI & MANAJEMEN USER
// ============================================================
function login(username, password) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === username && data[i][2] === hashPassword(password)) {
        return { 
          success: true, 
          user: { 
            id: data[i][0], 
            username: data[i][1], 
            nama: data[i][3], 
            role: data[i][4], 
            permissions: data[i][6] || '[]' 
          } 
        };
      }
    }
    return { success: false, message: 'Username atau password salah' };
  } catch (e) { return { success: false, message: e.message }; }
}

function getUsers() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({ 
        id: data[i][0], 
        username: data[i][1], 
        nama: data[i][3], 
        role: data[i][4], 
        createdAt: data[i][5], 
        permissions: data[i][6] || '[]' 
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function addUser(username, password, nama, role, permissions) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === username) return { success: false, message: 'Username sudah ada' };
    }
    const id = generateId();
    sheet.appendRow([id, username, hashPassword(password), nama, role || 'user', new Date().toISOString(), permissions || '[]']);
    return { success: true, id: id };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateUser(id, username, password, nama, role, permissions) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i + 1, 2).setValue(username);
        if (password) sheet.getRange(i + 1, 3).setValue(hashPassword(password));
        sheet.getRange(i + 1, 4).setValue(nama);
        sheet.getRange(i + 1, 5).setValue(role);
        sheet.getRange(i + 1, 7).setValue(permissions || '[]');
        return { success: true };
      }
    }
    return { success: false, message: 'User tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteUser(id) { 
  return deleteRow(CONFIG.SHEETS.USERS, id); 
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
  } catch (e) { return { success: false, message: e.message }; }
}

// ============================================================
// WORKFLOW APPROVAL
// ============================================================
function getPendingApprovals() {
  try {
    return { 
       success: true, 
       ijin: getIjin(),
       lembur: getLembur()
    };
  } catch (e) { return { success: false, message: e.message }; }
}

function processApprovalStatus(tipe, id, action, userNama, userRole, reason) {
  try {
    let sheetName = tipe === 'ijin' ? CONFIG.SHEETS.IJIN : CONFIG.SHEETS.LEMBUR;
    const sheet = getSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    
    let statusCol = tipe === 'ijin' ? 7 : 8; 
    let historyCol = tipe === 'ijin' ? 10 : 11; 

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        let currentStatus = data[i][statusCol - 1] === 'Pending' ? 'Pending HR' : data[i][statusCol - 1];
        let newStatus = '';
        
        if (action === 'Reject') {
          newStatus = 'Ditolak';
        } else if (action === 'Approve') {
          if (userRole === 'admin') newStatus = 'Disetujui';
          else if (currentStatus === 'Pending HR') newStatus = 'Pending TL';
          else if (currentStatus === 'Pending TL') newStatus = 'Pending VS';
          else if (currentStatus === 'Pending VS') newStatus = 'Pending SPV';
          else if (currentStatus === 'Pending SPV') newStatus = 'Disetujui';
          else newStatus = 'Disetujui'; 
        }
        
        sheet.getRange(i + 1, statusCol).setValue(newStatus);
        
        let historyRaw = data[i][historyCol - 1] || '[]';
        let historyArr = [];
        try { historyArr = JSON.parse(historyRaw); } catch(e) { historyArr = []; }
        
        historyArr.push({
           date: new Date().toISOString(), action: action, status: newStatus,
           by: userNama, role: userRole, reason: reason || ''
        });
        
        sheet.getRange(i + 1, historyCol).setValue(JSON.stringify(historyArr));
        return { success: true, newStatus: newStatus };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

// ============================================================
// DATA KARYAWAN
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
        telepon: data[i][4],
        email: data[i][5],
        tanggalMasuk: data[i][6] instanceof Date ? Utilities.formatDate(data[i][6], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][6] || ''),
        status: data[i][7] || 'Tetap',
        createdAt: data[i][8] instanceof Date ? data[i][8].toISOString() : String(data[i][8] || ''),
        tanggalSelesai: data[i][9] instanceof Date ? Utilities.formatDate(data[i][9], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][9] || '')
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function addKaryawan(nama, jabatan, telepon, email, tanggalMasuk, status, tanggalSelesai) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KARYAWAN);
    const id = generateId();
    sheet.appendRow([id, nama, jabatan, '', telepon, email, tanggalMasuk, status || 'Tetap', new Date().toISOString(), tanggalSelesai || '']);
    return { success: true, id: id };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateKaryawan(id, nama, jabatan, telepon, email, tanggalMasuk, status, tanggalSelesai) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KARYAWAN);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i + 1, 2, 1, 7).setValues([[nama, jabatan, '', telepon, email, tanggalMasuk, status]]);
        sheet.getRange(i + 1, 10).setValue(tanggalSelesai || '');
        return { success: true };
      }
    }
    return { success: false, message: 'Karyawan tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteKaryawan(id) { 
  return deleteRow(CONFIG.SHEETS.KARYAWAN, id); 
}

// ============================================================
// PENGAJUAN IJIN / CUTI
// ============================================================
function getIjin() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.IJIN);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({
        id: data[i][0],
        tanggal: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]),
        nama: data[i][2],
        jenis: data[i][3],
        keterangan: data[i][4],
        bukti: data[i][5],
        status: data[i][6],
        createdBy: data[i][7],
        createdAt: data[i][8] instanceof Date ? data[i][8].toISOString() : String(data[i][8]),
        history: data[i][9] || '[]'
      });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addIjin(tanggal, nama, jenis, keterangan, bukti, createdBy) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.IJIN);
    const historyArr = [{ date: new Date().toISOString(), action: 'Diajukan', status: 'Pending HR', by: createdBy, role: 'Pemohon', reason: '' }];
    sheet.appendRow([generateId(), tanggal, nama, jenis, keterangan, bukti, 'Pending HR', createdBy, new Date().toISOString(), JSON.stringify(historyArr)]);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteIjin(id) { 
  return deleteRow(CONFIG.SHEETS.IJIN, id); 
}

// ============================================================
// PENGAJUAN LEMBUR
// ============================================================
function getLembur() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.LEMBUR);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({
        id: data[i][0],
        tanggal: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]),
        nama: data[i][2],
        divisi: data[i][3],
        jamMulai: data[i][4] instanceof Date ? Utilities.formatDate(data[i][4], Session.getScriptTimeZone(), 'HH:mm') : String(data[i][4]),
        jamSelesai: data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), 'HH:mm') : String(data[i][5]),
        keterangan: data[i][6],
        status: data[i][7],
        createdBy: data[i][8],
        createdAt: data[i][9] instanceof Date ? data[i][9].toISOString() : String(data[i][9]),
        history: data[i][10] || '[]'
      });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addLembur(tanggal, nama, divisi, jamMulai, jamSelesai, keterangan, createdBy) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.LEMBUR);
    const historyArr = [{ date: new Date().toISOString(), action: 'Diajukan', status: 'Pending HR', by: createdBy, role: 'Pemohon', reason: '' }];
    sheet.appendRow([generateId(), tanggal, nama, divisi, jamMulai, jamSelesai, keterangan, 'Pending HR', createdBy, new Date().toISOString(), JSON.stringify(historyArr)]);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteLembur(id) { 
  return deleteRow(CONFIG.SHEETS.LEMBUR, id); 
}

// ============================================================
// DASHBOARD & KEUANGAN
// ============================================================
function getDashboardData() {
  try {
    const sG = getSaldoGudang(); 
    const sTB = getSaldoTeamBuilding(); 
    const kg = getKasGudang(); 
    const tb = getTeamBuilding(); 
    const lk = getLaporanKerja();
    let history = [];
    
    if (kg.success && kg.data) { 
      kg.data.forEach(k => { history.push({ tanggal: k.tanggal, tipe: k.tipe === 'IN' ? 'Kas Masuk' : 'Kas Keluar', keterangan: k.keterangan, nominal: k.nominal, kategori: 'Kas Gudang' }); }); 
    }
    if (tb.success && tb.data) { 
      tb.data.forEach(t => { history.push({ tanggal: t.tanggal, tipe: 'Team Building', keterangan: t.keterangan, nominal: t.nominal, kategori: 'Team Building' }); }); 
    }
    
    history.sort((a, b) => new Date(b.tanggal||0) - new Date(a.tanggal||0)); 
    history = history.slice(0, 20);
    
    const totalKasIn = (kg.data||[]).filter(k => k.tipe === 'IN').reduce((s, k) => s + k.nominal, 0); 
    const totalKasOut = (kg.data||[]).filter(k => k.tipe === 'OUT').reduce((s, k) => s + k.nominal, 0);
    
    return { 
      success: true, 
      saldoGudang: sG.saldo || 0, 
      saldoTB: sTB.saldo || 0, 
      history: history, 
      totalKasIn: totalKasIn, 
      totalKasOut: totalKasOut, 
      kasData: kg.data || [], 
      tbData: tb.data || [], 
      laporanData: lk.success ? lk.data : [] 
    };
  } catch (e) { return { success: false, message: e.message }; }
}

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
        createdAt: data[i][7]
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function addKasGudang(tanggal, tipe, keterangan, nominal, buktiUrl, createdBy) {
  try {
    getSheet(CONFIG.SHEETS.KAS_GUDANG).appendRow([generateId(), tanggal, tipe, keterangan, parseFloat(nominal), buktiUrl, createdBy, new Date().toISOString()]);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteKasGudang(id) { return deleteRow(CONFIG.SHEETS.KAS_GUDANG, id); }

function getSaldoGudang() {
  try {
    const r = getKasGudang();
    if (!r.success) return r;
    let s = 0;
    r.data.forEach(d => { s += d.tipe === 'IN' ? d.nominal : -d.nominal; });
    return { success: true, saldo: s };
  } catch (e) { return { success: false, message: e.message }; }
}

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
        createdAt: data[i][6],
        tipe: data[i][7] || 'Pengeluaran'
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function addTeamBuilding(tanggal, keterangan, nominal, buktiUrl, createdBy, tipe) {
  try {
    getSheet(CONFIG.SHEETS.TEAM_BUILDING).appendRow([generateId(), tanggal, keterangan, parseFloat(nominal), buktiUrl, createdBy, new Date().toISOString(), tipe || 'Pengeluaran']);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteTeamBuilding(id) { return deleteRow(CONFIG.SHEETS.TEAM_BUILDING, id); }

function getSaldoTeamBuilding() {
  try {
    const r = getTeamBuilding();
    if (!r.success) return r;
    let s = 0;
    r.data.forEach(d => { s += d.tipe === 'Pemasukan' ? d.nominal : -d.nominal; });
    return { success: true, saldo: s };
  } catch (e) { return { success: false, message: e.message }; }
}

function getExpense() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.EXPENSE);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({
        id: data[i][0],
        tanggal: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]),
        perusahaan: data[i][2],
        kategori: data[i][3],
        keterangan: data[i][4],
        nominal: parseFloat(data[i][5]) || 0,
        bank: data[i][6],
        rekening: data[i][7],
        createdBy: data[i][8],
        createdAt: data[i][9]
      });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addExpense(tanggal, perusahaan, kategori, keterangan, nominal, bank, rekening, createdBy) {
  try {
    getSheet(CONFIG.SHEETS.EXPENSE).appendRow([generateId(), tanggal, perusahaan, kategori, keterangan, parseFloat(nominal)||0, bank, rekening, createdBy, new Date().toISOString()]);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteExpense(id) { return deleteRow(CONFIG.SHEETS.EXPENSE, id); }

// ============================================================
// UPLOAD FILE
// ============================================================
function getOrCreateBuktiFolder(subFolderName) {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty('BUKTI_FOLDER_ID');
  let folder;
  if (folderId) { try { folder = DriveApp.getFolderById(folderId); } catch(e) { folderId = null; } }
  if (!folderId) { folder = DriveApp.createFolder('Gudang FCL - Bukti Invoice'); props.setProperty('BUKTI_FOLDER_ID', folder.getId()); }
  const name = subFolderName || 'Umum';
  const subs = folder.getFoldersByName(name);
  return subs.hasNext() ? subs.next() : folder.createFolder(name);
}

function uploadFileToDrive(base64Data, fileName, mimeType, folderName) {
  try {
    const folder = getOrCreateBuktiFolder(folderName);
    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, mimeType || 'application/octet-stream', fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: 'https://drive.google.com/file/d/' + file.getId() + '/view' };
  } catch (e) { return { success: false, message: e.message }; }
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
  } catch (e) { return { success: false, message: e.message }; }
}

function finalizeChunkedUpload(uploadId, fileName, mimeType, folderName) {
  try {
    const cache = CacheService.getScriptCache();
    const countStr = cache.get('meta_' + uploadId + '_count');
    if (!countStr) return { success: false, message: 'Upload session kedaluwarsa' };
    const totalChunks = parseInt(countStr);
    let fullBase64 = '';
    for (let i = 0; i < totalChunks; i++) {
      const key = 'chunk_' + uploadId + '_' + i;
      const isSplit = cache.get(key + '_split');
      if (isSplit) { fullBase64 += (cache.get(key + '_a') || '') + (cache.get(key + '_b') || ''); } 
      else { fullBase64 += cache.get(key) || ''; }
    }
    const folder = getOrCreateBuktiFolder(folderName);
    const decoded = Utilities.base64Decode(fullBase64);
    const blob = Utilities.newBlob(decoded, mimeType || 'application/octet-stream', fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: 'https://drive.google.com/file/d/' + file.getId() + '/view' };
  } catch (e) { return { success: false, message: e.message }; }
}

// ============================================================
// LAPORAN KERJA
// ============================================================
function getLaporanKerja() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.LAPORAN_KERJA);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({
        id: data[i][0],
        tanggal: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]),
        divisi: data[i][2],
        pic: data[i][3],
        totalOrang: parseInt(data[i][4]) || 0,
        perbantuan: parseFloat(data[i][5]) || 0,
        pengurangan: parseFloat(data[i][6]) || 0,
        jamLembur: parseFloat(data[i][7]) || 0,
        totalJamKerja: parseFloat(data[i][8]) || 0,
        kendala: data[i][9] || '-',
        totalStaff: parseInt(data[i][10]) || 0,
        totalAdmin: parseInt(data[i][11]) || 0,
        totalOrder: parseInt(data[i][12]) || 0,
        createdBy: data[i][13],
        sisaOrder: parseInt(data[i][15]) || 0
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function addLaporanKerja(tanggal, divisi, pic, totalOrang, perbantuan, pengurangan, jamLembur, totalJamKerja, kendala, totalStaff, totalAdmin, totalOrder, createdBy, sisaOrder) {
  try {
    getSheet(CONFIG.SHEETS.LAPORAN_KERJA).appendRow([
      generateId(), tanggal, divisi, pic, parseInt(totalOrang)||0, parseFloat(perbantuan)||0, parseFloat(pengurangan)||0, parseFloat(jamLembur)||0, parseFloat(totalJamKerja)||0, kendala, parseInt(totalStaff)||0, parseInt(totalAdmin)||0, parseInt(totalOrder)||0, createdBy, new Date().toISOString(), parseInt(sisaOrder)||0
    ]);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}
function deleteLaporanKerja(id) { return deleteRow(CONFIG.SHEETS.LAPORAN_KERJA, id); }

// ============================================================
// HANDOVER & KLAIM
// ============================================================
function getHandover() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.HANDOVER);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({
        id: data[i][0],
        tanggal: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]),
        pic: data[i][2],
        resi: data[i][3],
        pengerjaan: data[i][4],
        keterangan: data[i][5],
        status: data[i][6],
        createdBy: data[i][7],
        createdAt: data[i][8] instanceof Date ? data[i][8].toISOString() : String(data[i][8]||'')
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}
function addHandover(tanggal, pic, resi, pengerjaan, keterangan, createdBy) {
  try { getSheet(CONFIG.SHEETS.HANDOVER).appendRow([generateId(), tanggal, pic, resi, pengerjaan, keterangan, 'Pending', createdBy, new Date().toISOString()]); return { success: true }; } catch (e) { return { success: false, message: e.message }; }
}
function updateHandoverStatus(id, status) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.HANDOVER); const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (String(data[i][0]) === String(id)) { sheet.getRange(i + 1, 7).setValue(status); return { success: true }; } }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}
function deleteHandover(id) { return deleteRow(CONFIG.SHEETS.HANDOVER, id); }

function getKlaim() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KLAIM); const data = sheet.getDataRange().getValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({
        id: data[i][0],
        tanggal: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]),
        pic: data[i][2], resi: data[i][3], harga: parseFloat(data[i][4]) || 0, keterangan: data[i][5], status: data[i][6], createdBy: data[i][7],
        createdAt: data[i][8] instanceof Date ? data[i][8].toISOString() : String(data[i][8]||'')
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}
function addKlaim(tanggal, pic, resi, harga, keterangan, createdBy) {
  try { getSheet(CONFIG.SHEETS.KLAIM).appendRow([generateId(), tanggal, pic, resi, parseFloat(harga) || 0, keterangan, 'Pending', createdBy, new Date().toISOString()]); return { success: true }; } catch (e) { return { success: false, message: e.message }; }
}
function updateKlaimStatus(id, status) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KLAIM); const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (String(data[i][0]) === String(id)) { sheet.getRange(i + 1, 7).setValue(status); return { success: true }; } }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}
function deleteKlaim(id) { return deleteRow(CONFIG.SHEETS.KLAIM, id); }

// ============================================================
// SOP & STRUKTUR ORGANISASI
// ============================================================
function getSOP() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SOP); const data = sheet.getDataRange().getValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({ id: data[i][0], judul: data[i][1], konten: data[i][2], kategori: data[i][3], createdBy: data[i][4] });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}
function addSOP(judul, konten, kategori, createdBy) {
  try { getSheet(CONFIG.SHEETS.SOP).appendRow([generateId(), judul, konten, kategori, createdBy, new Date().toISOString()]); return { success: true }; } catch (e) { return { success: false, message: e.message }; }
}
function updateSOP(id, judul, konten, kategori) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SOP); const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (String(data[i][0]) === String(id)) { sheet.getRange(i + 1, 2, 1, 4).setValues([[judul, konten, kategori, new Date().toISOString()]]); return { success: true }; } }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}
function deleteSOP(id) { return deleteRow(CONFIG.SHEETS.SOP, id); }
function exportSOP() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SOP); const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: false, message: 'Tidak ada data SOP' };
    const doc = DocumentApp.create('SOP Gudang - FCL'); const body = doc.getBody();
    body.appendParagraph('SOP GUDANG — FCL').setHeading(DocumentApp.ParagraphHeading.TITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
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
        body.appendParagraph((idx + 1) + '. ' + sop.judul).setHeading(DocumentApp.ParagraphHeading.HEADING2);
        body.appendParagraph(sop.konten || '-');
        body.appendParagraph('');
      });
    });
    DriveApp.getFileById(doc.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: 'https://docs.google.com/document/d/' + doc.getId() + '/edit' };
  } catch (e) { return { success: false, message: e.message }; }
}

function getOrganisasi() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ORGANISASI); const data = sheet.getDataRange().getValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({ id: data[i][0], nama: data[i][1], jabatan: data[i][2], atasan: data[i][3], departemen: data[i][4], foto: data[i][5], urutan: data[i][6] });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}
function addOrganisasi(nama, jabatan, atasan, departemen, foto, urutan) {
  try { getSheet(CONFIG.SHEETS.ORGANISASI).appendRow([generateId(), nama, jabatan, atasan, departemen, foto, urutan || 0]); return { success: true }; } catch (e) { return { success: false, message: e.message }; }
}
function updateOrganisasi(id, nama, jabatan, atasan, departemen, foto, urutan) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ORGANISASI); const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (String(data[i][0]) === String(id)) { sheet.getRange(i + 1, 2, 1, 6).setValues([[nama, jabatan, atasan, departemen, foto, urutan]]); return { success: true }; } }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}
function deleteOrganisasi(id) { return deleteRow(CONFIG.SHEETS.ORGANISASI, id); }

// ============================================================
// INVENTORY: STOCK, INBOUND, OUTBOUND
// ============================================================
function generateSKU(nama) {
  const prefix = nama.replace(/[^A-Za-z]/g,'').toUpperCase().substring(0,3) || 'SKU';
  return prefix + '-' + String(Date.now()).slice(-5);
}

function generateNoSJ(prefix) {
  return prefix + '/' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') + '/' + String(Math.floor(Math.random()*9000)+1000);
}

function getStock() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.STOCK); const data = sheet.getDataRange().getValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({
        id: data[i][0], sku: data[i][1], nama: data[i][2], barcode: data[i][3], batch: data[i][4],
        expDate: data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][5]||''),
        satuan: data[i][6], stok: parseFloat(data[i][7])||0, stokMin: parseFloat(data[i][8])||0, kategori: data[i][9], lokasi: data[i][10]
      });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addStock(skuInput, nama, barcode, batch, expDate, satuan, stok, stokMin, kategori, lokasi) {
  try {
    const sku = (skuInput && skuInput.trim() !== '') ? skuInput.trim() : generateSKU(nama);
    const now = new Date().toISOString();
    getSheet(CONFIG.SHEETS.STOCK).appendRow([generateId(), sku, nama, barcode, batch, expDate, satuan, parseFloat(stok)||0, parseFloat(stokMin)||0, kategori, lokasi, now, now]);
    return { success: true, sku };
  } catch(e) { return { success: false, message: e.message }; }
}

function updateStock(id, sku, nama, barcode, batch, expDate, satuan, stok, stokMin, kategori, lokasi) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.STOCK); const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i+1, 2, 1, 10).setValues([[sku, nama, barcode, batch, expDate, satuan, parseFloat(stok)||0, parseFloat(stokMin)||0, kategori, lokasi]]);
        sheet.getRange(i+1, 13).setValue(new Date().toISOString());
        return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch(e) { return { success: false, message: e.message }; }
}
function deleteStock(id) { return deleteRow(CONFIG.SHEETS.STOCK, id); }

function updateStokQty(id, delta) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.STOCK); const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        const newStok = (parseFloat(data[i][7])||0) + delta;
        if (newStok < 0) return { success: false, message: 'Stok tidak cukup! Sisa: ' + (parseFloat(data[i][7])||0) };
        sheet.getRange(i+1,8).setValue(newStok);
        return { success: true };
      }
    }
    return { success: false, message: 'Barang tidak ditemukan' };
  } catch(e) { return { success: false, message: e.message }; }
}

function getSuratJalanMasuk() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SURAT_JALAN_MASUK); const data = sheet.getDataRange().getValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({ id: data[i][0], noSJ: data[i][1], tanggal: data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][2]||''), supplier: data[i][3], keterangan: data[i][4], createdBy: data[i][5], createdAt: data[i][6] instanceof Date ? data[i][6].toISOString() : String(data[i][6]||'') });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addSuratJalanMasuk(tanggal, supplier, keterangan, items, createdBy) {
  try {
    const noSJ = generateNoSJ('SJM'); const id = generateId();
    getSheet(CONFIG.SHEETS.SURAT_JALAN_MASUK).appendRow([id, noSJ, tanggal, supplier, keterangan, createdBy, new Date().toISOString()]);
    const detSheet = getSheet(CONFIG.SHEETS.SURAT_JALAN_MASUK_DETAIL);
    const parsedItems = typeof items === 'string' ? JSON.parse(items) : items;
    parsedItems.forEach(item => {
      detSheet.appendRow([generateId(), id, noSJ, item.stockId, item.sku, item.nama, parseFloat(item.qty)||0, item.satuan, item.batch||'', item.expDate||'']);
      updateStokQty(item.stockId, parseFloat(item.qty)||0);
    });
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function getSJMasukDetail(sjId) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SURAT_JALAN_MASUK_DETAIL); const data = sheet.getDataRange().getValues(); const result = [];
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
    const data = getSheet(sheetName).getDataRange().getValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(sjId)) result.push({ sku:data[i][4], nama:data[i][5], qty:parseFloat(data[i][6])||0, satuan:data[i][7], batch:data[i][8], expDate:data[i][9] });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function getSuratJalanKeluar() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SURAT_JALAN_KELUAR); const data = sheet.getDataRange().getValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({ id: data[i][0], noSJ: data[i][1], tanggal: data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][2]||''), tujuan: data[i][3], keterangan: data[i][4], createdBy: data[i][5], createdAt: data[i][6] instanceof Date ? data[i][6].toISOString() : String(data[i][6]||'') });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addSuratJalanKeluar(tanggal, tujuan, keterangan, items, createdBy) {
  try {
    const noSJ = generateNoSJ('SJK'); const id = generateId();
    getSheet(CONFIG.SHEETS.SURAT_JALAN_KELUAR).appendRow([id, noSJ, tanggal, tujuan, keterangan, createdBy, new Date().toISOString()]);
    const detSheet = getSheet(CONFIG.SHEETS.SURAT_JALAN_KELUAR_DETAIL);
    const parsedItems = typeof items === 'string' ? JSON.parse(items) : items;
    for (const item of parsedItems) {
      const res = updateStokQty(item.stockId, -(parseFloat(item.qty)||0));
      if (!res.success) return { success: false, message: res.message };
      detSheet.appendRow([generateId(), id, noSJ, item.stockId, item.sku, item.nama, parseFloat(item.qty)||0, item.satuan, item.batch||'', item.expDate||'']);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

// ============================================================
// ORDER & RETUR
// ============================================================
function getOrders() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ORDER); const data = sheet.getDataRange().getValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({
        id:data[i][0], noOrder:data[i][1],
        tanggal: data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][2]||''),
        pelanggan:data[i][3], alamat:data[i][4], status:data[i][5], totalItem:parseFloat(data[i][6])||0, keterangan:data[i][7], createdBy:data[i][8],
        createdAt: data[i][9] instanceof Date ? data[i][9].toISOString() : String(data[i][9]||''),
        sentAt: data[i][10] instanceof Date ? data[i][10].toISOString() : String(data[i][10]||'')
      });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addOrder(tanggal, pelanggan, alamat, keterangan, items, createdBy) {
  try {
    const noOrder = generateNoSJ('WHFCL'); const id = generateId();
    const parsedItems = typeof items === 'string' ? JSON.parse(items) : items;
    const totalItem = parsedItems.reduce((s,x) => s + (parseFloat(x.qty)||0), 0);
    getSheet(CONFIG.SHEETS.ORDER).appendRow([id, noOrder, tanggal, pelanggan, alamat, 'Pending', totalItem, keterangan, createdBy, new Date().toISOString(), '']);
    const detSheet = getSheet(CONFIG.SHEETS.ORDER_DETAIL);
    parsedItems.forEach(item => {
      detSheet.appendRow([generateId(), id, noOrder, item.stockId, item.sku, item.nama, parseFloat(item.qty)||0, item.satuan]);
    });
    return { success: true, noOrder: noOrder };
  } catch(e) { return { success: false, message: e.message }; }
}

function getOrderDetail(orderId) {
  try {
    const data = getSheet(CONFIG.SHEETS.ORDER_DETAIL).getDataRange().getValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(orderId)) result.push({ sku:data[i][4], nama:data[i][5], qty:parseFloat(data[i][6])||0, satuan:data[i][7] });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteOrder(id) { return deleteRow(CONFIG.SHEETS.ORDER, id); }

function kirimOrder(id) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ORDER); const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        if (data[i][5] === 'Terkirim') return { success: false, message: 'Sudah terkirim' };
        sheet.getRange(i+1, 6).setValue('Terkirim');
        sheet.getRange(i+1, 11).setValue(new Date().toISOString());
        found = true;
        break;
      }
    }
    if (!found) return { success: false, message: 'Order Tidak ditemukan' };
    const detData = getSheet(CONFIG.SHEETS.ORDER_DETAIL).getDataRange().getValues();
    for(let i=1; i<detData.length; i++) {
      if(String(detData[i][1]) === String(id)) { updateStokQty(detData[i][3], -(parseFloat(detData[i][6])||0)); }
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function getRetur() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.RETUR); const data = sheet.getDataRange().getValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({
        id: data[i][0], noRetur: data[i][1],
        tanggal: data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][2]||''),
        sumber: data[i][3], alasan: data[i][4], keterangan: data[i][5], createdBy: data[i][6],
        createdAt: data[i][7] instanceof Date ? data[i][7].toISOString() : String(data[i][7]||'')
      });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addRetur(tanggal, sumber, alasan, keterangan, items, createdBy) {
  try {
    const noRetur = generateNoSJ('RTR'); const id = generateId();
    getSheet(CONFIG.SHEETS.RETUR).appendRow([id, noRetur, tanggal, sumber, alasan, keterangan, createdBy, new Date().toISOString()]);
    const detSheet = getSheet(CONFIG.SHEETS.RETUR_DETAIL);
    const parsedItems = typeof items === 'string' ? JSON.parse(items) : items;
    for (const item of parsedItems) {
      const res = updateStokQty(item.stockId, parseFloat(item.qty)||0);
      if (!res.success) return { success: false, message: res.message };
      detSheet.appendRow([generateId(), id, noRetur, item.stockId, item.sku, item.nama, parseFloat(item.qty)||0, item.satuan, item.batch||'', item.expDate||'']);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function getReturDetail(returId) {
  try {
    const data = getSheet(CONFIG.SHEETS.RETUR_DETAIL).getDataRange().getValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(returId)) result.push({ sku:data[i][4], nama:data[i][5], qty:parseFloat(data[i][6])||0, satuan:data[i][7], batch:data[i][8], expDate:data[i][9] });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteRetur(id) { return deleteRow(CONFIG.SHEETS.RETUR, id); }

// ============================================================
// ANALISIS STOK & BULK IMPORT
// ============================================================
function importOrdersBulk(jsonString, createdBy) {
  try {
    const orders = JSON.parse(jsonString); let count = 0; const stData = getStock().data || [];
    for(let i=0; i<orders.length; i++) {
      const o = orders[i];
      const mappedItems = o.items.map(item => {
        let stId = ''; let stNama = item.sku; let stSatuan = 'PCS';
        const found = stData.find(s => s.sku === item.sku);
        if(found) { stId = found.id; stNama = found.nama; stSatuan = found.satuan; }
        return { stockId: stId, sku: item.sku, nama: stNama, qty: item.qty, satuan: stSatuan };
      });
      addOrder(o.tanggal, o.pelanggan, o.alamat, o.keterangan, mappedItems, createdBy);
      count++;
    }
    return { success: true, count };
  } catch(e) { return { success: false, message: e.message }; }
}

function getAnalisisStock() {
  try {
    const stockData = getStock().data || []; const result = [];
    const msInDay = 86400000; const now = new Date();
    const orderDetData = getSheet(CONFIG.SHEETS.ORDER_DETAIL).getDataRange().getValues();
    const orderData = getSheet(CONFIG.SHEETS.ORDER).getDataRange().getValues();
    const usageWeek = {}; const usageMonth = {};
    
    for (let i = 1; i < orderData.length; i++) {
      if(orderData[i][5] !== 'Terkirim') continue;
      const dDate = new Date(orderData[i][10] || orderData[i][2]);
      const diff = now - dDate;
      const isW = diff <= 7 * msInDay;
      const isM = diff <= 30 * msInDay;
      
      if (isM) {
        for (let j=1; j<orderDetData.length; j++) {
          if (String(orderDetData[j][1]) === String(orderData[i][0])) {
            const sId = orderDetData[j][3];
            const q = parseFloat(orderDetData[j][6])||0;
            if(isW) usageWeek[sId] = (usageWeek[sId]||0) + q;
            usageMonth[sId] = (usageMonth[sId]||0) + q;
          }
        }
      }
    }
    
    stockData.forEach(s => {
      const mw = usageWeek[s.id] || 0;
      const mm = usageMonth[s.id] || 0;
      const rata = (mm / 30).toFixed(1);
      const status = s.stok <= 0 ? 'Kritis' : (s.stok <= s.stokMin ? 'Rendah' : 'Aman');
      result.push({ sku: s.sku, nama: s.nama, stokSaat: s.stok, minggu: mw, bulan: mm, rataHarian: rata, satuan: s.satuan, statusStok: status });
    });
    
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

// ============================================================
// TUGAS PROJECT
// ============================================================

function getTugasProject() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.TUGAS_PROJECT);
    const data  = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({
        id:           data[i][0],
        judul:        data[i][1],
        assignee:     data[i][2],
        assigneeName: data[i][3],
        prioritas:    data[i][4],
        tanggalMulai: data[i][5] instanceof Date
          ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : String(data[i][5] || ''),
        deadline:     data[i][6] instanceof Date
          ? Utilities.formatDate(data[i][6], Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : String(data[i][6] || ''),
        targetHari:   data[i][7] ? parseInt(data[i][7]) : null,
        status:       data[i][8],
        kategori:     data[i][9],
        deskripsi:    data[i][10],
        createdBy:    data[i][11],
        createdAt:    data[i][12] instanceof Date ? data[i][12].toISOString() : String(data[i][12] || ''),
        updatedAt:    data[i][13] instanceof Date ? data[i][13].toISOString() : String(data[i][13] || ''),
        log:          data[i][14] ? String(data[i][14]) : '[]'
      });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addTugasProject(jsonString) {
  try {
    const p      = typeof jsonString === 'string' ? JSON.parse(jsonString) : jsonString;
    const id     = generateId();
    const now    = new Date().toISOString();
    const logEntry = JSON.stringify([{
      time: now,
      action: 'Tugas dibuat',
      by: p.createdBy || '-'
    }]);

    getSheet(CONFIG.SHEETS.TUGAS_PROJECT).appendRow([
      id,
      p.judul,
      p.assignee,
      p.assigneeName || p.assignee,
      p.prioritas || 'Sedang',
      p.tanggalMulai || '',
      p.deadline || '',
      p.targetHari || '',
      p.status || 'Todo',
      p.kategori || '',
      p.deskripsi || '',
      p.createdBy || '',
      now,
      now,
      logEntry
    ]);
    return { success: true, id: id };
  } catch(e) { return { success: false, message: e.message }; }
}

function updateTugasProject(jsonString) {
  try {
    const p     = typeof jsonString === 'string' ? JSON.parse(jsonString) : jsonString;
    const sheet = getSheet(CONFIG.SHEETS.TUGAS_PROJECT);
    const data  = sheet.getDataRange().getValues();
    const now   = new Date().toISOString();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(p.id)) continue;

      // Parse existing log
      let log = [];
      try { log = JSON.parse(data[i][14] || '[]'); } catch(ex) { log = []; }
      log.push({
        time: now,
        action: 'Tugas diperbarui (status: ' + p.status + ')',
        by: p.createdBy || '-'
      });

      sheet.getRange(i+1, 2).setValue(p.judul);
      sheet.getRange(i+1, 3).setValue(p.assignee);
      sheet.getRange(i+1, 4).setValue(p.assigneeName || p.assignee);
      sheet.getRange(i+1, 5).setValue(p.prioritas || 'Sedang');
      sheet.getRange(i+1, 6).setValue(p.tanggalMulai || '');
      sheet.getRange(i+1, 7).setValue(p.deadline || '');
      sheet.getRange(i+1, 8).setValue(p.targetHari || '');
      sheet.getRange(i+1, 9).setValue(p.status || 'Todo');
      sheet.getRange(i+1,10).setValue(p.kategori || '');
      sheet.getRange(i+1,11).setValue(p.deskripsi || '');
      sheet.getRange(i+1,14).setValue(now);
      sheet.getRange(i+1,15).setValue(JSON.stringify(log));

      return { success: true };
    }
    return { success: false, message: 'Tugas tidak ditemukan' };
  } catch(e) { return { success: false, message: e.message }; }
}

function updateTugasStatus(id, newStatus, updatedBy) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.TUGAS_PROJECT);
    const data  = sheet.getDataRange().getValues();
    const now   = new Date().toISOString();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(id)) continue;

      let log = [];
      try { log = JSON.parse(data[i][14] || '[]'); } catch(ex) { log = []; }
      log.push({ time: now, action: 'Status diubah ke: ' + newStatus, by: updatedBy || '-' });

      sheet.getRange(i+1, 9).setValue(newStatus);
      sheet.getRange(i+1,14).setValue(now);
      sheet.getRange(i+1,15).setValue(JSON.stringify(log));
      return { success: true };
    }
    return { success: false, message: 'Tugas tidak ditemukan' };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteTugasProject(id) {
  return deleteRow(CONFIG.SHEETS.TUGAS_PROJECT, id);
}
