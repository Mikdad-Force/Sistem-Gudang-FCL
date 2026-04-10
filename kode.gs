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
    TUGAS_PROJECT: 'TugasProject',
    ASSET: 'PengajuanAsset',
    STOCK_OPNAME: 'StockOpname',
    PACKING_LIST: 'PackingList',
    RIWAYAT_KARYAWAN: 'RiwayatKaryawan',
    SURAT_PERINGATAN: 'SuratPeringatan',
    TGL_MERAH: 'TglMerah',
    ASSET_WAREHOUSE: 'AssetWarehouse',
    BOOKING_MOBIL: 'BookingMobil',
    TUGAS_CONSUMABLE: 'TugasConsumable',
    ABSENSI_LEMBUR: 'AbsensiLembur'
  },
  DRIVE_FOLDER_ID: '14u5aMQltzyc7BCw3-87p25mqPeYf9weC'
};

// ============================================================
// ENTRY POINT
// ============================================================
function doGet(e) {
  // Trigger update untuk mengatasi kolom kosong (title/judul Header)
  try { ForceUpdateAllHeaders(); } catch(err) {}

  // Menggunakan 'Index' dengan I besar menyesuaikan default Google Apps Script
  var html = HtmlService.createHtmlOutputFromFile('Index'); 
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return html;
}

/**
 * Handle API requests from external domains (Netlify, etc.)
 */
function doPost(e) {
  const result = { success: false, message: 'Invalid Request' };
  
  try {
    const postData = JSON.parse(e.postData.contents);
    const funcName = postData.func;
    const args = postData.args || [];
    
    // Security check: Only allow direct execution of defined functions
    if (typeof this[funcName] === 'function') {
      const data = this[funcName].apply(null, args);
      return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      result.message = 'Function ' + funcName + ' not found or access denied.';
    }
  } catch (err) {
    result.message = 'API Error: ' + err.toString();
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
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

  setupSheet(ss, CONFIG.SHEETS.USERS, ['id', 'username', 'password', 'nama', 'role', 'createdAt', 'permissions', 'divisi']);
  setupSheet(ss, CONFIG.SHEETS.KAS_GUDANG, ['id', 'tanggal', 'tipe', 'keterangan', 'nominal', 'buktiUrl', 'createdBy', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.TEAM_BUILDING, ['id', 'tanggal', 'keterangan', 'nominal', 'buktiUrl', 'createdBy', 'createdAt', 'tipe']);
  setupSheet(ss, CONFIG.SHEETS.EXPENSE, ['id', 'tanggal', 'perusahaan', 'kategori', 'keterangan', 'nominal', 'bank', 'rekening', 'createdBy', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.KARYAWAN, ['id', 'nama', 'jabatan', 'cabang', 'telepon', 'email', 'tanggalMasuk', 'status', 'createdAt', 'tanggalSelesai', 'sisaCuti']);
  setupSheet(ss, CONFIG.SHEETS.IJIN, ['id', 'tanggal', 'nama', 'jenis', 'keterangan', 'bukti', 'status', 'createdBy', 'createdAt', 'history']);
  setupSheet(ss, CONFIG.SHEETS.LEMBUR, ['id', 'tanggal', 'nama', 'divisi', 'jumlahJam', 'keterangan', 'status', 'createdBy', 'createdAt', 'history']);
  setupSheet(ss, CONFIG.SHEETS.LAPORAN_KERJA, ['id', 'tanggal', 'divisi', 'pic', 'totalOrang', 'perbantuan', 'pengurangan', 'jamLembur', 'totalJamKerja', 'kendala', 'totalStaff', 'totalAdmin', 'totalOrder', 'createdBy', 'createdAt', 'sisaOrder', 'staffLemburNames', 'shift', 'totalPHL', 'jamKerjaPHL', 'totalPO', 'totalQty', 'totalInbound', 'pendapatanPotongBubble', 'pendapatanBuatBubble', 'alasanPengurangan']);
  setupSheet(ss, CONFIG.SHEETS.SOP, ['id', 'judul', 'konten', 'kategori', 'createdBy', 'updatedAt']);
  setupSheet(ss, CONFIG.SHEETS.ORGANISASI, ['id', 'nama', 'jabatan', 'atasan', 'departemen', 'foto', 'urutan']);
  setupSheet(ss, CONFIG.SHEETS.STOCK, ['id','sku','nama','barcode','batch','expDate','satuan','stok','stokMin','kategori','lokasi','createdAt','updatedAt']);
  setupSheet(ss, CONFIG.SHEETS.SURAT_JALAN_MASUK, ['id','noSJ','tanggal','supplier','keterangan','createdBy','createdAt']);
  setupSheet(ss, CONFIG.SHEETS.SURAT_JALAN_MASUK_DETAIL, ['id','sjId','noSJ','stockId','sku','nama','qty','satuan','batch','expDate']);
  setupSheet(ss, CONFIG.SHEETS.SURAT_JALAN_KELUAR, ['id','noSJ','tanggal','tujuan','keterangan','createdBy','createdAt']);
  setupSheet(ss, CONFIG.SHEETS.SURAT_JALAN_KELUAR_DETAIL, ['id','sjId','noSJ','stockId','sku','nama','qty','satuan','batch','expDate']);
  setupSheet(ss, CONFIG.SHEETS.ORDER, ['id','noOrder','tanggal','pelanggan','alamat','status','totalItem','keterangan','createdBy','createdAt','sentAt','buktiPacking']);
  setupSheet(ss, CONFIG.SHEETS.ORDER_DETAIL, ['id','orderId','noOrder','stockId','sku','nama','qty','satuan','batch','expDate','packedQty']);
  setupSheet(ss, CONFIG.SHEETS.RETUR, ['id','noRetur','tanggal','sumber','alasan','keterangan','createdBy','createdAt']);
  setupSheet(ss, CONFIG.SHEETS.RETUR_DETAIL, ['id','returId','noRetur','stockId','sku','nama','qty','satuan','batch','expDate']);
  setupSheet(ss, CONFIG.SHEETS.HANDOVER, ['id', 'tanggal', 'pic', 'resi', 'pengerjaan', 'keterangan', 'status', 'createdBy', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.KLAIM, ['id', 'tanggal', 'pic', 'resi', 'harga', 'keterangan', 'status', 'createdBy', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.TUGAS_PROJECT, ['id','judul','assignee','assigneeName','prioritas','tanggalMulai','deadline','targetHari','status','kategori','deskripsi','createdBy','createdAt','updatedAt','log']);
  setupSheet(ss, CONFIG.SHEETS.ASSET, ['id','tanggal','nama','jenisAsset','deskripsi','estimasiHarga','prioritas','bukti','status','createdBy','createdAt','history']);
  setupSheet(ss, CONFIG.SHEETS.STOCK_OPNAME, ['id','tanggal','stockId','sku','nama','lokasi','batch','expDate','stokSistem','stokFisik','selisih','status','catatan','createdBy','createdAt','approvedBy','approvedAt']);
  setupSheet(ss, CONFIG.SHEETS.PACKING_LIST, ['id','tanggal','noPL','keterangan','fileUrl','createdBy','createdAt']);
  setupSheet(ss, CONFIG.SHEETS.RIWAYAT_KARYAWAN, ['id','nama','jabatan','cabang','telepon','tanggalMasuk','tanggalResign','alasanResign','keterangan','createdBy','createdAt']);
  setupSheet(ss, CONFIG.SHEETS.SURAT_PERINGATAN, ['id','karyawanNama','karyawanId','jenisSP','alasan','tanggalSP','masaBerlaku','tanggalKadaluarsa','status','createdBy','createdAt']);
  setupSheet(ss, CONFIG.SHEETS.TUGAS_CONSUMABLE, ['id', 'tanggal', 'pemberiTugas', 'picName', 'targetPotong', 'targetBuat', 'actualPotong', 'actualBuat', 'status', 'catatan', 'createdAt', 'updatedAt']);
  setupSheet(ss, CONFIG.SHEETS.TGL_MERAH, ['id', 'tanggal', 'nama', 'divisi', 'jamEstimasi', 'createdBy', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.ASSET_WAREHOUSE, ['id', 'code', 'nama', 'tanggalMasuk', 'divisi', 'status', 'createdBy', 'createdAt', 'history', 'qty']);
  setupSheet(ss, CONFIG.SHEETS.BOOKING_MOBIL, ['id', 'tanggal', 'pic', 'jamBerangkat', 'tujuan', 'keterangan', 'rute', 'status', 'createdBy', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.ABSENSI_LEMBUR, ['id', 'tanggal', 'jam', 'nama', 'divisi', 'karyawanId', 'status', 'createdAt']);

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
  } else {
    // Update headers if already exist (untuk mencegah kolom baru yang masuk ke database kolom/kosong tanpa judul)
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
  }
  return sheet;
}

// Fungsi manual untuk trigger update header yang kosong di Google Sheets
function ForceUpdateAllHeaders() {
  setupDatabase();
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

function deleteRow(sheetName, id, secondaryId, secondaryCol) {
  try {
    const sheet = getSheet(sheetName); 
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const matchPrimary = id && String(data[i][0]) === String(id);
      const matchSecondary = secondaryId && secondaryCol !== undefined && String(data[i][secondaryCol]) === String(secondaryId);
      if (matchPrimary || matchSecondary) { 
        sheet.deleteRow(i + 1); 
        return { success: true }; 
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

function generateId() { 
  return Utilities.getUuid(); 
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
            permissions: data[i][6] || '[]',
            divisi: data[i][7] || '',
            isDefaultPassword: (password === '1')
          } 
        };
      }
    }
    return { success: false, message: 'Username atau password salah' };
  } catch (e) { return { success: false, message: e.message }; }
}

function adminResetPassword(newPassword, userId) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    const hashed = hashPassword(newPassword);
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(userId)) {
        sheet.getRange(i + 1, 3).setValue(hashed);
        return { success: true };
      }
    }
    return { success: false, message: 'User tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

function getUsers() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      result.push({ 
        id: data[i][0], 
        username: data[i][1], 
        nama: data[i][3], 
        role: data[i][4], 
        createdAt: data[i][5], 
        permissions: data[i][6] || '[]',
        divisi: data[i][7] || ''
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function importUsersBulk(userList) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    const existingUsernames = new Set(data.slice(1).map(row => String(row[1]).trim().toLowerCase()));
    
    let addedCount = 0;
    userList.forEach(u => {
      const username = String(u.Username || u.username || '').trim();
      const password = String(u.Password || u.password || '').trim();
      const nama = String(u.Nama_Lengkap || u['Nama Lengkap'] || u.nama || '').trim();
      const role = String(u.Role || u.role || 'user').trim();
      const permissions = '[]';
      
      if (username && password && !existingUsernames.has(username.toLowerCase())) {
        const id = generateId();
        sheet.appendRow([id, username, hashPassword(password), nama, role, new Date().toISOString(), permissions]);
        existingUsernames.add(username.toLowerCase());
        addedCount++;
      }
    });
    
    return { success: true, message: addedCount + ' akun berhasil diimpor.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addUser(username, password, nama, role, permissions) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === username) return { success: false, message: 'Username sudah ada' };
    }
    const id = generateId();
    sheet.appendRow([id, username, hashPassword(password), nama, role || 'user', new Date().toISOString(), permissions || '[]', arguments[5] || '']);
    return { success: true, id: id };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateUser(id, username, password, nama, role, permissions) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      // Find by ID or Username fallback
      if ((id && String(data[i][0]) === String(id)) || (!id && String(data[i][1]) === String(username))) {
        sheet.getRange(i + 1, 2).setValue(username);
        if (password) sheet.getRange(i + 1, 3).setValue(hashPassword(password));
        sheet.getRange(i + 1, 4).setValue(nama);
        sheet.getRange(i + 1, 5).setValue(role);
        sheet.getRange(i + 1, 7).setValue(permissions || '[]');
        sheet.getRange(i + 1, 8).setValue(arguments[6] || '');
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
       ijin: (getIjin().data || []),
       lembur: (getLembur().data || []),
       asset: (getAsset().data || []),
       stockOpname: (getStockOpname().data || [])
    };
  } catch (e) { return { success: false, message: e.message }; }
}

function processApprovalStatus(tipe, id, action, userNama, userRole, reason, pemohonNama, tanggal) {
  try {
    let sheetName, statusCol, historyCol, namaCol, tglCol;

    if (tipe === 'ijin') {
      sheetName = CONFIG.SHEETS.IJIN; statusCol = 7; historyCol = 10; namaCol = 3; tglCol = 2;
    } else if (tipe === 'lembur') {
      sheetName = CONFIG.SHEETS.LEMBUR; statusCol = 7; historyCol = 10; namaCol = 3; tglCol = 2;
    } else if (tipe === 'asset') {
      sheetName = CONFIG.SHEETS.ASSET; statusCol = 9; historyCol = 12; namaCol = 3; tglCol = 2;
    } else if (tipe === 'opname') {
      return approveStockOpname(id, action === 'Approve' ? 'Approved' : 'Rejected', userNama);
    } else {
      return { success: false, message: 'Tipe tidak dikenali' };
    }

    const sheet = getSheet(sheetName);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      // Logic Fallback Pencarian: ID atau (Nama + Tanggal)
      const rowId = String(data[i][0]).trim();
      const rowNama = String(data[i][namaCol - 1]).trim();
      const rowTgl = data[i][tglCol - 1] instanceof Date ? Utilities.formatDate(data[i][tglCol - 1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][tglCol - 1]).trim();

      const matchId = (id && rowId === String(id).trim());
      const matchBusinessKey = (pemohonNama && tanggal && rowNama === String(pemohonNama).trim() && rowTgl === String(tanggal).trim());

      if (matchId || matchBusinessKey) {
        let currentStatus = data[i][statusCol - 1];
        if (!currentStatus || currentStatus === 'Pending') currentStatus = 'Pending Team Leader'; // Fallback

        // BACKEND AUTHORIZATION CHECK
        const isAdmin = (userRole === 'admin' || userRole === 'Super Admin');
        const isTL = (userRole === 'Team Leader' || userRole === 'TL' || userRole.includes('Team Leader'));
        const isVice = (userRole === 'Vice Supervisor' || userRole === 'Vice SPV' || userRole === 'Vice VPV' || userRole.includes('Vice'));
        const isSPV = (userRole === 'Supervisor' || userRole === 'SPV' || userRole === 'Supervisor HR' || (userRole.includes('Supervisor') && !userRole.includes('Vice')));
        const isHR = (userRole === 'HR' || userRole === 'Supervisor HR' || userRole.includes('HR'));

        let authorized = isAdmin;
        if (currentStatus === 'Pending Team Leader' && (isTL || isAdmin)) authorized = true;
        if (currentStatus === 'Pending Vice Supervisor' && (isVice || isAdmin)) authorized = true;
        if (currentStatus === 'Pending Supervisor' && (isSPV || isAdmin)) authorized = true;
        if (currentStatus === 'Pending HR' && (isHR || isAdmin)) authorized = true;

        if (!authorized) return { success: false, message: 'Anda tidak memiliki wewenang untuk tahap approval ini (' + currentStatus + ').' };

        let newStatus = '';
        if (action === 'Reject') {
          newStatus = 'Ditolak';
        } else if (action === 'Approve') {
          if (isAdmin) newStatus = 'Disetujui'; // Admin bypasses flow
          else if (currentStatus === 'Pending Team Leader') newStatus = 'Pending Vice Supervisor';
          else if (currentStatus === 'Pending Vice Supervisor') newStatus = 'Pending Supervisor';
          else if (currentStatus === 'Pending Supervisor') newStatus = 'Pending HR';
          else if (currentStatus === 'Pending HR') newStatus = 'Disetujui';
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
        
        // AUTO DEDUCT CUTI (Hanya untuk Ijin Cuti)
        if (tipe === 'ijin' && action === 'Approve' && newStatus === 'Disetujui') {
          const jenis = String(data[i][3] || '').toLowerCase(); // index 3 = Jenis
          const namaKaryawan = data[i][2]; // index 2 = Nama
          if (jenis.includes('cuti')) {
            deductSisaCuti(namaKaryawan, 1);
          }
        }

        if (tipe === 'lembur') formatSheetLembur();
        
        return { success: true, newStatus: newStatus };
      }
    }
    return { success: false, message: 'Data tidak ditemukan (ID: '+id+', Nama: '+pemohonNama+', Tgl: '+tanggal+')' };
  } catch (e) { return { success: false, message: e.message }; }
}

/**
 * Mendapatkan semua data approval pending yang sudah dilengkapi dengan data Divisi
 * untuk keperluan Approval Massal.
 */
function getBulkApprovalData() {
  try {
    const res = getPendingApprovals();
    if (!res.success) return res;

    // Ambil mapping Nama -> Departemen dari Struktur Organisasi
    const orgRes = getOrganisasi();
    const orgMap = {};
    if (orgRes.success && orgRes.data) {
      orgRes.data.forEach(o => { if (o.nama) orgMap[o.nama.trim()] = o.departemen || 'Lainnya'; });
    }

    const processList = (list, type, module) => {
      return (list || []).map(item => {
        // Tentukan divisi: Prioritaskan field 'divisi' (Lembur), fallback ke Organisasi
        let div = item.divisi || orgMap[(item.nama || item.karyawan || '').trim()] || 'Lainnya';
        return {
          ...item,
          _type: type,
          _module: module,
          _divisi: div
        };
      });
    };

    return {
      success: true,
      data: [
        ...processList(res.ijin, 'Ijin/Cuti', 'ijin'),
        ...processList(res.lembur, 'Lembur', 'lembur'),
        ...processList(res.asset, 'Asset', 'asset'),
        ...processList(res.stockOpname, 'Stock Opname', 'opname')
      ]
    };
  } catch (e) { return { success: false, message: e.message }; }
}

/**
 * Memproses approval banyak item sekaligus (Batch)
 * @param {Array} items - List of {tipe, id, action, nama, tanggal}
 */
function processBatchApproval(items, userNama, userRole) {
  try {
    const results = { success: 0, failed: 0, errors: [] };
    
    // Batasi batch agar tidak timeout (maks 50)
    const batch = items.slice(0, 50);
    
    batch.forEach(item => {
      try {
        const res = processApprovalStatus(item.tipe, item.id, item.action, userNama, userRole, item.reason || '', item.nama, item.tanggal);
        if (res.success) results.success++;
        else {
          results.failed++;
          results.errors.push(item.nama + ': ' + res.message);
        }
      } catch(err) {
        results.failed++;
        results.errors.push(item.nama + ': ' + err.message);
      }
    });

    return { 
      success: true, 
      processed: results.success, 
      failed: results.failed, 
      errors: results.errors,
      total: items.length
    };
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
      if (data[i].join('').trim() === '') continue;
      result.push({
        id: String(data[i][0]),
        nama: data[i][1],
        jabatan: data[i][2],
        cabang: data[i][3] || '-',
        telepon: data[i][4],
        email: data[i][5],
        tanggalMasuk: data[i][6] instanceof Date ? Utilities.formatDate(data[i][6], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][6] || ''),
        status: data[i][7] || 'Tetap',
        createdAt: data[i][8] instanceof Date ? data[i][8].toISOString() : String(data[i][8] || ''),
        tanggalSelesai: data[i][9] instanceof Date ? Utilities.formatDate(data[i][9], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][9] || ''),
        sisaCuti: parseInt(data[i][10]) || 0
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function deductSisaCuti(nama, qty) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KARYAWAN);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === nama) {
        let current = parseInt(data[i][10]) || 0;
        sheet.getRange(i + 1, 11).setValue(Math.max(0, current - qty));
        return true;
      }
    }
  } catch(e) { Logger.log('Error deduct: ' + e.message); }
  return false;
}

function addKaryawan(nama, jabatan, cabang, telepon, email, tanggalMasuk, status, tanggalSelesai, sisaCuti) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KARYAWAN);
    const id = generateId();
    sheet.appendRow([id, nama, jabatan, cabang || '', telepon, email, tanggalMasuk, status || 'Tetap', new Date().toISOString(), tanggalSelesai || '', sisaCuti || 12]);
    return { success: true, id: id };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateKaryawan(id, nama, jabatan, cabang, telepon, email, tanggalMasuk, status, tanggalSelesai, sisaCuti) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KARYAWAN);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      // Find by ID or Name fallback
      if ((id && String(data[i][0]) === String(id)) || (!id && String(data[i][1]) === String(nama))) {
        sheet.getRange(i + 1, 2, 1, 7).setValues([[nama, jabatan, cabang || '', telepon, email, tanggalMasuk, status]]);
        sheet.getRange(i + 1, 10).setValue(tanggalSelesai || '');
        sheet.getRange(i + 1, 11).setValue(sisaCuti || 0);
        return { success: true };
      }
    }
    return { success: false, message: 'Karyawan tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteKaryawan(id) { 
  return deleteRow(CONFIG.SHEETS.KARYAWAN, id); 
}

function addBulkKaryawan(items) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KARYAWAN);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const now = new Date().toISOString();
    let updated = 0, added = 0;

    items.forEach(item => {
      let foundIndex = -1;
      for (let i = 1; i < data.length; i++) {
        // Match by ID if provided, otherwise by Name
        if ((item.id && String(data[i][0]) === String(item.id)) || (!item.id && String(data[i][1]).toLowerCase() === String(item.nama).toLowerCase())) {
          foundIndex = i;
          break;
        }
      }

      const rowData = [
        item.id || generateId(),
        item.nama,
        item.jabatan,
        item.cabang || '',
        item.telepon || '',
        item.email || '',
        item.tanggalMasuk || '',
        item.status || 'Tetap',
        now,
        item.tanggalSelesai || '',
        item.sisaCuti || 12
      ];

      if (foundIndex > -1) {
        sheet.getRange(foundIndex + 1, 1, 1, rowData.length).setValues([rowData]);
        updated++;
      } else {
        sheet.appendRow(rowData);
        added++;
      }
    });

    return { success: true, updated, added };
  } catch (e) { return { success: false, message: e.message }; }
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
      if (data[i].join('').trim() === '') continue;
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
    const historyArr = [{ date: new Date().toISOString(), action: 'Diajukan', status: 'Pending Team Leader', by: createdBy, role: 'Pemohon', reason: '' }];
    sheet.appendRow([generateId(), tanggal, nama, jenis, keterangan, bukti, 'Pending Team Leader', createdBy, new Date().toISOString(), JSON.stringify(historyArr)]);
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
      if (data[i].join('').trim() === '') continue;
      result.push({
        id: data[i][0],
        tanggal: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]),
        nama: data[i][2],
        divisi: data[i][3],
        jumlahJam: data[i][4],
        keterangan: data[i][5],
        status: data[i][6],
        createdBy: data[i][7],
        createdAt: data[i][8] instanceof Date ? data[i][8].toISOString() : String(data[i][8]),
        history: data[i][9] || '[]'
      });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addLembur(tanggal, nama, divisi, jumlahJam, keterangan, createdBy) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.LEMBUR);
    const historyArr = [{ date: new Date().toISOString(), action: 'Diajukan', status: 'Pending Team Leader', by: createdBy, role: 'Pemohon', reason: '' }];
    sheet.appendRow([generateId(), tanggal, nama, divisi, jumlahJam, keterangan, 'Pending Team Leader', createdBy, new Date().toISOString(), JSON.stringify(historyArr)]);
    formatSheetLembur(); // Rapihkan setelah tambah
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteLembur(id) { 
  const res = deleteRow(CONFIG.SHEETS.LEMBUR, id);
  if (res.success) formatSheetLembur(); // Rapihkan setelah hapus
  return res;
}

// ============================================================
// LEMBUR TANGGAL MERAH
// ============================================================
function getTglMerahData() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.TGL_MERAH);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      result.push({
        id: data[i][0],
        tanggal: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]),
        nama: data[i][2],
        divisi: data[i][3],
        jamEstimasi: data[i][4],
        createdBy: data[i][5],
        createdAt: data[i][6]
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function addTglMerahPersonel(tanggal, dataArray, createdBy) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.TGL_MERAH);
    const now = new Date().toISOString();
    dataArray.forEach(d => {
      sheet.appendRow([generateId(), tanggal, d.nama, d.divisi, d.jamEstimasi, createdBy, now]);
    });
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteTglMerah(id) {
  return deleteRow(CONFIG.SHEETS.TGL_MERAH, id);
}


// updateApprovalStatus didepresiasi, beralih ke processApprovalStatus

function updateLemburAdmin(id, jam, status, note, adminName) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.LEMBUR);
    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        let history = [];
        try { history = JSON.parse(data[i][9] || '[]'); } catch (e) { history = []; }
        history.push({ date: now, action: 'Edit Admin', status: status, by: adminName, role: 'admin', reason: note || 'Perubahan jam ' + data[i][4] + ' -> ' + jam });

        sheet.getRange(i+1, 5).setValue(jam);
        sheet.getRange(i+1, 7).setValue(status);
        sheet.getRange(i+1, 10).setValue(JSON.stringify(history));
        
        formatSheetLembur();
        return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

function formatSheetLembur() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.LEMBUR);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow === 0) return;

    // 1. Format Header
    const headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setFontWeight('bold')
               .setBackground('#1e293b')
               .setFontColor('#ffffff')
               .setHorizontalAlignment('center')
               .setVerticalAlignment('middle');

    // 2. Auto Resize Columns
    for (let i = 1; i <= lastCol; i++) {
      sheet.autoResizeColumn(i);
    }

    // 3. Border & Alignment untuk Data
    if (lastRow > 1) {
      const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
      dataRange.setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
      dataRange.setVerticalAlignment('middle');
      
      // Tengah untuk kolom tertentu (ID, Tanggal, Jam, Status)
      sheet.getRange(2, 1, lastRow - 1, 1).setHorizontalAlignment('center'); // ID
      sheet.getRange(2, 2, lastRow - 1, 1).setHorizontalAlignment('center'); // Tanggal
      sheet.getRange(2, 5, lastRow - 1, 1).setHorizontalAlignment('center'); // Jam
      sheet.getRange(2, 7, lastRow - 1, 1).setHorizontalAlignment('center'); // Status
    }

    // 4. Aktifkan Filter jika belum ada
    const filter = sheet.getFilter();
    if (filter) filter.remove();
    sheet.getRange(1, 1, lastRow, lastCol).createFilter();

    // 5. Freeze Header
    sheet.setFrozenRows(1);

  } catch (e) {
    Logger.log('Error formatSheetLembur: ' + e.message);
  }
}

// ============================================================
// ABSENSI LEMBUR (QR Code)
// ============================================================
function getAbsensiLembur(tanggal) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ABSENSI_LEMBUR);
    const data = sheet.getDataRange().getValues();
    const targetDate = tanggal || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
        const rowDate = data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]);
        if (rowDate === targetDate) {
            result.push({
                id: data[i][0],
                tanggal: rowDate,
                jam: data[i][2],
                nama: data[i][3],
                divisi: data[i][4],
                karyawanId: data[i][5],
                status: data[i][6],
                createdAt: data[i][7]
            });
        }
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function addAbsensiLembur(nama, divisi, karyawanId) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ABSENSI_LEMBUR);
    const now = new Date();
    const tanggal = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const jam = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
    
    // Check if already clocked in today
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
        const rowDate = data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]);
        if (rowDate === tanggal && String(data[i][3]) === String(nama)) {
            return { success: false, message: nama + ' sudah absen hari ini pada ' + data[i][2] };
        }
    }
    
    sheet.appendRow([generateId(), tanggal, jam, nama, divisi, karyawanId || '', 'Clock IN', now.toISOString()]);
    return { success: true, data: { nama, jam, tanggal } };
  } catch (e) { return { success: false, message: e.message }; }
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
      if (data[i].join('').trim() === '') continue;
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
    if(!r.success) return r;
    let s = 0; r.data.forEach(d => { s += d.tipe === 'IN' ? d.nominal : -d.nominal; });
    return { success: true, saldo: s };
  } catch(e) { return { success: false, message: e.message }; }
}

function getTeamBuilding() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.TEAM_BUILDING);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
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
  } catch(e) { return { success: false, message: e.message }; }
}

function addTeamBuilding(tanggal, keterangan, nominal, buktiUrl, createdBy, tipe) {
  try {
    getSheet(CONFIG.SHEETS.TEAM_BUILDING).appendRow([generateId(), tanggal, keterangan, parseFloat(nominal), buktiUrl, createdBy, new Date().toISOString(), tipe || 'Pengeluaran']);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteTeamBuilding(id) { return deleteRow(CONFIG.SHEETS.TEAM_BUILDING, id); }

function getSaldoTeamBuilding() {
  try {
    const r = getTeamBuilding();
    if(!r.success) return r;
    let s = 0; r.data.forEach(d => { s += d.tipe === 'Pemasukan' ? d.nominal : -d.nominal; });
    return { success: true, saldo: s };
  } catch(e) { return { success: false, message: e.message }; }
}

function getExpense() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.EXPENSE);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
        if (data[i].join('').trim() === '') continue;
        result.push({
            id:data[i][0], tanggal:data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]),
            perusahaan:data[i][2], kategori:data[i][3], keterangan:data[i][4], nominal:parseFloat(data[i][5])||0, bank:data[i][6], rekening:data[i][7], createdBy:data[i][8], createdAt:data[i][9]
        });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addExpense(tanggal, perusahaan, kategori, keterangan, nominal, bank, rekening, createdBy) {
  try { getSheet(CONFIG.SHEETS.EXPENSE).appendRow([generateId(), tanggal, perusahaan, kategori, keterangan, parseFloat(nominal)||0, bank, rekening, createdBy, new Date().toISOString()]); return { success: true }; } catch(e) { return { success: false, message: e.message }; }
}

function deleteExpense(id) { return deleteRow(CONFIG.SHEETS.EXPENSE, id); }

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
    let folder;
    if (folderName === 'Bukti Packing') {
      try {
        folder = DriveApp.getFolderById('1lE_NWzThv9MdODkmtYjScWz-Bb3N8ocA');
      } catch(e) {
        return { success: false, message: 'Gagal akses folder Bukti Packing: ' + e.message };
      }
    } else {
      folder = getOrCreateBuktiFolder(folderName);
    }
    
    const decoded = Utilities.base64Decode(fullBase64);
    const blob = Utilities.newBlob(decoded, mimeType || 'application/octet-stream', fileName);
    const file = folder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(err) {} // abaikan error permission domain
    return { success: true, url: file.getUrl() };
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
      if (data[i].join('').trim() === '') continue;
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
        sisaOrder: parseInt(data[i][15]) || 0,
        staffLemburNames: data[i][16] || '',
        shift: data[i][17] || 'Pagi',
        totalPHL: parseInt(data[i][18]) || 0,
        jamKerjaPHL: parseFloat(data[i][19]) || 0,
        totalPO: parseInt(data[i][20]) || 0,
        totalQty: parseInt(data[i][21]) || 0,
        totalInbound: parseInt(data[i][22]) || 0,
        pendapatanPotongBubble: parseFloat(data[i][23]) || 0,
        pendapatanBuatBubble: parseFloat(data[i][24]) || 0,
        alasanPengurangan: data[i][25] || ''
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function addLaporanKerja(tanggal, divisi, pic, totalOrang, perbantuan, pengurangan, jamLembur, totalJamKerja, kendala, totalStaff, totalAdmin, totalOrder, createdBy, sisaOrder, staffLemburNames, shift, totalPHL, jamKerjaPHL, totalPO, totalQty, totalInbound, pendapatanPotongBubble, pendapatanBuatBubble, alasanPengurangan) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.LAPORAN_KERJA);
    const data = sheet.getDataRange().getValues();
    const targetDateStr = String(tanggal).split('T')[0];
    const targetShift = shift || 'Pagi';

    for (let i = 1; i < data.length; i++) {
      const rowDate = data[i][1];
      const rowDateStr = rowDate instanceof Date ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(rowDate).split('T')[0];
      const rowDivisi = String(data[i][2]);
      const rowShift = String(data[i][17] || 'Pagi');

      if (rowDateStr === targetDateStr && rowDivisi === divisi && rowShift === targetShift) {
        return { success: false, message: `Laporan untuk ${divisi} (${targetShift}) pada tanggal ${targetDateStr} sudah ada. Silakan gunakan fitur Edit.` };
      }
    }

    sheet.appendRow([
      generateId(), tanggal, divisi, pic, parseInt(totalOrang)||0, parseFloat(perbantuan)||0, parseFloat(pengurangan)||0, parseFloat(jamLembur)||0, parseFloat(totalJamKerja)||0, kendala, parseInt(totalStaff)||0, parseInt(totalAdmin)||0, parseInt(totalOrder)||0, createdBy, new Date().toISOString(), parseInt(sisaOrder)||0, staffLemburNames || '', targetShift, parseInt(totalPHL)||0, parseFloat(jamKerjaPHL)||0, parseInt(totalPO)||0, parseInt(totalQty)||0, parseInt(totalInbound)||0, parseFloat(pendapatanPotongBubble)||0, parseFloat(pendapatanBuatBubble)||0, alasanPengurangan || ''
    ]);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}
function deleteLaporanKerja(id) { return deleteRow(CONFIG.SHEETS.LAPORAN_KERJA, id); }
function updateLaporanKerja(id, tanggal, divisi, pic, totalOrang, perbantuan, pengurangan, jamLembur, totalJamKerja, kendala, totalStaff, totalAdmin, totalOrder, createdBy, sisaOrder, staffLemburNames, shift, totalPHL, jamKerjaPHL, totalPO, totalQty, totalInbound, pendapatanPotongBubble, pendapatanBuatBubble, alasanPengurangan) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.LAPORAN_KERJA);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i + 1, 2, 1, 25).setValues([[
          tanggal, divisi, pic, parseInt(totalOrang)||0, parseFloat(perbantuan)||0, parseFloat(pengurangan)||0, parseFloat(jamLembur)||0, parseFloat(totalJamKerja)||0, kendala, parseInt(totalStaff)||0, parseInt(totalAdmin)||0, parseInt(totalOrder)||0, createdBy, new Date().toISOString(), parseInt(sisaOrder)||0, staffLemburNames || '', shift || 'Pagi', parseInt(totalPHL)||0, parseFloat(jamKerjaPHL)||0, parseInt(totalPO)||0, parseInt(totalQty)||0, parseInt(totalInbound)||0, parseFloat(pendapatanPotongBubble)||0, parseFloat(pendapatanBuatBubble)||0, alasanPengurangan || ''
        ]]);
        return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

// ============================================================
// HANDOVER & KLAIM
// ============================================================
function getHandover() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.HANDOVER);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
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
function updateHandoverStatus(id, status, resiFallback) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.HANDOVER); const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { 
      if ((id && String(data[i][0]) === String(id)) || (!id && resiFallback && String(data[i][3]) === String(resiFallback))) { 
        sheet.getRange(i + 1, 7).setValue(status); return { success: true }; 
      } 
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}
function deleteHandover(id) { return deleteRow(CONFIG.SHEETS.HANDOVER, id); }

function getKlaim() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KLAIM); const data = sheet.getDataRange().getValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
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
function updateKlaimStatus(id, status, resiFallback) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KLAIM); const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { 
      if ((id && String(data[i][0]) === String(id)) || (!id && resiFallback && String(data[i][3]) === String(resiFallback))) { 
        sheet.getRange(i + 1, 7).setValue(status); return { success: true }; 
      } 
    }
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
      if (data[i].join('').trim() === '') continue;
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
      if (data[i].join('').trim() === '') continue;
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
      if (data[i].join('').trim() === '') continue;
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
      if (data[i].join('').trim() === '') continue;
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
      // Find by ID or fallback to SKU if ID is empty/not matching
      if ((id && String(data[i][0]) === String(id)) || (!id && String(data[i][1]) === String(sku))) {
        sheet.getRange(i+1, 2, 1, 10).setValues([[sku, nama, barcode, batch, expDate, satuan, parseFloat(stok)||0, parseFloat(stokMin)||0, kategori, lokasi]]);
        sheet.getRange(i+1, 13).setValue(new Date().toISOString());
        return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan (ID/SKU tidak cocok)' };
  } catch(e) { return { success: false, message: e.message }; }
}
function deleteStock(id) { return deleteRow(CONFIG.SHEETS.STOCK, id); }

function updateStokQty(id, delta, skuFallback) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.STOCK); const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if ((id && String(data[i][0]) === String(id)) || (!id && skuFallback && String(data[i][1]) === String(skuFallback))) {
        const cur = parseFloat(data[i][7]) || 0;
        sheet.getRange(i+1, 8).setValue(cur + delta);
        sheet.getRange(i+1, 13).setValue(new Date().toISOString());
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
      if (data[i].join('').trim() === '') continue;
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
      updateStokQty(item.stockId, parseFloat(item.qty)||0, item.sku);
    });
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function getSJMasukDetail(sjId, noSJFallback) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SURAT_JALAN_MASUK_DETAIL); const data = sheet.getDataRange().getValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      const match = (sjId && String(data[i][1]) === String(sjId)) || (!sjId && noSJFallback && String(data[i][2]) === String(noSJFallback));
      if (!match) continue;
      result.push({ id:data[i][0], sjId:data[i][1], noSJ:data[i][2], stockId:data[i][3], sku:data[i][4], nama:data[i][5], qty:parseFloat(data[i][6])||0, satuan:data[i][7], batch:data[i][8], expDate:data[i][9] });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function getSJDetailData(sjId, tipe, noSJFallback) {
  try {
    const sheetName = tipe === 'masuk' ? CONFIG.SHEETS.SURAT_JALAN_MASUK_DETAIL : CONFIG.SHEETS.SURAT_JALAN_KELUAR_DETAIL;
    const data = getSheet(sheetName).getDataRange().getValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      const match = (sjId && String(data[i][1]) === String(sjId)) || (!sjId && noSJFallback && String(data[i][2]) === String(noSJFallback));
      if (match) result.push({ sku:data[i][4], nama:data[i][5], qty:parseFloat(data[i][6])||0, satuan:data[i][7], batch:data[i][8], expDate:data[i][9] });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function getSuratJalanKeluar() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SURAT_JALAN_KELUAR); const data = sheet.getDataRange().getValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
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
      const res = updateStokQty(item.stockId, -(parseFloat(item.qty)||0), item.sku);
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
      if (data[i].join('').trim() === '') continue;
      // Index 11 is Column L (buktiPacking)
      const bukti = data[i].length > 11 ? (data[i][11] || '') : '';
      result.push({
        id: String(data[i][0] || ''), 
        noOrder: String(data[i][1] || ''),
        tanggal: data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][2]||''),
        pelanggan: String(data[i][3] || ''), 
        alamat: String(data[i][4] || ''), 
        status: String(data[i][5] || ''), 
        totalItem: parseFloat(data[i][6]) || 0, 
        keterangan: String(data[i][7] || ''), 
        createdBy: String(data[i][8] || ''),
        createdAt: data[i][9] instanceof Date ? data[i][9].toISOString() : String(data[i][9]||''),
        sentAt: data[i][10] instanceof Date ? data[i][10].toISOString() : String(data[i][10]||''),
        buktiPacking: bukti,
        kategori: String(data[i][12] || ''),
        noResi: String(data[i][13] || '')
      });


    }
    // Sort latest first
    return { success: true, data: result.reverse() };
  } catch(e) { return { success: false, message: e.message }; }
}

function getOrdersWithDetails() {
  try {
    const ordRes = getOrders();
    if (!ordRes.success) return ordRes;
    
    const orders = ordRes.data;
    const detSheet = getSheet(CONFIG.SHEETS.ORDER_DETAIL);
    const detData = detSheet.getDataRange().getValues();
    
    // Create a map for quick lookup
    const detailsMap = {};
    for (let i = 1; i < detData.length; i++) {
      const orderId = String(detData[i][1]);
      if (!detailsMap[orderId]) detailsMap[orderId] = [];
      detailsMap[orderId].push({
        sku: detData[i][4],
        nama: detData[i][5],
        qty: parseFloat(detData[i][6]) || 0,
        satuan: detData[i][7],
        batch: detData[i][8] || '-',
        expDate: detData[i][9] || '-'
      });
    }
    
    // Merge details into orders
    orders.forEach(o => {
      o.items = detailsMap[String(o.id)] || [];
    });
    
    return { success: true, data: orders };
  } catch(e) { return { success: false, message: e.message }; }
}

function addOrder(tanggal, pelanggan, alamat, keterangan, items, createdBy, kategori, noResi) {
  try {
    let noOrder = (kategori === 'Marketplace' && pelanggan) ? pelanggan : generateNoSJ('ORD'); // Jika MP, pelanggan mungkin berisi No Order Custom
    // Namun sesuai permintaan: Marketplace -> Custom No Order, No Resi
    // Distributor/Store -> Auto No Order
    
    if (kategori === 'Marketplace') {
      // Kita asumsikan 'pelanggan' diisi dengan No Order Custom jika Marketplace? 
      // Atau kita tambah parameter? Mari kita konsisten dengan input frontend nanti.
      // Saya akan gunakan parameter 'pelanggan' sebagai Nama Pelanggan, dan No Order bisa dari parameter lain atau pinter-pinteran.
      // Revisi: Saya akan biarkan 'noOrder' dihandle di frontend atau di sini.
    }
    
    // Mari kita buat lebih eksplisit
    const id = generateId();
    
    // Logic untuk menentukan No. Order dan No. Resi khusus Marketplace
    let finalNoOrder = generateNoSJ('ORD');
    let finalNoResi = '';

    if (kategori === 'Marketplace' && noResi) {
      // Jika noResi dikirim sebagai objek {customNoOrder: ..., noResi: ...}
      if (typeof noResi === 'object') {
        finalNoOrder = noResi.customNoOrder || finalNoOrder;
        finalNoResi = noResi.noResi || '';
      } else {
        // Fallback jika noResi dikirim sebagai string (untuk backward compatibility)
        finalNoResi = noResi;
      }
    }
    
    const parsedItems = typeof items === 'string' ? JSON.parse(items) : items;
    const totalItem = parsedItems.reduce((s,x) => s + (parseFloat(x.qty)||0), 0);
    
    getSheet(CONFIG.SHEETS.ORDER).appendRow([
      id, finalNoOrder, tanggal, pelanggan, alamat, 'Pending', totalItem, keterangan, createdBy, 
      new Date().toISOString(), '', '', kategori || 'Distributor', finalNoResi || ''
    ]);

    
    const detSheet = getSheet(CONFIG.SHEETS.ORDER_DETAIL);
    parsedItems.forEach(item => {
      detSheet.appendRow([generateId(), id, finalNoOrder, item.stockId, item.sku, item.nama, parseFloat(item.qty)||0, item.satuan, item.batch||'', item.expDate||'', 0, item.lokasi || '']);
    });
    return { success: true, noOrder: finalNoOrder };
  } catch(e) { return { success: false, message: e.message }; }
}

function getOrderDetail(orderId, noOrderFallback) {
  try {
    const data = getSheet(CONFIG.SHEETS.ORDER_DETAIL).getDataRange().getValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      const match = (orderId && String(data[i][1]) === String(orderId)) || (!orderId && noOrderFallback && String(data[i][2]) === String(noOrderFallback));
      if (match) {
        result.push({ 
          id:data[i][0], orderId:data[i][1], noOrder:data[i][2], stockId:data[i][3], 
          sku:data[i][4], nama:data[i][5], qty:parseFloat(data[i][6])||0, 
          satuan:data[i][7], batch:data[i][8]||'-', expDate:data[i][9]||'-', 
          packedQty:parseFloat(data[i][10])||0,
          lokasi:data[i][11]||'-'
        });
      }
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function updateBuktiPackingUrl(orderId, noOrderFallback, url) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ORDER);
    const data = sheet.getDataRange().getValues();
    let updated = false;
    for (let i = 1; i < data.length; i++) {
      const matchId = orderId && String(data[i][0]).trim() === String(orderId).trim();
      const matchNo = !orderId && noOrderFallback && String(data[i][1]).trim() === String(noOrderFallback).trim();
      
      if (matchId || matchNo) {
        sheet.getRange(i+1, 12).setValue(url); // Col L (buktiPacking)
        updated = true;
        break;
      }
    }
    if (!updated) return { success: false, message: 'ID/No Order tidak ditemukan di database.' };
    return { success: true, url: url };
  } catch(e) { return { success: false, message: 'Simpan URL Error: ' + e.message }; }
}

function deleteOrder(id, noOrder) { return deleteRow(CONFIG.SHEETS.ORDER, id, noOrder, 1); }

function kirimOrder(id, noOrderFallback) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ORDER); const data = sheet.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 1; i < data.length; i++) {
      if ((id && String(data[i][0]) === String(id)) || (!id && noOrderFallback && String(data[i][1]) === String(noOrderFallback))) {
        if (data[i][5] === 'Terkirim') return { success: false, message: 'Sudah terkirim' };
        foundRow = i + 1;
        break;
      }
    }
    if (foundRow === -1) return { success: false, message: 'Order Tidak ditemukan' };
    
    sheet.getRange(foundRow, 6).setValue('Terkirim');
    sheet.getRange(foundRow, 11).setValue(new Date().toISOString());
    
    const rowId = data[foundRow-1][0];
    const rowNoOrder = data[foundRow-1][1];
    
    const detData = getSheet(CONFIG.SHEETS.ORDER_DETAIL).getDataRange().getValues();
    for(let i=1; i<detData.length; i++) {
      // Link by ID or noOrder
      const match = (rowId && String(detData[i][1]) === String(rowId)) || (rowNoOrder && String(detData[i][2]) === String(rowNoOrder));
      if(match) { updateStokQty(detData[i][3], -(parseFloat(detData[i][6])||0), detData[i][4]); }
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function getRetur() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.RETUR); const data = sheet.getDataRange().getValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
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
      const res = updateStokQty(item.stockId, parseFloat(item.qty)||0, item.sku);
      if (!res.success) return { success: false, message: res.message };
      detSheet.appendRow([generateId(), id, noRetur, item.stockId, item.sku, item.nama, parseFloat(item.qty)||0, item.satuan, item.batch||'', item.expDate||'']);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function getReturDetail(returId, noReturFallback) {
  try {
    const data = getSheet(CONFIG.SHEETS.RETUR_DETAIL).getDataRange().getValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      const match = (returId && String(data[i][1]) === String(returId)) || (!returId && noReturFallback && String(data[i][2]) === String(noReturFallback));
      if (match) result.push({ sku:data[i][4], nama:data[i][5], qty:parseFloat(data[i][6])||0, satuan:data[i][7], batch:data[i][8], expDate:data[i][9] });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteRetur(id, noRetur) { return deleteRow(CONFIG.SHEETS.RETUR, id, noRetur, 1); }

// ============================================================
// ANALISIS STOK & BULK IMPORT
// ============================================================
function importOrdersBulk(jsonString, createdBy) {
  try {
    const orders = JSON.parse(jsonString); 
    let count = 0; 
    const stData = getStock().data || [];

    for(let i=0; i < orders.length; i++) {
      const o = orders[i];
      // Items sudah lengkap dari frontend (stockId, sku, qty, batch, expDate)
      // Kita perlu melengkapi nama barang & satuan dari stockData untuk addOrder (OrderDetail)
      const mappedItems = o.items.map(item => {
        const found = stData.find(s => s.id === item.stockId);
        return {
          stockId: item.stockId,
          sku: item.sku,
          nama: found ? found.nama : item.sku,
          qty: item.qty,
          satuan: found ? found.satuan : 'PCS',
          batch: item.batch,
          expDate: item.expDate,
          lokasi: found ? found.lokasi : '-'
        };
      });

      // Kirim customNoOrder & noResi sebagai objek resiParam
      const resiParam = { 
        customNoOrder: o.customNoOrder || '', 
        noResi: o.noResi || '' 
      };
      
      addOrder(o.tanggal, o.pelanggan, o.alamat, 'Import Excel', mappedItems, createdBy, o.kategori, resiParam);
      count++;
    }
    return { success: true, count };
  } catch(e) { return { success: false, message: e.message }; }
}



function importInboundBulk(jsonString, createdBy) {
  try {
    const inbounds = JSON.parse(jsonString); let count = 0; const stData = getStock().data || [];
    for(let i=0; i<inbounds.length; i++) {
      const b = inbounds[i];
      const mappedItems = b.items.map(item => {
        let stId = ''; let stNama = item.sku; let stSatuan = 'PCS';
        const found = stData.find(s => s.sku === item.sku);
        if(found) { stId = found.id; stNama = found.nama; stSatuan = found.satuan; }
        return { stockId: stId, sku: item.sku, nama: stNama, qty: item.qty, satuan: stSatuan, batch: item.batch||'', expDate: item.expDate||'' };
      });
      addSuratJalanMasuk(b.tanggal, b.supplier, b.keterangan, mappedItems, createdBy);
      count++;
    }
    return { success: true, count };
  } catch(e) { return { success: false, message: e.message }; }
}

function importReturBulk(jsonString, createdBy) {
  try {
    const returns = JSON.parse(jsonString); let count = 0; const stData = getStock().data || [];
    for(let i=0; i<returns.length; i++) {
      const r = returns[i];
      const mappedItems = r.items.map(item => {
        let stId = ''; let stNama = item.sku; let stSatuan = 'PCS';
        const found = stData.find(s => s.sku === item.sku);
        if(found) { stId = found.id; stNama = found.nama; stSatuan = found.satuan; }
        return { stockId: stId, sku: item.sku, nama: stNama, qty: item.qty, satuan: stSatuan, batch: item.batch||'', expDate: item.expDate||'' };
      });
      addRetur(r.tanggal, r.sumber, r.alasan, r.keterangan, mappedItems, createdBy);
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
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      result.push({
        id: String(data[i][0]),
        judul: data[i][1] || '',
        assignee: data[i][2] || '',
        assigneeName: data[i][3] || '',
        prioritas: data[i][4] || 'Sedang',
        tanggalMulai: data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][5]||''),
        deadline: data[i][6] instanceof Date ? Utilities.formatDate(data[i][6], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][6]||''),
        targetHari: parseInt(data[i][7]) || 0,
        status: data[i][8] || 'Todo',
        kategori: data[i][9] || '',
        deskripsi: data[i][10] || '',
        createdBy: data[i][11] || '',
        createdAt: data[i][12] instanceof Date ? data[i][12].toISOString() : String(data[i][12]||''),
        updatedAt: data[i][13] instanceof Date ? data[i][13].toISOString() : String(data[i][13]||''),
        log: data[i][14] || '[]'
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
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

// ============================================================
// PENGAJUAN ASSET
// ============================================================
function getAsset() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ASSET);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      result.push({
        id: data[i][0],
        tanggal: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]),
        nama: data[i][2],
        jenisAsset: data[i][3],
        deskripsi: data[i][4],
        estimasiHarga: parseFloat(data[i][5]) || 0,
        prioritas: data[i][6],
        bukti: data[i][7],
        status: data[i][8],
        createdBy: data[i][9],
        createdAt: data[i][10] instanceof Date ? data[i][10].toISOString() : String(data[i][10] || ''),
        history: data[i][11] || '[]'
      });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addAsset(tanggal, nama, jenisAsset, deskripsi, estimasiHarga, prioritas, bukti, createdBy) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ASSET);
    const historyArr = [{ date: new Date().toISOString(), action: 'Diajukan', status: 'Pending Team Leader', by: createdBy, role: 'Pemohon', reason: '' }];
    sheet.appendRow([generateId(), tanggal, nama, jenisAsset, deskripsi, parseFloat(estimasiHarga)||0, prioritas, bukti || '', 'Pending Team Leader', createdBy, new Date().toISOString(), JSON.stringify(historyArr)]);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteAsset(id) {
  return deleteRow(CONFIG.SHEETS.ASSET, id);
}

// ============================================================
// STOCK OPNAME
// ============================================================
function getStockOpname() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.STOCK_OPNAME); const data = sheet.getDataRange().getValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      result.push({
        id:data[i][0], tanggal:data[i][1] instanceof Date ? data[i][1].toISOString().split('T')[0] : String(data[i][1]||''), 
        stockId:data[i][2], sku:data[i][3], nama:data[i][4], lokasi:data[i][5], batch:data[i][6], expDate:data[i][7],
        stokSistem:data[i][8], stokFisik:data[i][9], selisih:data[i][10], status:data[i][11], catatan:data[i][12],
        createdBy:data[i][13], createdAt:data[i][14], approvedBy:data[i][15], approvedAt:data[i][16]
      });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function submitStockOpname(tanggal, stockId, sku, nama, lokasi, batch, expDate, sistem, fisik, catatan, createdBy) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.STOCK_OPNAME);
    const selisih = parseFloat(fisik) - parseFloat(sistem);
    sheet.appendRow([generateId(), tanggal, stockId, sku, nama, lokasi||'-', batch||'-', expDate||'-', sistem, fisik, selisih, 'Pending', catatan||'', createdBy, new Date().toISOString(), '', '']);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function approveStockOpname(id, status, approvedBy) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.STOCK_OPNAME); const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        if (data[i][11] !== 'Pending') return { success: false, message: 'Sudah diproses' };
        sheet.getRange(i+1, 12).setValue(status); // Status: Column L
        sheet.getRange(i+1, 16).setValue(approvedBy); // ApprovedBy: Column P
        sheet.getRange(i+1, 17).setValue(new Date().toISOString()); // ApprovedAt: Column Q
        
        if (status === 'Approved') {
          const stockId = data[i][2];
          const fisik = data[i][9]; // FIXED: Fisik is at index 9
          const stSheet = getSheet(CONFIG.SHEETS.STOCK);
          const stData = stSheet.getDataRange().getValues();
          for(let j=1; j<stData.length; j++) {
            if(String(stData[j][0]) === String(stockId)) {
              stSheet.getRange(j+1, 8).setValue(fisik); // Update master stok ke angka fisik
              break;
            }
          }
        }
        return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch(e) { return { success: false, message: e.message }; }
}

// ============================================================
// PACKING LIST
// ============================================================
function getPackingList() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PACKING_LIST); const data = sheet.getDataRange().getValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      result.push({ 
        id:data[i][0], 
        tanggal:data[i][1] instanceof Date ? data[i][1].toISOString().split('T')[0] : String(data[i][1]||''), 
        noPL:data[i][2], keterangan:data[i][3], fileUrl:data[i][4], createdBy:data[i][5], createdAt:data[i][6] 
      });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addPackingList(tanggal, noPL, keterangan, fileUrl, createdBy) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PACKING_LIST);
    if (!sheet) {
      const ss = getSpreadsheet();
      setupSheet(ss, CONFIG.SHEETS.PACKING_LIST, ['id','tanggal','noPL','keterangan','fileUrl','createdBy','createdAt']);
    }
    const finalSheet = getSheet(CONFIG.SHEETS.PACKING_LIST);
    finalSheet.appendRow([generateId(), tanggal, noPL, keterangan, fileUrl, createdBy || 'User', new Date().toISOString()]);
    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) { return { success: false, message: 'Gagal Simpan ke Tabel: ' + e.message }; }
}

// UPLOAD HANDLER DUPLIKAT DIHAPUS - MENGGUNAKAN VERSI CACHESERVICE DI ATAS
// RIWAYAT KARYAWAN (RESIGN)
// ============================================================
function getRiwayatKaryawan() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.RIWAYAT_KARYAWAN);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      result.push({
        id: String(data[i][0]),
        nama: data[i][1],
        jabatan: data[i][2],
        cabang: data[i][3] || '-',
        telepon: data[i][4] || '-',
        tanggalMasuk: data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][5] || ''),
        tanggalResign: data[i][6] instanceof Date ? Utilities.formatDate(data[i][6], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][6] || ''),
        alasanResign: data[i][7] || '-',
        keterangan: data[i][8] || '',
        createdBy: data[i][9] || '',
        createdAt: data[i][10] instanceof Date ? data[i][10].toISOString() : String(data[i][10] || '')
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function addRiwayatKaryawan(karyawanId, nama, jabatan, cabang, telepon, tanggalMasuk, tanggalResign, alasanResign, keterangan, createdBy) {
  try {
    // 1. Simpan ke sheet RiwayatKaryawan
    const id = generateId();
    getSheet(CONFIG.SHEETS.RIWAYAT_KARYAWAN).appendRow([
      id, nama, jabatan, cabang || '', telepon || '',
      tanggalMasuk, tanggalResign, alasanResign, keterangan || '',
      createdBy, new Date().toISOString()
    ]);
    
    // 2. Hapus dari sheet Karyawan (Pindah resmi)
    deleteRow(CONFIG.SHEETS.KARYAWAN, karyawanId);
    
    return { success: true, id: id };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteRiwayatKaryawan(id) {
  return deleteRow(CONFIG.SHEETS.RIWAYAT_KARYAWAN, id);
}

// ============================================================
// SURAT PERINGATAN (SP)
// ============================================================
function getSuratPeringatan() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SURAT_PERINGATAN);
    const data = sheet.getDataRange().getValues();
    const result = [];
    const now = new Date();
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      const kadaluarsa = data[i][7] instanceof Date ? data[i][7] : new Date(data[i][7]);
      const sisaHari = Math.ceil((kadaluarsa - now) / 86400000);
      result.push({
        id: String(data[i][0]),
        karyawanNama: data[i][1],
        karyawanId: String(data[i][2] || ''),
        jenisSP: data[i][3],
        alasan: data[i][4],
        tanggalSP: data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][5] || ''),
        masaBerlaku: parseInt(data[i][6]) || 180,
        tanggalKadaluarsa: data[i][7] instanceof Date ? Utilities.formatDate(data[i][7], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][7] || ''),
        sisaHari: sisaHari,
        status: sisaHari <= 0 ? 'Kadaluarsa' : 'Aktif',
        createdBy: data[i][9] || '',
        createdAt: data[i][10] instanceof Date ? data[i][10].toISOString() : String(data[i][10] || '')
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function addSuratPeringatan(karyawanNama, karyawanId, jenisSP, alasan, tanggalSP, masaBerlaku, createdBy) {
  try {
    const id = generateId();
    const tglSP = new Date(tanggalSP);
    const tglKadaluarsa = new Date(tglSP);
    tglKadaluarsa.setDate(tglKadaluarsa.getDate() + parseInt(masaBerlaku));
    const tglKadStr = Utilities.formatDate(tglKadaluarsa, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    getSheet(CONFIG.SHEETS.SURAT_PERINGATAN).appendRow([
      id, karyawanNama, karyawanId || '', jenisSP, alasan,
      tanggalSP, parseInt(masaBerlaku), tglKadStr,
      'Aktif', createdBy, new Date().toISOString()
    ]);
    return { success: true, id: id };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteSuratPeringatan(id) {
  return deleteRow(CONFIG.SHEETS.SURAT_PERINGATAN, id);
}

// BATCH DATA FOR FASTER LOADING
function getKaryawanFullData() {
  try {
    return {
      success: true,
      karyawan: getKaryawan(),
      riwayat: getRiwayatKaryawan(),
      sp: getSuratPeringatan(),
      ijin: getIjin()
    };
  } catch (e) { return { success: false, message: e.message }; }
}

// ============================================================
// ASSET WAREHOUSE
// ============================================================
function getAssetWarehouseData() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ASSET_WAREHOUSE);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
        if (data[i].join('').trim() === '') continue;
        result.push({
            id: data[i][0],
            code: data[i][1],
            nama: data[i][2],
            tanggalMasuk: data[i][3] instanceof Date ? Utilities.formatDate(data[i][3], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][3]||''),
            divisi: data[i][4],
            status: data[i][5] || 'Aktif',
            createdBy: data[i][6],
            createdAt: data[i][7],
            history: data[i][8],
            qty: data[i][9] || 1
        });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function addAssetWarehouse(codePrefix, nama, tanggalMasuk, divisi, status, createdBy, qty) {
  try {
    const id = generateId();
    const randomNum = Math.floor(10000 + Math.random() * 90000);
    const code = codePrefix ? `${codePrefix}-${randomNum}` : `AW-${randomNum}`;
    const createdAt = new Date().toISOString();
    const history = `🛒 Dibuat oleh ${createdBy} pada ${createdAt} (Tgl Masuk: ${tanggalMasuk})`;
    
    const sheet = getSheet(CONFIG.SHEETS.ASSET_WAREHOUSE);
    sheet.appendRow([id, code, nama, tanggalMasuk, divisi, status || 'Aktif', createdBy, createdAt, history, qty || 1]);
    return { success: true, code: code };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateAssetWarehouse(id, nama, tanggalMasuk, status, userNama, qty) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ASSET_WAREHOUSE);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i + 1, 3).setValue(nama);
        sheet.getRange(i + 1, 4).setValue(tanggalMasuk);
        sheet.getRange(i + 1, 6).setValue(status);
        sheet.getRange(i + 1, 10).setValue(qty || 1);
        
        let oldHist = data[i][8] || '';
        const now = new Date().toLocaleString('id-ID');
        const entry = `✏️ Diperbarui oleh ${userNama} pada ${now}`;
        sheet.getRange(i + 1, 9).setValue(oldHist ? oldHist + '\n' + entry : entry);
        
        return { success: true };
      }
    }
    return { success: false, message: 'Asset tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

function moveAssetWarehouse(assetId, targetDivisi, userNama) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ASSET_WAREHOUSE);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(assetId)) {
        const oldDiv = data[i][4];
        const oldHist = data[i][8] || '';
        const now = new Date().toLocaleString('id-ID');
        const entry = `📦 Dipindah dari ${oldDiv} ke ${targetDivisi} oleh ${userNama} pada ${now}`;
        
        sheet.getRange(i + 1, 5).setValue(targetDivisi);
        sheet.getRange(i + 1, 9).setValue(oldHist ? oldHist + '\n' + entry : entry);
        return { success: true };
      }
    }
    return { success: false, message: 'Asset tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteAssetWarehouse(id) {
  return deleteRow(CONFIG.SHEETS.ASSET_WAREHOUSE, id);
}

// ============================================================
// BOOKING MOBIL
// ============================================================
function getBookingMobil() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.BOOKING_MOBIL);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, data: [] };
    
    // Ambil data mulai dari baris 2 kolom 1 s/d baris terakhir kolom 10
    const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const result = [];
    
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row[0]) continue; // Skip jika ID kosong
        
        result.push({
            id: String(row[0]),
            tanggal: row[1] instanceof Date ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(row[1]||''),
            pic: String(row[2] || '-'),
            jamBerangkat: (function(val) {
                if (!val) return '-';
                const str = val instanceof Date ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm') : String(val);
                const match = str.match(/(\d{2}:\d{2})/);
                return match ? match[1] : str;
            })(row[3]),
            tujuan: String(row[4] || '-'),
            keterangan: String(row[5] || '-'),
            rute: String(row[6] || '-'),
            status: String(row[7] || 'Belum Jalan'),
            createdBy: String(row[8] || '-'),
            createdAt: row[9]
        });
    }
    return { success: true, data: result, totalOnSheet: lastRow - 1 };
  } catch (e) { return { success: false, message: e.message }; }
}

function addBookingMobil(tanggal, pic, jamBerangkat, tujuan, keterangan, rute, createdBy) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.BOOKING_MOBIL);
    const data = sheet.getDataRange().getValues();
    
    // Check for overlap (Same Date and Same Time)
    for (let i = 1; i < data.length; i++) {
      const rowTgl = data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]);
      const rowJam = String(data[i][3]);
      
      if (rowTgl === String(tanggal) && rowJam === String(jamBerangkat)) {
        return { success: false, message: 'Waktu tersebut sudah di-booking (' + rowJam + '). Silakan pilih waktu lain.' };
      }
    }

    const id = generateId();
    const createdAt = new Date().toISOString();
    sheet.appendRow([id, tanggal, pic, jamBerangkat, tujuan, keterangan, rute, 'Belum Jalan', createdBy, createdAt]);
    return { success: true, id: id };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteBookingMobil(id) {
  return deleteRow(CONFIG.SHEETS.BOOKING_MOBIL, id);
}

function updateBookingStatus(id, newStatus) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.BOOKING_MOBIL);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(id)) {
            sheet.getRange(i + 1, 8).setValue(newStatus);
            return { success: true };
        }
    }
    return { success: false, message: 'ID tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}
// ============================================================
// MODUL TUGAS CONSUMABLE
// ============================================================
function getTugasConsumable() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.TUGAS_CONSUMABLE);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      result.push({
        id: data[i][0],
        tanggal: data[i][1] ? Utilities.formatDate(new Date(data[i][1]), Session.getScriptTimeZone(), "yyyy-MM-dd") : '',
        pemberiTugas: data[i][2],
        picName: data[i][3],
        targetPotong: parseInt(data[i][4]) || 0,
        targetBuat: parseInt(data[i][5]) || 0,
        actualPotong: parseInt(data[i][6]) || 0,
        actualBuat: parseInt(data[i][7]) || 0,
        status: data[i][8],
        catatan: data[i][9],
        createdAt: data[i][10],
        updatedAt: data[i][11],
        finishedBy: data[i][12] || '-'
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function addTugasConsumable(data) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.TUGAS_CONSUMABLE);
    sheet.appendRow([
      generateId(), 
      data.tanggal, 
      data.pemberiTugas, 
      data.picName, 
      parseInt(data.targetPotong)||0, 
      parseInt(data.targetBuat)||0, 
      0, 0, 'Pending', '', 
      new Date().toISOString(), 
      new Date().toISOString(),
      '' // finishedBy
    ]);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateTugasConsumable(data) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.TUGAS_CONSUMABLE);
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]) === String(data.id)) {
            const row = i + 1;
            // Jika status Selesai, update hasil aktual
            if (data.status === 'Selesai') {
                sheet.getRange(row, 7).setValue(parseInt(data.actualPotong) || 0);
                sheet.getRange(row, 8).setValue(parseInt(data.actualBuat) || 0);
                sheet.getRange(row, 9).setValue('Selesai');
                sheet.getRange(row, 10).setValue(data.catatan || '');
                sheet.getRange(row, 12).setValue(new Date().toISOString()); // updatedAt
                sheet.getRange(row, 13).setValue(data.finishedBy || ''); // finishedBy
            } else {
                // Update basic info (Edit mode)
                sheet.getRange(row, 2).setValue(data.tanggal);
                sheet.getRange(row, 4).setValue(data.picName);
                sheet.getRange(row, 5).setValue(parseInt(data.targetPotong) || 0);
                sheet.getRange(row, 6).setValue(parseInt(data.targetBuat) || 0);
                sheet.getRange(row, 12).setValue(new Date().toISOString());
            }
            return { success: true };
        }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteTugasConsumable(id) {
  return deleteRow(CONFIG.SHEETS.TUGAS_CONSUMABLE, id);
}

function addBulkTugasConsumable(rows) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.TUGAS_CONSUMABLE);
    const now = new Date().toISOString();
    const rowsToAdd = rows.map(r => [
      generateId(), r.tanggal, r.pemberiTugas, r.picName, parseInt(r.targetPotong)||0, parseInt(r.targetBuat)||0, 0, 0, 'Pending', '', now, now
    ]);
    if (rowsToAdd.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, 12).setValues(rowsToAdd);
    }
    return { success: true, count: rowsToAdd.length };
  } catch (e) { return { success: false, message: e.message }; }
}

// ============================================================
// MODUL ABSENSI LEMBUR (SERVER SIDE)
// ============================================================
function getAbsensiLembur() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ABSENSI_LEMBUR);
    const data = sheet.getDataRange().getValues();
    const result = [];
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      const rowTgl = data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]);
      
      // Filter hanya untuk hari ini
      if (rowTgl === today) {
        result.push({
          id: data[i][0],
          tanggal: rowTgl,
          jam: data[i][2],
          nama: data[i][3],
          divisi: data[i][4],
          karyawanId: data[i][5],
          status: data[i][6],
          createdAt: data[i][7]
        });
      }
    }
    // Urutkan berdasarkan jam terbaru di atas
    result.sort((a, b) => b.jam.localeCompare(a.jam));
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function addAbsensiLembur(nama, divisi, karyawanId) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ABSENSI_LEMBUR);
    const data = sheet.getDataRange().getValues();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const jam = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss');

    // Cek duplikasi hari ini
    for (let i = 1; i < data.length; i++) {
      const rowTgl = data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]);
      const rowNama = String(data[i][3]).toLowerCase();
      if (rowTgl === today && rowNama === String(nama).toLowerCase()) {
        return { success: false, message: nama + ' sudah Clock IN hari ini.' };
      }
    }

    const id = generateId();
    sheet.appendRow([
      id, 
      today, 
      jam, 
      nama, 
      divisi || '-', 
      karyawanId || '', 
      'Clock IN', 
      new Date().toISOString()
    ]);
    
    return { success: true, message: 'Clock IN Berhasil: ' + nama, data: { id, today, jam, nama } };
  } catch (e) { return { success: false, message: e.message }; }
}
