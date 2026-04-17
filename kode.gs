// ============================================================
// GUDANG FCL - Google Apps Script Backend
// Code.gs - Main Server-Side Logic
// ============================================================

const CONFIG = {
  SPREADSHEET_ID: '1lde5La49rhI5NElJNtpaGP7ZMFcS9n28ZNRy6YyhU3s', // <-- GANTI INI DENGAN ID DARI URL SHEET ANDA
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
    ASSET_WAREHOUSE: 'AssetWarehouse',
    STOCK_OPNAME: 'StockOpname',
    PACKING_LIST: 'PackingList',
    RIWAYAT_KARYAWAN: 'RiwayatKaryawan',
    SURAT_PERINGATAN: 'SuratPeringatan',
    TUGAS_CONSUMABLE: 'TugasConsumable',
    TGL_MERAH: 'TglMerah',
    BOOKING_MOBIL: 'BookingMobil',
    ABSENSI_LEMBUR: 'AbsensiLembur',
    WAREHOUSE_MAP: 'WarehouseMap',
    ABSENSI_KARYAWAN: 'AbsensiKaryawan',
    JADWAL_SHIFT: 'JadwalShift',
    JADWAL_ROSTER: 'JadwalRoster'
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

// (Fungsi doPost lama yang bertumpuk sudah dipindahkan dan digabung ke paling bawah)

// ============================================================
// SETUP DATABASE
// ============================================================
function setupDatabase() {
  let ss;
  const props = PropertiesService.getScriptProperties();
  
  // Override dengan CONFIG ID agar tidak nyasar ke file lama
  let ssId = CONFIG.SPREADSHEET_ID || props.getProperty('SPREADSHEET_ID');
  
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
  setupSheet(ss, CONFIG.SHEETS.KARYAWAN, ['id', 'nama', 'jabatan', 'cabang', 'telepon', 'email', 'tanggalMasuk', 'status', 'createdAt', 'tanggalSelesai', 'sisaCuti', 'fingerprintId']);
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
  setupSheet(ss, CONFIG.SHEETS.ASSET_WAREHOUSE, ['id', 'code', 'nama', 'tanggalMasuk', 'divisi', 'status', 'createdBy', 'createdAt', 'history', 'qty', 'zoneId']);
  setupSheet(ss, CONFIG.SHEETS.WAREHOUSE_MAP, ['id', 'configJson', 'updatedAt']);
  setupSheet(ss, CONFIG.SHEETS.BOOKING_MOBIL, ['id', 'tanggal', 'pic', 'jamBerangkat', 'tujuan', 'keterangan', 'rute', 'status', 'createdBy', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.ABSENSI_LEMBUR, ['id', 'tanggal', 'jam', 'nama', 'divisi', 'karyawanId', 'status', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.ABSENSI_KARYAWAN, ['id', 'tanggal', 'jam', 'karyawanId', 'nama', 'divisi', 'jabatan', 'tipe', 'sumber', 'fingerprintId', 'status', 'keterangan', 'createdAt']);
  setupSheet(ss, CONFIG.SHEETS.JADWAL_SHIFT, ['id', 'namaJadwal', 'divisi', 'shiftType', 'jamMasuk', 'jamPulang', 'toleransiMenit', 'aktif', 'createdAt', 'updatedAt']);
  setupSheet(ss, CONFIG.SHEETS.SETTINGS, ['key', 'value', 'updatedAt']);

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
  // PAKSA GUNAKAN ID DARI CONFIG (Abaikan Cache/Properties yang nyasar)
  if (CONFIG.SPREADSHEET_ID && CONFIG.SPREADSHEET_ID !== '') {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }

  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID');
  if (!ssId) { 
    // Cek apakah script ini menempel di sheet tertentu
    const activeSs = SpreadsheetApp.getActiveSpreadsheet();
    if (activeSs) {
      ssId = activeSs.getId();
      props.setProperty('SPREADSHEET_ID', ssId);
    } else {
      const result = setupDatabase(); ssId = result.spreadsheetId; 
    }
  }
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

/**
 * Global Utility: Periksa apakah user memiliki hak akses ke menu tertentu.
 */
function checkPermission(username, permissionKey) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === username) {
        if (data[i][4] === 'admin') return true;
        const perms = JSON.parse(data[i][6] || '[]');
        return perms.includes(permissionKey);
      }
    }
    return false;
  } catch (e) { return false; }
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
        sisaCuti: parseInt(data[i][10]) || 0,
        fingerprintId: String(data[i][11] || '')
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

function addKaryawan(nama, jabatan, cabang, telepon, email, tanggalMasuk, status, tanggalSelesai, sisaCuti, fingerprintId) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KARYAWAN);
    const id = generateId();
    sheet.appendRow([id, nama, jabatan, cabang || '', telepon, email, tanggalMasuk, status || 'Tetap', new Date().toISOString(), tanggalSelesai || '', sisaCuti || 12, normalizeFingerprintId(fingerprintId) || '']);
    return { success: true, id: id };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateKaryawan(id, nama, jabatan, cabang, telepon, email, tanggalMasuk, status, tanggalSelesai, sisaCuti, fingerprintId) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KARYAWAN);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      // Find by ID or Name fallback
      if ((id && String(data[i][0]) === String(id)) || (!id && String(data[i][1]) === String(nama))) {
        sheet.getRange(i + 1, 2, 1, 7).setValues([[nama, jabatan, cabang || '', telepon, email, tanggalMasuk, status]]);
        sheet.getRange(i + 1, 10).setValue(tanggalSelesai || '');
        sheet.getRange(i + 1, 11).setValue(sisaCuti || 0);
        sheet.getRange(i + 1, 12).setValue(normalizeFingerprintId(fingerprintId) || '');
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
        item.sisaCuti || 12,
        normalizeFingerprintId(item.fingerprintId) || ''
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
    
    // Helper untuk memastikan data aman dikirim via JSON (google.script.run)
    const sanitize = (arr) => (arr || []).map(obj => {
      if (!obj || typeof obj !== 'object') return obj;
      const newObj = {};
      for (let key in obj) {
        const val = obj[key];
        if (val instanceof Date) {
          // Jika tahun di bawah 1920, biasanya Google Sheets menganggapnya sebagai "Time" (Jam)
          // Kita format ulang menjadi HH:mm:ss agar tidak muncul 1899-12-30...
          if (val.getFullYear() < 1920) {
            newObj[key] = Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm:ss');
          } else {
            newObj[key] = val.toISOString();
          }
        } else if (val !== null && typeof val === 'object' && !(val instanceof Array)) {
          newObj[key] = String(val);
        } else {
          newObj[key] = val;
        }
      }
      return newObj;
    });

    if (kg.success && kg.data) { 
      kg.data.forEach(k => { history.push({ tanggal: k.tanggal, tipe: k.tipe === 'IN' ? 'Kas Masuk' : 'Kas Keluar', keterangan: k.keterangan, nominal: k.nominal, kategori: 'Kas Gudang' }); }); 
    }
    if (tb.success && tb.data) { 
      tb.data.forEach(t => { history.push({ tanggal: t.tanggal, tipe: 'Team Building', keterangan: t.keterangan, nominal: t.nominal, kategori: 'Team Building' }); }); 
    }
    
    history.sort((a, b) => new Date(b.tanggal||0) - new Date(a.tanggal||0)); 
    history = history.slice(0, 20);
    
    const totalKasIn = (kg.data||[]).filter(k => k.tipe === 'IN').reduce((s, k) => s + (k.nominal || 0), 0); 
    const totalKasOut = (kg.data||[]).filter(k => k.tipe === 'OUT').reduce((s, k) => s + (k.nominal || 0), 0);
    
    const absensiLembur = getAbsensiLembur();
    
    return { 
      success: true, 
      saldoGudang: sG.saldo || 0, 
      saldoTB: sTB.saldo || 0, 
      history: sanitize(history), 
      totalKasIn: totalKasIn, 
      totalKasOut: totalKasOut, 
      kasData: sanitize(kg.data), 
      tbData: sanitize(tb.data), 
      laporanData: sanitize(lk.success ? lk.data : []),
      absensiLembur: sanitize(absensiLembur.success ? absensiLembur.data : [])
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

// ============================================================
// MOVE STOCK - Pindah Stok ke Lokasi Lain
// ============================================================
function moveStock(stockId, jumlah, lokasiTujuan, keterangan, createdBy) {
  try {
    jumlah = parseFloat(jumlah) || 0;
    if (!stockId || jumlah <= 0 || !lokasiTujuan) {
      return { success: false, message: 'Parameter tidak lengkap' };
    }

    const sheet = getSheet(CONFIG.SHEETS.STOCK);
    const data = sheet.getDataRange().getValues();
    let foundRow = -1;
    let stockRow = null;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(stockId)) {
        foundRow = i + 1; // 1-indexed sheet row
        stockRow = data[i];
        break;
      }
    }

    if (foundRow === -1) return { success: false, message: 'Stok tidak ditemukan' };

    const stokSaatIni = parseFloat(stockRow[7]) || 0;
    if (jumlah > stokSaatIni) {
      return { success: false, message: `Jumlah (${jumlah}) melebihi stok tersedia (${stokSaatIni})` };
    }

    const now = new Date().toISOString();

    if (jumlah === stokSaatIni) {
      // Pindah seluruh stok: cukup update lokasi di baris yang ada
      sheet.getRange(foundRow, 11).setValue(lokasiTujuan); // Kolom 11 = Lokasi
      sheet.getRange(foundRow, 13).setValue(now);
    } else {
      // Pindah sebagian: kurangi stok asal, buat baris baru di lokasi tujuan
      // 1. Kurangi stok di lokasi asal
      sheet.getRange(foundRow, 8).setValue(stokSaatIni - jumlah); // Kolom 8 = Stok
      sheet.getRange(foundRow, 13).setValue(now);

      // 2. Cek apakah sudah ada baris dengan SKU yang sama di lokasi tujuan
      const skuAsal = String(stockRow[1]);
      const batchAsal = String(stockRow[4] || '');
      let targetRow = -1;

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][1]) === skuAsal &&
            String(data[i][10]) === lokasiTujuan &&
            String(data[i][4] || '') === batchAsal) {
          targetRow = i + 1;
          break;
        }
      }

      if (targetRow !== -1) {
        // Tambah ke baris yang sudah ada di lokasi tujuan
        const existingStok = parseFloat(data[targetRow - 1][7]) || 0;
        sheet.getRange(targetRow, 8).setValue(existingStok + jumlah);
        sheet.getRange(targetRow, 13).setValue(now);
      } else {
        // Buat entri baru di lokasi tujuan
        const newRow = [
          generateId(),
          skuAsal,                  // SKU
          stockRow[2],              // Nama
          stockRow[3] || '',        // Barcode
          stockRow[4] || '',        // Batch
          stockRow[5] || '',        // Exp Date
          stockRow[6] || '',        // Satuan
          jumlah,                   // Stok
          stockRow[8] || 0,         // Stok Min
          stockRow[9] || '',        // Kategori
          lokasiTujuan,             // Lokasi (baru)
          now,                      // Created At
          now                       // Updated At
        ];
        sheet.appendRow(newRow);
      }
    }

    return { success: true, message: `${jumlah} unit berhasil dipindah ke ${lokasiTujuan}` };
  } catch (e) {
    return { success: false, message: 'Error moveStock: ' + e.message };
  }
}


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

function getSJMasukWithDetails() {
  try {
    const sjRes = getSuratJalanMasuk(); if (!sjRes.success) return sjRes;
    const sheet = getSheet(CONFIG.SHEETS.SURAT_JALAN_MASUK_DETAIL);
    const detData = sheet.getLastRow() > 1 ? sheet.getDataRange().getValues() : [];
    const map = {};
    for (let i = 1; i < detData.length; i++) {
      const row = detData[i];
      if (!row || row.length < 5) continue;
      const sjId = String(row[1]); if (!map[sjId]) map[sjId] = [];
      map[sjId].push({ 
        sku: row[4], 
        nama: row[5], 
        qty: parseFloat(row[6])||0, 
        satuan: row[7], 
        batch: row[8] || '-', 
        expDate: row[9] || '-', 
        lokasi: row[10] || '-' 
      });
    }
    sjRes.data.forEach(d => d.items = map[String(d.id)] || []);
    return sjRes;
  } catch(e) { return { success: false, message: "Err SJM Detail: " + e.message }; }
}

function addSuratJalanMasuk(tanggal, supplier, keterangan, items, createdBy) {
  try {
    const noSJ = generateNoSJ('SJM'); const id = generateId(); const now = new Date().toISOString();
    getSheet(CONFIG.SHEETS.SURAT_JALAN_MASUK).appendRow([id, noSJ, tanggal, supplier, keterangan, createdBy, now]);
    const detSheet = getSheet(CONFIG.SHEETS.SURAT_JALAN_MASUK_DETAIL);
    const parsedItems = typeof items === 'string' ? JSON.parse(items) : items;
    parsedItems.forEach(item => {
      detSheet.appendRow([generateId(), id, noSJ, item.stockId, item.sku, item.nama, parseFloat(item.qty)||0, item.satuan, item.batch||'', item.expDate||'', item.lokasi||'']);
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
    const data = getSheet(sheetName).getDataRange().getDisplayValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      const match = (sjId && String(data[i][1]) === String(sjId)) || (!sjId && noSJFallback && String(data[i][2]) === String(noSJFallback));
      if (match) result.push({ 
        sku:data[i][4], 
        nama:data[i][5], 
        qty:parseFloat(data[i][6])||0, 
        satuan:data[i][7], 
        batch:data[i][8]||'-', 
        expDate:data[i][9]||'-',
        lokasi:data[i][10]||'-'
      });
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

function getSJKeluarWithDetails() {
  try {
    const sjRes = getSuratJalanKeluar(); if (!sjRes.success) return sjRes;
    const sheet = getSheet(CONFIG.SHEETS.SURAT_JALAN_KELUAR_DETAIL);
    const detData = sheet.getLastRow() > 1 ? sheet.getDataRange().getValues() : [];
    const map = {};
    for (let i = 1; i < detData.length; i++) {
      const row = detData[i];
      if (!row || row.length < 5) continue;
      const sjId = String(row[1]); if (!map[sjId]) map[sjId] = [];
      map[sjId].push({ 
        sku: row[4], 
        nama: row[5], 
        qty: parseFloat(row[6])||0, 
        satuan: row[7], 
        batch: row[8] || '-', 
        expDate: row[9] || '-', 
        lokasi: row[10] || '-' 
      });
    }
    sjRes.data.forEach(d => d.items = map[String(d.id)] || []);
    return sjRes;
  } catch(e) { return { success: false, message: "Err SJK Detail: " + e.message }; }
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
      detSheet.appendRow([generateId(), id, noSJ, item.stockId, item.sku, item.nama, parseFloat(item.qty)||0, item.satuan, item.batch||'', item.expDate||'', item.lokasi||'']);
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
    
    // Sederhanakan: Jangan muat detail dulu untuk tes stabilitas
    ordRes.data.forEach(o => o.items = []);
    return ordRes;
  } catch(e) { return { success: false, message: "Err Order Detail Test: " + e.message }; }
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
    const data = getSheet(CONFIG.SHEETS.ORDER_DETAIL).getDataRange().getDisplayValues(); const result = [];
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

/**
 * Mendapatkan detail lengkap (Header & Item) dalam 1 kali hit server
 * untuk mempercepat proses Cetak SJ dan menghindari blokir popup browser.
 */
function getOrderDetailFull(orderId, noOrderFallback) {
  try {
    const ss = getSpreadsheet();
    
    // 1. Ambil Header Order
    const orderSheet = ss.getSheetByName(CONFIG.SHEETS.ORDER);
    const orderData = orderSheet.getDataRange().getValues();
    let orderHeader = null;
    
    for (let i = 1; i < orderData.length; i++) {
      const match = (orderId && String(orderData[i][0]) === String(orderId)) || 
                    (!orderId && noOrderFallback && String(orderData[i][1]) === String(noOrderFallback));
      if (match) {
        orderHeader = {
          id: String(orderData[i][0] || ''),
          noOrder: String(orderData[i][1] || ''),
          tanggal: orderData[i][2] instanceof Date ? Utilities.formatDate(orderData[i][2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(orderData[i][2]||''),
          pelanggan: String(orderData[i][3] || ''),
          alamat: String(orderData[i][4] || ''),
          status: String(orderData[i][5] || ''),
          totalItem: parseFloat(orderData[i][6]) || 0,
          keterangan: String(orderData[i][7] || ''),
          kategori: String(orderData[i][12] || ''),
          noResi: String(orderData[i][13] || '')
        };
        break;
      }
    }
    
    if (!orderHeader) return { success: false, message: 'Header order tidak ditemukan' };
    
    // 2. Ambil Item Detail
    const detailRes = getOrderDetail(orderId, noOrderFallback);
    if (!detailRes.success) return detailRes;
    
    return {
      success: true,
      header: orderHeader,
      items: detailRes.data
    };
  } catch(e) {
    return { success: false, message: 'Server Error: ' + e.message };
  }
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

function getReturWithDetails() {
  try {
    const retRes = getRetur(); if (!retRes.success) return retRes;
    const sheet = getSheet(CONFIG.SHEETS.RETUR_DETAIL);
    const detData = sheet.getLastRow() > 1 ? sheet.getDataRange().getValues() : [];
    const map = {};
    for (let i = 1; i < detData.length; i++) {
      const row = detData[i];
      if (!row || row.length < 5) continue;
      const retId = String(row[1]); if (!map[retId]) map[retId] = [];
      map[retId].push({ 
        sku: row[4], 
        nama: row[5], 
        qty: parseFloat(row[6])||0, 
        satuan: row[7], 
        batch: row[8] || '-', 
        expDate: row[9] || '-', 
        lokasi: row[11] || '-' 
      });
    }
    retRes.data.forEach(d => d.items = map[String(d.id)] || []);
    return retRes;
  } catch(e) { return { success: false, message: "Err Retur Detail: " + e.message }; }
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
    const data = getSheet(CONFIG.SHEETS.RETUR_DETAIL).getDataRange().getDisplayValues(); const result = [];
    for (let i = 1; i < data.length; i++) {
      const match = (returId && String(data[i][1]) === String(returId)) || (!returId && noReturFallback && String(data[i][2]) === String(noReturFallback));
      if (match) result.push({ 
        sku:data[i][4], 
        nama:data[i][5], 
        qty:parseFloat(data[i][6])||0, 
        satuan:data[i][7], 
        batch:data[i][8]||'-', 
        expDate:data[i][9]||'-',
        lokasi:data[i][11]||'-'
      });
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
    const stockData = getStock().data || []; 
    const result = [];
    const now = new Date();
    
    // Hitung batas Tanggal (Senin minggu ini dan tanggal 1 bulan ini)
    const startOfWeek = new Date(now);
    const day = now.getDay();
    const diffToMon = now.getDate() - day + (day === 0 ? -6 : 1);
    startOfWeek.setDate(diffToMon);
    startOfWeek.setHours(0,0,0,0);
    
    const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    startOfMonth.setHours(0,0,0,0);

    const orderData = getSheet(CONFIG.SHEETS.ORDER).getDataRange().getValues();
    const orderDetData = getSheet(CONFIG.SHEETS.ORDER_DETAIL).getDataRange().getValues();
    
    // Mapping detail order ke dalam Map untuk efisiensi (O(N+M))
    const detailMap = {};
    for (let j = 1; j < orderDetData.length; j++) {
      const orderId = String(orderDetData[j][1]);
      if (!detailMap[orderId]) detailMap[orderId] = [];
      detailMap[orderId].push({
        stockId: orderDetData[j][3],
        qty: parseFloat(orderDetData[j][6]) || 0
      });
    }

    const usageWeek = {}; 
    const usageMonth = {};
    
    for (let i = 1; i < orderData.length; i++) {
      if (orderData[i][5] !== 'Terkirim') continue;
      
      const dDate = new Date(orderData[i][10] || orderData[i][2]);
      const isW = dDate >= startOfWeek;
      const isM = dDate >= startOfMonth;
      
      if (isM) {
        const orderId = String(orderData[i][0]);
        const items = detailMap[orderId] || [];
        items.forEach(item => {
          if (isW) usageWeek[item.stockId] = (usageWeek[item.stockId] || 0) + item.qty;
          usageMonth[item.stockId] = (usageMonth[item.stockId] || 0) + item.qty;
        });
      }
    }
    
    stockData.forEach(s => {
      const mw = usageWeek[s.id] || 0;
      const mm = usageMonth[s.id] || 0;
      const daysElapsed = Math.max(1, now.getDate()); 
      const rata = (mm / daysElapsed).toFixed(1);
      const status = s.stok <= 0 ? 'Kritis' : (s.stok <= (s.stokMin || 0) ? 'Rendah' : 'Aman');
      result.push({ 
        sku: s.sku, 
        nama: s.nama, 
        stokSaat: s.stok, 
        minggu: mw, 
        bulan: mm, 
        rataHarian: rata, 
        satuan: s.satuan, 
        statusStok: status 
      });
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
        noPL:data[i][2], 
        noOrder: data[i][3] || '-',
        supplier: data[i][4] || '-',
        keterangan:data[i][5], 
        fileUrl:data[i][6], 
        createdBy:data[i][7], 
        createdAt:data[i][8] 
      });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addPackingList(tanggal, noPL, noOrder, supplier, keterangan, fileUrl, createdBy) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PACKING_LIST);
    if (!sheet) {
      const ss = getSpreadsheet();
      setupSheet(ss, CONFIG.SHEETS.PACKING_LIST, ['id','tanggal','noPL','noOrder','supplier','keterangan','fileUrl','createdBy','createdAt']);
    }
    const finalSheet = getSheet(CONFIG.SHEETS.PACKING_LIST);
    finalSheet.appendRow([generateId(), tanggal, noPL, noOrder || '-', supplier || '-', keterangan, fileUrl, createdBy || 'User', new Date().toISOString()]);
    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) { return { success: false, message: 'Gagal Simpan ke Tabel: ' + e.message }; }
}

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
            qty: data[i][9] || 1,
            zoneId: data[i][10] || ''
        });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function addAssetWarehouse(codePrefix, nama, tanggalMasuk, divisi, status, createdBy, qty, zoneId) {
  try {
    const id = generateId();
    const randomNum = Math.floor(10000 + Math.random() * 90000);
    const code = codePrefix ? `${codePrefix}-${randomNum}` : `AW-${randomNum}`;
    const createdAt = new Date().toISOString();
    const history = `🛒 Dibuat oleh ${createdBy} pada ${createdAt} (Tgl Masuk: ${tanggalMasuk})`;
    
    const sheet = getSheet(CONFIG.SHEETS.ASSET_WAREHOUSE);
    sheet.appendRow([id, code, nama, tanggalMasuk, divisi, status || 'Aktif', createdBy, createdAt, history, qty || 1, zoneId || '']);
    return { success: true, code: code };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateAssetWarehouse(id, nama, tanggalMasuk, status, userNama, qty, zoneId) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ASSET_WAREHOUSE);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i + 1, 3).setValue(nama);
        sheet.getRange(i + 1, 4).setValue(tanggalMasuk);
        sheet.getRange(i + 1, 6).setValue(status);
        sheet.getRange(i + 1, 10).setValue(qty || 1);
        if (zoneId !== undefined) sheet.getRange(i + 1, 11).setValue(zoneId || '');
        
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

function moveAssetWarehouse(assetId, targetDivisi, targetZoneId, userNama) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ASSET_WAREHOUSE);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(assetId)) {
        const oldDiv = data[i][4];
        const oldZone = data[i][10] || '';
        const oldHist = data[i][8] || '';
        const now = new Date().toLocaleString('id-ID');
        
        const divChanged = oldDiv !== targetDivisi;
        const zoneChanged = oldZone !== targetZoneId;
        
        if (!divChanged && !zoneChanged) return { success: true };
        
        let textChange = [];
        if (divChanged) textChange.push(`divisi dari ${oldDiv} ke ${targetDivisi}`);
        if (zoneChanged) {
           const zoneNameStr = targetZoneId ? `zona baru` : `Tanpa Zona`;
           textChange.push(`lokasi ke ${zoneNameStr}`);
        }
        
        const entry = `📦 Dipindah ${textChange.join(' dan ')} oleh ${userNama} pada ${now}`;
        
        sheet.getRange(i + 1, 5).setValue(targetDivisi);
        sheet.getRange(i + 1, 11).setValue(targetZoneId || '');
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

// Map Data Sync
function getWarehouseMapData() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.WAREHOUSE_MAP);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [] };
    // We only take the first row of data (index 1) as our primary config
    let config = [];
    try {
      config = JSON.parse(data[1][1] || '[]');
    } catch(e) { config = []; }
    return { success: true, data: config };
  } catch(e) { return { success: false, message: e.message }; }
}

function saveWarehouseMapData(jsonData) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.WAREHOUSE_MAP);
    const data = sheet.getDataRange().getValues();
    const updatedBy = 'system';
    const now = new Date().toISOString();
    
    if (data.length > 1) {
      // Update existing
      sheet.getRange(2, 2).setValue(jsonData);
      sheet.getRange(2, 3).setValue(now);
    } else {
      // Create new
      sheet.appendRow(['MAIN_CONFIG', jsonData, now]);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
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
function getAbsensiLembur(startDate, endDate) {
  try {
    const ss = getSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.ABSENSI_LEMBUR);
    if (!sheet) return { success: true, data: [] };
    
    const data = sheet.getDataRange().getValues();
    const result = [];
    const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    
    // Konversi parameter tanggal ke objek Date untuk perbandingan
    const startObj = startDate ? new Date(startDate) : null;
    const endObj = endDate ? new Date(endDate) : null;
    if (startObj) startObj.setHours(0, 0, 0, 0);
    if (endObj) endObj.setHours(23, 59, 59, 999);

    for (let i = 1; i < data.length; i++) {
      if (!data[i][1]) continue; // Skip jika tanggal kosong
      
      let rowTglObj;
      let rowTglStr = '';
      
      if (data[i][1] instanceof Date) {
        rowTglObj = data[i][1];
        rowTglStr = Utilities.formatDate(rowTglObj, tz, 'yyyy-MM-dd');
      } else {
        rowTglStr = String(data[i][1]).trim().split('T')[0];
        rowTglObj = new Date(rowTglStr);
      }
      
      // Logika Filter
      let isMatch = false;
      if (startObj && endObj) {
        const compareDate = new Date(rowTglStr);
        compareDate.setHours(0, 0, 0, 0);
        isMatch = (compareDate >= startObj && compareDate <= endObj);
      } else {
        isMatch = (rowTglStr === todayStr);
      }
      
      if (isMatch) {
        result.push({
          id: data[i][0],
          tanggal: rowTglStr,
          jam: data[i][2],
          nama: data[i][3],
          divisi: data[i][4],
          karyawanId: data[i][5],
          status: data[i][6],
          createdAt: data[i][7]
        });
      }
    }
    // Urutkan berdasarkan tanggal & jam terbaru di atas
    result.sort((a, b) => {
      const dateTimeA = String(a.tanggal) + ' ' + String(a.jam);
      const dateTimeB = String(b.tanggal) + ' ' + String(b.jam);
      return dateTimeB.localeCompare(dateTimeA);
    });
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.messsage }; }
}

// ============================================================
// MODUL ABSENSI KARYAWAN 
// ============================================================

function getAbsensiKaryawan(tanggal, divisi, requesterUsername) {
  if (requesterUsername && !checkPermission(requesterUsername, 'absensiKaryawan')) {
    return { success: false, message: 'Akses Ditolak: Anda tidak memiliki izin untuk melihat Absensi Karyawan.' };
  }
  try {
    const ss = getSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.ABSENSI_KARYAWAN);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: [] };

    const todayStr = tanggal || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const data = sheet.getDataRange().getValues();
    const result = [];

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      let rawTgl = data[i][1];
      let rowTgl = "";
      
      if (rawTgl instanceof Date) {
        rowTgl = Utilities.formatDate(rawTgl, tz, 'yyyy-MM-dd');
      } else {
        // Jika string, coba bersihkan dan normalisasi
        let s = String(rawTgl).trim().split('T')[0];
        // Jika format DD/MM/YYYY
        if (s.includes('/')) {
          let parts = s.split('/');
          if (parts[0].length === 4) rowTgl = parts[0] + '-' + parts[1].padStart(2, '0') + '-' + parts[2].padStart(2, '0');
          else rowTgl = parts[2] + '-' + parts[1].padStart(2, '0') + '-' + parts[0].padStart(2, '0');
        } else {
          rowTgl = s;
        }
      }

      if (rowTgl !== todayStr) continue;
      if (divisi && data[i][5] !== divisi) continue;

      let rawJam = data[i][2];
      let jamStr = "";
      if (rawJam instanceof Date) {
        jamStr = Utilities.formatDate(rawJam, tz, 'HH:mm:ss');
      } else {
        jamStr = String(rawJam || '').trim();
      }

      result.push({
        id: data[i][0],
        tanggal: rowTgl,
        jam: jamStr,
        karyawanId: data[i][3],
        nama: data[i][4],
        divisi: data[i][5],
        jabatan: data[i][6],
        tipe: data[i][7],
        sumber: data[i][8],
        fingerprintId: data[i][9],
        status: data[i][10],
        keterangan: data[i][11],
        createdAt: String(data[i][12] || '')
      });
    }

    // Sortir jam descending (terbaru di atas) - Pastikan dibandingkan sebagai string
    result.sort((a, b) => String(b.jam || '').localeCompare(String(a.jam || '')));
    return { success: true, data: result };
  } catch (e) { 
    return { success: false, message: "Error getAbsensiKaryawan: " + e.message }; 
  }
}

function addAbsensiKaryawan(karyawanId, nama, divisi, jabatan, tipe, jam, tanggal, keterangan, requesterUsername) {
  if (requesterUsername && !checkPermission(requesterUsername, 'absensiKaryawan')) {
    return { success: false, message: 'Akses Ditolak: Anda tidak memiliki izin untuk mengelola Absensi Karyawan.' };
  }
  try {
    const ss = getSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.ABSENSI_KARYAWAN);
    const now = new Date();
    const tglStr = tanggal || Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    const jamStr = jam || Utilities.formatDate(now, tz, 'HH:mm:ss');
    const statusInfo = _hitungStatusAbsensi(divisi, tipe, jamStr, nama, tglStr);
    const id = generateId();
    sheet.appendRow([
      id, tglStr, jamStr, karyawanId, nama, divisi, jabatan || '',
      tipe, 'manual', '', statusInfo.status, keterangan || '', now.toISOString()
    ]);
    return { success: true, id: id, status: statusInfo.status };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteAbsensiKaryawan(id, requesterUsername) {
  if (requesterUsername && !checkPermission(requesterUsername, 'absensiKaryawan')) {
    return { success: false, message: 'Akses Ditolak: Anda tidak memiliki izin untuk menghapus data Absensi.' };
  }
  return deleteRow(CONFIG.SHEETS.ABSENSI_KARYAWAN, id);
}

function normalizeFingerprintId(fpId) {
  var raw = String(fpId || '').trim();
  if (!raw) return '';
  // Strip common prefixes
  raw = raw.replace(/^fp[\-_]*/i, '');
  // Clean special chars
  raw = raw.replace(/[^0-9A-Z]/gi, '');
  // Strip leading zeros if numeric part only
  if (/^\d+$/.test(raw)) {
    raw = String(parseInt(raw, 10));
  }
  return raw.toUpperCase();
}

function syncFingerprintData(records) {
  try {
    if (!records || !Array.isArray(records) || records.length === 0) {
      return { success: false, message: 'Tidak ada data fingerprint yang diterima.' };
    }

    const ss = getSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.ABSENSI_KARYAWAN);
    const now = new Date();
    const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

    // Cek duplikat berdasarkan fingerprintId + jam + tanggal
    const existing = sheet.getLastRow() > 1 ? sheet.getDataRange().getValues() : [];
    const existingKeys = new Set();
    for (let i = 1; i < existing.length; i++) {
      const tgl = existing[i][1] instanceof Date
        ? Utilities.formatDate(existing[i][1], tz, 'yyyy-MM-dd')
        : String(existing[i][1] || '').split('T')[0];
      const fpKey = normalizeFingerprintId(existing[i][9]);
      const tipeKey = String(existing[i][7] || '').toUpperCase();
      existingKeys.add(tgl + '_' + fpKey + '_' + String(existing[i][2]) + '_' + tipeKey);
    }

    // Lookup data karyawan
    const karSheet = ss.getSheetByName(CONFIG.SHEETS.KARYAWAN);
    const karyawanMap = {};
    if (karSheet) {
      const kData = karSheet.getDataRange().getValues();
      for (let i = 1; i < kData.length; i++) {
        if (kData[i][0]) {
          const karId = String(kData[i][0]);
          const rawFpId = String(kData[i][11] || '').trim();
          const fpId = normalizeFingerprintId(rawFpId);
          const info = { id: karId, nama: kData[i][1], jabatan: kData[i][2], divisi: kData[i][3], fingerprintId: rawFpId };
          karyawanMap[karId] = info;
          if (fpId) karyawanMap[fpId] = info;
        }
      }
    }

    const rows = [];
    let added = 0, skipped = 0;

    records.forEach(function(rec) {
      const rawFpId = String(rec.fingerprintId || rec.finger_id || '').trim();
      const fpId    = normalizeFingerprintId(rawFpId);
      const karId   = String(rec.karyawanId || rec.user_id || '').trim();
      const tglStr  = String(rec.tanggal || todayStr).split('T')[0];
      const jamStr  = String(rec.jam || rec.time || '');
      const tipe    = String(rec.tipe || 'IN').toUpperCase();
      const dupKey  = tglStr + '_' + (fpId || karId) + '_' + jamStr + '_' + tipe;

      if (existingKeys.has(dupKey)) { skipped++; return; }
      existingKeys.add(dupKey);

      // Lookup by karyawanId first, then by normalized fingerprint ID
      let karData = (karId && karyawanMap[karId]) ? karyawanMap[karId] : (fpId && karyawanMap[fpId]) ? karyawanMap[fpId] : {};
      const finalKarId = karData.id || karId;
      const nama    = rec.nama    || karData.nama    || ('FP-' + (rawFpId || karId));
      const divisi  = rec.divisi  || karData.divisi  || 'Tidak Diketahui';
      const jabatan = rec.jabatan || karData.jabatan || '';
      const statusInfo = _hitungStatusAbsensi(divisi, tipe, jamStr, nama, tglStr);

      rows.push([
        generateId(), tglStr, jamStr, finalKarId, nama, divisi, jabatan,
        tipe, 'fingerprint', rawFpId || karData.fingerprintId || '', statusInfo.status, '', now.toISOString()
      ]);
      added++;
    });

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return { success: true, added: added, skipped: skipped, total: records.length };
  } catch (e) { return { success: false, message: e.message }; }
}

function _hitungStatusAbsensi(divisi, tipe, jam, nama, tanggal) {
  try {
    // 1. Prioritas: Cek di Roster Bulanan
    if (nama && tanggal) {
       const settingsRes = getRosterSettings();
       const settings = settingsRes.success ? settingsRes.data : { pagiIn:"08:00", pagiOut:"17:00", malamIn:"20:00", malamOut:"05:00" };
       const tglDate = new Date(tanggal);
       const monthYear = Utilities.formatDate(tglDate, Session.getScriptTimeZone(), "yyyy-MM");
       const dayNum = String(tglDate.getDate());
       
       const rosterSheet = getSheet(CONFIG.SHEETS.JADWAL_ROSTER);
       const rData = rosterSheet.getDataRange().getValues();
       const rHeaders = rData[0];
       
       for (let i = 1; i < rData.length; i++) {
         let rBulan = rData[i][0];
         if (rBulan instanceof Date) rBulan = Utilities.formatDate(rBulan, Session.getScriptTimeZone(), "yyyy-MM");
         
         if (String(rBulan).trim() === monthYear && String(rData[i][1]).trim() === nama) {
            const dCol = rHeaders.map(String).indexOf(dayNum);
            if (dCol !== -1) {
               const shiftVal = String(rData[i][dCol]).toUpperCase();
               let shiftIn = "", shiftOut = "";
               
               if (shiftVal === 'PAGI' || shiftVal.includes('PAGI')) {
                 shiftIn = settings.pagiIn; shiftOut = settings.pagiOut;
               } else if (shiftVal === 'MALAM' || shiftVal.includes('MALAM')) {
                 shiftIn = settings.malamIn; shiftOut = settings.malamOut;
               } else if (shiftVal === 'OFF') {
                 return { status: tipe === 'IN' ? 'Masuk (OFF)' : 'Pulang (OFF)' };
               }
               
               if (shiftIn && shiftOut) {
                 return _compareAttendanceTime(tipe, jam, shiftIn, shiftOut, settings.toleransi || 0); 
               }
            }
            break;
         }
       }
    }

    // 2. Fallback: Jadwal Shift Divisi
    const jadwalRes = getJadwalShift(divisi);
    if (!jadwalRes.success || jadwalRes.data.length === 0) {
      return { status: tipe === 'IN' ? 'Hadir' : 'Pulang' };
    }
    const aktifList = jadwalRes.data.filter(function(j) {
      return String(j.aktif).toLowerCase() === 'ya' || j.aktif === true || String(j.aktif).toLowerCase() === 'true';
    });
    const jadwal = aktifList[0];
    if (!jadwal) return { status: tipe === 'IN' ? 'Hadir' : 'Pulang' };

    return _compareAttendanceTime(tipe, jam, jadwal.jamMasuk, jadwal.jamPulang, parseInt(jadwal.toleransiMenit) || 0);

  } catch (e) { return { status: tipe === 'IN' ? 'Hadir' : 'Pulang' }; }
}

/**
 * Helper untuk membandingkan jam absen dengan jadwal
 */
function _compareAttendanceTime(tipe, jamAbsen, jamMasuk, jamPulang, toleransi) {
  const jamParts   = String(jamAbsen).split(':');
  const jamMnt     = (parseInt(jamParts[0] || 0) * 60) + parseInt(jamParts[1] || 0);
  const masukParts = String(jamMasuk).split(':');
  const masukMnt   = (parseInt(masukParts[0] || 0) * 60) + parseInt(masukParts[1] || 0);
  const pulangParts = String(jamPulang).split(':');
  const pulangMnt  = (parseInt(pulangParts[0] || 0) * 60) + parseInt(pulangParts[1] || 0);

  if (tipe === 'IN') {
    return { status: jamMnt <= masukMnt + toleransi ? 'Hadir' : 'Terlambat' };
  } else {
    // Normalisasi untuk shift malam (jika pulang jam 05:00 pagi besoknya)
    let effectivePulangMnt = pulangMnt;
    // Jika jam masuk > jam pulang (misal 20:00 -> 05:00), maka pulang dianggap hari berikutnya
    // Namun perbandingan menit biasanya cukup jika kita asumsikan absen dilakukan di rentang yang wajar
    return { status: jamMnt >= effectivePulangMnt - toleransi ? 'Pulang' : 'Pulang Awal' };
  }
}

// ============================================================
// MODUL JADWAL SHIFT
// ============================================================

function getJadwalShift(divisi, requesterUsername) {
  if (requesterUsername && !checkPermission(requesterUsername, 'jadwalShift')) {
    return { success: false, message: 'Akses Ditolak: Anda tidak memiliki izin untuk melihat Jadwal Shift.' };
  }
  try {
    const sheet = getSheet(CONFIG.SHEETS.JADWAL_SHIFT);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: [] };
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (divisi && data[i][2] !== divisi) continue;
      result.push({
        id: data[i][0],
        namaJadwal: data[i][1],
        divisi: data[i][2],
        shiftType: data[i][3],
        jamMasuk: data[i][4],
        jamPulang: data[i][5],
        toleransiMenit: data[i][6],
        aktif: data[i][7],
        createdAt: data[i][8],
        updatedAt: data[i][9]
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function saveJadwalShift(id, namaJadwal, divisi, shiftType, jamMasuk, jamPulang, toleransiMenit, aktif) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.JADWAL_SHIFT);
    const now = new Date().toISOString();
    if (id) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(id)) {
          sheet.getRange(i + 1, 2, 1, 9).setValues([[namaJadwal, divisi, shiftType, jamMasuk, jamPulang, toleransiMenit, aktif, data[i][8], now]]);
          return { success: true };
        }
      }
      return { success: false, message: 'Jadwal tidak ditemukan.' };
    } else {
      const newId = generateId();
      sheet.appendRow([newId, namaJadwal, divisi, shiftType, jamMasuk, jamPulang, toleransiMenit, aktif, now, now]);
      return { success: true, id: newId };
    }
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteJadwalShift(id, requesterUsername) {
  if (requesterUsername && !checkPermission(requesterUsername, 'jadwalShift')) {
    return { success: false, message: 'Akses Ditolak: Anda tidak memiliki izin untuk menghapus Jadwal Shift.' };
  }
  return deleteRow(CONFIG.SHEETS.JADWAL_SHIFT, id);
}

// ============================================================
// LAPORAN ABSENSI
// ============================================================

function getLaporanAbsensi(tanggal, divisi, requesterUsername) {
  if (requesterUsername && !checkPermission(requesterUsername, 'absensiKaryawan')) {
    return { success: false, message: 'Akses Ditolak: Izin ditolak untuk Laporan Absensi.' };
  }
  try {
    const ss = getSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const tglStr = tanggal || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    // 1. Ambil data Karyawan (Daftar Master yang seharusnya hadir)
    const karSheet = ss.getSheetByName(CONFIG.SHEETS.KARYAWAN);
    const karData  = karSheet ? karSheet.getDataRange().getValues() : [];
    const allKaryawanMap = {};
    const filteredKaryawan = [];
    
    for (let i = 1; i < karData.length; i++) {
      if (!karData[i][0]) continue;
      const karId = String(karData[i][0]).split('.')[0].trim(); // Normalisasi ID
      const statusK = String(karData[i][7] || '');
      const kDiv = String(karData[i][3] || 'Lainnya');
      
      const kInfo = {
        id: karId,
        nama: karData[i][1],
        jabatan: karData[i][2],
        divisi: kDiv
      };
      
      allKaryawanMap[karId] = kInfo;
      
      // Filter karyawan yang ditampilkan berdasarkan pilihan divisi (kecuali Resign)
      if (statusK === 'Resign' || statusK === 'Non-Aktif') continue;
      if (!divisi || kDiv === divisi) {
        filteredKaryawan.push(kInfo);
      }
    }

    // 2. Ambil SEMUA data absensi hari ini + besok (Khusus untuk Shift Malam lintas hari)
    const absenRes  = getAbsensiKaryawan(tglStr, '', requesterUsername);
    
    // Hitung tanggal besok
    const currentTgl = new Date(tglStr + 'T12:00:00'); // Use noon to avoid TZ issues
    currentTgl.setDate(currentTgl.getDate() + 1);
    const tomorrowStr = Utilities.formatDate(currentTgl, tz, 'yyyy-MM-dd');
    const tomorrowRes = getAbsensiKaryawan(tomorrowStr, '', requesterUsername);

    const absenData = [
      ...(absenRes.success ? absenRes.data : []),
      ...(tomorrowRes.success ? tomorrowRes.data : [])
    ];
    
    const logsByKarId = {};
    const unmappedLogs = [];
    const usedLogIds = new Set();

    absenData.forEach(function(a) {
      if (!a.karyawanId && !a.fingerprintId) {
        unmappedLogs.push(a);
      } else {
        const idToMap = String(a.karyawanId || a.fingerprintId).split('.')[0].trim();
        if (allKaryawanMap[idToMap]) {
          if (!logsByKarId[idToMap]) logsByKarId[idToMap] = [];
          logsByKarId[idToMap].push(a);
        } else {
          unmappedLogs.push(a);
        }
      }
    });

    // 3. Ambil data Roster untuk bulan ini (untuk filter OFF)
    const rosterSheet = ss.getSheetByName(CONFIG.SHEETS.JADWAL_ROSTER);
    const roData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];
    const roHeaders = roData.length > 0 ? roData[0] : [];
    const monthYear = tglStr.substring(0, 7);
    const dayNumStr = String(parseInt(tglStr.substring(8, 10)));
    const dayCol = roHeaders.map(String).indexOf(dayNumStr);
    
    const rosterMap = {};
    if (dayCol !== -1) {
      for (let i = 1; i < roData.length; i++) {
        let rBulanCell = roData[i][0];
        let rBulanStr = "";
        
        if (rBulanCell instanceof Date) {
          rBulanStr = Utilities.formatDate(rBulanCell, tz, "yyyy-MM");
        } else {
          // Handle string or numeric date formats robustly
          rBulanStr = String(rBulanCell).trim().substring(0, 7);
        }

        if (rBulanStr === monthYear) {
          const rNama = String(roData[i][1]).trim().toLowerCase();
          rosterMap[rNama] = String(roData[i][dayCol]).trim().toUpperCase();
        }
      }
    }

    const sudahAbsen  = [];
    const belumAbsen  = [];
    const rekapDivisi = {};

    // 4. Proses Karyawan Terfilter (Master)
    filteredKaryawan.forEach(function(k) {
      // CEK ROSTER: Jika OFF atau LIBUR, jangan masukkan ke laporan (kecuali jika ada log aktivitas)
      const shiftHariIni = rosterMap[k.nama.trim().toLowerCase()] || "";
      const logs = logsByKarId[k.id] || [];
      
      const isOff = (shiftHariIni === "OFF" || shiftHariIni === "LIBUR" || shiftHariIni === "L");
      
      if (isOff && logs.length === 0) {
        return; // SKIP: Karyawan sedang OFF dan tidak ada aktivitas absen
      }

      const div = k.divisi;
      if (!rekapDivisi[div]) rekapDivisi[div] = { divisi: div, total: 0, hadir: 0, terlambat: 0, alfa: 0 };
      rekapDivisi[div].total++;

      if (logs.length > 0) {
        // Urutkan logs secara ASCENDING (terawal ke terakhir)
        const sortedLogs = logs.slice().sort((a,b) => String(a.jam || '').localeCompare(String(b.jam || '')));
        // Identifikasi log Masuk vs Pulang secara ketat
        let masuks = sortedLogs.filter(l => {
          let t = String(l.tipe || '').toUpperCase();
          let s = String(l.status || '').toLowerCase();
          // Masuk harus di hari yang sama dengan tanggal laporan
          return (l.tanggal === tglStr) && (t === 'IN' || s.includes('hadir') || s.includes('terlambat'));
        });
        
        let pulangs = sortedLogs.filter(l => {
          let t = String(l.tipe || '').toUpperCase();
          let s = String(l.status || '').toLowerCase();
          
          if (shiftHariIni === 'MALAM') {
            // Untuk shift malam, pulang bisa hari ini (malam) atau besok pagi (sebelum jam 11:00)
            const isToday = (l.tanggal === tglStr);
            const isTomorrowMorning = (l.tanggal === tomorrowStr && (l.jam || '00:00') < '11:00');
            return (isToday || isTomorrowMorning) && (t === 'OUT' || s.includes('pulang'));
          } else {
            // Untuk shift pagi/biasa, pulang harus di hari yang sama
            return (l.tanggal === tglStr) && (t === 'OUT' || s.includes('pulang'));
          }
        });

        // Jika tidak ada tipe yang jelas, fallback ke urutan waktu (hanya jika benar-benar tidak ada label)
        if (masuks.length === 0 && pulangs.length === 0) {
          const todayLogs = sortedLogs.filter(l => l.tanggal === tglStr);
          if (todayLogs.length > 0) {
            masuks = [todayLogs[0]];
            if (todayLogs.length > 1) pulangs = [todayLogs[todayLogs.length - 1]];
          }
        }
        
        // Jam Masuk = Masuk yang paling pertama (absen pertama mereka)
        let inLog = masuks.length > 0 ? masuks[0] : null;
        
        // Jam Pulang = Pulang yang paling terakhir
        let outLog = pulangs.length > 0 ? pulangs[pulangs.length - 1] : null;

        // Tentukan Status Gabungan
        let st = "Hadir";
        if (inLog && outLog) {
           st = (inLog.status === outLog.status) ? inLog.status : inLog.status + ' / ' + outLog.status;
        } else if (inLog) {
           st = inLog.status;
        } else if (outLog) {
           st = outLog.status;
        }

        sudahAbsen.push(Object.assign({}, k, { 
          tanggal: tglStr,
          shift: shiftHariIni || '-',
          inLog: inLog,
          outLog: outLog,
          logs: logs, 
          statusAbsen: st 
        }));
        
        if (st.includes('Terlambat')) rekapDivisi[div].terlambat++;
        else rekapDivisi[div].hadir++;
      } else {
        // Jika tidak ada log sama sekali
        belumAbsen.push(Object.assign({}, k, { 
          statusAbsen: 'Alfa', 
          shift: shiftHariIni || '-',
          logs: [] 
        }));
        rekapDivisi[div].alfa++;
      }
    });

    // 4. Tambahkan Log Unmapped (ID tidak dikenal atau ID tidak terdaftar)
    // agar Admin tahu ada orang yang absen tapi datanya tidak singkon
    unmappedLogs.forEach(function(a) {
      if (divisi && a.divisi && a.divisi !== divisi) return;
      
      sudahAbsen.push({
        id: a.karyawanId || 'UNMAPPED',
        tanggal: tglStr,
        nama: a.nama + ' (ID Tidak Terdaftar)',
        jabatan: a.jabatan || '-',
        divisi: a.divisi || 'Tidak Diketahui',
        statusAbsen: a.status || 'Hadir',
        logs: [a],
        inLog: a,
        outLog: null,
        isUnmapped: true
      });
    });

    return {
      success: true,
      tanggal: tglStr,
      sudahAbsen: sudahAbsen,
      belumAbsen: belumAbsen,
      rekapDivisi: Object.values(rekapDivisi),
      totalKaryawan: filteredKaryawan.length,
      totalHadir: sudahAbsen.filter(x => !x.isUnmapped).length,
      totalAlfa: belumAbsen.length,
      unmappedCount: unmappedLogs.length
    };
  } catch (e) { return { success: false, message: e.message }; }
}

function getDivisiList() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KARYAWAN);
    const data  = sheet.getDataRange().getValues();
    const divSet = new Set();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][3]) divSet.add(String(data[i][3]).trim());
    }
    return { success: true, data: Array.from(divSet).sort() };
  } catch (e) { return { success: false, message: e.message }; }
}

/**
 * Fungsi Pemeliharaan: Sinkronkan masal data nama & divisi di log absensi lama
 * Berguna jika ada perbaikan data di Master Karyawan atau upload fingerprint ID baru.
 */
function repairAbsensiData(requesterUsername) {
  if (requesterUsername && !checkPermission(requesterUsername, 'aksesRepairAbsensi')) {
    return { success: false, message: 'Anda tidak memiliki hak akses untuk menjalankan perbaikan data absensi masal.' };
  }
  try {
    const ss = getSpreadsheet();
    const karSheet = ss.getSheetByName(CONFIG.SHEETS.KARYAWAN);
    const absSheet = ss.getSheetByName(CONFIG.SHEETS.ABSENSI_KARYAWAN);
    if (!karSheet || !absSheet) return { success: false, message: 'Sheet tidak ditemukan.' };

    // 1. Build Karyawan Lookup Map
    const kData = karSheet.getDataRange().getValues();
    const kMap = {};
    for (let i = 1; i < kData.length; i++) {
      if (!kData[i][0]) continue;
      const uuid = String(kData[i][0]);
      const fpId = normalizeFingerprintId(kData[i][11]);
      const info = [kData[i][1], kData[i][3] || 'Lainnya', kData[i][2] || '']; // [nama, divisi, jabatan]
      kMap[uuid] = info;
      if (fpId) kMap[fpId] = info;
    }

    // 2. Scan & Update Absensi
    const aRange = absSheet.getDataRange();
    const aData = aRange.getValues();
    let updatedCount = 0;

    for (let j = 1; j < aData.length; j++) {
      const rowNum = j + 1;
      const karId = String(aData[j][3]);
      const fpIdLog = normalizeFingerprintId(aData[j][9]);
      const currentNama = String(aData[j][4] || '');
      
      // Cari kecocokan di master
      const match = kMap[karId] || kMap[fpIdLog];
      
      if (match) {
        // Cek apakah perlu diupdate (Nama masih placeholder FP- atau Nama berbeda)
        const isPlaceholder = currentNama.startsWith('FP-') || currentNama === '';
        const isNameDiff = currentNama !== match[0];
        const isDivDiff = String(aData[j][5]) !== match[1];

        if (isPlaceholder || isNameDiff || isDivDiff) {
          // Update Nama (E), Divisi (F), Jabatan (G) -> Kolom 5, 6, 7
          absSheet.getRange(rowNum, 5, 1, 3).setValues([[match[0], match[1], match[2]]]);
          updatedCount++;
        }
      }
    }

    return { success: true, message: 'Berhasil menyinkronkan ' + updatedCount + ' baris data absensi.' };
  } catch (e) { return { success: false, message: e.message }; }
}


// ============================================================
// SINGLE POST HANDLER PINTAR (MENANGANI JSON & RAW TEXT X900)
// ============================================================
function doPost(e) {
  // 1. Pastikan ada data yang masuk
  if (!e || !e.postData || !e.postData.contents) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'No data' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var rawContent = e.postData.contents;

  // 2. Coba baca sebagai JSON (Untuk Web ERP Anda)
  try {
    var body = JSON.parse(rawContent);
    var func = body.func;
    var args = body.args || [];
    var result;

    const context = typeof globalThis !== 'undefined' ? globalThis : this;

    if (func === 'syncFingerprintData') {
      result = syncFingerprintData(args[0] || []);
    } else if (func === 'addAbsensiKaryawan') {
      result = addAbsensiKaryawan.apply(null, args);
    } else if (func && typeof context[func] === 'function') {
      result = context[func].apply(null, args);
    } else {
      result = { success: false, message: 'Fungsi tidak dikenal: ' + func };
    }

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    // ===========================================================
    // JIKA BUKAN JSON (ERROR PARSE), INI PASTI DARI MESIN X900
    // ===========================================================
    try {
      // Paksa sistem untuk memanggil ID yang ada di CONFIG
      var ss = getSpreadsheet();
      
      // Buat atau cari sheet untuk log data mentah dari mesin
      var logSheetName = "LogMesinX900";
      var logSheet = ss.getSheetByName(logSheetName);
      
      if (!logSheet) {
        logSheet = ss.insertSheet(logSheetName);
        logSheet.appendRow(["Waktu Terima", "Data Mentah (Raw Text) dari Mesin"]);
        logSheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#1a3a5c").setFontColor("#ffffff");
      }

      // Simpan apa adanya ke sheet agar kita tahu strukturnya
      logSheet.appendRow([new Date(), rawContent]);

      // PARSING OTOMATIS DATA X900
      // Contoh format: "123 2024-11-20 08:00:00 0" (ID Tanggal Jam Status)
      // Status 0=Check-In, 1=Check-Out (Umumnya)
      var rows_raw = rawContent.split('\n');
      var records_to_sync = [];
      
      rows_raw.forEach(function(line) {
        var cleanLine = line.trim();
        if (!cleanLine) return;
        
        // Split by space, tab, or comma
        var parts = cleanLine.split(/[\s,]+/);
        if (parts.length >= 3) {
          var fpId = parts[0];
          var tgl = parts[1];
          var jam = parts[2];
          var statusX = parts[3] || "0"; // Default 0 (IN)
          
          records_to_sync.push({
            fingerprintId: fpId,
            tanggal: tgl,
            jam: jam,
            tipe: (statusX == "0" || statusX.toLowerCase() == "in") ? "IN" : "OUT"
          });
        }
      });

      if (records_to_sync.length > 0) {
        syncFingerprintData(records_to_sync);
      }

      // WAJIB KEMBALIKAN TEKS "OK" AGAR MESIN ABSENSI MERASA SUKSES
      return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);

    } catch (logErr) {
      // Jika terjadi error saat menyimpan ke sheet, tetap balas OK agar mesin tidak hang
      return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
    }
  }
}

// ============================================================
// FUNGSI TESTER (JALANKAN MANUAL DARI EDITOR)
// ============================================================
function TestLogMesin() {
  var dummyEvent = {
    postData: {
      contents: "Ini adalah tes simulasi data tembakan dari Mesin X900 \n 123 2024-11-20 08:00:00 1"
    }
  };
  // Panggil doPost seolah-olah Netlify yang mengirim
  doPost(dummyEvent);
  Logger.log("Tes Selesai. Silakan cek Google Sheet Anda, apakah sheet LogMesinX900 sudah muncul?");
}

function getSpreadsheetUrl() {
  try {
    var ss = getSpreadsheet();
    return { success: true, url: ScriptApp.getService().getUrl(), spreadsheetUrl: ss.getUrl() };
  } catch (e) {
    return { success: false, url: '', message: e.message };
  }
}

// ============================================================
// SHIFT ROSTER MANAGEMENT (MONTHLY MATRIX)
// ============================================================

/**
 * Mendapatkan data roster untuk bulan tertentu
 * @param {string} monthYear Format "YYYY-MM"
 */
function getShiftRoster(monthYear, requesterUsername) {
  if (requesterUsername && !checkPermission(requesterUsername, 'jadwalShift')) {
    return { success: false, message: 'Akses Ditolak: Anda tidak memiliki izin untuk melihat Jadwal Roster.' };
  }

  if (!monthYear) monthYear = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM");
  
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEETS.JADWAL_ROSTER);
    
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.JADWAL_ROSTER);
      // Header: Bulan, Nama, 1..31
      const headers = ["Bulan", "Nama"];
      for (let i = 1; i <= 31; i++) headers.push(i.toString());
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.setFrozenColumns(2);
      return { success: true, data: [], monthYear: monthYear, headers: headers };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    // 1. Ambil data attendance untuk bulan ini
    const attMap = {};
    try {
      const attSheet = ss.getSheetByName(CONFIG.SHEETS.ABSENSI_KARYAWAN);
      if (attSheet) {
        const attData = attSheet.getDataRange().getValues();
        const tz = ss.getSpreadsheetTimeZone();
        
        for (let i = 1; i < attData.length; i++) {
          let rawTgl = attData[i][1];
          let tglStr = "";
          if (rawTgl instanceof Date) tglStr = Utilities.formatDate(rawTgl, tz, "yyyy-MM");
          else tglStr = String(rawTgl).trim().split("-").slice(0,2).join("-");
          
          if (tglStr === monthYear) {
            const nama = String(attData[i][4]).trim();
            const rawD = attData[i][1];
            const d = (rawD instanceof Date) ? rawD.getDate() : parseInt(String(rawD).split("-")[2]);
            const status = String(attData[i][10] || "Hadir");
            
            if (!attMap[nama]) attMap[nama] = {};
            // Utamakan status 'Terlambat' atau yang lebih detail jika ada double log
            if (!attMap[nama][d] || status === "Terlambat") {
              attMap[nama][d] = status;
            }
          }
        }
      }
    } catch (e) { console.error("Error attMap:", e); }

    const filtered = rows.filter(row => {
      let cellValue = row[0];
      if (cellValue instanceof Date) {
        cellValue = Utilities.formatDate(cellValue, "GMT+7", "yyyy-MM");
      } else {
        cellValue = String(cellValue).trim();
      }
      return cellValue === monthYear;
    }).map(row => {
      const obj = { nama: row[1] };
      for (let i = 2; i < headers.length; i++) {
        obj[headers[i]] = row[i];
      }
      return obj;
    });

    return { success: true, data: filtered, attendanceMap: attMap, monthYear: monthYear, headers: headers };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Menyimpan data roster dari impor Excel
 * @param {string} monthYear Format "YYYY-MM"
 * @param {Array} rosterData [{nama: "...", "1": "Pagi", ...}]
 */
function importShiftRoster(monthYear, rosterData, requesterUsername) {
  if (requesterUsername && !checkPermission(requesterUsername, 'jadwalShift')) return { success: false, message: "Akses Ditolak" };
  
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEETS.JADWAL_ROSTER);
    
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.JADWAL_ROSTER);
      const headers = ["Bulan", "Nama"];
      for (let i = 1; i <= 31; i++) headers.push(i.toString());
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.setFrozenColumns(2);
    }

    // Ambil data lama
    const range = sheet.getDataRange();
    const values = range.getValues();
    const headers = values[0];
    
    // Hapus data lama untuk bulan yang sama
    const newValues = [headers];
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] !== monthYear) {
        newValues.push(values[i]);
      }
    }

    // Tambah data baru
    rosterData.forEach(item => {
      // Gunakan prefix ' agar tidak otomatis diubah jadi tanggal oleh Sheets
      const row = [monthYear, item.name || item.nama]; 
      for (let d = 1; d <= 31; d++) {
        row.push(item[d.toString()] || "");
      }
      newValues.push(row);
    });

    // Tulis ulang
    sheet.clear();
    sheet.getRange(1, 1, newValues.length, headers.length).setValues(newValues);

    return { success: true, message: "Berhasil mengimpor " + rosterData.length + " data roster." };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Memperbarui satu sel dalam Roster Shift
 */
function updateRosterCell(monthYear, name, day, newValue, requesterUsername) {
  if (requesterUsername && !checkPermission(requesterUsername, 'jadwalShift')) return { success: false, message: "Akses Ditolak" };
  
  try {
    const sheet = getSheet(CONFIG.SHEETS.JADWAL_ROSTER);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Cari kolom hari (Mapping ke string untuk penanganan angka di Sheet)
    const colIndex = headers.map(String).indexOf(String(day));
    if (colIndex === -1) return { success: false, message: "Kolom tanggal tidak ditemukan (Header: " + day + ")" };
    
    for (let i = 1; i < data.length; i++) {
       // Normalisasi cellValue bulan
       let rowBulan = data[i][0];
       if (rowBulan instanceof Date) rowBulan = Utilities.formatDate(rowBulan, "GMT+7", "yyyy-MM");
       
       if (String(rowBulan).trim() === monthYear && String(data[i][1]).trim() === name) {
         sheet.getRange(i + 1, colIndex + 1).setValue(newValue);
         return { success: true };
       }
    }
    return { success: false, message: "Data baris tidak ditemukan" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Mengambil setting jam shift global untuk Roster
 */
function getRosterSettings() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SETTINGS);
    const data = (sheet.getLastRow() > 0) ? sheet.getDataRange().getValues() : [];
    const result = {};
    for (let i = 1; i < data.length; i++) {
        if (data[i][0]) result[data[i][0]] = data[i][1];
    }
    return { 
      success: true, 
      data: {
        pagiIn: result.roster_pagi_in || "08:00",
        pagiOut: result.roster_pagi_out || "17:00",
        malamIn: result.roster_malam_in || "20:00",
        malamOut: result.roster_malam_out || "05:00",
        toleransi: parseInt(result.roster_toleransi) || 0
      }
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Menyimpan setting jam shift global untuk Roster
 */
function saveRosterSettings(settings, requesterUsername) {
  if (requesterUsername && !checkPermission(requesterUsername, 'jadwalShift')) return { success: false, message: "Akses Ditolak" };
  
  try {
    const sheet = getSheet(CONFIG.SHEETS.SETTINGS);
    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();
    
    const upsert = (key, val) => {
      let found = false;
      for (let i = 1; i < data.length; i++) {
         if (data[i][0] === key) {
           sheet.getRange(i + 1, 2).setValue(val);
           sheet.getRange(i + 1, 3).setValue(now);
           found = true;
           break;
         }
      }
      if (!found) sheet.appendRow([key, val, now]);
    };
    
    upsert('roster_pagi_in', settings.pagiIn);
    upsert('roster_pagi_out', settings.pagiOut);
    upsert('roster_malam_in', settings.malamIn);
    upsert('roster_malam_out', settings.malamOut);
    upsert('roster_toleransi', settings.toleransi);
      return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Mendapatkan ringkasan absensi pribadi untuk user yang sedang login hari ini
 * @param {string} nama Nama lengkap pengguna
 */
function getMyAttendanceToday(nama) {
  if (!nama) return { success: false, message: "Nama tidak valid" };
  
  try {
    const ss = getSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const now = new Date();
    const todayStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
    const monthYear = Utilities.formatDate(now, tz, "yyyy-MM");
    const dayDate = now.getDate();
    
    // Formatting tanggal Indonesia
    const days = ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
    const months = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
    const tglFriendly = days[now.getDay()] + ", " + dayDate + " " + months[now.getMonth()] + " " + now.getFullYear();

    let result = {
      tanggal: tglFriendly,
      shift: "-",
      in: "--:--",
      out: "--:--",
      status: "Belum Absen",
      timestamp: now.getTime()
    };

    const cleanNama = String(nama).trim().toLowerCase();

    // 1. Cari Jadwal Shift (Roster)
    const rosterSheet = ss.getSheetByName(CONFIG.SHEETS.JADWAL_ROSTER);
    if (rosterSheet) {
      const rData = rosterSheet.getDataRange().getValues();
      const rHeaders = rData[0];
      const nameCol = 1; // Nama di kolom B
      const dayColStart = 2; // Tanggal 1 di kolom C
      
      for (let i = 1; i < rData.length; i++) {
        let rMonth = rData[i][0];
        if (rMonth instanceof Date) rMonth = Utilities.formatDate(rMonth, tz, "yyyy-MM");
        
        if (String(rMonth) === monthYear && String(rData[i][nameCol]).trim().toLowerCase() === cleanNama) {
          for (let j = dayColStart; j < rHeaders.length; j++) {
            if (String(rHeaders[j]) == String(dayDate)) {
              result.shift = String(rData[i][j] || "-").toUpperCase();
              break;
            }
          }
          break;
        }
      }
    }

    // 2. Cari Log Absensi
    const attSheet = ss.getSheetByName(CONFIG.SHEETS.ABSENSI_KARYAWAN);
    if (attSheet) {
      const aData = attSheet.getDataRange().getValues();
      let logs = [];
      
      for (let i = 1; i < aData.length; i++) {
        let aTgl = aData[i][1];
        let aTglStr = "";
        if (aTgl instanceof Date) aTglStr = Utilities.formatDate(aTgl, tz, "yyyy-MM-dd");
        else if (aTgl) aTglStr = String(aTgl).split("T")[0];
        
        if (aTglStr === todayStr && String(aData[i][4]).trim().toLowerCase() === cleanNama) {
          let jam = aData[i][2];
          let jamStr = "";
          if (jam instanceof Date) jamStr = Utilities.formatDate(jam, tz, "HH:mm");
          else if (jam) jamStr = String(jam).substring(0, 5);
          
          if (jamStr && jamStr !== "00:00") logs.push(jamStr);
        }
      }

      if (logs.length > 0) {
        logs.sort();
        result.in = logs[0];
        result.status = "Sudah Absen";
        // Absen pertama masuk, absen terakhir hari ini pulang (jika berbeda)
        if (logs.length > 1) {
          result.out = logs[logs.length - 1];
        } else if (result.shift === "OFF") {
          result.status = "OFF (Masuk?)";
        }
      }
    }

    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: "Error Server: " + e.toString() };
  }
}
