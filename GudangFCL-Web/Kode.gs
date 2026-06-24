// ============================================================
// GUDANG FCL - Google Apps Script Backend
// Code.gs - Main Server-Side Logic
// ============================================================

const CONFIG = {
  SPREADSHEET_ID: '1lde5La49rhI5NElJNtpaGP7ZMFcS9n28ZNRy6YyhU3s', // Database utama aplikasi
  DISTRIBUTOR_QUEUE_SPREADSHEET_ID: '1etetQhe3HGaE8eR2iUHOnUcCLXkdhHrnfzjMRG-tcfA', // Spreadsheet khusus Antrian Distributor
  SHEETS: {
    USERS: 'Users',
    KAS_GUDANG: 'KasGudang',
    TEAM_BUILDING: 'TeamBuilding',
    EXPENSE: 'Expense',
    KARYAWAN: 'Karyawan',
    KPI_KARYAWAN: 'KPI Karyawan',
    KPI_QUESTIONS: 'KPI Questions',
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
    DISTRIBUTOR_QUEUE: 'Antrian Distributor',
    DISTRIBUTOR_QUEUE_FOCALSKIN: 'ANTRIAN FOCALSKIN',
    DISTRIBUTOR_QUEUE_MISTINE: 'ANTRIAN MISTINE',
    DISTRIBUTOR_QUEUE_SBY: 'ANTRIAN SBY',
    LATE_SHIPMENT: 'LateShipment',
    RETUR: 'Retur',
    RETUR_DETAIL: 'ReturDetail',
    STOCK_CONTROL: 'Stock Control',
    STOCK_CONTROL_DETAIL: 'StockControlDetail',
    KLAIM: 'Klaim',
    KLAIM_DETAIL: 'KlaimDetail',
    TUGAS_PROJECT: 'TugasProject',
    ASSET: 'PengajuanAsset',
    ASSET_WAREHOUSE: 'AssetWarehouse',
    STOCK_OPNAME: 'StockOpname',
    PACKING_LIST: 'PackingList',
    RIWAYAT_KARYAWAN: 'RiwayatKaryawan',
    SURAT_PERINGATAN: 'SuratPeringatan',
    TGL_MERAH: 'TglMerah',
    BOOKING_MOBIL: 'BookingMobil',
    WAREHOUSE_MAP: 'WarehouseMap',
    ABSENSI_KARYAWAN: 'AbsensiKaryawan',
    JADWAL_SHIFT: 'JadwalShift',
    JADWAL_ROSTER: 'JadwalRoster',
    ASSET_AUDIT_LOG: 'AssetAuditLog',
    AUDIT_REPORTS: 'AuditReports',
    INVENTORY_CONTROL: 'InventoryControl',
    INVENTORY_MONITORING: 'InventoryMonitoring',
    BOOKING_MOBIL_DETAIL: 'BookingMobilDetail',
    PETTY_CASH: 'PettyCash',
    PETTY_CASH_PERIOD: 'PettyCashPeriod',
    PAYMENT_GUDANG: 'PaymentGudang',
    PAYMENT_GUDANG_PARTICIPANTS: 'PaymentGudangParticipants',
    MIDTRANS_CONFIG: 'MidtransConfig',
    PAYMENT_KOL_INSTANT: 'PaymentKOLInstant'
  },
  DRIVE_FOLDER_ID: '14u5aMQltzyc7BCw3-87p25mqPeYf9weC',
  BOOKING_PAYMENT_FOLDER_ID: '1rPn5Fq0KvwKCgx1rlCCBm2C1fGrfM6Oo',  // Folder untuk Bukti Pembayaran Booking Mobil
  PETTY_CASH_FOLDER_ID: '1d1jqWSHxHQkEXu2bmhFDTTHuA36afRPd'          // Folder utama Petty Cash Drive
};

const SHEET_HEADERS = {
  [CONFIG.SHEETS.USERS]: ['id', 'username', 'password', 'nama', 'role', 'createdAt', 'permissions', 'divisi'],
  [CONFIG.SHEETS.KAS_GUDANG]: ['id', 'tanggal', 'tipe', 'keterangan', 'nominal', 'buktiUrl', 'createdBy', 'createdAt'],
  [CONFIG.SHEETS.TEAM_BUILDING]: ['id', 'tanggal', 'keterangan', 'nominal', 'buktiUrl', 'createdBy', 'createdAt', 'tipe'],
  [CONFIG.SHEETS.EXPENSE]: ['id', 'tanggal', 'perusahaan', 'kategori', 'keterangan', 'nominal', 'bank', 'rekening', 'createdBy', 'createdAt'],
  [CONFIG.SHEETS.KARYAWAN]: ['id', 'nama', 'jabatan', 'cabang', 'telepon', 'email', 'tanggalMasuk', 'status', 'createdAt', 'tanggalSelesai', 'sisaCuti', 'fingerprintId'],
  [CONFIG.SHEETS.KPI_KARYAWAN]: ['id', 'username', 'nama', 'divisi', 'tanggal', 'totalPoints', 'grade', 'answersJson', 'submittedAt', 'createdAt'],
  [CONFIG.SHEETS.KPI_QUESTIONS]: ['id', 'divisi', 'question', 'opt1Label', 'opt1Point', 'opt2Label', 'opt2Point', 'opt3Label', 'opt3Point', 'opt4Label', 'opt4Point'],
  [CONFIG.SHEETS.IJIN]: ['id', 'tanggal', 'nama', 'jenis', 'keterangan', 'bukti', 'status', 'createdBy', 'createdAt', 'history'],
  [CONFIG.SHEETS.LEMBUR]: ['id', 'tanggal', 'nama', 'divisi', 'jumlahJam', 'keterangan', 'status', 'createdBy', 'createdAt', 'history'],
  [CONFIG.SHEETS.LAPORAN_KERJA]: ['id', 'tanggal', 'divisi', 'pic', 'totalOrang', 'perbantuan', 'pengurangan', 'jamLembur', 'totalJamKerja', 'kendala', 'totalStaff', 'totalAdmin', 'totalOrder', 'createdBy', 'createdAt', 'sisaOrder', 'staffLemburNames', 'shift', 'totalPHL', 'jamKerjaPHL', 'totalPO', 'totalQty', 'totalInbound', 'pendapatanPotongBubble', 'pendapatanBuatBubble', 'alasanPengurangan'],
  [CONFIG.SHEETS.SOP]: ['id', 'judul', 'konten', 'kategori', 'createdBy', 'updatedAt'],
  [CONFIG.SHEETS.ORGANISASI]: ['id', 'nama', 'jabatan', 'atasan', 'departemen', 'foto', 'urutan'],
  [CONFIG.SHEETS.STOCK]: ['id','sku','nama','barcode','batch','expDate','satuan','stok','stokMin','kategori','lokasi','createdAt','updatedAt'],
  [CONFIG.SHEETS.SURAT_JALAN_MASUK]: ['id','noSJ','tanggal','supplier','keterangan','createdBy','createdAt'],
  [CONFIG.SHEETS.SURAT_JALAN_MASUK_DETAIL]: ['id','sjId','noSJ','stockId','sku','nama','qty','satuan','batch','expDate'],
  [CONFIG.SHEETS.SURAT_JALAN_KELUAR]: ['id','noSJ','tanggal','tujuan','keterangan','createdBy','createdAt'],
  [CONFIG.SHEETS.SURAT_JALAN_KELUAR_DETAIL]: ['id','sjId','noSJ','stockId','sku','nama','qty','satuan','batch','expDate'],
  [CONFIG.SHEETS.ORDER]: ['id','noOrder','tanggal','pelanggan','alamat','status','totalItem','keterangan','createdBy','createdAt','sentAt','buktiPacking','kategori','noResi'],
  [CONFIG.SHEETS.ORDER_DETAIL]: ['id','orderId','noOrder','stockId','sku','nama','qty','satuan','batch','expDate','packedQty','lokasi'],
  [CONFIG.SHEETS.LATE_SHIPMENT]: ['id','poNumber','sourceSheet','rowNumber','keterangan','status','createdBy','createdAt','history'],
  [CONFIG.SHEETS.RETUR]: ['id','noRetur','tanggal','sumber','alasan','keterangan','createdBy','createdAt'],
  [CONFIG.SHEETS.RETUR_DETAIL]: ['id','returId','noRetur','stockId','sku','nama','qty','satuan','batch','expDate'],
  [CONFIG.SHEETS.STOCK_CONTROL]: ['id', 'tanggal', 'pic', 'area', 'kategori', 'alasan', 'karyawan', 'status', 'createdBy', 'createdAt'],
  [CONFIG.SHEETS.STOCK_CONTROL_DETAIL]: ['id', 'stockControlId', 'lokasi', 'sku', 'batch', 'exp', 'stockMabang', 'stockTtx', 'stockFisik', 'selisihMabang', 'selisihTtx', 'aksi', 'alasan'],
  [CONFIG.SHEETS.KLAIM]: ['id', 'tanggal', 'pic', 'resi', 'harga', 'keterangan', 'status', 'createdBy', 'createdAt'],
  [CONFIG.SHEETS.KLAIM_DETAIL]: ['id', 'klaimId', 'sku', 'harga'],
  [CONFIG.SHEETS.TUGAS_PROJECT]: ['id','judul','assignee','assigneeName','prioritas','tanggalMulai','deadline','targetHari','status','kategori','deskripsi','createdBy','createdAt','updatedAt','log'],
  [CONFIG.SHEETS.ASSET]: ['id','tanggal','nama','jenisAsset','deskripsi','estimasiHarga','prioritas','bukti','status','createdBy','createdAt','history'],
  [CONFIG.SHEETS.STOCK_OPNAME]: ['id','tanggal','stockId','sku','nama','lokasi','batch','expDate','stokSistem','stokFisik','selisih','status','catatan','createdBy','createdAt','approvedBy','approvedAt'],
  [CONFIG.SHEETS.PACKING_LIST]: ['id','tanggal','noPL','keterangan','fileUrl','createdBy','createdAt'],
  [CONFIG.SHEETS.RIWAYAT_KARYAWAN]: ['id','nama','jabatan','cabang','telepon','tanggalMasuk','tanggalResign','alasanResign','keterangan','createdBy','createdAt'],
  [CONFIG.SHEETS.SURAT_PERINGATAN]: ['id','karyawanNama','karyawanId','jenisSP','alasan','tanggalSP','masaBerlaku','tanggalKadaluarsa','status','createdBy','createdAt'],
  [CONFIG.SHEETS.TGL_MERAH]: ['id', 'tanggal', 'nama', 'divisi', 'jamEstimasi', 'createdBy', 'createdAt'],
  [CONFIG.SHEETS.ASSET_WAREHOUSE]: ['id', 'kategori', 'code', 'nama', 'divisi', 'tanggalMasuk', 'status', 'createdBy', 'createdAt', 'history', 'qty', 'zoneId'],
  [CONFIG.SHEETS.WAREHOUSE_MAP]: ['id', 'configJson', 'updatedAt'],
  [CONFIG.SHEETS.BOOKING_MOBIL]: ['id', 'tanggal', 'pic', 'jamBerangkat', 'tujuan', 'keterangan', 'rute', 'status', 'createdBy', 'createdAt', 'parkir', 'tol', 'bensin', 'pkbm', 'lainLain', 'totalBiaya', 'buktiPembayaranUrl', 'driverNotes', 'jamMulaiPerjalanan', 'jamTibaTujuan', 'jamKembaliWarehouse', 'jamSampaiWarehouse'],
  [CONFIG.SHEETS.BOOKING_MOBIL_DETAIL]: ['id', 'bookingId', 'tanggal', 'namaCustomer', 'noPo', 'totalCartoon', 'parkir', 'tol', 'pkbm', 'lainLain', 'keterangan', 'buktiUrls'],
  [CONFIG.SHEETS.PETTY_CASH_PERIOD]: ['id', 'nama', 'tanggalMulai', 'tanggalSelesai', 'saldoAwal', 'keterangan', 'status', 'createdBy', 'createdAt'],
  [CONFIG.SHEETS.PETTY_CASH]: ['id', 'periodId', 'tanggal', 'kategori', 'keterangan', 'nominal', 'tipe', 'buktiUrl', 'createdBy', 'createdAt'],
  [CONFIG.SHEETS.PAYMENT_GUDANG]: ['id', 'nama', 'deskripsi', 'hargaPerOrang', 'deadline', 'status', 'createdBy', 'createdAt', 'midtransOrderId', 'midtransStatus', 'tipe'],
  [CONFIG.SHEETS.PAYMENT_GUDANG_PARTICIPANTS]: ['id', 'paymentId', 'karyawanId', 'namaKaryawan', 'statusBayar', 'tanggalBayar', 'metodeBayar', 'buktiUrl', 'midtransTransactionId'],
  [CONFIG.SHEETS.MIDTRANS_CONFIG]: ['id', 'serverKey', 'clientKey', 'isProduction', 'updatedBy', 'updatedAt'],
  [CONFIG.SHEETS.PAYMENT_KOL_INSTANT]: ['id', 'tanggal', 'noOrder', 'noResi', 'harga', 'createdBy', 'createdAt'],
  [CONFIG.SHEETS.ABSENSI_KARYAWAN]: ['id', 'tanggal', 'jam', 'karyawanId', 'nama', 'divisi', 'jabatan', 'tipe', 'sumber', 'fingerprintId', 'status', 'keterangan', 'createdAt'],
  [CONFIG.SHEETS.JADWAL_SHIFT]: ['id', 'namaJadwal', 'divisi', 'shiftType', 'jamMasuk', 'jamPulang', 'toleransiMenit', 'aktif', 'createdAt', 'updatedAt'],
  [CONFIG.SHEETS.ASSET_AUDIT_LOG]: ['id','assetId','tanggal','kondisi','catatan','petugas','createdAt','statusApproval','approvedBy','approvedAt'],
  [CONFIG.SHEETS.AUDIT_REPORTS]: ['id','tanggal','auditor','totalAsset','terscan','minus','status','createdBy','createdAt','history'],
  [CONFIG.SHEETS.INVENTORY_CONTROL]: ['id', 'tanggalPengerjaan', 'cycleCount', 'lokasi', 'sku', 'batch', 'exp', 'stockTTX', 'stockMabang', 'stockFisik', 'selisihMabang', 'selisihTTX', 'action', 'keterangan', 'createdBy', 'createdAt'],
  [CONFIG.SHEETS.INVENTORY_MONITORING]: ['id', 'areaPosisi', 'keterangan', 'updatedAt'],
  [CONFIG.SHEETS.SETTINGS]: ['key', 'value', 'updatedAt'],
  [CONFIG.SHEETS.JADWAL_ROSTER]: ['Bulan', 'Nama', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31']
};

const SUPABASE_URL = 'https://hnofmrmwkropijhpexpx.supabase.co';
const SUPABASE_KEY = 'sb_publishable_nu_0fpE0B_dnH3Z6d6dkIA_rWE9xGV8';
const SUPABASE_SYNC_SHEETS = Object.keys(SHEET_HEADERS);

function formatSupabaseValue(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (value === undefined) return null;
  if (value === null) return null;
  return value;
}

function buildSupabaseRowObject(sheetName, rowValues) {
  const headers = SHEET_HEADERS[sheetName];
  if (!headers) return null;
  const row = {};
  headers.forEach(function (key, index) {
    row[key] = formatSupabaseValue(rowValues[index]);
  });
  return row;
}

function supabaseUpsertRow(sheetName, rowValues) {
  try {
    if (!SUPABASE_URL || !SUPABASE_KEY) {
      Logger.log('❌ Supabase config belum diatur. Pastikan SUPABASE_URL dan SUPABASE_KEY sudah benar.');
      return { success: false, message: 'Supabase config belum diatur.' };
    }

    const rowObject = buildSupabaseRowObject(sheetName, rowValues);
    if (!rowObject) {
      Logger.log('❌ Tidak dapat membuat object row untuk sheet: ' + sheetName);
      return { success: false, message: 'Sheet tidak dikenali: ' + sheetName };
    }

    const conflictKey = getSupabaseConflictKey(sheetName);
    const tableName = getSupabaseTableName(sheetName);
    const apiUrl = SUPABASE_URL + '/rest/v1/' + encodeURIComponent(tableName) + '?on_conflict=' + encodeURIComponent(conflictKey);
    
    Logger.log('🔄 Syncing to Supabase: ' + tableName + ' (from sheet: ' + sheetName + ')');
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'apikey': SUPABASE_KEY,
        'Authorization': 'Bearer ' + SUPABASE_KEY,
        'Accept': 'application/json',
        'Prefer': 'resolution=merge-duplicates,return=minimal'
      },
      payload: JSON.stringify(rowObject),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const status = response.getResponseCode();
    if (status >= 200 && status < 300) {
      Logger.log('✅ Sync berhasil untuk: ' + tableName);
      return { success: true };
    }

    const body = response.getContentText();
    Logger.log('❌ Supabase sync failed for ' + tableName + ': HTTP ' + status + ' - ' + body);
    
    // Parse error message dari Supabase
    try {
      const errorJson = JSON.parse(body);
      if (errorJson.message) {
        return { success: false, message: 'Sync failed: ' + errorJson.message };
      }
    } catch (e) {}
    
    return { success: false, message: 'Sync failed (' + status + ')' };
  } catch (e) {
    Logger.log('❌ Supabase sync error for ' + sheetName + ': ' + e.message);
    return { success: false, message: e.message };
  }
}

function getSupabaseConflictKey(sheetName) {
  const headers = SHEET_HEADERS[sheetName];
  if (!headers || headers.length === 0) return 'id';
  return headers[0];
}

// Kirim banyak baris sekaligus dalam 1 request, jauh lebih cepat daripada 1 request per baris.
// Dipakai oleh syncSheetRangeToSupabase untuk migrasi/sync massal (setupDatabase, syncAllSheetsToSupabase).
function supabaseUpsertBatch(sheetName, rowObjects) {
  try {
    if (!SUPABASE_URL || !SUPABASE_KEY) {
      Logger.log('❌ Supabase config belum diatur.');
      return { success: false, message: 'Supabase config belum diatur.', synced: 0 };
    }
    if (!rowObjects || rowObjects.length === 0) return { success: true, synced: 0 };

    const conflictKey = getSupabaseConflictKey(sheetName);
    const tableName = getSupabaseTableName(sheetName);
    const apiUrl = SUPABASE_URL + '/rest/v1/' + encodeURIComponent(tableName) + '?on_conflict=' + encodeURIComponent(conflictKey);

    Logger.log('🔄 Syncing batch to Supabase: ' + tableName + ' (' + rowObjects.length + ' rows)');

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'apikey': SUPABASE_KEY,
        'Authorization': 'Bearer ' + SUPABASE_KEY,
        'Accept': 'application/json',
        'Prefer': 'resolution=merge-duplicates,return=minimal'
      },
      payload: JSON.stringify(rowObjects),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const status = response.getResponseCode();
    if (status >= 200 && status < 300) {
      Logger.log('✅ Batch sync berhasil untuk: ' + tableName + ' (' + rowObjects.length + ' rows)');
      return { success: true, synced: rowObjects.length };
    }

    const body = response.getContentText();
    Logger.log('❌ Supabase batch sync failed for ' + tableName + ': HTTP ' + status + ' - ' + body);
    return { success: false, message: 'Batch sync failed (' + status + ')', synced: 0 };
  } catch (e) {
    Logger.log('❌ Supabase batch sync error for ' + sheetName + ': ' + e.message);
    return { success: false, message: e.message, synced: 0 };
  }
}

function getSupabaseTableName(sheetName) {
  if (!sheetName) return '';
  return String(sheetName)
    .trim()
    .replace(/\s+/g, '_')
    .replace(/[^a-zA-Z0-9_]/g, '')
    .toLowerCase();
}

function isSupabaseHeaderRow(sheetName, rowValues) {
  const headers = SHEET_HEADERS[sheetName];
  if (!headers || !Array.isArray(rowValues)) return false;
  return headers.every(function (header, index) {
    return String(rowValues[index] || '').trim() === String(header || '').trim();
  });
}

// ============================================================
// syncRowUpdate — dipanggil setelah update sel tertentu di Sheet
// (setValues, setValue untuk update status, approval, dll)
// Mengambil ulang seluruh baris dari Sheet lalu upsert ke Supabase.
// ============================================================
function syncRowUpdate(sheetName, rowIndex) {
  try {
    if (!sheetName || !rowIndex || rowIndex <= 1) return;
    if (SUPABASE_SYNC_SHEETS.indexOf(sheetName) === -1) return;
    const sheet = getSheet(sheetName);
    if (!sheet) return;
    const headers = SHEET_HEADERS[sheetName];
    const headerCount = headers ? headers.length : sheet.getLastColumn();
    const rowValues = sheet.getRange(rowIndex, 1, 1, headerCount).getValues()[0];
    if (!rowValues || rowValues.every(v => v === '' || v === null || v === undefined)) return;
    const res = supabaseUpsertRow(sheetName, rowValues);
    if (!res.success) {
      Logger.log('⚠️ syncRowUpdate gagal ' + sheetName + ' baris ' + rowIndex + ': ' + (res.message || ''));
    }
  } catch (e) {
    Logger.log('❌ syncRowUpdate error ' + sheetName + ': ' + e.message);
  }
}

// ============================================================
// syncCellUpdate — dipanggil setelah update 1 kolom di baris tertentu
// Ambil seluruh baris terbaru lalu upsert ke Supabase.
// Lebih ringan dari syncRowUpdate karena tidak perlu tahu semua header.
// ============================================================
function syncCellUpdate(sheetName, id, updatedFields) {
  try {
    if (!sheetName || !id || !SUPABASE_URL || !SUPABASE_KEY) return;
    const tableName = getSupabaseTableName(sheetName);
    const conflictKey = getSupabaseConflictKey(sheetName);
    const apiUrl = SUPABASE_URL + '/rest/v1/' + encodeURIComponent(tableName) +
      '?' + encodeURIComponent(conflictKey) + '=eq.' + encodeURIComponent(id);
    UrlFetchApp.fetch(apiUrl, {
      method: 'patch',
      contentType: 'application/json',
      headers: {
        'apikey': SUPABASE_KEY,
        'Authorization': 'Bearer ' + SUPABASE_KEY,
        'Prefer': 'return=minimal'
      },
      payload: JSON.stringify(updatedFields),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('❌ syncCellUpdate error ' + sheetName + ': ' + e.message);
  }
}

function syncSheetRowToSupabase(sheetName, rowValues, sheet, row) {
  if (SUPABASE_SYNC_SHEETS.indexOf(sheetName) === -1) return { success: false, message: 'Sheet tidak di-sync ke Supabase.' };
  if (!rowValues || !Array.isArray(rowValues) || rowValues.every(function (value) { return value === '' || value === null || value === undefined; })) {
    return { success: false, message: 'Row kosong atau tidak valid.' };
  }
  if (isSupabaseHeaderRow(sheetName, rowValues)) {
    return { success: false, message: 'Header row tidak disinkronkan.' };
  }

  const headers = SHEET_HEADERS[sheetName] || [];
  if (headers.length > 0 && headers[0] === 'id' && !rowValues[0]) {
    rowValues[0] = generateId();
    if (sheet && row && sheet.getRange) {
      try {
        sheet.getRange(row, 1).setValue(rowValues[0]);
      } catch (e) {
        Logger.log('Gagal menulis id otomatis ke sheet: ' + e.message);
      }
    }
  }

  return supabaseUpsertRow(sheetName, rowValues);
}

function syncSheetRangeToSupabase(sheetName, rows, sheet) {
  if (!Array.isArray(rows) || rows.length === 0) return { success: true, synced: 0 };

  const BATCH_SIZE = 500; // jumlah baris per request - aman untuk ukuran payload & jauh lebih cepat
  const headers = SHEET_HEADERS[sheetName] || [];
  let synced = 0;
  let batch = [];

  for (let i = 0; i < rows.length; i++) {
    const rowValues = rows[i];
    const rowNumber = i + 2; // baris aktual di sheet (1-indexed + header)

    if (!rowValues || !Array.isArray(rowValues) || rowValues.every(function (value) { return value === '' || value === null || value === undefined; })) {
      continue; // skip baris kosong
    }
    if (isSupabaseHeaderRow(sheetName, rowValues)) continue;

    // Auto-generate id kalau kolom id kosong (sama seperti logic syncSheetRowToSupabase)
    if (headers.length > 0 && headers[0] === 'id' && !rowValues[0]) {
      rowValues[0] = generateId();
      if (sheet && sheet.getRange) {
        try { sheet.getRange(rowNumber, 1).setValue(rowValues[0]); } catch (e) {
          Logger.log('Gagal menulis id otomatis ke sheet: ' + e.message);
        }
      }
    }

    const rowObject = buildSupabaseRowObject(sheetName, rowValues);
    if (rowObject) batch.push(rowObject);

    if (batch.length >= BATCH_SIZE) {
      const result = supabaseUpsertBatch(sheetName, batch);
      if (result.success) synced += result.synced;
      batch = [];
    }
  }

  if (batch.length > 0) {
    const result = supabaseUpsertBatch(sheetName, batch);
    if (result.success) synced += result.synced;
  }

  return { success: true, synced: synced };
}

function syncAllSheetsToSupabase() {
  const props = PropertiesService.getScriptProperties();
  const startTime = Date.now();
  const MAX_RUNTIME_MS = 5 * 60 * 1000; // berhenti sebelum limit 6 menit Apps Script, lalu lanjut otomatis

  Logger.log('🔄 Memulai sinkronisasi semua sheet ke Supabase...');
  const ss = getSpreadsheet();

  // Kalau ada progress tersisa dari eksekusi sebelumnya yang terhenti karena timeout, lanjutkan dari situ
  let report;
  let sheetsToProcess;
  const savedReport = props.getProperty('SYNC_REPORT');
  if (savedReport) {
    report = JSON.parse(savedReport);
    sheetsToProcess = JSON.parse(props.getProperty('SYNC_REMAINING_SHEETS') || '[]');
    Logger.log('▶️  Melanjutkan sync sebelumnya, sisa ' + sheetsToProcess.length + ' sheet.');
  } else {
    report = { total: SUPABASE_SYNC_SHEETS.length, synced: 0, failed: 0, empty: 0, details: {} };
    sheetsToProcess = SUPABASE_SYNC_SHEETS.slice();
  }

  if (!SUPABASE_URL || !SUPABASE_KEY) {
    Logger.log('❌ ERROR: Supabase URL atau API Key belum dikonfigurasi!');
    Logger.log('   Update SUPABASE_URL dan SUPABASE_KEY di bagian atas Kode.gs');
    report.error = 'Supabase config tidak lengkap';
    props.deleteProperty('SYNC_REPORT');
    props.deleteProperty('SYNC_REMAINING_SHEETS');
    return report;
  }

  while (sheetsToProcess.length > 0) {
    if (Date.now() - startTime > MAX_RUNTIME_MS) {
      Logger.log('⏰ Mendekati limit waktu eksekusi, menjadwalkan lanjutan otomatis dalam 10 detik...');
      props.setProperty('SYNC_REPORT', JSON.stringify(report));
      props.setProperty('SYNC_REMAINING_SHEETS', JSON.stringify(sheetsToProcess));
      scheduleSyncContinuation();
      report.continuing = true;
      return report;
    }

    const sheetName = sheetsToProcess.shift();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log('⚠️  Sheet tidak ditemukan: ' + sheetName);
      report.details[sheetName] = { status: 'sheet_not_found' };
      report.failed++;
      continue;
    }

    const values = sheet.getDataRange().getValues();
    if (!values || values.length <= 1) {
      Logger.log('⏭️  Sheet kosong (hanya header atau tidak ada data): ' + sheetName);
      report.details[sheetName] = { status: 'empty', rows: 0 };
      report.empty++;
      continue;
    }

    const rows = values.slice(1);
    const result = syncSheetRangeToSupabase(sheetName, rows, sheet);

    if (result.success) {
      Logger.log('✅ Berhasil sync: ' + sheetName + ' (' + result.synced + ' baris)');
      report.details[sheetName] = { status: 'success', synced: result.synced };
      report.synced++;
    } else {
      Logger.log('❌ Gagal sync: ' + sheetName + ' - ' + (result.error || 'Unknown error'));
      report.details[sheetName] = { status: 'failed', error: result.error };
      report.failed++;
    }
  }

  // Semua sheet sudah diproses - bersihkan progress tersimpan & trigger lanjutan kalau ada
  props.deleteProperty('SYNC_REPORT');
  props.deleteProperty('SYNC_REMAINING_SHEETS');
  removeSyncContinuationTriggers();

  Logger.log('📊 Sinkronisasi selesai: ' + report.synced + '/' + report.total + ' sheet berhasil');
  Logger.log('   - Synced: ' + report.synced);
  Logger.log('   - Failed: ' + report.failed);
  Logger.log('   - Empty: ' + report.empty);

  return report;
}

// Menjadwalkan pemanggilan ulang syncAllSheetsToSupabase beberapa detik kemudian,
// supaya proses sync bisa lanjut otomatis tanpa kena limit 6 menit per eksekusi.
function scheduleSyncContinuation() {
  removeSyncContinuationTriggers(); // hindari trigger dobel
  ScriptApp.newTrigger('syncAllSheetsToSupabase')
    .timeBased()
    .after(10 * 1000)
    .create();
}

function removeSyncContinuationTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'syncAllSheetsToSupabase') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

const DISTRIBUTOR_QUEUE_HEADERS = [
  'No',
  'Order queue time',
  'PIC Sales',
  'Nama Distributor',
  'alamat',
  'No. HP',
  'PO number',
  'No Mabang',
  'Metode Pengiriman',
  'Ongkir dibayar oleh',
  'Note',
  'Time',
  'Status',
  'Jumlah Dus',
  'Total Pcs',
  'PACKER',
  'Validation',
  'Tanggal selesai packing',
  'Ship date',
  'Status Mabang',
  'GDRIVE',
  'Delivery Bill',
  'nomor resi',
  'Bukti Pengiriman',
  'Harga Ongkir',
  'Harga Ongkir Ekspedisi'
];

// ============================================================
// ENTRY POINT
// ============================================================
function doGet(e) {
  // Trigger update untuk mengatasi kolom kosong (title/judul Header) - DINONAKTIFKAN untuk performa
  // try { ForceUpdateAllHeaders(); } catch(err) {}
  try { setupDistributorQueueDatabase(); } catch (err) { console.error('Error setupDistributorQueueDatabase:', err); }
  try { setupReturnDistributorSheet(); } catch (err) { console.error('Error setupReturnDistributorSheet:', err); }
  try { setupPettyCashSheets(); } catch (err) { console.error('Error setupPettyCashSheets:', err); }
  try { ensureKpiKaryawanSheet(); } catch (err) { console.error('Error ensureKpiKaryawanSheet:', err); }
  try { ensureKpiQuestionsSheet(); } catch (err) { console.error('Error ensureKpiQuestionsSheet:', err); }

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
  Logger.log('🚀 Memulai setup database...');
  
  let ss;
  const props = PropertiesService.getScriptProperties();
  
  // Override dengan CONFIG ID agar tidak nyasar ke file lama
  let ssId = CONFIG.SPREADSHEET_ID || props.getProperty('SPREADSHEET_ID');
  
  if (!ssId) {
    Logger.log('📝 Membuat spreadsheet baru...');
    ss = SpreadsheetApp.create('Gudang FCL - Database');
    ssId = ss.getId();
    props.setProperty('SPREADSHEET_ID', ssId);
  } else {
    ss = SpreadsheetApp.openById(ssId);
  }

  Logger.log('📋 Setup sheet struktur...');
  Object.entries(SHEET_HEADERS).forEach(([sheetName, headers]) => {
    setupSheet(ss, sheetName, headers);
  });

  // Ensure installable triggers exist for onEdit syncing
  Logger.log('⚙️  Setup trigger otomatis...');
  ensureTriggers();

  const usersSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
  if (usersSheet.getLastRow() <= 1) {
    Logger.log('👤 Membuat user admin default...');
    usersSheet.appendRow([generateId(), 'admin', hashPassword('admin123'), 'Administrator', 'admin', new Date().toISOString(), '[]']);
  }

  let supabaseSyncReport = null;
  try {
    Logger.log('☁️  Mulai sinkronisasi ke Supabase...');
    supabaseSyncReport = syncAllSheetsToSupabase();
    Logger.log('✅ Setup database selesai dengan sinkronisasi!');
  } catch (err) {
    Logger.log('❌ ERROR saat Supabase sync: ' + err.message);
    supabaseSyncReport = { error: err.message, success: false };
  }
  
  return { 
    success: true, 
    spreadsheetId: ssId, 
    url: ss.getUrl(), 
    supabaseSync: supabaseSyncReport,
    message: 'Database setup selesai. Periksa Execution log untuk detail sinkronisasi.'
  };
}

// ============================================================
// TEST SUPABASE CONNECTIVITY & TABLE EXISTENCE
// ============================================================
function testSupabaseConnection() {
  Logger.log('🧪 Testing Supabase connection...\n');
  
  const report = {
    timestamp: new Date().toISOString(),
    config: {
      url_configured: !!SUPABASE_URL,
      key_configured: !!SUPABASE_KEY,
      url: SUPABASE_URL ? SUPABASE_URL.substring(0, 50) + '...' : 'NOT SET',
      key: SUPABASE_KEY ? SUPABASE_KEY.substring(0, 20) + '...' : 'NOT SET'
    },
    tables: {}
  };

  if (!SUPABASE_URL || !SUPABASE_KEY) {
    Logger.log('❌ FAIL: Supabase URL atau API Key belum dikonfigurasi!');
    Logger.log('   Update di bagian atas Kode.gs:');
    Logger.log('   const SUPABASE_URL = \'...\';');
    Logger.log('   const SUPABASE_KEY = \'...\';\n');
    return report;
  }

  Logger.log('✅ Config OK\n');

  // Test setiap tabel untuk melihat apakah ada di Supabase
  Logger.log('🔍 Checking tables...\n');
  
  Object.keys(SHEET_HEADERS).forEach(function(sheetName) {
    const tableName = getSupabaseTableName(sheetName);
    const testUrl = SUPABASE_URL + '/rest/v1/' + encodeURIComponent(tableName) + '?limit=1';
    
    try {
      const options = {
        method: 'get',
        headers: {
          'apikey': SUPABASE_KEY,
          'Authorization': 'Bearer ' + SUPABASE_KEY,
          'Accept': 'application/json'
        },
        muteHttpExceptions: true,
        timeout: 10
      };

      const response = UrlFetchApp.fetch(testUrl, options);
      const status = response.getResponseCode();

      if (status === 200) {
        Logger.log('✅ ' + sheetName + ' → ' + tableName + ' (TABLE FOUND)');
        report.tables[sheetName] = { status: 'OK', tableName: tableName, code: status };
      } else if (status === 404) {
        Logger.log('❌ ' + sheetName + ' → ' + tableName + ' (TABLE NOT FOUND)');
        report.tables[sheetName] = { status: 'NOT_FOUND', tableName: tableName, code: status, 
          message: 'Buat tabel ini di Supabase SQL Editor dengan nama: ' + tableName };
      } else {
        Logger.log('⚠️  ' + sheetName + ' → ' + tableName + ' (HTTP ' + status + ')');
        report.tables[sheetName] = { status: 'ERROR', tableName: tableName, code: status };
      }
    } catch (e) {
      Logger.log('❌ ' + sheetName + ' → ' + tableName + ' (ERROR: ' + e.message + ')');
      report.tables[sheetName] = { status: 'ERROR', tableName: tableName, error: e.message };
    }
  });

  Logger.log('\n📊 Summary:');
  let found = 0, notFound = 0, error = 0;
  Object.values(report.tables).forEach(function(t) {
    if (t.status === 'OK') found++;
    else if (t.status === 'NOT_FOUND') notFound++;
    else error++;
  });
  
  Logger.log('   ✅ Found: ' + found);
  Logger.log('   ❌ Not Found: ' + notFound);
  Logger.log('   ⚠️  Errors: ' + error);
  
  if (notFound > 0) {
    Logger.log('\n💡 ACTION: Run SQL script di Supabase SQL Editor:');
    Logger.log('   1. Buka: https://app.supabase.com/project/_/sql');
    Logger.log('   2. Paste file: SUPABASE_CREATE_TABLES.sql');
    Logger.log('   3. Run');
    Logger.log('   4. Jalankan testSupabaseConnection() lagi\n');
  }

  return report;
}

function ensureKpiKaryawanSheet() {
  const ss = getSpreadsheet();
  return setupSheet(ss, CONFIG.SHEETS.KPI_KARYAWAN, ['id', 'username', 'nama', 'divisi', 'tanggal', 'totalPoints', 'grade', 'answersJson', 'submittedAt', 'createdAt']);
}

function ensureKpiQuestionsSheet() {
  const ss = getSpreadsheet();
  const sheet = setupSheet(ss, CONFIG.SHEETS.KPI_QUESTIONS, ['id', 'divisi', 'question', 'opt1Label', 'opt1Point', 'opt2Label', 'opt2Point', 'opt3Label', 'opt3Point', 'opt4Label', 'opt4Point']);
  if (sheet && sheet.getLastRow() <= 1) {
    const sample = [
      ['WQ1', 'Warehouse', 'Langkah pertama yang benar saat menerima barang masuk?', 'Periksa dokumen, jumlah, dan kondisi barang', 25, 'Langsung simpan barang tanpa pemeriksaan', 0, 'Tunggu supervisor memeriksa terlebih dahulu', 10, 'Tarik pallet tanpa membuka segel', 5],
      ['WQ2', 'Warehouse', 'Bagaimana cara menjaga akurasi stok gudang?', 'Melakukan pengecekan berkala dan pencatatan rapi', 25, 'Hanya mengandalkan ingatan saja', 0, 'Memindahkan barang tanpa update sistem', 5, 'Menunggu audit tahunan', 10],
      ['IQ1', 'Inbound', 'Apa prioritas utama operasional inbound?', 'Segera input barang ke sistem setelah pemeriksaan', 25, 'Biarkan menumpuk sebelum diproses', 0, 'Tunggu supervisor memberi instruksi', 10, 'Simpan dulu lalu input data besok', 5],
      ['DQ1', 'Distributor', 'Saat menyiapkan paket distributor, apa yang harus diperiksa?', 'Kondisi barang, alamat, dan kode resi', 25, 'Hanya memastikan barang ada', 0, 'Cukup kirim tanpa cek ulang', 5, 'Tanyakan ke kolega nanti', 10],
      ['HRQ1', 'HR', 'Bagaimana Anda membantu menjaga kedisiplinan tim?', 'Memberikan reminder dan monitoring kehadiran', 25, 'Biarkan setiap orang mengatur sendiri', 0, 'Hanya catat masalah bila terjadi', 10, 'Serahkan sepenuhnya pada mesin presensi', 5],
      ['OQ1', 'Operasional', 'Apa yang dilakukan saat ada selisih Stock Opname?', 'Lakukan pengecekan ulang dan laporkan ke supervisor', 25, 'Abaikan selisih jika kecil', 0, 'Ubah data tanpa konfirmasi', 5, 'Tunggu audit berikutnya', 10]
    ];
    sample.forEach(row => sheet.appendRow(row));
  }
  return sheet;
}

function getKpiQuestions(divisi) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KPI_QUESTIONS);
    if (!sheet) return { success: true, data: [] };
    const data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) return { success: true, data: [] };
    const target = String(divisi || '').trim().toLowerCase();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || row.join('').trim() === '') continue;
      const rowDivision = String(row[1] || '').trim() || 'General';
      const rowDivisionLower = rowDivision.toLowerCase();
      if (target && rowDivisionLower !== target && rowDivisionLower !== 'general') continue;
      const options = [];
      for (let j = 3; j <= 9; j += 2) {
        const label = String(row[j] || '').trim();
        const point = parseInt(row[j + 1], 10) || 0;
        if (label) options.push({ label, point });
      }
      result.push({
        id: row[0],
        divisi: rowDivision,
        question: String(row[2] || '').trim(),
        options: options
      });
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addBulkKpiQuestions(items) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KPI_QUESTIONS) || ensureKpiQuestionsSheet();
    const data = sheet.getDataRange().getValues();
    let updated = 0, added = 0;

    items.forEach(item => {
      if (!item.question) return;
      const rowData = [
        item.id || generateId(),
        item.divisi || 'General',
        item.question,
        item.opt1Label || '',
        item.opt1Point || 0,
        item.opt2Label || '',
        item.opt2Point || 0,
        item.opt3Label || '',
        item.opt3Point || 0,
        item.opt4Label || '',
        item.opt4Point || 0
      ];

      let foundIndex = -1;
      for (let i = 1; i < data.length; i++) {
        if (!data[i] || data[i].join('').trim() === '') continue;
        const existingId = String(data[i][0] || '').trim();
        const existingQuestion = String(data[i][2] || '').trim().toLowerCase();
        if ((item.id && existingId && String(existingId) === String(item.id)) ||
            (!item.id && existingQuestion === String(item.question || '').trim().toLowerCase())) {
          foundIndex = i;
          break;
        }
      }

      if (foundIndex > -1) {
        sheet.getRange(foundIndex + 1, 1, 1, rowData.length).setValues([rowData]);
        updated++;
      } else {
        sheet.appendRow(rowData);
        syncSheetRowToSupabase(CONFIG.SHEETS.KPI_QUESTIONS, rowData);
        added++;
      }
    });

    return { success: true, updated: updated, added: added };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addKpiKaryawan(username, nama, divisi, tanggal, totalPoints, grade, answers) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KPI_KARYAWAN) || ensureKpiKaryawanSheet();
    if (!sheet) return { success: false, message: 'Sheet KPI Karyawan tidak tersedia' };
    const id = generateId();
    const answersJson = typeof answers === 'string' ? answers : JSON.stringify(answers || []);
    const row = [id, username || '', nama || '', divisi || '', tanggal || '', totalPoints || 0, grade || '', answersJson, new Date().toLocaleString(), new Date().toISOString()];
    sheet.appendRow(row);
    syncSheetRowToSupabase(CONFIG.SHEETS.KPI_KARYAWAN, row);
    return { success: true, id: id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getKpiKaryawan(username) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KPI_KARYAWAN);
    if (!sheet) return { success: true, data: [] };
    const data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) return { success: true, data: [] };
    const lowerUsername = String(username || '').toLowerCase();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row) continue;
      const rowUsername = String(row[1] || '').toLowerCase();
      if (!lowerUsername || rowUsername === lowerUsername) {
        let answers = [];
        try { answers = typeof row[7] === 'string' ? JSON.parse(row[7]) : row[7] || []; } catch (err) { answers = []; }
        result.push({
          id: row[0],
          username: row[1],
          nama: row[2],
          divisi: row[3],
          tanggal: row[4],
          totalPoints: row[5],
          grade: row[6],
          answers: answers,
          submittedAt: row[8],
          createdAt: row[9]
        });
      }
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateKpiKaryawan(id, nama, divisi, tanggal, totalPoints, grade, answers) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KPI_KARYAWAN);
    if (!sheet) return { success: false, message: 'Sheet KPI Karyawan tidak tersedia' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        const answersJson = typeof answers === 'string' ? answers : JSON.stringify(answers || []);
        sheet.getRange(i + 1, 3).setValue(nama || data[i][2] || '');
        sheet.getRange(i + 1, 4).setValue(divisi || data[i][3] || '');
        sheet.getRange(i + 1, 5).setValue(tanggal || data[i][4] || '');
        sheet.getRange(i + 1, 6).setValue(totalPoints != null ? totalPoints : data[i][5]);
        sheet.getRange(i + 1, 7).setValue(grade || data[i][6] || '');
        sheet.getRange(i + 1, 8).setValue(answersJson);
        return { success: true };
      }
    }
    return { success: false, message: 'Data KPI tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function setupSheet(ss, sheetName, headers) {
  // SAFETY GUARD: Prevent creation of default "SheetXXX" if name is missing or invalid
  if (!sheetName || typeof sheetName !== 'string' || sheetName.trim() === "" || sheetName === "undefined" || sheetName === "null") {
    console.error('Skipping setupSheet: sheetName is invalid ->', sheetName);
    return null;
  }
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

// Migrasi data lama: pastikan kolom 'kategori' ada dan isi default untuk baris yang kosong
function migrateAssetWarehouseKategori() {
  try {
    const ss = getSpreadsheet();
    const sheet = getSheet(CONFIG.SHEETS.ASSET_WAREHOUSE);
    if (!sheet) return { success: false, message: 'Sheet AssetWarehouse tidak ditemukan' };

    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    if (!data || data.length <= 1) return { success: true, message: 'Tidak ada data untuk dimigrasi' };

    // Pastikan header sesuai urutan baru
    const desiredHeader = ['id', 'kategori', 'code', 'nama', 'divisi', 'tanggalMasuk', 'status', 'createdBy', 'createdAt', 'history', 'qty', 'zoneId'];
    sheet.getRange(1, 1, 1, desiredHeader.length).setValues([desiredHeader]);

    let updated = 0;
    for (let i = 1; i < data.length; i++) {
      const row = data[i] || [];
      const val = (row && row[1]) !== undefined ? row[1] : '';
      if (val === null || String(val || '').trim() === '') {
        sheet.getRange(i + 1, 2).setValue('Lain-lain');
        updated++;
      }
      // Ensure qty exists at col 11
      if ((row[10] === undefined || row[10] === '') ) {
        sheet.getRange(i + 1, 11).setValue(1);
      }
    }

    return { success: true, message: 'Migrasi selesai. Baris yang diperbarui: ' + updated };
  } catch (e) { return { success: false, message: e.message }; }
}

/**
 * Menghapus sheet otomatis yang tidak terpakai (Sheet1, Sheet2, dll)
 * Peringatan: Gunakan dengan hati-hati, pastikan tidak ada data di sheet tersebut.
 */
function cleanupJunkSheets() {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  const pattern = /^Sheet\d+$/i;
  let count = 0;
  
  sheets.forEach(sheet => {
    const name = sheet.getName();
    // Hanya hapus jika namanya cocok pola Sheet + angka DAN barisnya kosong (hanya header atau kosong sama sekali)
    if (pattern.test(name)) {
      if (sheet.getLastRow() <= 1) {
        ss.deleteSheet(sheet);
        count++;
      }
    }
  });
  return { success: true, message: count + " sheet sampah berhasil dihapus." };
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

function getSheetHeaders(sheetName) {
  return SHEET_HEADERS[sheetName] || null;
}

function getSheet(sheetName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    const headers = getSheetHeaders(sheetName);
    if (headers) {
      sheet = setupSheet(ss, sheetName, headers);
    }
  }
  return sheet;
}

function getDistributorQueueSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.DISTRIBUTOR_QUEUE_SPREADSHEET_ID);
}

function setupDistributorQueueDatabase() {
  const ss = getDistributorQueueSpreadsheet();
  setupSheet(ss, CONFIG.SHEETS.DISTRIBUTOR_QUEUE, DISTRIBUTOR_QUEUE_HEADERS);
  
  // Setup 3 sheet baru dengan template header yang sama
  setupSheet(ss, CONFIG.SHEETS.DISTRIBUTOR_QUEUE_FOCALSKIN, DISTRIBUTOR_QUEUE_HEADERS);
  setupSheet(ss, CONFIG.SHEETS.DISTRIBUTOR_QUEUE_MISTINE, DISTRIBUTOR_QUEUE_HEADERS);
  setupSheet(ss, CONFIG.SHEETS.DISTRIBUTOR_QUEUE_SBY, DISTRIBUTOR_QUEUE_HEADERS);
  
  return { success: true, url: ss.getUrl(), spreadsheetId: ss.getId() };
}

function getDistributorQueueSheet() {
  const ss = getDistributorQueueSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DISTRIBUTOR_QUEUE);

  if (!sheet) {
    setupDistributorQueueDatabase();
    sheet = ss.getSheetByName(CONFIG.SHEETS.DISTRIBUTOR_QUEUE);
  }

  if (!sheet) {
    throw new Error('Sheet Antrian Distributor gagal dibuat di spreadsheet distributor.');
  }

  setupSheet(ss, CONFIG.SHEETS.DISTRIBUTOR_QUEUE, DISTRIBUTOR_QUEUE_HEADERS);
  return sheet;
}

function deleteRow(sheetName, id, secondaryId, secondaryCol) {
  try {
    const sheet = getSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    let deletedId = null;
    for (let i = 1; i < data.length; i++) {
      const matchPrimary = id && String(data[i][0]) === String(id);
      const matchSecondary = secondaryId && secondaryCol !== undefined && String(data[i][secondaryCol]) === String(secondaryId);
      if (matchPrimary || matchSecondary) {
        deletedId = String(data[i][0] || id || '');
        sheet.deleteRow(i + 1);
        break;
      }
    }
    if (deletedId === null) return { success: false, message: 'Data tidak ditemukan' };

    // Langsung hapus dari Supabase agar Web App tidak perlu menunggu onChange trigger
    if (deletedId && SUPABASE_URL && SUPABASE_KEY) {
      try {
        const tableName = getSupabaseTableName(sheetName);
        const conflictKey = getSupabaseConflictKey(sheetName);
        const apiUrl = SUPABASE_URL + '/rest/v1/' + encodeURIComponent(tableName) +
          '?' + encodeURIComponent(conflictKey) + '=eq.' + encodeURIComponent(deletedId);
        UrlFetchApp.fetch(apiUrl, {
          method: 'delete',
          headers: {
            'apikey': SUPABASE_KEY,
            'Authorization': 'Bearer ' + SUPABASE_KEY,
            'Prefer': 'return=minimal'
          },
          muteHttpExceptions: true
        });
      } catch (syncErr) {
        Logger.log('⚠️  deleteRow Supabase sync error for ' + sheetName + ': ' + syncErr.message);
        // Tidak gagalkan operasi utama jika Supabase sync error
      }
    }

    return { success: true };
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
        const bulkUserRow = [id, username, hashPassword(password), nama, role, new Date().toISOString(), permissions];
        sheet.appendRow(bulkUserRow);
        syncSheetRowToSupabase(CONFIG.SHEETS.USERS, bulkUserRow);
        existingUsernames.add(username.toLowerCase());
        addedCount++;
      }
    });
    
    return { success: true, message: addedCount + ' akun berhasil diimpor.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addUser(username, password, nama, role, permissions, divisi) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === username) return { success: false, message: 'Username sudah ada' };
    }
    const id = generateId();
    const row = [id, username, hashPassword(password), nama, role || 'user', new Date().toISOString(), permissions || '[]', divisi || ''];
    sheet.appendRow(row);
    syncSheetRowToSupabase(CONFIG.SHEETS.USERS, row);
    return { success: true, id: id };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateUser(id, username, password, nama, role, permissions, divisi) {
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
        sheet.getRange(i + 1, 8).setValue(divisi || '');
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
    // Menggunakan Promise-like pattern untuk eksekusi paralel
    const results = {
      ijin: [],
      lembur: [],
      asset: [],
      stockOpname: [],
      assetOpname: [],
      stockControl: []
    };
    
    // Ambil data secara paralel untuk performa optimal
    try {
      const ijinResult = getIjin();
      results.ijin = ijinResult.success ? ijinResult.data : [];
    } catch (e) {
      console.error('Error loading ijin:', e);
    }
    
    try {
      const lemburResult = getLembur();
      results.lembur = lemburResult.success ? lemburResult.data : [];
    } catch (e) {
      console.error('Error loading lembur:', e);
    }
    
    try {
      const assetResult = getAsset();
      results.asset = assetResult.success ? assetResult.data : [];
    } catch (e) {
      console.error('Error loading asset:', e);
    }
    
    try {
      const opnameResult = getStockOpname();
      results.stockOpname = opnameResult.success ? opnameResult.data : [];
    } catch (e) {
      console.error('Error loading stock opname:', e);
    }
    
    try {
      const assetOpnameResult = getAssetOpnamePendingApprovals();
      results.assetOpname = assetOpnameResult.success ? assetOpnameResult.data : [];
    } catch (e) {
      console.error('Error loading asset opname:', e);
    }
    
    try {
      const stockControlResult = getStockControl();
      results.stockControl = stockControlResult.success ? stockControlResult.data : [];
    } catch (e) {
      console.error('Error loading stock control:', e);
    }
    
    return { 
      success: true, 
      ...results
    };
  } catch (e) { 
    console.error('Error in getPendingApprovals:', e);
    return { 
      success: false, 
      message: e.message,
      ijin: [],
      lembur: [],
      asset: [],
      stockOpname: [],
      assetOpname: [],
      stockControl: []
    }; 
  }
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
    } else if (tipe === 'assetopname') {
      return approveAssetOpname(id, action, userNama);
    } else if (tipe === 'stockcontrol') {
      return updateStockControlStatus(id, action === 'Approve' ? 'Disetujui' : 'Ditolak');
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
        
        // Cek otorisasi berdasarkan status saat ini
        if (currentStatus === 'Pending Team Leader' && (isTL || isAdmin)) authorized = true;
        else if (currentStatus === 'Pending Vice Supervisor' && (isVice || isAdmin)) authorized = true;
        else if (currentStatus === 'Pending Supervisor' && (isSPV || isAdmin)) authorized = true;
        else if (currentStatus === 'Pending HR' && (isHR || isAdmin)) authorized = true;

        if (!authorized) return { success: false, message: 'Anda tidak memiliki wewenang untuk tahap approval ini (' + currentStatus + ').' };

        let newStatus = '';
        if (action === 'Reject') {
          newStatus = 'Ditolak';
        } else if (action === 'Approve') {
          if (isAdmin) {
            newStatus = 'Disetujui'; // Admin bypasses flow
          } else {
            // Alur Sinkron: TL -> Vice SPV -> SPV -> HR -> Disetujui
            if (currentStatus === 'Pending Team Leader') newStatus = 'Pending Vice Supervisor';
            else if (currentStatus === 'Pending Vice Supervisor') newStatus = 'Pending Supervisor';
            else if (currentStatus === 'Pending Supervisor') newStatus = 'Pending HR';
            else if (currentStatus === 'Pending HR') newStatus = 'Disetujui';
            else newStatus = 'Disetujui';
          }
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
        
        // ── Sync baris yang diupdate ke Supabase ──
        syncRowUpdate(sheetName, i + 1);
        
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

function getApprovalDashboardData(userRole) {
  try {
    const res = getPendingApprovals();
    if (!res.success) return res;

    const isAdmin = (userRole === 'admin' || userRole === 'Super Admin');
    const isTL = (userRole === 'Team Leader' || userRole === 'TL' || userRole.includes('Team Leader'));
    const isVice = (userRole === 'Vice Supervisor' || userRole === 'Vice SPV' || userRole === 'Vice VPV' || userRole.includes('Vice'));
    const isSPV = (userRole === 'Supervisor' || userRole === 'SPV' || userRole === 'Supervisor HR' || (userRole.includes('Supervisor') && !userRole.includes('Vice')));
    const isHR = (userRole === 'HR' || userRole === 'Supervisor HR' || userRole.includes('HR'));

    const filterPending = (list) => {
      return (list || []).filter(item => {
        const status = item.status || 'Pending Team Leader';
        
        if (isAdmin) return status.toLowerCase().includes('pending');
        
        if (status === 'Pending Team Leader' && isTL) return true;
        if (status === 'Pending Vice Supervisor' && isVice) return true;
        if (status === 'Pending Supervisor' && isSPV) return true;
        if (status === 'Pending HR' && isHR) return true;
        
        return false;
      });
    };

    return {
      success: true,
      data: {
        lembur: filterPending(res.lembur),
        ijin: filterPending(res.ijin),
        asset: filterPending(res.asset),
        stockOpname: filterPending(res.stockOpname),
        assetOpname: filterPending(res.assetOpname)
      }
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
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
    const row = [id, nama, jabatan, cabang || '', telepon, email, tanggalMasuk, status || 'Tetap', new Date().toISOString(), tanggalSelesai || '', sisaCuti || 12, normalizeFingerprintId(fingerprintId) || ''];
    sheet.appendRow(row);
    syncSheetRowToSupabase(CONFIG.SHEETS.KARYAWAN, row);
    return { success: true, id: id };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateKaryawan(id, nama, jabatan, cabang, telepon, email, tanggalMasuk, status, tanggalSelesai, sisaCuti, fingerprintId) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KARYAWAN);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if ((id && String(data[i][0]) === String(id)) || (!id && String(data[i][1]) === String(nama))) {
        sheet.getRange(i + 1, 2, 1, 7).setValues([[nama, jabatan, cabang || '', telepon, email, tanggalMasuk, status]]);
        sheet.getRange(i + 1, 10).setValue(tanggalSelesai || '');
        sheet.getRange(i + 1, 11).setValue(sisaCuti || 0);
        sheet.getRange(i + 1, 12).setValue(normalizeFingerprintId(fingerprintId) || '');
        syncRowUpdate(CONFIG.SHEETS.KARYAWAN, i + 1);
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
        syncSheetRowToSupabase(CONFIG.SHEETS.KARYAWAN, rowData);
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
    const row = [generateId(), tanggal, nama, jenis, keterangan, bukti, 'Pending Team Leader', createdBy, new Date().toISOString(), JSON.stringify(historyArr)];
    sheet.appendRow(row);
    syncSheetRowToSupabase(CONFIG.SHEETS.IJIN, row);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteIjin(id) { 
  return deleteRow(CONFIG.SHEETS.IJIN, id); 
}

// ============================================================
// PENGAJUAN LEMBUR
// ============================================================
function getLembur(bulanFilter) {
  try {
    const ss = getSpreadsheet();
    const tz = Session.getScriptTimeZone();
    
    // 1. Fetch Roster Settings
    let settings = { pagiIn:"08:00", pagiOut:"17:00", malamIn:"20:00", malamOut:"05:00" };
    try {
      const settingsRes = getRosterSettings();
      if (settingsRes && settingsRes.success && settingsRes.data) {
        settings = settingsRes.data;
      }
    } catch (e) {
      console.warn('Gagal mengambil roster settings, menggunakan default:', e.message);
    }
    
    // 2. Fetch Division Maps from multiple fallbacks
    const divMap = {};
    
    const orgSheet = ss.getSheetByName(CONFIG.SHEETS.ORGANISASI);
    if (orgSheet) {
      const orgData = orgSheet.getDataRange().getValues();
      for (let i = 1; i < orgData.length; i++) {
        if (orgData[i][1]) {
          divMap[String(orgData[i][1]).trim().toLowerCase()] = String(orgData[i][2] || 'Lainnya');
        }
      }
    }
    
    const karSheet = ss.getSheetByName(CONFIG.SHEETS.KARYAWAN);
    if (karSheet) {
      const karData = karSheet.getDataRange().getValues();
      for (let i = 1; i < karData.length; i++) {
        if (karData[i][1]) {
          const n = String(karData[i][1]).trim().toLowerCase();
          if (!divMap[n]) divMap[n] = String(karData[i][3] || 'Lainnya');
        }
      }
    }
    
    const uSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
    if (uSheet) {
      const uData = uSheet.getDataRange().getValues();
      for (let i = 1; i < uData.length; i++) {
        if (uData[i][3]) {
          const n = String(uData[i][3]).trim().toLowerCase();
          if (!divMap[n]) divMap[n] = String(uData[i][7] || 'Lainnya');
        }
      }
    }
    
    // 3. Fetch Jadwal Shift (active shifts for fallback)
    const jsSheet = ss.getSheetByName(CONFIG.SHEETS.JADWAL_SHIFT);
    const shiftList = [];
    if (jsSheet && jsSheet.getLastRow() > 1) {
      const jsData = jsSheet.getDataRange().getValues();
      for (let i = 1; i < jsData.length; i++) {
        if (!jsData[i][0]) continue;
        const aktif = jsData[i][7];
        if (String(aktif).toLowerCase() === 'ya' || aktif === true || String(aktif).toLowerCase() === 'true') {
          shiftList.push({
            divisi: String(jsData[i][2]),
            jamMasuk: String(jsData[i][4]),
            jamPulang: String(jsData[i][5])
          });
        }
      }
    }
    
    // Helper to get scheduled clock in/out for a record
    const rosterSheet = ss.getSheetByName(CONFIG.SHEETS.JADWAL_ROSTER);
    const rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];
    const rosterHeaders = rosterData.length > 0 ? rosterData[0] : [];
    
    function getScheduledOutServer(nama, tanggal, divisi) {
      if (nama && tanggal && rosterData.length > 1) {
        const tglDate = new Date(tanggal + 'T12:00:00');
        const monthYear = Utilities.formatDate(tglDate, tz, "yyyy-MM");
        const dayNum = String(tglDate.getDate());
        
        for (let i = 1; i < rosterData.length; i++) {
          let rBulan = rosterData[i][0];
          if (rBulan instanceof Date) rBulan = Utilities.formatDate(rBulan, tz, "yyyy-MM");
          
          if (String(rBulan).trim() === monthYear && String(rosterData[i][1]).trim() === nama) {
            const dCol = rosterHeaders.map(String).indexOf(dayNum);
            if (dCol !== -1) {
              const shiftVal = String(rosterData[i][dCol]).toUpperCase();
              if (shiftVal === 'PAGI' || shiftVal.includes('PAGI')) {
                return settings.pagiOut || "17:00";
              } else if (shiftVal === 'MALAM' || shiftVal.includes('MALAM')) {
                return settings.malamOut || "05:00";
              } else if (shiftVal === 'OFF') {
                return "OFF";
              }
            }
            break;
          }
        }
      }
      
      const activeShift = shiftList.filter(s => s.divisi === divisi)[0];
      if (activeShift) {
        let jp = activeShift.jamPulang;
        if (jp instanceof Date) return Utilities.formatDate(jp, tz, 'HH:mm');
        return String(jp || "17:00");
      }
      return settings.pagiOut || "17:00";
    }
    
    function getScheduledInServer(scheduledOut, divisi) {
      if (scheduledOut === settings.pagiOut) return settings.pagiIn;
      if (scheduledOut === settings.malamOut) return settings.malamIn;
      const activeShift = shiftList.filter(s => s.divisi === divisi && String(s.jamPulang) === scheduledOut)[0];
      if (activeShift) {
        let jm = activeShift.jamMasuk;
        if (jm instanceof Date) return Utilities.formatDate(jm, tz, 'HH:mm');
        return String(jm || "08:00");
      }
      return "08:00";
    }

    // 4. Load all Absensi logs into memory map to prevent database hits in loop
    const absSheet = ss.getSheetByName(CONFIG.SHEETS.ABSENSI_KARYAWAN);
    const absLogsByName = {};
    if (absSheet) {
      const absData = absSheet.getDataRange().getValues();
      for (let i = 1; i < absData.length; i++) {
        if (!absData[i][1] || !absData[i][4]) continue;
        const nameKey = String(absData[i][4]).trim().toLowerCase();
        if (!absLogsByName[nameKey]) absLogsByName[nameKey] = [];
        
        let rowTgl = "";
        let rawTgl = absData[i][1];
        if (rawTgl instanceof Date) {
          rowTgl = Utilities.formatDate(rawTgl, tz, 'yyyy-MM-dd');
        } else {
          let s = String(rawTgl).trim().split('T')[0];
          if (s.includes('/')) {
            let parts = s.split('/');
            if (parts[0].length === 4) rowTgl = parts[0] + '-' + parts[1].padStart(2, '0') + '-' + parts[2].padStart(2, '0');
            else rowTgl = parts[2] + '-' + parts[1].padStart(2, '0') + '-' + parts[0].padStart(2, '0');
          } else {
            rowTgl = s;
          }
        }
        
        let rawJam = absData[i][2];
        let jamStr = "";
        if (rawJam instanceof Date) {
          jamStr = Utilities.formatDate(rawJam, tz, 'HH:mm');
        } else {
          let match = String(rawJam || '').trim().match(/^(\d{1,2}):(\d{2})/);
          if (match) {
            jamStr = match[1].padStart(2, '0') + ':' + match[2];
          } else {
            jamStr = String(rawJam || '').trim().substring(0, 5);
          }
        }
        
        absLogsByName[nameKey].push({
          tanggal: rowTgl,
          jam: jamStr,
          tipe: String(absData[i][7] || '').trim().toUpperCase()
        });
      }
    }

    // Helper to search absensi in-memory
    function getAbsensiInOutServer(nama, tanggal, scheduledIn, scheduledOut) {
      const nameKey = String(nama || '').trim().toLowerCase();
      const logs = absLogsByName[nameKey] || [];
      
      let inTime = '-';
      let outTime = '-';
      let outDateStr = tanggal;
      
      const inMnt = _parseTimeToMinutes(scheduledIn || "08:00");
      const outMnt = _parseTimeToMinutes(scheduledOut || "17:00");
      const isNightShift = (inMnt > outMnt);
      
      const tglDate = new Date(tanggal + 'T12:00:00');
      tglDate.setDate(tglDate.getDate() + 1);
      const nextDayStr = Utilities.formatDate(tglDate, tz, 'yyyy-MM-dd');
      
      logs.forEach(log => {
        const isClockIn = (log.tipe === 'IN' || log.tipe.includes('MASUK') || log.tipe.includes('IN'));
        const isClockOut = (log.tipe === 'OUT' || log.tipe.includes('PULANG') || log.tipe.includes('OUT'));
        
        if (isClockIn && log.tanggal === tanggal) {
          if (inTime === '-' || log.jam < inTime) inTime = log.jam;
        }
        if (isClockOut) {
          if (isNightShift) {
            if (log.tanggal === nextDayStr) {
              if (outTime === '-' || log.jam > outTime) {
                outTime = log.jam;
                outDateStr = nextDayStr;
              }
            }
          } else {
            if (log.tanggal === tanggal) {
              if (outTime === '-' || log.jam > outTime) {
                outTime = log.jam;
                outDateStr = tanggal;
              }
            }
          }
        }
      });
      
      return { inTime, outTime, outDateStr };
    }

    // 5. Load Overtime requests
    const sheet = getSheet(CONFIG.SHEETS.LEMBUR);
    if (!sheet) return { success: false, message: "Sheet Lembur tidak ditemukan" };
    
    const data = (sheet.getLastRow() > 0) ? sheet.getDataRange().getValues() : [];
    const result = [];
    if (data.length <= 1) return { success: true, data: [] };

    // Filter bulan di server jika parameter bulanFilter ada (format YYYY-MM)
    let filterYear = null, filterMonth = null;
    if (bulanFilter && bulanFilter.trim() !== "" && bulanFilter.includes('-')) {
      const parts = bulanFilter.split('-');
      filterYear = parseInt(parts[0]);
      filterMonth = parseInt(parts[1]);
    }

    // Gunakan cache untuk scheduled times agar tidak mencari berulang kali dalam loop
    const scheduledCache = {};

    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      
      const rawTgl = data[i][1];
      let tgl = "";
      let rowYear = 0, rowMonth = 0;

      if (rawTgl instanceof Date) {
        tgl = Utilities.formatDate(rawTgl, tz, 'yyyy-MM-dd');
        rowYear = rawTgl.getFullYear();
        rowMonth = rawTgl.getMonth() + 1;
      } else {
        tgl = String(rawTgl);
        const parts = tgl.split('-');
        if (parts.length >= 2) {
          rowYear = parseInt(parts[0]);
          rowMonth = parseInt(parts[1]);
        }
      }

      // OPTIMASI: Hanya proses data yang sesuai filter bulan
      if (filterYear && filterMonth) {
        if (rowYear !== filterYear || rowMonth !== filterMonth) continue;
      }
      
      const nama = data[i][2];
      const nameKey = String(nama || '').trim().toLowerCase();
      const divisi = data[i][3] || divMap[nameKey] || '';
      
      // Sinkronisasi data absensi
      const schOut = getScheduledOutServer(nama, tgl, divisi);
      const schIn = getScheduledInServer(schOut, divisi);
      const abs = getAbsensiInOutServer(nama, tgl, schIn, schOut);

      result.push({
        id: data[i][0],
        tanggal: tgl,
        nama: nama,
        divisi: divisi,
        jumlahJam: data[i][4],
        keterangan: data[i][5],
        status: data[i][6],
        createdBy: data[i][7],
        createdAt: data[i][8] instanceof Date ? data[i][8].toISOString() : String(data[i][8]),
        history: data[i][9] || '[]',
        inTime: abs.inTime,
        outTime: abs.outTime
      });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function addLembur(tanggal, nama, divisi, jumlahJam, keterangan, createdBy) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.LEMBUR);
    const data = sheet.getDataRange().getValues();
    
    // Pastikan Divisi memakai Jabatan dengan melakukan lookup ke database Karyawan
    let realDivisi = divisi;
    try {
      const karSheet = getSheet(CONFIG.SHEETS.KARYAWAN);
      const karData = karSheet.getDataRange().getValues();
      for (let j = 1; j < karData.length; j++) {
        if (String(karData[j][1]).trim().toLowerCase() === nama.trim().toLowerCase()) {
          realDivisi = karData[j][2] || divisi;
          break;
        }
      }
    } catch(e) {}

    let existingRowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      const rowTanggal = data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]);
      const rowNama = String(data[i][2]).trim().toLowerCase();
      const rowStatus = String(data[i][6]);
      
      if (rowTanggal === tanggal && rowNama === nama.trim().toLowerCase() && rowStatus !== 'Dibatalkan') {
        existingRowIdx = i + 1;
        break;
      }
    }
    
    const now = new Date().toISOString();
    if (existingRowIdx !== -1) {
      // Akumulasikan/update jam jika sudah ada baris aktif di hari yang sama
      const oldJam = parseFloat(data[existingRowIdx - 1][4]) || 0;
      const newJam = parseFloat(jumlahJam) || 0;
      
      sheet.getRange(existingRowIdx, 4).setValue(realDivisi);
      sheet.getRange(existingRowIdx, 5).setValue(newJam);
      
      // Gabungkan keterangan jika berbeda
      let oldKet = String(data[existingRowIdx - 1][5]);
      let newKet = oldKet.includes(keterangan) ? oldKet : (oldKet ? oldKet + " | " + keterangan : keterangan);
      sheet.getRange(existingRowIdx, 6).setValue(newKet);
      
      let historyArr = [];
      try { historyArr = JSON.parse(data[existingRowIdx - 1][9] || '[]'); } catch(e) { historyArr = []; }
      historyArr.push({ date: now, action: 'Diupdate (Penggabungan)', status: data[existingRowIdx - 1][6], by: createdBy, role: 'Pemohon', reason: `Akumulasi Jam: ${oldJam} -> ${newJam}` });
      sheet.getRange(existingRowIdx, 10).setValue(JSON.stringify(historyArr));
    } else {
      // Tambah baris baru jika belum ada
      const historyArr = [{ date: now, action: 'Diajukan', status: 'Pending Team Leader', by: createdBy, role: 'Pemohon', reason: '' }];
      const row = [generateId(), tanggal, nama, realDivisi, jumlahJam, keterangan, 'Pending Team Leader', createdBy, now, JSON.stringify(historyArr)];
      sheet.appendRow(row);
      syncSheetRowToSupabase(CONFIG.SHEETS.LEMBUR, row);
    }
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
    const existingData = sheet.getDataRange().getValues();
    const now = new Date().toISOString();
    
    dataArray.forEach(d => {
      let existingRow = -1;
      for (let i = 1; i < existingData.length; i++) {
        const rowTanggal = existingData[i][1] instanceof Date ? Utilities.formatDate(existingData[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(existingData[i][1]);
        const rowNama = String(existingData[i][2]).trim().toLowerCase();
        
        if (rowTanggal === tanggal && rowNama === d.nama.trim().toLowerCase()) {
          existingRow = i + 1;
          break;
        }
      }
      
      if (existingRow !== -1) {
        sheet.getRange(existingRow, 4).setValue(d.divisi);
        sheet.getRange(existingRow, 5).setValue(d.jamEstimasi);
        sheet.getRange(existingRow, 6).setValue(createdBy);
        sheet.getRange(existingRow, 7).setValue(now);
        // Sync update row ke Supabase
        const updatedRow = sheet.getRange(existingRow, 1, 1, 7).getValues()[0];
        syncSheetRowToSupabase(CONFIG.SHEETS.TGL_MERAH, updatedRow);
      } else {
        const newRow = [generateId(), tanggal, d.nama, d.divisi, d.jamEstimasi, createdBy, now];
        sheet.appendRow(newRow);
        syncSheetRowToSupabase(CONFIG.SHEETS.TGL_MERAH, newRow);
      }
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
    
    return { 
      success: true, 
      saldoGudang: sG.saldo || 0, 
      saldoTB: sTB.saldo || 0, 
      history: sanitize(history), 
      totalKasIn: totalKasIn, 
      totalKasOut: totalKasOut, 
      kasData: sanitize(kg.data), 
      tbData: sanitize(tb.data), 
      laporanData: sanitize(lk.success ? lk.data : [])
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
    const row = [generateId(), tanggal, tipe, keterangan, parseFloat(nominal), buktiUrl, createdBy, new Date().toISOString()];
    getSheet(CONFIG.SHEETS.KAS_GUDANG).appendRow(row);
    const syncRes = syncSheetRowToSupabase(CONFIG.SHEETS.KAS_GUDANG, row);
    if (!syncRes.success) {
      Logger.log('Supabase sync failed for addKasGudang: ' + syncRes.message);
    }
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
    const row = [generateId(), tanggal, keterangan, parseFloat(nominal), buktiUrl, createdBy, new Date().toISOString(), tipe || 'Pengeluaran'];
    getSheet(CONFIG.SHEETS.TEAM_BUILDING).appendRow(row);
    const syncRes = syncSheetRowToSupabase(CONFIG.SHEETS.TEAM_BUILDING, row);
    if (!syncRes.success) {
      Logger.log('Supabase sync failed for addTeamBuilding: ' + syncRes.message);
    }
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
  try {
    const row = [generateId(), tanggal, perusahaan, kategori, keterangan, parseFloat(nominal) || 0, bank, rekening, createdBy, new Date().toISOString()];
    getSheet(CONFIG.SHEETS.EXPENSE).appendRow(row);
    const syncRes = syncSheetRowToSupabase(CONFIG.SHEETS.EXPENSE, row);
    if (!syncRes.success) {
      Logger.log('Supabase sync failed for addExpense: ' + syncRes.message);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
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
    
    // Validasi ukuran chunk
    if (!chunkData || chunkData.length === 0) {
      return { success: false, message: 'Chunk data kosong' };
    }
    
    // Support chunk size hingga 30KB (sangat aman untuk file besar)
    if (chunkData.length <= 30000) { 
      cache.put(key, chunkData, 21600); 
      cache.put('meta_' + id + '_count', String(chunkIndex + 1), 21600); 
    } else { 
      // Split jika lebih besar dari 30KB
      const half = Math.ceil(chunkData.length / 2); 
      cache.put(key + '_a', chunkData.substring(0, half), 21600); 
      cache.put(key + '_b', chunkData.substring(half), 21600); 
      cache.put(key + '_split', '1', 21600); 
      cache.put('meta_' + id + '_count', String(chunkIndex + 1), 21600); 
    }
    return { success: true, uploadId: id };
  } catch (e) { 
    Logger.log('uploadChunk error: ' + e.message);
    return { success: false, message: e.message }; 
  }
}

function finalizeChunkedUpload(uploadId, fileName, mimeType, folderName, periodName) {
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
    
    // Khusus untuk Petty Cash, gunakan folder per periode
    if (folderName === 'PettyCash' && periodName) {
      try {
        folder = getPettyCashPeriodFolder(periodName);
        Logger.log('File akan disimpan ke folder Petty Cash periode: ' + periodName);
      } catch(e) {
        Logger.log('Error accessing Petty Cash period folder, using default: ' + e.message);
        folder = getOrCreateBuktiFolder(folderName);
      }
    } else if (folderName === 'Bukti Packing') {
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
    
    const fileId = file.getId();
    Logger.log('File uploaded successfully with ID: ' + fileId);
    
    // PENTING: Set sharing permission agar bisa diakses publik
    // Gunakan metode yang lebih reliable
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      Logger.log('File sharing permission set to ANYONE_WITH_LINK');
      
      // Tambahan: Set sharing secara eksplisit dengan Drive API
      Drive.Permissions.insert(
        {
          'type': 'anyone',
          'role': 'reader',
          'withLink': true
        },
        fileId,
        {
          'supportsAllDrives': true
        }
      );
      Logger.log('Drive API permission set successfully');
    } catch(err) {
      Logger.log('Warning: Tidak bisa set sharing permission: ' + err.message);
    }
    
    // Generate multiple URL formats untuk compatibility
    const fileUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;
    const thumbnailUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1000';
    const directUrl = 'https://lh3.googleusercontent.com/d/' + fileId;
    const embedUrl = 'https://drive.google.com/file/d/' + fileId + '/preview';
    
    Logger.log('File URL: ' + fileUrl);
    Logger.log('Thumbnail URL: ' + thumbnailUrl);
    Logger.log('Direct URL: ' + directUrl);
    Logger.log('Embed URL: ' + embedUrl);
    
    // Tunggu sebentar agar permission propagate
    Utilities.sleep(500);
    
    return { 
      success: true, 
      url: fileUrl,
      fileId: fileId,
      thumbnailUrl: thumbnailUrl,
      directUrl: directUrl,
      embedUrl: embedUrl
    };
  } catch (e) { 
    Logger.log('finalizeChunkedUpload error: ' + e.message);
    return { success: false, message: e.message }; 
  }
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
        alasanPengurangan: data[i][25] || '',
        alasanLembur: data[i][26] || ''
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function addLaporanKerja(tanggal, divisi, pic, totalOrang, perbantuan, pengurangan, jamLembur, totalJamKerja, kendala, totalStaff, totalAdmin, totalOrder, createdBy, sisaOrder, staffLemburNames, shift, totalPHL, jamKerjaPHL, totalPO, totalQty, totalInbound, pendapatanPotongBubble, pendapatanBuatBubble, alasanPengurangan, alasanLembur) {
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

    const row = [
      generateId(), tanggal, divisi, pic, parseInt(totalOrang)||0, parseFloat(perbantuan)||0, parseFloat(pengurangan)||0, parseFloat(jamLembur)||0, parseFloat(totalJamKerja)||0, kendala, parseInt(totalStaff)||0, parseInt(totalAdmin)||0, parseInt(totalOrder)||0, createdBy, new Date().toISOString(), parseInt(sisaOrder)||0, staffLemburNames || '', targetShift, parseInt(totalPHL)||0, parseFloat(jamKerjaPHL)||0, parseInt(totalPO)||0, parseInt(totalQty)||0, parseInt(totalInbound)||0, parseFloat(pendapatanPotongBubble)||0, parseFloat(pendapatanBuatBubble)||0, alasanPengurangan || '', alasanLembur || ''
    ];
    sheet.appendRow(row);
    syncSheetRowToSupabase(CONFIG.SHEETS.LAPORAN_KERJA, row);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}
function importLaporanBulk(dataArr, username) {
  try {
    if (!dataArr || !dataArr.length) return { success: false, message: 'Tidak ada data untuk diimpor.' };
    const sheet = getSheet(CONFIG.SHEETS.LAPORAN_KERJA);
    const existingData = sheet.getDataRange().getValues();
    const existingKeys = new Set();
    
    for (let i = 1; i < existingData.length; i++) {
       const row = existingData[i];
       const dateStr = row[1] instanceof Date ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(row[1]).split('T')[0];
       const key = `${dateStr}_${row[2]}_${row[17] || 'Pagi'}`;
       existingKeys.add(key);
    }
    
    let count = 0;
    const now = new Date().toISOString();
    
    dataArr.forEach(d => {
      const targetDateStr = String(d.tanggal).split('T')[0];
      const key = `${targetDateStr}_${d.divisi}_${d.shift || 'Pagi'}`;
      
      if (!existingKeys.has(key)) {
        const bulkRow = [
          generateId(), d.tanggal, d.divisi, d.pic,
          parseInt(d.totalOrang) || 0,
          parseFloat(d.perbantuan) || 0,
          parseFloat(d.pengurangan) || 0,
          parseFloat(d.jamLembur) || 0,
          parseFloat(d.totalJamKerja) || 0,
          d.kendala || '-',
          parseInt(d.totalStaff) || 0,
          parseInt(d.totalAdmin) || 0,
          parseInt(d.totalOrder) || 0,
          username, now, 0, '', d.shift || 'Pagi',
          parseInt(d.totalPHL) || 0,
          parseFloat(d.jamKerjaPHL) || 0,
          0, parseInt(d.totalQty) || 0, 0, 0, 0,
          d.alasanPengurangan || '', d.alasanLembur || ''
        ];
        sheet.appendRow(bulkRow);
        syncSheetRowToSupabase(CONFIG.SHEETS.LAPORAN_KERJA, bulkRow);
        existingKeys.add(key);
        count++;
      }
    });
    return { success: true, count: count };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteLaporanKerja(id) { return deleteRow(CONFIG.SHEETS.LAPORAN_KERJA, id); }
function updateLaporanKerja(id, tanggal, divisi, pic, totalOrang, perbantuan, pengurangan, jamLembur, totalJamKerja, kendala, totalStaff, totalAdmin, totalOrder, createdBy, sisaOrder, staffLemburNames, shift, totalPHL, jamKerjaPHL, totalPO, totalQty, totalInbound, pendapatanPotongBubble, pendapatanBuatBubble, alasanPengurangan, alasanLembur) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.LAPORAN_KERJA);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i + 1, 2, 1, 26).setValues([[
          tanggal, divisi, pic, parseInt(totalOrang)||0, parseFloat(perbantuan)||0, parseFloat(pengurangan)||0, parseFloat(jamLembur)||0, parseFloat(totalJamKerja)||0, kendala, parseInt(totalStaff)||0, parseInt(totalAdmin)||0, parseInt(totalOrder)||0, createdBy, new Date().toISOString(), parseInt(sisaOrder)||0, staffLemburNames || '', shift || 'Pagi', parseInt(totalPHL)||0, parseFloat(jamKerjaPHL)||0, parseInt(totalPO)||0, parseInt(totalQty)||0, parseInt(totalInbound)||0, parseFloat(pendapatanPotongBubble)||0, parseFloat(pendapatanBuatBubble)||0, alasanPengurangan || '', alasanLembur || ''
        ]]);
        return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

// ============================================================
// STOCK CONTROL (REPLACING HANDOVER)
// ============================================================
/**
 * Get Master Stock Control Records
 */
function getStockControl() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.STOCK_CONTROL);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      result.push({
        id: data[i][0],
        tanggal: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]),
        pic: data[i][2],
        area: data[i][3],
        kategori: data[i][4],
        alasan: data[i][5],
        karyawan: data[i][6],
        status: data[i][7],
        createdBy: data[i][8],
        createdAt: data[i][9] instanceof Date ? data[i][9].toISOString() : String(data[i][9]||''),
        syncLog: data[i][10] || ''
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}
// Register globally
this['getStockControl'] = getStockControl;

function getStockControlStats() {
  try {
    const masterSheet = getSheet(CONFIG.SHEETS.STOCK_CONTROL);
    const detailSheet = getSheet(CONFIG.SHEETS.STOCK_CONTROL_DETAIL);
    const masters = masterSheet.getDataRange().getValues();
    const details = detailSheet.getDataRange().getValues();

    const opnameMasters = [];
    for (let i = 1; i < masters.length; i++) {
      if (masters[i][4] === 'Stock Opname') {
        opnameMasters.push({ id: masters[i][0], tanggal: masters[i][1] });
      }
    }

    const detailMap = {};
    for (let j = 1; j < details.length; j++) {
      const mid = details[j][1];
      if (!detailMap[mid]) detailMap[mid] = [];
      detailMap[mid].push({
        sm: details[j][9], // selisihMabang
        st: details[j][10] // selisihTtx
      });
    }

    let totalItems = 0;
    let correctItems = 0;
    const trend = {};

    opnameMasters.forEach(m => {
      const items = detailMap[m.id] || [];
      const tgl = m.tanggal instanceof Date ? Utilities.formatDate(m.tanggal, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(m.tanggal);
      
      if (!trend[tgl]) trend[tgl] = { total: 0, correct: 0 };

      items.forEach(it => {
        totalItems++;
        trend[tgl].total++;
        if (Number(it.sm) === 0 && Number(it.st) === 0) {
          correctItems++;
          trend[tgl].correct++;
        }
      });
    });

    return {
      success: true,
      stats: {
        totalItems,
        correctItems,
        accuracy: totalItems > 0 ? (correctItems / totalItems) * 100 : 0
      },
      trend: Object.entries(trend).sort((a,b) => a[0].localeCompare(b[0])).map(([t, v]) => ({
        tanggal: t,
        accuracy: v.total > 0 ? (v.correct / v.total) * 100 : 0
      }))
    };
  } catch (e) { return { success: false, message: e.message }; }
}
this['getStockControlStats'] = getStockControlStats;

/**
 * Get Detail Items for a specific Master Record
 */
function getStockControlDetail(masterId) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.STOCK_CONTROL_DETAIL);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(masterId)) {
        result.push({
          id: data[i][0],
          lokasi: data[i][2],
          sku: data[i][3],
          batch: data[i][4],
          exp: data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][5]||''),
          stockMabang: data[i][6],
          stockTtx: data[i][7],
          stockFisik: data[i][8],
          selisihMabang: data[i][9],
          selisihTtx: data[i][10],
          aksi: data[i][11],
          alasan: data[i][12],
          stockFisikStaff: data[i][13] || 0,
          selisihMabangStaff: data[i][14] || 0,
          selisihTtxStaff: data[i][15] || 0
        });
      }
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}
// Register globally
this['getStockControlDetail'] = getStockControlDetail;

/**
 * Save or Update Stock Control Record
 */
function saveStockControl(id, tanggal, pic, area, kategori, alasan, karyawan, items, createdBy) {
  try {
    const masterSheet = getSheet(CONFIG.SHEETS.STOCK_CONTROL);
    const detailSheet = getSheet(CONFIG.SHEETS.STOCK_CONTROL_DETAIL);
    const masterData = masterSheet.getDataRange().getValues();
    const detailData = detailSheet.getDataRange().getValues();
    const now = new Date().toISOString();
    const nowLocal = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    
    let masterId = id || generateId();
    let isUpdate = !!id;
    let masterRowIndex = -1;
    let oldSyncLog = '';
    
    // Status Logic
    let status = 'Selesai';
    if (kategori === 'Stock Opname' && items && items.length > 0) {
      const needsAdjustment = items.some(it => it.aksi === 'Adjust Stock');
      if (needsAdjustment) status = 'Menunggu Approval';
    }

    if (isUpdate) {
      for (let i = 1; i < masterData.length; i++) {
        if (String(masterData[i][0]) === String(masterId)) {
          masterRowIndex = i + 1;
          oldSyncLog = String(masterData[i][10] || '');
          masterSheet.getRange(masterRowIndex, 2, 1, 7).setValues([[tanggal, pic, area, kategori, alasan, karyawan, status]]);
          break;
        }
      }
      if (masterRowIndex === -1) throw new Error('Data tidak ditemukan');
      
      // Delete old details
      for (let j = detailData.length - 1; j >= 1; j--) {
        if (String(detailData[j][1]) === String(masterId)) {
          detailSheet.deleteRow(j + 1);
        }
      }
    } else {
      const masterRow = [masterId, tanggal, pic, area, kategori, alasan, karyawan, status, createdBy, now, ''];
      masterSheet.appendRow(masterRow);
      syncSheetRowToSupabase(CONFIG.SHEETS.STOCK_CONTROL, masterRow);
      masterRowIndex = masterSheet.getLastRow();
    }

    // Save details and calculate variance
    let hasVariance = false;
    if (items && Array.isArray(items)) {
      items.forEach(it => {
        const m = parseFloat(it.m) || 0;
        const ttx = parseFloat(it.ttx) || 0;
        const f = parseFloat(it.f) || 0;
        const fs = parseFloat(it.fStaff) || 0;
        const sm = f - m;
        const st = f - ttx;
        const sms = fs ? (fs - m) : 0;
        const sts = fs ? (fs - ttx) : 0;
        
        if (sm !== 0 || st !== 0 || sms !== 0 || sts !== 0) hasVariance = true;
        
        const detRow = [
          generateId(), masterId, it.lokasi, it.sku, it.batch, it.exp, m, ttx, f, sm, st, it.aksi, it.alasan,
          fs, sms, sts
        ];
        detailSheet.appendRow(detRow);
        syncSheetRowToSupabase(CONFIG.SHEETS.STOCK_CONTROL_DETAIL, detRow);
      });
    }

    // Update log
    const logType = isUpdate ? '🔄 Diperbarui' : '✨ Dibuat';
    const newLog = `${logType}: ${nowLocal}. ${hasVariance ? 'Ada Selisih' : 'Cocok'}.`;
    masterSheet.getRange(masterRowIndex, 11).setValue(newLog + (oldSyncLog ? '\n' + oldSyncLog : ''));

    return { success: true, message: isUpdate ? 'Berhasil memperbarui laporan' : 'Berhasil menyimpan laporan', masterId: masterId };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
this['saveStockControl'] = saveStockControl;

function updateStockControlStatus(id, status) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.STOCK_CONTROL); 
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { 
      if (String(data[i][0]) === String(id)) { 
        sheet.getRange(i + 1, 8).setValue(status); 
        return { success: true }; 
      } 
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

this['updateStockControlStatus'] = updateStockControlStatus;

function deleteStockControl(id) { 
  const res = deleteRow(CONFIG.SHEETS.STOCK_CONTROL, id);
  if (res.success) {
    // Delete details too
    const sheet = getSheet(CONFIG.SHEETS.STOCK_CONTROL_DETAIL);
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][1]) === String(id)) {
        sheet.deleteRow(i + 1);
      }
    }
  }
  return res;
}
this['deleteStockControl'] = deleteStockControl;

function importStockControl(tanggal, pic, area, kategori, alasan, karyawan, items, username) {
  return addStockControl(tanggal, pic, area, kategori, alasan, karyawan, items, username);
}
this['importStockControl'] = importStockControl;

/**
 * Validasi & Hitung Ulang Stock untuk semua laporan dalam kategori dan rentang waktu tertentu.
 * Memperbarui selisihMabang dan selisihTtx pada setiap detail item,
 * serta memperbarui status master jika ada perbedaan yang perlu dikoreksi.
 */
function validateAndRecalculateStock(kategori, tglMulai, tglAkhir, username) {
  try {
    const masterSheet = getSheet(CONFIG.SHEETS.STOCK_CONTROL);
    const detailSheet = getSheet(CONFIG.SHEETS.STOCK_CONTROL_DETAIL);
    const masterData = masterSheet.getDataRange().getValues();
    const detailData = detailSheet.getDataRange().getValues();

    const dateFrom = new Date(tglMulai);
    const dateTo = new Date(tglAkhir);
    dateTo.setHours(23, 59, 59); // Include entire last day

    let masterUpdated = 0;
    let detailUpdated = 0;
    let totalItems = 0;
    let correctItems = 0;

    // Build detail row index keyed by masterId for fast lookup
    const detailRowMap = {}; // masterId -> [{rowIndex, data}]
    for (let j = 1; j < detailData.length; j++) {
      if (detailData[j].join('').trim() === '') continue;
      const mid = String(detailData[j][1]);
      if (!detailRowMap[mid]) detailRowMap[mid] = [];
      detailRowMap[mid].push({ rowIndex: j + 1, row: detailData[j] });
    }

    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i].join('').trim() === '') continue;

      // Filter by kategori
      const rowKategori = String(masterData[i][4] || '');
      if (kategori && rowKategori !== kategori) continue;

      // Filter by date range
      const tglRaw = masterData[i][1];
      const tglDate = tglRaw instanceof Date ? tglRaw : new Date(tglRaw);
      if (isNaN(tglDate.getTime()) || tglDate < dateFrom || tglDate > dateTo) continue;

      const masterId = String(masterData[i][0]);
      const items = detailRowMap[masterId] || [];
      let masterHasVariance = false;

      items.forEach(item => {
        totalItems++;
        const row = item.row;
        const m = parseFloat(row[6]) || 0;   // stockMabang
        const ttx = parseFloat(row[7]) || 0; // stockTtx
        const f = parseFloat(row[8]) || 0;   // stockFisik

        const newSelisihMabang = f - m;
        const newSelisihTtx = f - ttx;

        const oldSelisihMabang = parseFloat(row[9]) || 0;
        const oldSelisihTtx = parseFloat(row[10]) || 0;

        // Recalculate Staff if exists
        const fs = parseFloat(row[13]) || 0; // stockFisikStaff
        if (fs) {
          detailSheet.getRange(item.rowIndex, 15).setValue(fs - m); // col O
          detailSheet.getRange(item.rowIndex, 16).setValue(fs - ttx); // col P
        }

        // Recalculate and update if different
        if (newSelisihMabang !== oldSelisihMabang || newSelisihTtx !== oldSelisihTtx) {
          detailSheet.getRange(item.rowIndex, 10).setValue(newSelisihMabang); // col J = selisihMabang
          detailSheet.getRange(item.rowIndex, 11).setValue(newSelisihTtx);   // col K = selisihTtx
          detailUpdated++;
          masterHasVariance = true;
        } else if (newSelisihMabang !== 0 || newSelisihTtx !== 0) {
          masterHasVariance = true;
        }

        if (newSelisihMabang === 0 && newSelisihTtx === 0) correctItems++;
      });

      // Update master status based on variance
      if (items.length > 0) {
        const currentStatus = String(masterData[i][7] || '');
        let newStatus = masterHasVariance ? 'Menunggu Approval' : 'Selesai';
        // Don't downgrade already-approved records
        if (currentStatus === 'Disetujui' || currentStatus === 'Ditolak') newStatus = currentStatus;

        if (newStatus !== currentStatus) {
          masterSheet.getRange(i + 1, 8).setValue(newStatus);
          masterUpdated++;
        }
      }
    }

    const accuracy = totalItems > 0 ? ((correctItems / totalItems) * 100).toFixed(1) : '0.0';
    return {
      success: true,
      message: `${detailUpdated} item diperbarui, ${masterUpdated} laporan dikoreksi. Akurasi: ${accuracy}% (${correctItems}/${totalItems})`,
      detailUpdated,
      masterUpdated,
      totalItems,
      correctItems,
      accuracy: parseFloat(accuracy)
    };
  } catch (e) { return { success: false, message: e.message }; }
}
this['validateAndRecalculateStock'] = validateAndRecalculateStock;

/**
 * Recalculate variance for a single Stock Control record
 */
function recalculateSingleStockControl(masterId) {
  try {
    const masterSheet = getSheet(CONFIG.SHEETS.STOCK_CONTROL);
    const detailSheet = getSheet(CONFIG.SHEETS.STOCK_CONTROL_DETAIL);
    const masterData = masterSheet.getDataRange().getValues();
    const detailData = detailSheet.getDataRange().getValues();
    
    let hasVariance = false;
    let updatedCount = 0;
    let varianceChanged = false;
    
    for (let j = 1; j < detailData.length; j++) {
      if (String(detailData[j][1]) === String(masterId)) {
        const m = parseFloat(detailData[j][6]) || 0;   // stockMabang
        const ttx = parseFloat(detailData[j][7]) || 0; // stockTtx
        const f = parseFloat(detailData[j][8]) || 0;   // stockFisik
        
        const oldSm = parseFloat(detailData[j][9]) || 0;
        const oldSt = parseFloat(detailData[j][10]) || 0;
        
        const sm = f - m;
        const st = f - ttx;
        
        // Staff Recalculate
        const fs = parseFloat(detailData[j][13]) || 0;
        if (fs) {
          detailSheet.getRange(j + 1, 15).setValue(fs - m);
          detailSheet.getRange(j + 1, 16).setValue(fs - ttx);
        }

        if (sm !== 0 || st !== 0) hasVariance = true;
        if (sm !== oldSm || st !== oldSt) varianceChanged = true;
        
        detailSheet.getRange(j + 1, 10).setValue(sm); // col J
        detailSheet.getRange(j + 1, 11).setValue(st); // col K
        updatedCount++;
      }
    }
    
    
    // Update status and sync log
    let logSummary = '';
    for (let i = 1; i < masterData.length; i++) {
      if (String(masterData[i][0]) === String(masterId)) {
        const currentStatus = String(masterData[i][7] || '');
        const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
        
        let warnPrefix = '';
        if (varianceChanged) warnPrefix = '⚠️ PERINGATAN: DATA BERUBAH! ';
        
        logSummary = `🔄 Sinkronisasi: ${now}. ${warnPrefix}${hasVariance ? 'Ada Selisih' : 'Cocok'}.`;
        
        // Update Log (col K = index 10)
        masterSheet.getRange(i + 1, 11).setValue(logSummary);

        if (hasVariance && (currentStatus === 'Selesai' || currentStatus === 'Pending' || !currentStatus)) {
          masterSheet.getRange(i + 1, 8).setValue('Menunggu Approval');
        } else if (!hasVariance && (currentStatus === 'Menunggu Approval' || currentStatus === 'Pending' || !currentStatus)) {
          masterSheet.getRange(i + 1, 8).setValue('Selesai');
        }
        break;
      }
    }
    
    return { success: true, message: logSummary + ' Berhasil sinkronisasi ' + updatedCount + ' item.', varianceChanged: varianceChanged };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
this['recalculateSingleStockControl'] = recalculateSingleStockControl;

/**
 * Bulk Recalculate variance for multiple Stock Control records
 */
function bulkRecalculateStockControl(ids) {
  try {
    if (!ids || !Array.isArray(ids) || ids.length === 0) return { success: false, message: 'ID tidak valid' };
    
    const masterSheet = getSheet(CONFIG.SHEETS.STOCK_CONTROL);
    const detailSheet = getSheet(CONFIG.SHEETS.STOCK_CONTROL_DETAIL);
    const masterData = masterSheet.getDataRange().getValues();
    const detailData = detailSheet.getDataRange().getValues();
    
    const idSet = new Set(ids.map(String));
    let masterUpdated = 0;
    let detailUpdated = 0;
    
    // Map details by masterId
    const detailMap = {};
    for (let j = 1; j < detailData.length; j++) {
      const mid = String(detailData[j][1]);
      if (idSet.has(mid)) {
        if (!detailMap[mid]) detailMap[mid] = [];
        detailMap[mid].push({ rowIndex: j + 1, row: detailData[j] });
      }
    }
    
    ids.forEach(masterId => {
      const items = detailMap[masterId] || [];
      let hasVariance = false;
      
      items.forEach(item => {
        const m = parseFloat(item.row[6]) || 0;
        const ttx = parseFloat(item.row[7]) || 0;
        const f = parseFloat(item.row[8]) || 0;
        const sm = f - m;
        const st = f - ttx;
        
        if (sm !== 0 || st !== 0) hasVariance = true;
        
        detailSheet.getRange(item.rowIndex, 10).setValue(sm);
        detailSheet.getRange(item.rowIndex, 11).setValue(st);
        detailUpdated++;
      });
      
      // Update master status and sync log
      for (let i = 1; i < masterData.length; i++) {
        if (String(masterData[i][0]) === String(masterId)) {
          const currentStatus = String(masterData[i][7] || '');
          let newStatus = currentStatus;
          
          const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
          const logSummary = `🔄 Sinkronisasi: ${now}. ${hasVariance ? '⚠️ Ada Selisih' : '✅ Cocok'}.`;
          
          masterSheet.getRange(i + 1, 11).setValue(logSummary); // col K

          if (hasVariance && (currentStatus === 'Selesai' || currentStatus === 'Pending' || !currentStatus)) {
            newStatus = 'Menunggu Approval';
          } else if (!hasVariance && (currentStatus === 'Menunggu Approval' || currentStatus === 'Pending' || !currentStatus)) {
            newStatus = 'Selesai';
          }
          
          if (newStatus !== currentStatus) {
            masterSheet.getRange(i + 1, 8).setValue(newStatus);
            masterUpdated++;
          }
          break;
        }
      }
    });
    
    return { success: true, message: `Sinkronisasi selesai. ${detailUpdated} item diperbarui, ${masterUpdated} status laporan dikoreksi.` };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
this['bulkRecalculateStockControl'] = bulkRecalculateStockControl;



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
function addKlaim(tanggal, pic, resi, harga, keterangan, items, createdBy) {
  try {
    const klaimId = generateId();
    const sheetKlaim = getSheet(CONFIG.SHEETS.KLAIM);
    const klaimRow = [klaimId, tanggal, pic, resi, parseFloat(harga) || 0, keterangan, 'Pending', createdBy, new Date().toISOString()];
    sheetKlaim.appendRow(klaimRow);
    syncSheetRowToSupabase(CONFIG.SHEETS.KLAIM, klaimRow);
    
    // Simpan rincian SKU jika ada
    if (items && Array.isArray(items) && items.length > 0) {
      const sheetDetail = getSheet(CONFIG.SHEETS.KLAIM_DETAIL);
      items.forEach(item => {
        const detailRow = [generateId(), klaimId, item.sku, parseFloat(item.harga) || 0];
        sheetDetail.appendRow(detailRow);
        syncSheetRowToSupabase(CONFIG.SHEETS.KLAIM_DETAIL, detailRow);
      });
    }
    
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function getKlaimDetail(klaimId) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KLAIM_DETAIL);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(klaimId)) {
        result.push({
          sku: data[i][2],
          harga: parseFloat(data[i][3]) || 0
        });
      }
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateKlaim(id, tanggal, pic, resi, harga, keterangan, items, updatedBy) {
  try {
    const sheetKlaim = getSheet(CONFIG.SHEETS.KLAIM);
    const dataKlaim = sheetKlaim.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < dataKlaim.length; i++) {
      if (String(dataKlaim[i][0]) === String(id)) {
        sheetKlaim.getRange(i + 1, 2, 1, 5).setValues([[tanggal, pic, resi, parseFloat(harga) || 0, keterangan]]);
        found = true;
        break;
      }
    }
    if (!found) return { success: false, message: 'Data Klaim tidak ditemukan' };

    const sheetDetail = getSheet(CONFIG.SHEETS.KLAIM_DETAIL);
    const dataDetail = sheetDetail.getDataRange().getValues();
    for (let j = dataDetail.length - 1; j >= 1; j--) {
      if (String(dataDetail[j][1]) === String(id)) {
        sheetDetail.deleteRow(j + 1);
      }
    }

    if (items && Array.isArray(items) && items.length > 0) {
      items.forEach(item => {
        const detRow = [generateId(), id, item.sku, parseFloat(item.harga) || 0];
        sheetDetail.appendRow(detRow);
        syncSheetRowToSupabase(CONFIG.SHEETS.KLAIM_DETAIL, detRow);
      });
    }

    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}
function updateKlaimStatus(id, status, resiFallback) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.KLAIM); const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if ((id && String(data[i][0]) === String(id)) || (!id && resiFallback && String(data[i][3]) === String(resiFallback))) {
        sheet.getRange(i + 1, 7).setValue(status);
        syncRowUpdate(CONFIG.SHEETS.KLAIM, i + 1);
        return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}
function deleteKlaim(id) { return deleteRow(CONFIG.SHEETS.KLAIM, id); }

// ============================================================
// SOP & STRUKTUR ORGANISASI
// ============================================================
// SOP & STRUKTUR ORGANISASI
// ============================================================
function getSOP() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SOP); 
    const data = sheet.getDataRange().getValues(); 
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      result.push({ 
        id: data[i][0], 
        judul: data[i][1], 
        konten: data[i][2], 
        kategori: data[i][3], 
        createdBy: data[i][4],
        updatedAt: data[i][5] || '' // Column 6 for updatedAt
      });
    }
    
    return { success: true, data: result };
  } catch (e) { 
    return { success: false, message: e.message }; 
  }
}

function addSOP(judul, konten, kategori, createdBy) {
  try {
    const row = [generateId(), judul, konten, kategori, createdBy, new Date().toISOString()];
    getSheet(CONFIG.SHEETS.SOP).appendRow(row);
    syncSheetRowToSupabase(CONFIG.SHEETS.SOP, row);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
function updateSOP(id, judul, konten, kategori) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SOP); 
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) { 
      if (String(data[i][0]) === String(id)) { 
        sheet.getRange(i + 1, 2, 1, 4).setValues([[judul, konten, kategori, new Date().toISOString()]]); 
        return { success: true }; 
      } 
    }
    
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) { 
    return { success: false, message: e.message }; 
  }
}

function deleteSOP(id) { 
  return deleteRow(CONFIG.SHEETS.SOP, id); 
}
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
  try {
    const row = [generateId(), nama, jabatan, atasan, departemen, foto, urutan || 0];
    getSheet(CONFIG.SHEETS.ORGANISASI).appendRow(row);
    syncSheetRowToSupabase(CONFIG.SHEETS.ORGANISASI, row);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
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
    const row = [generateId(), sku, nama, barcode, batch, expDate, satuan, parseFloat(stok)||0, parseFloat(stokMin)||0, kategori, lokasi, now, now];
    getSheet(CONFIG.SHEETS.STOCK).appendRow(row);
    syncSheetRowToSupabase(CONFIG.SHEETS.STOCK, row);
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

function getBarangByCode(code) {
  try {
    if (!code) return { success: false, message: 'Kode kosong' };
    const stock = getStock();
    if (!stock.success) return stock;
    
    const results = stock.data.filter(s => 
      String(s.sku).toLowerCase() === String(code).toLowerCase() || 
      String(s.barcode).toLowerCase() === String(code).toLowerCase()
    );
    
    if (results.length === 0) return { success: false, message: 'Barang tidak ditemukan: ' + code };
    return { success: true, data: results };
  } catch(e) { return { success: false, message: e.message }; }
}

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
        syncSheetRowToSupabase(CONFIG.SHEETS.STOCK, newRow);
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
      result.push({ id: data[i][0], noSJ: data[i][1], tanggal: data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd') : String(data[i][2]||''), supplier: data[i][3], keterangan: data[i][4], createdBy: data[i][5], createdAt: data[i][6] instanceof Date ? data[i][6].toISOString() : String(data[i][6]||'') });
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
    const masterRow = [id, noSJ, tanggal, supplier, keterangan, createdBy, now];
    getSheet(CONFIG.SHEETS.SURAT_JALAN_MASUK).appendRow(masterRow);
    syncSheetRowToSupabase(CONFIG.SHEETS.SURAT_JALAN_MASUK, masterRow);
    const detSheet = getSheet(CONFIG.SHEETS.SURAT_JALAN_MASUK_DETAIL);
    const parsedItems = typeof items === 'string' ? JSON.parse(items) : items;
    parsedItems.forEach(item => {
      const detRow = [generateId(), id, noSJ, item.stockId, item.sku, item.nama, parseFloat(item.qty)||0, item.satuan, item.batch||'', item.expDate||'', item.lokasi||''];
      detSheet.appendRow(detRow);
      syncSheetRowToSupabase(CONFIG.SHEETS.SURAT_JALAN_MASUK_DETAIL, detRow);
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
      result.push({ id: data[i][0], noSJ: data[i][1], tanggal: data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd') : String(data[i][2]||''), tujuan: data[i][3], keterangan: data[i][4], createdBy: data[i][5], createdAt: data[i][6] instanceof Date ? data[i][6].toISOString() : String(data[i][6]||'') });
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
    const masterRow = [id, noSJ, tanggal, tujuan, keterangan, createdBy, new Date().toISOString()];
    getSheet(CONFIG.SHEETS.SURAT_JALAN_KELUAR).appendRow(masterRow);
    syncSheetRowToSupabase(CONFIG.SHEETS.SURAT_JALAN_KELUAR, masterRow);
    const detSheet = getSheet(CONFIG.SHEETS.SURAT_JALAN_KELUAR_DETAIL);
    const parsedItems = typeof items === 'string' ? JSON.parse(items) : items;
    for (const item of parsedItems) {
      const res = updateStokQty(item.stockId, -(parseFloat(item.qty)||0), item.sku);
      if (!res.success) return { success: false, message: res.message };
      const detRow = [generateId(), id, noSJ, item.stockId, item.sku, item.nama, parseFloat(item.qty)||0, item.satuan, item.batch||'', item.expDate||'', item.lokasi||''];
      detSheet.appendRow(detRow);
      syncSheetRowToSupabase(CONFIG.SHEETS.SURAT_JALAN_KELUAR_DETAIL, detRow);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

// ============================================================
// ORDER & RETUR
// ============================================================
function parseDistributorQueueDate(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) return new Date(value.getTime());

  const str = String(value).trim();
  if (!str) return null;

  let match = str.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T\s](\d{2}):(\d{2})(?::(\d{2}))?)?$/);
  if (match) {
    return new Date(
      Number(match[1]),
      Number(match[2]) - 1,
      Number(match[3]),
      Number(match[4] || 0),
      Number(match[5] || 0),
      Number(match[6] || 0)
    );
  }

  match = str.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:[T\s](\d{2}):(\d{2})(?::(\d{2}))?)?$/);
  if (match) {
    return new Date(
      Number(match[3]),
      Number(match[2]) - 1,
      Number(match[1]),
      Number(match[4] || 0),
      Number(match[5] || 0),
      Number(match[6] || 0)
    );
  }

  const parsed = new Date(str);
  return isNaN(parsed) ? null : parsed;
}

function startOfDayLocal(dateObj) {
  return new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
}

function addDaysLocal(dateObj, days) {
  const next = new Date(dateObj.getTime());
  next.setDate(next.getDate() + days);
  return next;
}

function diffDaysLocal(baseDate, compareDate) {
  const ms = startOfDayLocal(compareDate).getTime() - startOfDayLocal(baseDate).getTime();
  return Math.floor(ms / 86400000);
}

function isSameDayLocal(dateA, dateB) {
  return startOfDayLocal(dateA).getTime() === startOfDayLocal(dateB).getTime();
}

function getStartOfWeekLocal(dateObj) {
  const current = startOfDayLocal(dateObj);
  const day = current.getDay();
  const diff = day === 0 ? -6 : 1 - day; // Senin sebagai awal minggu
  current.setDate(current.getDate() + diff);
  return current;
}

function isWithinCurrentWeekLocal(targetDate, referenceDate) {
  const weekStart = getStartOfWeekLocal(referenceDate);
  const weekEnd = addDaysLocal(weekStart, 6);
  const target = startOfDayLocal(targetDate);
  return target.getTime() >= weekStart.getTime() && target.getTime() <= weekEnd.getTime();
}

function isWithinCurrentMonthLocal(targetDate, referenceDate) {
  return targetDate.getFullYear() === referenceDate.getFullYear() && targetDate.getMonth() === referenceDate.getMonth();
}

function formatDistributorQueueDate(value, withTime) {
  const dateObj = parseDistributorQueueDate(value);
  if (!dateObj) return String(value || '');
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone() || 'Asia/Jakarta', withTime ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd');
}

function evaluateDistributorQueueSLA(item, slaSettings) {
  if (!slaSettings) {
    slaSettings = getDistributorQueueSLASettings().data || { dueDays: 1, ruleDescription: 'SLA H+1 dari Order queue time' };
  }
  const orderDate = parseDistributorQueueDate(item.orderQueueTime) || parseDistributorQueueDate(item.timeWib);
  if (!orderDate) {
    return {
      status: 'Tanggal order kosong',
      isLate: false,
      lateDays: 0,
      dueDate: '',
      completionDate: '',
      completionSource: '',
      ruleDescription: slaSettings.ruleDescription
    };
  }

  const dueDate = addDaysLocal(startOfDayLocal(orderDate), Number(slaSettings.dueDays || 1));
  const packingDate = parseDistributorQueueDate(item.tanggalSelesaiPacking);
  const completionDate = packingDate;
  const today = startOfDayLocal(new Date());

  let isLate = false;
  let lateDays = 0;
  let status = 'Pending';
  let completionSource = '';

  if (completionDate) {
    const finalDate = startOfDayLocal(completionDate);
    isLate = finalDate.getTime() > dueDate.getTime();
    lateDays = isLate ? diffDaysLocal(dueDate, finalDate) : 0;
    status = isLate ? 'Late' : 'On Time';
    completionSource = 'Tanggal selesai packing';
  } else if (today.getTime() > dueDate.getTime()) {
    isLate = true;
    lateDays = diffDaysLocal(dueDate, today);
    status = 'Late';
  }

  return {
    status: status,
    isLate: isLate,
    lateDays: lateDays,
    dueDate: formatDistributorQueueDate(dueDate, false),
    completionDate: completionDate ? formatDistributorQueueDate(completionDate, false) : '',
    completionSource: completionSource,
    ruleDescription: slaSettings.ruleDescription
  };
}

function getDistributorQueueSLASettings() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SETTINGS);
    const data = sheet.getDataRange().getValues();
    const settings = {
      dueDays: 1,
      ruleDescription: 'SLA H+1 dari Order queue time'
    };
    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0] || '').trim();
      const value = data[i][1];
      if (!key) continue;
      if (key === 'distributorQueueSlaDays') settings.dueDays = parseInt(value, 10) || 1;
      if (key === 'distributorQueueSlaRule') settings.ruleDescription = String(value || settings.ruleDescription);
    }
    return { success: true, data: settings };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getDistributorQueueLateNotesMap() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SETTINGS);
    const data = sheet.getDataRange().getValues();
    const map = {};
    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0] || '').trim();
      const value = String(data[i][1] || '');
      if (key.indexOf('distributorQueueLateNote:') === 0) {
        map[key.replace('distributorQueueLateNote:', '')] = value;
      }
    }
    return map;
  } catch (e) {
    return {};
  }
}

function saveDistributorQueueSLASettings(settings) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SETTINGS);
    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();
    const upsert = (key, val) => {
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0] || '') === key) {
          sheet.getRange(i + 1, 2).setValue(val);
          sheet.getRange(i + 1, 3).setValue(now);
          found = true;
          break;
        }
      }
      if (!found) sheet.appendRow([key, val, now]);
    };
    upsert('distributorQueueSlaDays', Number(settings.dueDays) || 1);
    upsert('distributorQueueSlaRule', String(settings.ruleDescription || 'SLA H+1 dari Order queue time'));
    return { success: true, message: 'Aturan SLA berhasil disimpan.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function saveDistributorQueueLateNote(poNumber, rowNumber, note, sourceSheet, createdBy) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.LATE_SHIPMENT);
    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();
    
    // Cari data existing berdasarkan PO Number dan Source Sheet
    let found = false;
    let existingId = null;
    let existingStatus = 'Pending Team Leader';
    let existingHistory = [];
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1] || '') === String(poNumber).trim() && 
          String(data[i][2] || '') === String(sourceSheet).trim()) {
        existingId = data[i][0];
        existingStatus = data[i][5] || 'Pending Team Leader';
        try {
          existingHistory = JSON.parse(data[i][8] || '[]');
        } catch(e) {
          existingHistory = [];
        }
        
        // Update keterangan
        sheet.getRange(i + 1, 5).setValue(note || ''); // kolom keterangan
        
        // Tambah history
        existingHistory.push({
          date: now,
          action: 'Update Keterangan',
          by: createdBy,
          keterangan: note
        });
        sheet.getRange(i + 1, 9).setValue(JSON.stringify(existingHistory)); // kolom history
        
        found = true;
        break;
      }
    }
    
    // Jika belum ada, buat baru
    if (!found) {
      const id = generateId();
      const history = [{
        date: now,
        action: 'Create',
        by: createdBy,
        keterangan: note
      }];
      const lateRow = [
        id, poNumber, sourceSheet, rowNumber,
        note || '', 'Pending Team Leader', createdBy,
        now, JSON.stringify(history)
      ];
      sheet.appendRow(lateRow);
      syncSheetRowToSupabase(CONFIG.SHEETS.LATE_SHIPMENT, lateRow);
    }
    
    return { success: true, message: 'Keterangan late shipment berhasil disimpan.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getLateShipmentData() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.LATE_SHIPMENT);
    const data = sheet.getDataRange().getValues();
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      
      let history = [];
      try {
        history = JSON.parse(data[i][8] || '[]');
      } catch(e) {
        history = [];
      }
      
      result.push({
        id: String(data[i][0] || ''),
        poNumber: String(data[i][1] || ''),
        sourceSheet: String(data[i][2] || ''),
        rowNumber: String(data[i][3] || ''),
        keterangan: String(data[i][4] || ''),
        status: String(data[i][5] || ''),
        createdBy: String(data[i][6] || ''),
        createdAt: data[i][7] instanceof Date ? data[i][7].toISOString() : String(data[i][7] || ''),
        history: history
      });
    }
    
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function approveLateShipment(id, action, userNama, userRole, reason) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.LATE_SHIPMENT);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        let currentStatus = String(data[i][5] || 'Pending Team Leader');
        
        // Authorization check
        const isAdmin = (userRole === 'admin' || userRole === 'Super Admin');
        const isTL = (userRole === 'Team Leader' || userRole === 'TL' || userRole.includes('Team Leader'));
        const isVice = (userRole === 'Vice Supervisor' || userRole === 'Vice SPV' || userRole.includes('Vice'));
        const isSPV = (userRole === 'Supervisor' || userRole === 'SPV' || (userRole.includes('Supervisor') && !userRole.includes('Vice')));
        const isSales = (userRole === 'Sales' || userRole.includes('Sales'));
        
        let authorized = isAdmin;
        if (currentStatus === 'Pending Team Leader' && (isTL || isAdmin)) authorized = true;
        if (currentStatus === 'Pending Vice Supervisor' && (isVice || isAdmin)) authorized = true;
        if (currentStatus === 'Pending Supervisor' && (isSPV || isAdmin)) authorized = true;
        if (currentStatus === 'Pending Sales' && (isSales || isAdmin)) authorized = true;
        
        if (!authorized) {
          return { success: false, message: 'Anda tidak memiliki wewenang untuk tahap approval ini (' + currentStatus + ').' };
        }
        
        // Determine new status
        let newStatus = '';
        if (action === 'Reject') {
          newStatus = 'Ditolak';
        } else if (action === 'Approve') {
          if (isAdmin) {
            newStatus = 'Disetujui';
          } else if (currentStatus === 'Pending Team Leader') {
            newStatus = 'Pending Vice Supervisor';
          } else if (currentStatus === 'Pending Vice Supervisor') {
            newStatus = 'Pending Supervisor';
          } else if (currentStatus === 'Pending Supervisor') {
            newStatus = 'Pending Sales';
          } else if (currentStatus === 'Pending Sales') {
            newStatus = 'Disetujui';
          } else {
            newStatus = 'Disetujui';
          }
        }
        
        // Update status
        sheet.getRange(i + 1, 6).setValue(newStatus);
        
        // Update history
        let history = [];
        try {
          history = JSON.parse(data[i][8] || '[]');
        } catch(e) {
          history = [];
        }
        
        history.push({
          date: new Date().toISOString(),
          action: action,
          status: newStatus,
          by: userNama,
          role: userRole,
          reason: reason || ''
        });
        
        sheet.getRange(i + 1, 9).setValue(JSON.stringify(history));
        
        return { success: true, newStatus: newStatus, message: 'Approval berhasil diproses.' };
      }
    }
    
    return { success: false, message: 'Data tidak ditemukan.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getDistributorQueueLateNotesMap() {
  try {
    const lateData = getLateShipmentData();
    if (!lateData.success) return {};
    
    const map = {};
    lateData.data.forEach(function(item) {
      const key = item.poNumber + '|' + item.sourceSheet;
      map[key] = {
        keterangan: item.keterangan,
        status: item.status,
        id: item.id
      };
    });
    
    return map;
  } catch (e) {
    console.error('Error getting late notes map:', e);
    return {};
  }
}

function mapDistributorQueueRow(row, rowNumber) {
  try {
    const item = {
      rowNumber: rowNumber,
      no: String(row[0] || ''), // Kolom No
      orderQueueTime: formatDistributorQueueDate(row[1], false),
      picSales: String(row[2] || ''),
      namaDistributor: String(row[3] || ''),
      alamat: String(row[4] || ''),
      noHp: String(row[5] || ''),
      poNumber: String(row[6] || ''),
      noMabang: String(row[7] || ''),
      metodePengiriman: String(row[8] || ''),
      ongkirDibayarOleh: String(row[9] || ''),
      note: String(row[10] || ''),
      timeWib: formatDistributorQueueDate(row[11], true),
      statusGudang: String(row[12] || ''),
      jumlahDus: String(row[13] || ''),
      totalPcs: String(row[14] || ''),
      packer: String(row[15] || ''),
      validation: String(row[16] || ''),
      tanggalSelesaiPacking: formatDistributorQueueDate(row[17], false),
      shipDate: formatDistributorQueueDate(row[18], false),
      statusMabang: String(row[19] || ''),
      gdrive: String(row[20] || ''),
      deliveryBill: String(row[21] || ''),
      nomorResi: String(row[22] || ''),
      buktiPengiriman: String(row[23] || '')
    };
    
    // OPTIMASI: Skip SLA evaluation untuk loading lebih cepat
    // item.sla = evaluateDistributorQueueSLA(item);
    item.sla = {
      status: 'Pending',
      isLate: false,
      dueDate: '',
      completionDate: '',
      daysRemaining: 0
    };
    
    return item;
  } catch (e) {
    console.error('Error mapping row ' + rowNumber + ':', e);
    return null;
  }
}

function buildDistributorQueueRowValues(payload) {
  return [
    payload.no || '',
    payload.orderQueueTime || '',
    payload.picSales || '',
    payload.namaDistributor || '',
    payload.alamat || '',
    payload.noHp || '',
    payload.poNumber || '',
    payload.noMabang || '',
    payload.metodePengiriman || '',
    payload.ongkirDibayarOleh || '',
    payload.note || '',
    payload.timeWib || '',
    payload.statusGudang || '',
    payload.jumlahDus || '',
    payload.totalPcs || '',
    payload.packer || '',
    payload.validation || '',
    payload.tanggalSelesaiPacking || '',
    payload.shipDate || '',
    payload.statusMabang || '',
    payload.gdrive || '',
    payload.deliveryBill || '',
    payload.nomorResi || '',
    payload.buktiPengiriman || '',
    payload.hargaOngkir || ''
  ];
}


// ============================================================
// STATUS HELPERS - Antrian Distributor Column L
// ============================================================
function normalizeStatus(s) {
  // Lowercase, trim, remove extra spaces
  return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function isStatusTerkirim(s) {
  var n = normalizeStatus(s).replace(/\s/g, '');
  return n === 'terkirim' || n.includes('terkirim') || n.includes('shipped') || n.includes('dikirim') || n.includes('sent');
}

function isStatusReadyPickup(s) {
  var n = normalizeStatus(s).replace(/\s/g, '');
  // Matches: "Ready Pickup", "Ready To Pickup", "Ready", "Pickup", "Siap", "Siap Pickup", etc.
  return n === 'readypickup' || n === 'ready' || n === 'pickup' ||
         n.includes('readypickup') || n.includes('readytopickup') ||
         n.includes('ready') || n.includes('pickup') || n.includes('siap');
}

// Debug: call this from Apps Script editor to test column L values
function debugDistributorQueueStatus() {
  var sheet = getDistributorQueueSheet();
  var data = sheet.getDataRange().getValues();
  var results = [];
  for (var i = 1; i < data.length && i <= 20; i++) {
    var row = data[i];
    if (row.join('').trim() === '') continue;
    var raw = String(row[11] || '');
    results.push({
      row: i + 1,
      poNumber: String(row[5] || ''),
      statusRaw: raw,
      statusNorm: normalizeStatus(raw),
      isTerkirim: isStatusTerkirim(raw),
      isReadyPickup: isStatusReadyPickup(raw)
    });
  }
  Logger.log(JSON.stringify(results, null, 2));
  return results;
}

// Fast dashboard-only endpoint (no SLA evaluation, no note map) for quick first paint
function getDistributorQueueDashboardFast() {
  try {
    const ss = getDistributorQueueSpreadsheet();
    const allSheets = [
      CONFIG.SHEETS.DISTRIBUTOR_QUEUE,
      CONFIG.SHEETS.DISTRIBUTOR_QUEUE_FOCALSKIN,
      CONFIG.SHEETS.DISTRIBUTOR_QUEUE_MISTINE,
      CONFIG.SHEETS.DISTRIBUTOR_QUEUE_SBY
    ];
    
    var total = 0, terkirim = 0, readyPickup = 0, selesai = 0,
        belumSelesai = 0, belumDikerjakan = 0, late = 0;
    var today = startOfDayLocal(new Date());
    var poHariIni = 0, poMingguIni = 0, poBulanIni = 0;

    allSheets.forEach(function(sheetName) {
      try {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) return;
        
        const data = sheet.getDataRange().getValues();

        for (var i = 1; i < data.length; i++) {
          var row = data[i];
          if (row.join('').trim() === '') continue;
          
          total++;

          var statusGudang = String(row[12] || '').toLowerCase().trim(); // Kolom M (index 12) - Status
          var shipDate     = row[18] ? String(row[18]).trim() : ''; // Kolom S (index 18) - Ship date
          var packingDate  = row[17] ? String(row[17]).trim() : ''; // Kolom R (index 17) - Tanggal selesai packing
          var orderDate    = parseDistributorQueueDate(row[1]) || parseDistributorQueueDate(row[11]); // Kolom B (index 1) - Order queue time, Kolom L (index 11) - Time

          // totalTerkirim from Status column M
          if (isStatusTerkirim(statusGudang)) terkirim++;
          // totalReadyToPickup from Status column M
          if (isStatusReadyPickup(statusGudang)) readyPickup++;

          // selesai / belumSelesai / belumDikerjakan
          if (shipDate || packingDate) {
            selesai++;
          } else {
            var sNorm = statusGudang.replace(/\s+/g, '');
            // Belum Dikerjakan: kolom M kosong
            if (statusGudang === '') {
              belumDikerjakan++;
            // Belum Selesai: kolom M = Picking
            } else if (sNorm.includes('picking')) {
              belumSelesai++;
            }
          }

          // PO counts
          if (orderDate) {
            if (isSameDayLocal(orderDate, today)) poHariIni++;
            if (isWithinCurrentWeekLocal(orderDate, today)) poMingguIni++;
            if (isWithinCurrentMonthLocal(orderDate, today)) poBulanIni++;
          }
        }
      } catch (err) {
        console.error('Error reading sheet ' + sheetName + ' in fast dashboard:', err);
      }
    });

    return {
      success: true,
      dashboard: {
        total: total,
        selesai: selesai,
        belumSelesai: belumSelesai,
        belumDikerjakan: belumDikerjakan,
        late: 0, // will be filled by full load
        totalTerkirim: terkirim,
        totalReadyToPickup: readyPickup,
        poKeluarHariIni: poHariIni,
        poKeluarMingguIni: poMingguIni,
        poKeluarBulanIni: poBulanIni,
        lateItems: [],
        pending: 0,
        onTime: 0,
        dueToday: 0
      }
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * OPTIMIZED untuk 30,000+ baris data
 * Load data terbaru dari semua sheet secara merata
 */
function getDistributorQueueData() {
  try {
    const ss = getDistributorQueueSpreadsheet();
    const sheets = [
      { name: CONFIG.SHEETS.DISTRIBUTOR_QUEUE, label: 'Antrian Distributor' },
      { name: CONFIG.SHEETS.DISTRIBUTOR_QUEUE_FOCALSKIN, label: 'ANTRIAN FOCALSKIN' },
      { name: CONFIG.SHEETS.DISTRIBUTOR_QUEUE_MISTINE, label: 'ANTRIAN MISTINE' },
      { name: CONFIG.SHEETS.DISTRIBUTOR_QUEUE_SBY, label: 'ANTRIAN SBY' }
    ];
    
    const rows = [];
    const MAX_ITEMS_PER_SHEET = 150; // 150 items per sheet = 600 total (4 sheets)
    
    console.log('Loading max', MAX_ITEMS_PER_SHEET, 'items per sheet from', sheets.length, 'sheets');
    
    // Load Late Shipment notes map
    const lateNotesMap = {};
    const slaSettings = getDistributorQueueSLASettings().data;
    try {
      const lateSheet = getSheet(CONFIG.SHEETS.LATE_SHIPMENT);
      const lateData = lateSheet.getDataRange().getValues();
      for (let i = 1; i < lateData.length; i++) {
        const po = String(lateData[i][1] || '').trim();
        const source = String(lateData[i][2] || '').trim();
        if (po) {
          lateNotesMap[po + '|' + source] = {
            id: lateData[i][0],
            note: lateData[i][4],
            status: lateData[i][5]
          };
        }
      }
    } catch (err) { console.error('Error loading late notes:', err); }

    // Baca dari semua sheet dengan jatah yang sama
    for (let s = 0; s < sheets.length; s++) {
      const sheetConfig = sheets[s];
      let sheetItemCount = 0;
      
      try {
        const sheet = ss.getSheetByName(sheetConfig.name);
        if (!sheet) {
          console.log('Sheet not found:', sheetConfig.name);
          continue;
        }
        
        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) {
          console.log('Sheet empty:', sheetConfig.name);
          continue;
        }
        
        // Hitung berapa baris yang perlu dibaca (max 150 per sheet)
        const startRow = Math.max(2, lastRow - MAX_ITEMS_PER_SHEET + 1);
        const numRows = lastRow - startRow + 1;
        
        console.log('Sheet:', sheetConfig.name, 'Total rows:', lastRow, 'Reading rows', startRow, 'to', lastRow);
        
        // Baca hanya range yang diperlukan
        const data = sheet.getRange(startRow, 1, numRows, 24).getValues();
        
        // Proses dari bawah ke atas (data terbaru)
        for (let i = data.length - 1; i >= 0; i--) {
          if (sheetItemCount >= MAX_ITEMS_PER_SHEET) break;
          if (data[i].join('').trim() === '') continue;
          
          try {
            const poNum = String(data[i][6] || '');
            const sourceLabel = sheetConfig.label;
            const lateNote = lateNotesMap[poNum + '|' + sourceLabel] || { id: '', note: '', status: '' };

            const item = {
              rowNumber: startRow + i,
              sourceSheet: sourceLabel,
              no: String(data[i][0] || ''),
              orderQueueTime: formatDistributorQueueDate(data[i][1], false),
              picSales: String(data[i][2] || ''),
              namaDistributor: String(data[i][3] || ''),
              alamat: String(data[i][4] || ''),
              noHp: String(data[i][5] || ''),
              poNumber: poNum,
              noMabang: String(data[i][7] || ''),
              metodePengiriman: String(data[i][8] || ''),
              ongkirDibayarOleh: String(data[i][9] || ''),
              note: String(data[i][10] || ''),
              timeWib: formatDistributorQueueDate(data[i][11], true),
              statusGudang: String(data[i][12] || ''),
              jumlahDus: String(data[i][13] || ''),
              totalPcs: String(data[i][14] || ''),
              packer: String(data[i][15] || ''),
              validation: String(data[i][16] || ''),
              tanggalSelesaiPacking: formatDistributorQueueDate(data[i][17], false),
              shipDate: formatDistributorQueueDate(data[i][18], false),
              statusMabang: String(data[i][19] || ''),
              gdrive: String(data[i][20] || ''),
              deliveryBill: String(data[i][21] || ''),
              nomorResi: String(data[i][22] || ''),
              buktiPengiriman: String(data[i][23] || ''),
              catatanLate: lateNote.note || '',
              catatanLateStatus: lateNote.status || '',
              catatanLateId: lateNote.id || '',
              sla: { status: 'Pending', isLate: false, dueDate: '', completionDate: '', daysRemaining: 0 }
            };
            
            // Calculate SLA
              item.sla = evaluateDistributorQueueSLA(item, slaSettings);
            
            rows.push(item);
            sheetItemCount++;
          } catch (err) {
            console.error('Error mapping row:', err);
          }
        }
        
        console.log('Sheet:', sheetConfig.name, 'Loaded', sheetItemCount, 'items');
        
      } catch (err) {
        console.error('Error reading sheet ' + sheetConfig.name + ':', err);
      }
    }
    
    console.log('Total items loaded from all sheets:', rows.length);
    
    // Sort by order queue time
    rows.sort(function(a, b) {
      const dateA = parseDistributorQueueDate(a.orderQueueTime) || new Date(0);
      const dateB = parseDistributorQueueDate(b.orderQueueTime) || new Date(0);
      return dateB.getTime() - dateA.getTime();
    });
    
    // Hitung dashboard
    const today = startOfDayLocal(new Date());
    const lateItems = rows.filter(function(item) { return item.sla && item.sla.isLate; });
    
    const dashboard = {
      total: rows.length,
      selesai: rows.filter(function(item) { return !!(item.shipDate || item.tanggalSelesaiPacking); }).length,
      belumSelesai: rows.filter(function(item) { return String(item.statusGudang || '').toLowerCase().includes('picking'); }).length,
      belumDikerjakan: rows.filter(function(item) { return String(item.statusGudang || '').trim() === ''; }).length,
      totalReadyToPickup: rows.filter(function(item) { 
        var s = String(item.statusGudang || '').toLowerCase();
        return s.includes('ready') || s.includes('pickup') || s.includes('siap');
      }).length,
      totalTerkirim: rows.filter(function(item) { 
        var s = String(item.statusGudang || '').toLowerCase();
        return s.includes('terkirim') || s.includes('delivered');
      }).length,
      poKeluarHariIni: rows.filter(function(item) {
        const orderDate = parseDistributorQueueDate(item.orderQueueTime);
        return orderDate ? isSameDayLocal(orderDate, today) : false;
      }).length,
      poKeluarMingguIni: rows.filter(function(item) {
        const orderDate = parseDistributorQueueDate(item.orderQueueTime);
        return orderDate ? isWithinCurrentWeekLocal(orderDate, today) : false;
      }).length,
      poKeluarBulanIni: rows.filter(function(item) {
        const orderDate = parseDistributorQueueDate(item.orderQueueTime);
        return orderDate ? isWithinCurrentMonthLocal(orderDate, today) : false;
      }).length,
      pending: rows.filter(function(item) { return item.sla && item.sla.status === 'Pending'; }).length,
      onTime: rows.filter(function(item) { return item.sla && item.sla.status === 'On Time'; }).length,
      late: lateItems.length,
      dueToday: 0,
      lateItems: lateItems
    };
    
    console.log('Success! Returning', rows.length, 'items');
    console.log('Dashboard:', JSON.stringify(dashboard));
    return { success: true, data: rows, dashboard: dashboard };
    
  } catch (e) {
    console.error('FATAL ERROR in getDistributorQueueData:', e);
    return { success: false, message: e.message, data: [] };
  }
}

function saveDistributorQueue(payload, updatedBy) {
  try {
    const parsed = typeof payload === 'string' ? JSON.parse(payload) : (payload || {});
    
    // Auto-numbering: Jika kolom "No" kosong, isi dengan nomor urut otomatis
    if (!parsed.no || parsed.no.trim() === '') {
      const sheet = getDistributorQueueSheet();
      const lastRow = sheet.getLastRow();
      parsed.no = String(lastRow); // Nomor urut berdasarkan baris terakhir
    }
    
    const rowValues = buildDistributorQueueRowValues(parsed);
    const sheet = getDistributorQueueSheet();
    const rowNumber = Number(parsed.rowNumber || 0);

    if (rowNumber > 1 && rowNumber <= sheet.getLastRow()) {
      sheet.getRange(rowNumber, 1, 1, DISTRIBUTOR_QUEUE_HEADERS.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      message: rowNumber > 1 ? 'Antrian distributor berhasil diperbarui.' : 'Antrian distributor berhasil ditambahkan.',
      updatedBy: updatedBy || ''
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getLateShipmentDashboard() {
  const result = getDistributorQueueData();
  if (!result.success) return result;
  return { success: true, dashboard: result.dashboard };
}

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
        tanggal: data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd') : String(data[i][2]||''),
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

function getOrdersBySku(sku) {
  try {
    sku = String(sku || '').toLowerCase().trim();
    if (!sku) return { success: true, data: [] };

    const ordRes = getOrders();
    if (!ordRes.success) return ordRes;

    const ordMap = {};
    ordRes.data.forEach(o => { ordMap[String(o.id)] = o; });

    const detSheet = getSheet(CONFIG.SHEETS.ORDER_DETAIL);
    if (!detSheet) return { success: true, data: [] };
    const data = detSheet.getDataRange().getValues();
    const matched = {};

    for (let i = 1; i < data.length; i++) {
      if (!data[i] || data[i].join('').trim() === '') continue;
      const orderId = String(data[i][1] || '');
      const skuCell = String(data[i][4] || '').toLowerCase();
      const nameCell = String(data[i][5] || '').toLowerCase();
      if ((skuCell && skuCell.indexOf(sku) !== -1) || (nameCell && nameCell.indexOf(sku) !== -1)) {
        matched[orderId] = true;
      }
    }

    const result = [];
    Object.keys(matched).forEach(id => { if (ordMap[id]) result.push(ordMap[id]); });
    // Return newest first
    return { success: true, data: result.reverse() };
  } catch (e) { return { success: false, message: e.message }; }
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
    
    const orderRow = [
      id, finalNoOrder, tanggal, pelanggan, alamat, 'Pending', totalItem, keterangan, createdBy, 
      new Date().toISOString(), '', '', kategori || 'Distributor', finalNoResi || ''
    ];
    getSheet(CONFIG.SHEETS.ORDER).appendRow(orderRow);
    syncSheetRowToSupabase(CONFIG.SHEETS.ORDER, orderRow);

    const detSheet = getSheet(CONFIG.SHEETS.ORDER_DETAIL);
    parsedItems.forEach(item => {
      const detRow = [generateId(), id, finalNoOrder, item.stockId, item.sku, item.nama, parseFloat(item.qty)||0, item.satuan, item.batch||'', item.expDate||'', 0, item.lokasi || ''];
      detSheet.appendRow(detRow);
      syncSheetRowToSupabase(CONFIG.SHEETS.ORDER_DETAIL, detRow);
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
          tanggal: orderData[i][2] instanceof Date ? Utilities.formatDate(orderData[i][2], SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd') : String(orderData[i][2]||''),
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

/**
 * Memperbarui jumlah barang yang sudah dipacking/divalidasi
 * @param {string} orderId ID Order
 * @param {Array|string} validationItems Array [{id: detailId, packedQty: number}]
 */
/**
 * FINALISASI VALIDASI & KIRIM ORDER (All-in-One Optimization)
 * Menggabungkan update packed qty, status order, bukti packing, dan potong stok dalam 1 call.
 */
function finalizeOrderValidation(orderId, noOrder, validationDataJson, photoUrl) {
  try {
    const orderSheet = getSheet(CONFIG.SHEETS.ORDER);
    const detailSheet = getSheet(CONFIG.SHEETS.ORDER_DETAIL);
    const stockSheet = getSheet(CONFIG.SHEETS.STOCK);
    
    const orderData = orderSheet.getDataRange().getValues();
    const detailData = detailSheet.getDataRange().getValues();
    const stockData = stockSheet.getDataRange().getValues();
    
    const items = typeof validationDataJson === 'string' ? JSON.parse(validationDataJson) : validationDataJson;
    const now = new Date().toISOString();

    // 1. Update Order Detail (Packed Qty)
    for (let i = 1; i < detailData.length; i++) {
      const detailId = String(detailData[i][0]);
      const match = items.find(it => String(it.id) === detailId);
      if (match) {
        detailSheet.getRange(i + 1, 11).setValue(parseFloat(match.packedQty) || 0);
      }
    }

    // 2. Cari & Update Order Utama
    let orderRow = -1;
    for (let i = 1; i < orderData.length; i++) {
      if (String(orderData[i][0]) === String(orderId) || String(orderData[i][1]) === String(noOrder)) {
        orderRow = i + 1;
        break;
      }
    }
    if (orderRow === -1) throw new Error('Order tidak ditemukan');

    orderSheet.getRange(orderRow, 6).setValue('Terkirim'); // Status
    orderSheet.getRange(orderRow, 11).setValue(now); // SentAt
    if (photoUrl) orderSheet.getRange(orderRow, 12).setValue(photoUrl); // BuktiPacking

    // 3. Potong Stok
    const rowId = orderData[orderRow-1][0];
    const rowNoOrder = orderData[orderRow-1][1];
    
    for (let i = 1; i < detailData.length; i++) {
      const match = (rowId && String(detailData[i][1]) === String(rowId)) || (rowNoOrder && String(detailData[i][2]) === String(rowNoOrder));
      if (match) {
        const stockId = String(detailData[i][3]);
        const skuFallback = String(detailData[i][4]);
        const qtyToDeduct = parseFloat(detailData[i][6]) || 0;
        
        // Cari baris di stockSheet
        for (let j = 1; j < stockData.length; j++) {
          if ((stockId && String(stockData[j][0]) === stockId) || (!stockId && skuFallback && String(stockData[j][1]) === skuFallback)) {
            const currentStok = parseFloat(stockData[j][7]) || 0;
            stockSheet.getRange(j + 1, 8).setValue(currentStok - qtyToDeduct);
            stockSheet.getRange(j + 1, 13).setValue(now);
            break;
          }
        }
      }
    }

    return { success: true };
  } catch (e) {
    return { success: false, message: 'Finalize Error: ' + e.message };
  }
}

function updatePackedQty(orderId, validationItems) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ORDER_DETAIL);
    const data = sheet.getDataRange().getValues();
    const items = typeof validationItems === 'string' ? JSON.parse(validationItems) : validationItems;
    
    let updatedCount = 0;
    for (let i = 1; i < data.length; i++) {
      const detailId = String(data[i][0]);
      const match = items.find(it => String(it.id) === detailId);
      
      if (match) {
        sheet.getRange(i + 1, 11).setValue(parseFloat(match.packedQty) || 0); // Column K
        updatedCount++;
      }
    }
    return { success: true, count: updatedCount };
  } catch (e) { return { success: false, message: e.message }; }
}

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
        tanggal: data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd') : String(data[i][2]||''),
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
    const masterRow = [id, noRetur, tanggal, sumber, alasan, keterangan, createdBy, new Date().toISOString()];
    getSheet(CONFIG.SHEETS.RETUR).appendRow(masterRow);
    syncSheetRowToSupabase(CONFIG.SHEETS.RETUR, masterRow);
    const detSheet = getSheet(CONFIG.SHEETS.RETUR_DETAIL);
    const parsedItems = typeof items === 'string' ? JSON.parse(items) : items;
    for (const item of parsedItems) {
      const res = updateStokQty(item.stockId, parseFloat(item.qty)||0, item.sku);
      if (!res.success) return { success: false, message: res.message };
      const detRow = [generateId(), id, noRetur, item.stockId, item.sku, item.nama, parseFloat(item.qty)||0, item.satuan, item.batch||'', item.expDate||''];
      detSheet.appendRow(detRow);
      syncSheetRowToSupabase(CONFIG.SHEETS.RETUR_DETAIL, detRow);
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
        tanggalMulai: data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd') : String(data[i][5]||''),
        deadline: data[i][6] instanceof Date ? Utilities.formatDate(data[i][6], SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd') : String(data[i][6]||''),
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

    const row = [
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
    ];
    getSheet(CONFIG.SHEETS.TUGAS_PROJECT).appendRow(row);
    syncSheetRowToSupabase(CONFIG.SHEETS.TUGAS_PROJECT, row);
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
        tanggal: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd') : String(data[i][1]),
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
    const row = [generateId(), tanggal, nama, jenisAsset, deskripsi, parseFloat(estimasiHarga)||0, prioritas, bukti || '', 'Pending Team Leader', createdBy, new Date().toISOString(), JSON.stringify(historyArr)];
    sheet.appendRow(row);
    syncSheetRowToSupabase(CONFIG.SHEETS.ASSET, row);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function updateAsset(id, tanggal, nama, jenisAsset, deskripsi, estimasiHarga, prioritas, bukti, updatedBy) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ASSET);
    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(id)) continue;
      const creator = String(data[i][9] || '');
      if (String(updatedBy) !== creator) {
        return { success: false, message: 'Hanya pembuat pengajuan yang dapat mengedit asset ini.' };
      }
      const status = String(data[i][8] || '');
      const finalStatuses = ['Disetujui', 'Tolak', 'Ditolak', 'Approved', 'Rejected', 'Disetujui Admin', 'Approved Admin'];
      if (finalStatuses.includes(status)) {
        return { success: false, message: 'Pengajuan tidak dapat diedit karena sudah selesai.' };
      }
      sheet.getRange(i + 1, 2).setValue(tanggal);
      sheet.getRange(i + 1, 3).setValue(nama);
      sheet.getRange(i + 1, 4).setValue(jenisAsset);
      sheet.getRange(i + 1, 5).setValue(deskripsi);
      sheet.getRange(i + 1, 6).setValue(parseFloat(estimasiHarga) || 0);
      sheet.getRange(i + 1, 7).setValue(prioritas);
      sheet.getRange(i + 1, 8).setValue(bukti || '');
      let historyRaw = data[i][11] || '[]';
      let historyArr = [];
      try { historyArr = JSON.parse(historyRaw); } catch (e) { historyArr = []; }
      historyArr.push({ date: now, action: 'Diedit', status: status, by: updatedBy, role: 'Pemohon', reason: 'Update data pengajuan' });
      sheet.getRange(i + 1, 12).setValue(JSON.stringify(historyArr));
      return { success: true };
    }
    return { success: false, message: 'Asset tidak ditemukan' };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteAsset(id) {
  return deleteRow(CONFIG.SHEETS.ASSET, id);
}

/**
 * Tambahkan komentar ke riwayat pengajuan asset (tidak mengubah status kecuali direquest)
 */
function addAssetComment(id, comment, by, role) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ASSET);
    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        let historyRaw = data[i][11] || '[]';
        let historyArr = [];
        try { historyArr = JSON.parse(historyRaw); } catch(e) { historyArr = []; }
        historyArr.push({ date: now, action: 'Komentar', status: data[i][8] || '', by: by || '', role: role || '', reason: comment || '' });
        sheet.getRange(i+1, 12).setValue(JSON.stringify(historyArr));
        return { success: true };
      }
    }
    return { success: false, message: 'Asset tidak ditemukan' };
  } catch(e) { return { success: false, message: e.message }; }
}

/**
 * Update status pengajuan asset secara eksplisit (bisa mundur/maju), dan catat history
 */
function updateAssetStatus(id, status, by, role, reason) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ASSET);
    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i+1, 9).setValue(status);
        let historyRaw = data[i][11] || '[]';
        let historyArr = [];
        try { historyArr = JSON.parse(historyRaw); } catch(e) { historyArr = []; }
        historyArr.push({ date: now, action: 'Status diubah', status: status, by: by || '', role: role || '', reason: reason || '' });
        sheet.getRange(i+1, 12).setValue(JSON.stringify(historyArr));
        return { success: true };
      }
    }
    return { success: false, message: 'Asset tidak ditemukan' };
  } catch(e) { return { success: false, message: e.message }; }
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
    const row = [generateId(), tanggal, stockId, sku, nama, lokasi||'-', batch||'-', expDate||'-', sistem, fisik, selisih, 'Pending', catatan||'', createdBy, new Date().toISOString(), '', ''];
    sheet.appendRow(row);
    syncSheetRowToSupabase(CONFIG.SHEETS.STOCK_OPNAME, row);
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
    const row = [generateId(), tanggal, noPL, noOrder || '-', supplier || '-', keterangan, fileUrl, createdBy || 'User', new Date().toISOString()];
    finalSheet.appendRow(row);
    syncSheetRowToSupabase(CONFIG.SHEETS.PACKING_LIST, row);
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
    const row = [
      id, nama, jabatan, cabang || '', telepon || '',
      tanggalMasuk, tanggalResign, alasanResign, keterangan || '',
      createdBy, new Date().toISOString()
    ];
    getSheet(CONFIG.SHEETS.RIWAYAT_KARYAWAN).appendRow(row);
    syncSheetRowToSupabase(CONFIG.SHEETS.RIWAYAT_KARYAWAN, row);
    
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
    const row = [
      id, karyawanNama, karyawanId || '', jenisSP, alasan,
      tanggalSP, parseInt(masaBerlaku), tglKadStr,
      'Aktif', createdBy, new Date().toISOString()
    ];
    getSheet(CONFIG.SHEETS.SURAT_PERINGATAN).appendRow(row);
    syncSheetRowToSupabase(CONFIG.SHEETS.SURAT_PERINGATAN, row);
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

function getLemburFullData() {
  try {
    return {
      success: true,
      lembur: getLembur(""),
      laporan: getLaporanKerja(),
      tglMerah: getTglMerahData()
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
            kategori: data[i][1] || '',
            code: data[i][2] || '',
            nama: data[i][3] || '',
            divisi: data[i][4] || '',
            tanggalMasuk: data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][5]||''),
            status: data[i][6] || 'Aktif',
            createdBy: data[i][7] || '',
            createdAt: data[i][8] || '',
            history: data[i][9] || '',
            qty: data[i][10] || 1,
            zoneId: data[i][11] || ''
        });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function getAssetAuditStatus() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ASSET_AUDIT_LOG);
    if (!sheet) return { success: true, data: {} };
    const data = sheet.getDataRange().getValues();
    const result = {};
    for (let i = 1; i < data.length; i++) {
      const assetId = String(data[i][1]);
      const date = data[i][2];
      if (!result[assetId] || new Date(date) > new Date(result[assetId])) {
        result[assetId] = date instanceof Date ? Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(date);
      }
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function addAssetWarehouse(codePrefix, nama, tanggalMasuk, divisi, status, createdBy, qty, zoneId, kategori) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ASSET_WAREHOUSE);
    const data = sheet.getDataRange().getValues();
    const createdAt = new Date().toISOString();
    const count = parseInt(qty) || 1;
    
    if (String(status || '').toLowerCase() === 'rusak') {
      divisi = '';
      zoneId = '';
    }

    const basePrefix = (codePrefix && codePrefix.trim() !== '') ? codePrefix.trim() : nama;
    const prefix = basePrefix + "-";
    
    let maxNum = 0;
    for (let j = 1; j < data.length; j++) {
      const existingCode = String(data[j][2] || '');
      if (existingCode.startsWith(prefix)) {
        const numPart = parseInt(existingCode.split('-').pop());
        if (!isNaN(numPart) && numPart > maxNum) maxNum = numPart;
      }
    }
    
    for (let i = 0; i < count; i++) {
        const id = generateId();
        const currentNum = maxNum + i + 1;
        const code = `${basePrefix}-${currentNum}`;
        const history = `🛒 Dibuat oleh ${createdBy} pada ${createdAt} (Tgl Masuk: ${tanggalMasuk})`;
        const row = [id, kategori || '', code, nama, divisi, tanggalMasuk, status || 'Aktif', createdBy, createdAt, history, 1, zoneId || ''];
        sheet.appendRow(row);
        syncSheetRowToSupabase(CONFIG.SHEETS.ASSET_WAREHOUSE, row);
    }
    
    return { success: true, message: `${count} unit asset ${nama} berhasil ditambahkan.` };
  } catch (e) { return { success: false, message: e.message }; }
}

function importAssetWarehouseBulk(items, userNama) {
  try {
    if (!items || !items.length) return { success: false, message: 'Tidak ada data untuk diimpor.' };
    const sheet = getSheet(CONFIG.SHEETS.ASSET_WAREHOUSE);
    const data = sheet.getDataRange().getValues();
    let imported = 0;

    items.forEach(item => {
      const codePrefix = item.CodePrefix || item.codeprefix || item.codePrefix || item['Code Prefix'] || item['code prefix'] || '';
      const nama = item.Nama || item.nama || item.Name || item.name || '';
      const tanggalMasuk = item.TanggalMasuk || item.tanggalMasuk || item['Tanggal Masuk'] || '';
      const divisi = item.Divisi || item.divisi || item.Division || item.division || 'Gudang';
      const status = item.Status || item.status || 'Aktif';
      const qty = item.Qty || item.qty || item.quantity || 1;
      const zoneId = item.ZoneId || item.zoneId || item.Zone || item.zone || '';
      const kategori = item.Kategori || item.kategori || item.Category || item.category || item['Kategori'] || item['kategori'] || item['Category'] || item['category'] || 'Lain-lain';

      if (!nama) return;
      addAssetWarehouse(codePrefix, nama, tanggalMasuk, divisi, status, userNama || 'System', qty, zoneId, kategori);
      imported += parseInt(qty) || 1;
    });

    return { success: true, message: `Berhasil mengimpor ${imported} asset.` };
  } catch (e) { return { success: false, message: e.message }; }
}

function addAssetAudit(assetId, kondisi, catatan, petugas) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ASSET_AUDIT_LOG); 
    if (!sheet) {
      const ss = getSpreadsheet();
      setupSheet(ss, CONFIG.SHEETS.ASSET_AUDIT_LOG, ['id','assetId','tanggal','kondisi','catatan','petugas','createdAt','statusApproval','approvedBy','approvedAt']);
    }
    const finalSheet = getSheet(CONFIG.SHEETS.ASSET_AUDIT_LOG);
    const id = generateId();
    const now = new Date().toISOString();
    const tanggal = now.split('T')[0];
    
    // Status Approval default: Pending
    finalSheet.appendRow([id, assetId, tanggal, kondisi, catatan || '', petugas, now, 'Pending', '', '']);
    syncSheetRowToSupabase(CONFIG.SHEETS.ASSET_AUDIT_LOG, [id, assetId, tanggal, kondisi, catatan || '', petugas, now, 'Pending', '', '']);
    
    // Status di master asset HANYA berubah jika disetujui, jadi jangan update di sini.
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function getAuditSyncData() {
  try {
    return {
      success: true,
      assets: getAssetWarehouseData().data || [],
      logs: getAssetAuditLogs().data || []
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getAssetAuditLogs() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ASSET_AUDIT_LOG);
    if (!sheet) return { success: true, data: [] };
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      result.push({
        id: data[i][0],
        assetId: String(data[i][1]),
        tanggal: data[i][2],
        kondisi: data[i][3],
        catatan: data[i][4],
        petugas: data[i][5],
        createdAt: data[i][6],
        statusApproval: data[i][7] || 'Approved',
        approvedBy: data[i][8] || '',
        approvedAt: data[i][9] || ''
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function approveAssetAudit(auditId, status, approver) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ASSET_AUDIT_LOG);
    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(auditId)) {
        sheet.getRange(i + 1, 8, 1, 3).setValues([[status, approver, now]]);
        
        // Jika disetujui (Approved), baru update status di master asset
        if (status === 'Approved') {
          const assetId = data[i][1];
          const kondisi = data[i][3];
          const awSheet = getSheet(CONFIG.SHEETS.ASSET_WAREHOUSE);
          const awData = awSheet.getDataRange().getValues();
          for (let j = 1; j < awData.length; j++) {
            if (String(awData[j][0]) === String(assetId)) {
              // status is column 7 in new layout
              awSheet.getRange(j + 1, 7).setValue(kondisi);
              break;
            }
          }
        }
        return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateAssetWarehouse(id, nama, tanggalMasuk, status, userNama, qty, zoneId, kategori) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ASSET_WAREHOUSE);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        // columns: 1=id,2=kategori,3=code,4=nama,5=divisi,6=tanggalMasuk,7=status,8=createdBy,9=createdAt,10=history,11=qty,12=zoneId
        sheet.getRange(i + 1, 4).setValue(nama);
        sheet.getRange(i + 1, 6).setValue(tanggalMasuk);
        sheet.getRange(i + 1, 7).setValue(status);
        sheet.getRange(i + 1, 11).setValue(qty || 1);
        sheet.getRange(i + 1, 2).setValue(kategori || 'Lain-lain');

        if (String(status || '').toLowerCase() === 'rusak') {
          sheet.getRange(i + 1, 5).setValue('');
          sheet.getRange(i + 1, 12).setValue('');
        } else if (zoneId !== undefined) {
          sheet.getRange(i + 1, 12).setValue(zoneId || '');
        }

        let oldHist = data[i][9] || '';
        const now = new Date().toLocaleString('id-ID');
        const entry = `✏️ Diperbarui oleh ${userNama} pada ${now}`;
        sheet.getRange(i + 1, 10).setValue(oldHist ? oldHist + '\n' + entry : entry);

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
        const oldZone = data[i][11] || '';
        const oldHist = data[i][9] || '';
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
          sheet.getRange(i + 1, 12).setValue(targetZoneId || '');
          sheet.getRange(i + 1, 10).setValue(oldHist ? oldHist + '\n' + entry : entry);
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
// AUDIT REPORTS
// ============================================================
function getAuditReports() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.AUDIT_REPORTS);
    if (!sheet) return { success: true, data: [] };
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      result.push({
        id: data[i][0],
        tanggal: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]),
        auditor: data[i][2],
        totalAsset: data[i][3],
        terscan: data[i][4],
        minus: data[i][5],
        status: data[i][6],
        createdBy: data[i][7],
        createdAt: data[i][8],
        history: data[i][9] || '',
        missingAssets: data[i][10] || '[]',
        negativeList: data[i][11] || '[]'
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function generateAuditReport(auditorName) {
  try {
    const assetSheet = getSheet(CONFIG.SHEETS.ASSET_WAREHOUSE);
    const auditLogSheet = getSheet(CONFIG.SHEETS.AUDIT_LOGS);
    
    if (!assetSheet || !auditLogSheet) throw new Error("Sheet data tidak ditemukan");
    
    const assets = assetSheet.getDataRange().getValues().slice(1).filter(r => r.join('').trim() !== '');
    const logs = auditLogSheet.getDataRange().getValues().slice(1).filter(r => r.join('').trim() !== '');
    
    // Only count assets that are "Aktif" if needed, but usually audit includes all
    const totalAsset = assets.length;
    const auditedIds = new Set(logs.map(l => l[1])); // assetId is col 2 (index 1)
    const terscan = auditedIds.size;
    const minus = totalAsset - terscan;
    
    // Identify Missing Assets
    const missingAssetsList = assets.filter(a => !auditedIds.has(a[0])).map(a => `${a[3]} (${a[2]})`); // Nama (Code)
    
    const reportSheet = getSheet(CONFIG.SHEETS.AUDIT_REPORTS);
    const id = 'REP-' + Date.now();
    const now = new Date();
    const nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    const tanggal = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    const history = JSON.stringify([{
      action: "Laporan digenerate",
      by: auditorName,
      time: nowStr
    }]);
    
    reportSheet.appendRow([id, tanggal, auditorName, totalAsset, terscan, minus, 'Pending', auditorName, nowStr, history, JSON.stringify(missingAssetsList)]);
    syncSheetRowToSupabase(CONFIG.SHEETS.AUDIT_REPORTS, [id, tanggal, auditorName, totalAsset, terscan, minus, 'Pending', auditorName, nowStr, history, JSON.stringify(missingAssetsList)]);
    
    return { success: true, id: id };
  } catch (e) { return { success: false, message: e.message }; }
}

/**
 * Generate a detailed opname report for a specific Asset Opname session.
 * Returns report id and list of negative-stock items (where system qty > physical qty).
 */
function generateOpnameReport(sessionId, auditorName) {
  try {
    const ss = getSpreadsheet();
    const assetSheet = getSheet(CONFIG.SHEETS.ASSET_WAREHOUSE);
    const logSheet = ss.getSheetByName('AssetOpnameLog');
    const reportSheet = getSheet(CONFIG.SHEETS.AUDIT_REPORTS);

    if (!assetSheet || !logSheet || !reportSheet) throw new Error('Sheet data tidak ditemukan');

    const assets = assetSheet.getDataRange().getValues().slice(1).filter(r => r.join('').trim() !== '');
    const logs = logSheet.getDataRange().getValues().slice(1).filter(r => r.join('').trim() !== '');

    // Fetch session details to get the divisi
    const sesSheet = ss.getSheetByName('AssetOpnameSession');
    let sessionDivisi = 'Semua';
    if (sesSheet) {
      const sesData = sesSheet.getDataRange().getValues();
      for (let i = 1; i < sesData.length; i++) {
        if (String(sesData[i][0]) === String(sessionId)) {
          sessionDivisi = sesData[i][2] || 'Semua';
          break;
        }
      }
    }

    // Filter logs for this session
    const sessionLogs = logs.filter(l => String(l[1]) === String(sessionId));
    const auditedIds = new Set(sessionLogs.map(l => String(l[2])));

    // Filter assets by session division to calculate correct stats
    const filteredAssets = assets.filter(a => {
      if (sessionDivisi && sessionDivisi !== 'Semua' && String(a[4] || '') !== sessionDivisi) return false;
      return true;
    });

    const totalAsset = filteredAssets.length;
    const terscan = auditedIds.size;
    const minusCount = totalAsset - terscan;

    // Missing assets (not scanned in this session, but belong to this session's division)
    const missingAssetsList = filteredAssets.filter(a => !auditedIds.has(String(a[0]))).map(a => `${a[3]} (${a[2]})`);

    // Negative stock details: compare system qty vs scanned qty per asset
    const negativeList = [];
    for (let i = 0; i < sessionLogs.length; i++) {
      const row = sessionLogs[i];
      const assetId = String(row[2]);
      const qtyFisik = parseFloat(row[6]) || 0; // log: qty fisik at index 6

      // find asset in assets
      const assetRow = assets.find(a => String(a[0]) === assetId);
      if (!assetRow) continue;
      const systemQty = parseFloat(assetRow[10]) || 0; // ASSET_WAREHOUSE qty column (index 10)
      const diff = systemQty - qtyFisik; // positive if system has more than physical (minus)
      if (diff > 0) {
        negativeList.push({
          assetId: assetRow[0],
          code: assetRow[2],
          nama: assetRow[3],
          divisi: assetRow[4],
          systemQty: systemQty,
          physicalQty: qtyFisik,
          difference: diff
        });
      }
    }

    // Append report to AuditReports (re-using existing audit reports sheet)
    const id = sessionId; // Link report ID directly to sessionId for easy tracking
    const now = new Date();
    const nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    const tanggal = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const history = JSON.stringify([{ action: 'Opname report generated', by: auditorName, time: nowStr }]);

    // Check if report row already exists to avoid duplication
    let existingRowIndex = -1;
    const reportData = reportSheet.getDataRange().getValues();
    for (let r = 1; r < reportData.length; r++) {
      if (String(reportData[r][0]) === String(id)) {
        existingRowIndex = r + 1;
        break;
      }
    }

    const rowData = [id, tanggal, auditorName, totalAsset, terscan, minusCount, 'Pending', auditorName, nowStr, history, JSON.stringify(missingAssetsList), JSON.stringify(negativeList)];
    if (existingRowIndex > 0) {
      reportSheet.getRange(existingRowIndex, 1, 1, rowData.length).setValues([rowData]);
      syncSheetRowToSupabase(CONFIG.SHEETS.AUDIT_REPORTS, rowData);
    } else {
      reportSheet.appendRow(rowData);
      syncSheetRowToSupabase(CONFIG.SHEETS.AUDIT_REPORTS, rowData);
    }

    return { success: true, id: id, negative: negativeList, missing: missingAssetsList };
  } catch (e) { return { success: false, message: e.message }; }
}

function approveAuditReport(reportId, approverName) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.AUDIT_REPORTS);
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    const nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(reportId)) {
        let history = [];
        try { history = JSON.parse(data[i][9] || '[]'); } catch(e) {}
        history.push({
          action: "Laporan disetujui",
          by: approverName,
          time: nowStr
        });
        
        sheet.getRange(i + 1, 7).setValue('Approved');
        sheet.getRange(i + 1, 10).setValue(JSON.stringify(history));
        return { success: true };
      }
    }
    return { success: false, message: 'Laporan tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

function bulkSyncAssetCodes() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.ASSET_WAREHOUSE);
    const range = sheet.getDataRange();
    const data = range.getValues();
    if (data.length <= 1) return { success: true, message: 'Tidak ada data untuk disinkronisasi.' };
    
    const header = data[0];
    const rows = data.slice(1);
    
    const newRows = [];
    const nameCounters = {};
    let splitCount = 0;
    let syncCount = 0;

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        if (row.join('').trim() === '') continue;
        
        const currentCode = String(row[1] || '').trim();
        const nama = String(row[2]);
        
        // Tentukan prefix: PRIORITASKAN kode yang sudah ada, jika tidak ada (kosong) baru ambil dari nama
        let basePrefix = '';
        if (currentCode !== '') {
            basePrefix = currentCode.includes('-') ? currentCode.split('-')[0] : currentCode;
        } else {
            basePrefix = nama;
        }
        
        const qty = parseInt(row[9]) || 1;
        
        // Inisialisasi counter untuk prefix ini jika belum ada
        if (!nameCounters[basePrefix]) nameCounters[basePrefix] = 0;
        
        // Proses per unit (Pecah baris jika Qty > 1)
        for (let q = 0; q < qty; q++) {
            nameCounters[basePrefix]++;
            const newCode = `${basePrefix}-${nameCounters[basePrefix]}`;
            
            // Buat row baru
            const newRow = [...row];
            // Selalu beri ID baru jika baris dipecah, untuk baris pertama tetap gunakan ID lama
            if (q > 0) {
                newRow[0] = generateId(); 
                splitCount++;
            }
            
            newRow[1] = newCode; // Kode Baru
            newRow[9] = 1;       // Qty diset ke 1 per baris
            
            newRows.push(newRow);
            syncCount++;
        }
    }
    
    // Tulis ulang seluruh data ke sheet
    sheet.clearContents();
    const output = [header, ...newRows];
    sheet.getRange(1, 1, output.length, header.length).setValues(output);
    
    return { 
        success: true, 
        count: syncCount,
        message: `✅ Sinkronisasi Selesai! ${syncCount} unit sekarang memiliki kode unik. ${splitCount} baris baru dibuat dari pemisahan Qty.` 
    };
  } catch (e) { return { success: false, message: e.message }; }
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
    
    // Ambil data mulai dari baris 2 kolom 1 s/d baris terakhir kolom 22 (termasuk jam-jam baru)
    const data = sheet.getRange(2, 1, lastRow - 1, 22).getValues();
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
            createdAt: row[9],
            parkir: Number(row[10] || 0),
            tol: Number(row[11] || 0),
            bensin: Number(row[12] || 0),
            pkbm: Number(row[13] || 0),
            lainLain: Number(row[14] || 0),
            totalBiaya: Number(row[15] || 0),
            buktiPembayaranUrl: String(row[16] || ''),
            driverNotes: String(row[17] || ''),
            jamMulaiPerjalanan: String(row[18] || ''),
            jamTibaTujuan: String(row[19] || ''),
            jamKembaliWarehouse: String(row[20] || ''),
            jamSampaiWarehouse: String(row[21] || '')
        });
    }
    return { success: true, data: result, totalOnSheet: lastRow - 1 };
  } catch (e) { return { success: false, message: e.message }; }
}

function getBookingMobilById(id) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.BOOKING_MOBIL);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        return {
          success: true,
          data: {
            id: String(data[i][0]),
            tanggal: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]||''),
            pic: data[i][2],
            jamBerangkat: data[i][3],
            tujuan: data[i][4],
            keterangan: data[i][5],
            rute: data[i][6],
            status: data[i][7],
            createdBy: data[i][8],
            createdAt: data[i][9],
            parkir: data[i][10],
            tol: data[i][11],
            bensin: data[i][12],
            pkbm: data[i][13],
            lainLain: data[i][14],
            totalBiaya: data[i][15],
            buktiPembayaranUrl: data[i][16],
            driverNotes: data[i][17]
          }
        };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

function addBookingMobil(tanggal, pic, jamBerangkat, tujuan, keterangan, rute, createdBy, parkir, tol, bensin, pkbm, lainLain, details) {
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
    const totalBiaya = (Number(parkir) || 0) + (Number(tol) || 0) + (Number(bensin) || 0) + (Number(pkbm) || 0) + (Number(lainLain) || 0);
    
    const masterRow = [
      id, tanggal, pic, jamBerangkat, tujuan, keterangan, rute, 'Belum Jalan', createdBy, createdAt,
      Number(parkir) || 0, Number(tol) || 0, Number(bensin) || 0, Number(pkbm) || 0, Number(lainLain) || 0,
      totalBiaya, '', ''
    ];
    sheet.appendRow(masterRow);
    syncSheetRowToSupabase(CONFIG.SHEETS.BOOKING_MOBIL, masterRow);

    // Simpan Detail PO jika ada
    if (details && Array.isArray(details)) {
      const detailSheet = getSheet(CONFIG.SHEETS.BOOKING_MOBIL_DETAIL);
      details.forEach(det => {
        const detRow = [
          generateId(), id,
          det.tanggal || tanggal, det.namaCustomer || '', det.noPo || '',
          Number(det.totalCartoon) || 0, Number(det.parkir) || 0, Number(det.tol) || 0,
          Number(det.pkbm) || 0, Number(det.lainLain) || 0, det.keterangan || '', ''
        ];
        detailSheet.appendRow(detRow);
        syncSheetRowToSupabase(CONFIG.SHEETS.BOOKING_MOBIL_DETAIL, detRow);
      });
      SpreadsheetApp.flush();
    }
    
    return { success: true, id: id };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteBookingMobil(id) {
  const res = deleteRow(CONFIG.SHEETS.BOOKING_MOBIL, id);
  if (res.success) {
    // Hapus juga detailnya
    try {
      const sheetDetail = getSheet(CONFIG.SHEETS.BOOKING_MOBIL_DETAIL);
      const dataDetail = sheetDetail.getDataRange().getValues();
      // Hapus dari bawah ke atas agar index tidak berantakan
      for (let i = dataDetail.length - 1; i >= 1; i--) {
        if (String(dataDetail[i][1]) === String(id)) {
          sheetDetail.deleteRow(i + 1);
        }
      }
    } catch(err) { console.warn('Failed to delete booking details:', err); }
  }
  return res;
}

function getBookingMobilDetail(bookingId) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.BOOKING_MOBIL_DETAIL);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(bookingId)) {
        const parkir = Number(data[i][6]) || 0;
        const tol = Number(data[i][7]) || 0;
        const pkbm = Number(data[i][8]) || 0;
        const lainLain = Number(data[i][9]) || 0;
        result.push({
          id: String(data[i][0]),
          bookingId: String(data[i][1]),
          tanggal: data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][2]||''),
          namaCustomer: String(data[i][3] || ''),
          noPo: String(data[i][4] || ''),
          totalCartoon: Number(data[i][5]) || 0,
          parkir: parkir,
          tol: tol,
          pkbm: pkbm,
          lainLain: lainLain,
          totalBiayaPo: parkir + tol + pkbm + lainLain,
          keterangan: String(data[i][10] || ''),
          buktiUrls: String(data[i][11] || '')
        });
      }
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function getBookingMobilDetailMaster() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.BOOKING_MOBIL_DETAIL);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      result.push({
        id: String(data[i][0]),
        buktiUrls: String(data[i][11] || '')
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateBookingMobilDetailBiaya(detailId, parkir, tol, pkbm, lainLain) {
  try {
    const detailSheet = getSheet(CONFIG.SHEETS.BOOKING_MOBIL_DETAIL);
    const detailData = detailSheet.getDataRange().getValues();
    let bookingId = null;
    let noPo = null;

    // Helper: Convert any value to safe numeric (remove letters, keep only digits)
    const sanitizeNumber = (val) => {
      if (typeof val === 'number') return Math.max(0, val);
      const str = String(val || '').trim();
      if (!str) return 0;
      // Remove all non-digit characters, keep only numbers
      const numStr = str.replace(/[^0-9]/g, '');
      const num = parseInt(numStr, 10);
      return isNaN(num) ? 0 : Math.max(0, num);
    };

    // Sanitize all inputs - convert any letters to 0
    const cleanParkir = sanitizeNumber(parkir);
    const cleanTol = sanitizeNumber(tol);
    const cleanPkbm = sanitizeNumber(pkbm);
    const cleanLainLain = sanitizeNumber(lainLain);

    // 1. Update baris detail & extract noPo
    for (let i = 1; i < detailData.length; i++) {
      if (String(detailData[i][0]) === String(detailId)) {
        bookingId = String(detailData[i][1]);
        noPo = String(detailData[i][4] || ''); // Extract noPo for sync
        detailSheet.getRange(i + 1, 7).setValue(cleanParkir);   // parkir
        detailSheet.getRange(i + 1, 8).setValue(cleanTol);      // tol
        detailSheet.getRange(i + 1, 9).setValue(cleanPkbm);     // pkbm
        detailSheet.getRange(i + 1, 10).setValue(cleanLainLain); // lainLain
        
        // Update data array local agar kalkulasi total di bawah menggunakan nilai baru
        detailData[i][6] = cleanParkir;
        detailData[i][7] = cleanTol;
        detailData[i][8] = cleanPkbm;
        detailData[i][9] = cleanLainLain;
        
        break;
      }
    }

    if (!bookingId) return { success: false, message: 'Detail tidak ditemukan' };

    // 2. Trigger sync to ensure proper cross-sheet synchronization
    if (noPo) {
      syncPOToBookingMobilDetail(noPo, cleanParkir, cleanTol, cleanPkbm, cleanLainLain);
    }

    // 3. Hitung ulang total biaya untuk bookingId tersebut menggunakan array lokal yang sudah update
    let totalDetailBiaya = 0;
    for (let i = 1; i < detailData.length; i++) {
      if (String(detailData[i][1]) === String(bookingId)) {
        totalDetailBiaya += (Number(detailData[i][6]) || 0) + 
                           (Number(detailData[i][7]) || 0) + 
                           (Number(detailData[i][8]) || 0) + 
                           (Number(detailData[i][9]) || 0);
      }
    }

    // Ambil data master
    const masterSheet = getSheet(CONFIG.SHEETS.BOOKING_MOBIL);
    const masterData = masterSheet.getDataRange().getValues();
    for (let i = 1; i < masterData.length; i++) {
      if (String(masterData[i][0]) === String(bookingId)) {
        // Biaya Driver (Master) + Biaya Detail PO
        const parkirMaster = Number(masterData[i][10]) || 0;
        const tolMaster = Number(masterData[i][11]) || 0;
        const bensinMaster = Number(masterData[i][12]) || 0;
        const pkbmMaster = Number(masterData[i][13]) || 0;
        const lainLainMaster = Number(masterData[i][14]) || 0;
        
        const finalTotal = parkirMaster + tolMaster + bensinMaster + pkbmMaster + lainLainMaster + totalDetailBiaya;
        
        masterSheet.getRange(i + 1, 16).setValue(finalTotal); // Kolom 16 = totalBiaya
        
        // Force apply all pending updates to sheet so subsequent reads see the new values
        SpreadsheetApp.flush();
        
        return { success: true, finalTotal: finalTotal };
      }
    }

    SpreadsheetApp.flush();
    return { success: true }; // Harusnya tidak sampai sini jika bookingId valid
  } catch (e) { return { success: false, message: e.message }; }
}

// Ensure installable triggers exist for automatic syncing via onEdit
// ============================================================
// REAL-TIME SYNC: GOOGLE SHEET → SUPABASE
// Setiap perubahan di Sheet langsung di-push ke Supabase.
// Trigger: onEdit (setiap cell berubah) + onChange (struktur berubah)
// ============================================================

function ensureTriggers() {
  try {
    const current = ScriptApp.getProjectTriggers();
    const ss = getSpreadsheet();

    const hasOnEdit   = current.some(t => t.getHandlerFunction() === 'onEdit');
    const hasOnChange = current.some(t => t.getHandlerFunction() === 'onChange');

    if (!hasOnEdit) {
      ScriptApp.newTrigger('onEdit').forSpreadsheet(ss).onEdit().create();
      Logger.log('✅ Trigger onEdit dibuat');
    }
    if (!hasOnChange) {
      ScriptApp.newTrigger('onChange').forSpreadsheet(ss).onChange().create();
      Logger.log('✅ Trigger onChange dibuat');
    }
  } catch (e) {
    Logger.log('⚠️ ensureTriggers error: ' + e.message);
  }
}

// ============================================================
// onEdit — dipanggil setiap kali user mengedit 1+ sel di Sheet
// Sync baris yang diedit langsung ke Supabase (real-time, ~1-2 detik)
// ============================================================
function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const sheet     = e.range.getSheet();
    if (!sheet) return;

    const sheetName = sheet.getName();
    const row       = e.range.getRow();
    const col       = e.range.getColumn();
    const numRows   = e.range.getNumRows();

    // Skip baris header
    if (row <= 1) return;

    // ── BookingMobilDetail: sanitasi angka Parkir/Tol/PKBM/Lain-Lain ──
    if (sheetName === CONFIG.SHEETS.BOOKING_MOBIL_DETAIL) {
      if (col >= 7 && col <= 10) {
        const rowValues = sheet.getRange(row, 1, 1, 12).getValues()[0];
        const noPo = String(rowValues[4] || '').trim();
        if (noPo) {
          const san = (v) => {
            if (typeof v === 'number') return Math.max(0, v);
            const n = parseInt(String(v || '').replace(/[^0-9]/g, ''), 10);
            return isNaN(n) ? 0 : Math.max(0, n);
          };
          const parkir   = san(rowValues[6]);
          const tol      = san(rowValues[7]);
          const pkbm     = san(rowValues[8]);
          const lainLain = san(rowValues[9]);
          if (parkir   !== rowValues[6]) sheet.getRange(row, 7).setValue(parkir);
          if (tol      !== rowValues[7]) sheet.getRange(row, 8).setValue(tol);
          if (pkbm     !== rowValues[8]) sheet.getRange(row, 9).setValue(pkbm);
          if (lainLain !== rowValues[9]) sheet.getRange(row, 10).setValue(lainLain);
          syncPOToBookingMobilDetail(noPo, parkir, tol, pkbm, lainLain);
        }
      }
    }

    // ── Sync setiap baris yang diedit ke Supabase ──
    if (SUPABASE_SYNC_SHEETS.indexOf(sheetName) === -1) return;

    const headers = SHEET_HEADERS[sheetName];
    const headerCount = headers ? headers.length : sheet.getLastColumn();

    // Support multi-baris diedit sekaligus (paste, fill-down, dll)
    const endRow = Math.min(row + numRows - 1, sheet.getLastRow());

    for (let r = row; r <= endRow; r++) {
      if (r <= 1) continue; // skip header
      const rowVals = sheet.getRange(r, 1, 1, headerCount).getValues()[0];

      // Skip baris kosong
      if (!rowVals || rowVals.every(v => v === '' || v === null || v === undefined)) continue;
      // Skip kalau row adalah baris header
      if (isSupabaseHeaderRow(sheetName, rowVals)) continue;

      const res = syncSheetRowToSupabase(sheetName, rowVals, sheet, r);
      if (!res.success) {
        Logger.log('⚠️ onEdit sync gagal ' + sheetName + ' baris ' + r + ': ' + (res.message || ''));
      } else {
        Logger.log('✅ onEdit sync: ' + sheetName + ' baris ' + r);
      }
    }

  } catch (err) {
    try { Logger.log('❌ onEdit error: ' + err.message); } catch (_) {}
  }
}

// ============================================================
// onChange — dipanggil saat struktur sheet berubah
// (tambah/hapus baris/kolom, paste banyak baris, dll)
// Lebih berat dari onEdit — gunakan untuk perubahan massal
// ============================================================
function onChange(e) {
  try {
    if (!e || !e.changeType) return;

    const changeType = e.changeType.toString().toUpperCase();
    Logger.log('📋 onChange dipanggil, changeType: ' + changeType);

    // EDIT sudah ditangani onEdit — skip agar tidak double-sync
    if (changeType === 'EDIT') return;

    if (changeType === 'REMOVE_ROW' || changeType === 'REMOVE_COLUMN') {
      // Hapus baris di Supabase yang sudah tidak ada di Sheet
      Logger.log('🗑️  onChange: mendeteksi penghapusan, mulai syncDeletedRowsToSupabase...');
      syncDeletedRowsToSupabase();
      return;
    }

    if (changeType === 'INSERT_ROW' || changeType === 'OTHER' || changeType === 'INSERT_GRID') {
      // Sync ulang semua sheet untuk menangkap baris baru yang ditambahkan lewat bulk paste
      Logger.log('📥 onChange INSERT/OTHER: mulai syncAllSheetsToSupabase...');
      syncAllSheetsToSupabase();
      return;
    }

    // Fallback: sync semua
    Logger.log('🔄 onChange fallback sync...');
    syncAllSheetsToSupabase();
    syncDeletedRowsToSupabase();

  } catch (err) {
    try { Logger.log('❌ onChange error: ' + err.message); } catch (_) {}
  }
}

// Synchronize Parkir/Tol/PKBM/Lain-Lain values from Detail Rincian PO to BookingMobilDetail
// Assumes that noPo in Detail Rincian PO corresponds to the same NoPo in BookingMobilDetail rows.
// After updating detail rows, it also recomputes the master booking total cost.
function syncPOToBookingMobilDetail(noPo, parkir, tol, pkbm, lainLain) {
  // Log input for debugging/integration tracing
  try {
    Logger.log('syncPOToBookingMobilDetail called - NoPo: ' + String(noPo) + 
      ' Parkir=' + parkir + ', Tol=' + tol + ', PKBM=' + pkbm + ', Lain=' + lainLain);
  } catch(e) { /* ignore logging errors */ }
  try {
    const detailSheet = getSheet(CONFIG.SHEETS.BOOKING_MOBIL_DETAIL);
    const detailData = detailSheet.getDataRange().getValues();

    // Helper: Convert any value to safe numeric (remove letters, keep only digits)
    const sanitizeNumber = (val) => {
      if (typeof val === 'number') return Math.max(0, val);
      const str = String(val || '').trim();
      if (!str) return 0;
      // Remove all non-digit characters, keep only numbers
      const numStr = str.replace(/[^0-9]/g, '');
      const num = parseInt(numStr, 10);
      return isNaN(num) ? 0 : Math.max(0, num);
    };

    // Normalize numeric inputs - convert letters to 0
    const pParkir = sanitizeNumber(parkir);
    const pTol = sanitizeNumber(tol);
    const pPkbm = sanitizeNumber(pkbm);
    const pLain = sanitizeNumber(lainLain);

    // Track affected bookings for total recomputation
    const affectedBookings = new Set();

    // Update all detail rows with matching noPo
    for (let r = 1; r < detailData.length; r++) {
      const row = detailData[r];
      const rowNoPo = String(row[4] || '').trim();
      if (rowNoPo === String(noPo).trim()) {
        // Update parkir/tol/pkbm/lainLain in detail row (columns 6-9, 0-based idx)
        detailSheet.getRange(r + 1, 7).setValue(pParkir); // parkir
        detailSheet.getRange(r + 1, 8).setValue(pTol);    // tol
        detailSheet.getRange(r + 1, 9).setValue(pPkbm);   // pkbm
        detailSheet.getRange(r + 1, 10).setValue(pLain);   // lainLain
        // Track bookingId for master recalculation
        const bookingId = String(row[1] || '');
        if (bookingId) affectedBookings.add(bookingId);
      }
    }

    // Recalculate totals for each affected booking
    if (affectedBookings.size > 0) {
      const masterSheet = getSheet(CONFIG.SHEETS.BOOKING_MOBIL);
      const masterData = masterSheet.getDataRange().getValues();

      // For each affected booking, compute new total by summing detail rows and master costs
      affectedBookings.forEach(bId => {
        // Master costs
        let parkirMaster = 0, tolMaster = 0, bensinMaster = 0, pkbmMaster = 0, lainMaster = 0;
        for (let i = 1; i < masterData.length; i++) {
          if (String(masterData[i][0]) === String(bId)) {
            parkirMaster = Number(masterData[i][10]) || 0;   // Parkir
            tolMaster = Number(masterData[i][11]) || 0;      // Tol
            bensinMaster = Number(masterData[i][12]) || 0;   // Bensin
            pkbmMaster = Number(masterData[i][13]) || 0;     // PKBM
            lainMaster = Number(masterData[i][14]) || 0;     // Lain-Lain
            break;
          }
        }

        // Sum detail rows for this booking
        let detailSum = 0;
        for (let i = 1; i < detailData.length; i++) {
          const dRow = detailData[i];
          if (String(dRow[1] || '') === String(bId)) {
            detailSum += (Number(dRow[6]) || 0) + (Number(dRow[7]) || 0) + (Number(dRow[8]) || 0) + (Number(dRow[9]) || 0);
          }
        }

        const finalTotal = parkirMaster + tolMaster + bensinMaster + pkbmMaster + lainMaster + detailSum;
        // Update master total (col 16)
        for (let i = 1; i < masterData.length; i++) {
          if (String(masterData[i][0]) === String(bId)) {
            masterSheet.getRange(i + 1, 16).setValue(finalTotal);
            break;
          }
        }
      });
    }

    SpreadsheetApp.flush();
    return { success: true, updatedDetailNoPo: noPo, updatedBookings: Array.from(affectedBookings) };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ── Kode di bawah ini sudah dipindah ke blok ensureTriggers/onEdit/onChange di atas ──

// Sheet yang TIDAK aman untuk delete-sync karena kolom pertamanya bukan id unik per baris
// (JadwalRoster pakai 'Bulan' yang berulang untuk banyak baris/nama).
const SUPABASE_SKIP_DELETE_SYNC = [CONFIG.SHEETS.JADWAL_ROSTER];

// Mengecek setiap sheet: baris yang ada di Supabase tapi sudah tidak ada lagi di sheet akan dihapus.
function syncDeletedRowsToSupabase() {
  if (!SUPABASE_URL || !SUPABASE_KEY) return;
  const ss = getSpreadsheet();

  SUPABASE_SYNC_SHEETS.forEach(function (sheetName) {
    if (SUPABASE_SKIP_DELETE_SYNC.indexOf(sheetName) !== -1) return;

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    const conflictKey = getSupabaseConflictKey(sheetName); // biasanya 'id'

    let currentIds = [];
    if (lastRow > 1) {
      currentIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
        .map(function (r) { return r[0]; })
        .filter(function (v) { return v !== '' && v !== null && v !== undefined; })
        .map(String);
    }

    const result = supabaseDeleteRowsNotIn(sheetName, conflictKey, currentIds);
    if (!result.success) {
      Logger.log('⚠️  Gagal cek/hapus baris usang untuk ' + sheetName + ': ' + result.message);
    }
  });
}

// Menghapus baris di tabel Supabase yang nilai conflictKey-nya TIDAK ADA di currentIds
// (artinya baris itu sudah dihapus dari Google Sheet).
function supabaseDeleteRowsNotIn(sheetName, conflictKey, currentIds) {
  try {
    const tableName = getSupabaseTableName(sheetName);
    let apiUrl;

    if (currentIds.length === 0) {
      // Sheet sudah tidak ada data sama sekali -> kosongkan tabel ini juga
      apiUrl = SUPABASE_URL + '/rest/v1/' + encodeURIComponent(tableName);
    } else {
      const idList = currentIds.map(function (id) {
        return String(id).replace(/[(),]/g, ''); // sanitasi ringan agar tidak merusak sintaks filter
      }).join(',');
      const filter = encodeURIComponent(conflictKey) + '=not.in.(' + idList + ')';
      apiUrl = SUPABASE_URL + '/rest/v1/' + encodeURIComponent(tableName) + '?' + filter;
    }

    const options = {
      method: 'delete',
      headers: {
        'apikey': SUPABASE_KEY,
        'Authorization': 'Bearer ' + SUPABASE_KEY,
        'Prefer': 'return=minimal'
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const status = response.getResponseCode();
    if (status >= 200 && status < 300) {
      return { success: true };
    }
    return { success: false, message: 'HTTP ' + status + ' - ' + response.getContentText() };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Clean all invalid (non-numeric) values in BookingMobilDetail sheet
// Replaces any letters/text with 0 to ensure data consistency
function cleanInvalidBiayaValues() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.BOOKING_MOBIL_DETAIL);
    const data = sheet.getDataRange().getValues();
    let cleanedCount = 0;

    // Helper: Convert any value to safe numeric
    const sanitizeNumber = (val) => {
      if (typeof val === 'number') return Math.max(0, val);
      const str = String(val || '').trim();
      if (!str) return 0;
      const numStr = str.replace(/[^0-9]/g, '');
      const num = parseInt(numStr, 10);
      return isNaN(num) ? 0 : Math.max(0, num);
    };

    // Clean columns 7-10 (Parkir, Tol, PKBM, Lain-Lain)
    for (let r = 1; r < data.length; r++) {
      for (let col = 7; col <= 10; col++) {
        const val = data[r][col - 1];
        const clean = sanitizeNumber(val);
        if (clean !== val) {
          sheet.getRange(r + 1, col).setValue(clean);
          cleanedCount++;
        }
      }
    }

    SpreadsheetApp.flush();
    Logger.log('Cleaned ' + cleanedCount + ' invalid cost values in BookingMobilDetail');
    return { success: true, cleanedCount: cleanedCount };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateBookingMobilDetailBukti(detailId, urlsJson) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.BOOKING_MOBIL_DETAIL);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(detailId)) {
        sheet.getRange(i + 1, 12).setValue(urlsJson); // buktiUrls (JSON array string)
        return { success: true };
      }
    }
    return { success: false, message: 'Detail tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateBookingStatus(id, newStatus) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.BOOKING_MOBIL);
    const data = sheet.getDataRange().getValues();
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(id)) {
            // Update status (kolom 8, index 7)
            sheet.getRange(i + 1, 8).setValue(newStatus);
            
            // Update timestamp sesuai status
            if (newStatus === 'Sedang Dalam Perjalanan Ke Tujuan') {
              sheet.getRange(i + 1, 19).setValue(now); // jamMulaiPerjalanan
            } else if (newStatus === 'Sudah Tiba Di Tempat Tujuan') {
              sheet.getRange(i + 1, 20).setValue(now); // jamTibaTujuan
            } else if (newStatus === 'Kembali Ke Warehouse JKT') {
              sheet.getRange(i + 1, 21).setValue(now); // jamKembaliWarehouse
            } else if (newStatus === 'Sudah Sampai Di Warehouse') {
              sheet.getRange(i + 1, 22).setValue(now); // jamSampaiWarehouse
            }
            
            SpreadsheetApp.flush();
            return { success: true, timestamp: now };
        }
    }
    return { success: false, message: 'ID tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

// Update Bukti Pembayaran URL untuk Booking Mobil
function updateBuktiPembayaranBookingUrl(bookingId, buktiUrl) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.BOOKING_MOBIL);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(bookingId)) {
        sheet.getRange(i + 1, 17).setValue(buktiUrl); // Kolom 17 = buktiPembayaranUrl
        return { success: true };
      }
    }
    return { success: false, message: 'ID booking tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

// Update Driver Notes untuk Booking Mobil
function updateDriverNotesBooking(bookingId, driverNotes) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.BOOKING_MOBIL);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(bookingId)) {
        sheet.getRange(i + 1, 17).setValue(driverNotes); // Kolom 17 = driverNotes
        return { success: true };
      }
    }
    return { success: false, message: 'ID booking tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

// Update Biaya Booking Mobil (Parkir, Tol, Bensin, PKBM, Lain-lain)
function updateBiayaBooking(bookingId, parkir, tol, bensin, pkbm, lainLain) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.BOOKING_MOBIL);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(bookingId)) {
        const p = Number(parkir) || 0;
        const t = Number(tol) || 0;
        const b = Number(bensin) || 0;
        const pk = Number(pkbm) || 0;
        const l = Number(lainLain) || 0;
        const total = p + t + b + pk + l;
        
        sheet.getRange(i + 1, 11).setValue(p);      // Kolom 11 = parkir
        sheet.getRange(i + 1, 12).setValue(t);      // Kolom 12 = tol
        sheet.getRange(i + 1, 13).setValue(b);      // Kolom 13 = bensin
        sheet.getRange(i + 1, 14).setValue(pk);     // Kolom 14 = pkbm
        sheet.getRange(i + 1, 15).setValue(l);      // Kolom 15 = lainLain
        sheet.getRange(i + 1, 16).setValue(total);  // Kolom 16 = totalBiaya
        return { success: true, totalBiaya: total };
      }
    }
    return { success: false, message: 'ID booking tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

function uploadChunkBookingPayment(chunkData, chunkIndex, uploadId) {
  try {
    const cache = CacheService.getScriptCache();
    const id = uploadId || Utilities.getUuid();
    const key = 'payment_chunk_' + id + '_' + chunkIndex;
    
    // Google Apps Script cache limit is 100KB per key. Chunk size is 90KB.
    // If somehow it's larger, we split it.
    if (chunkData.length <= 90000) {
      cache.put(key, chunkData, 21600);
      cache.put('payment_meta_' + id + '_count', String(chunkIndex + 1), 21600);
    } else {
      const half = Math.ceil(chunkData.length / 2);
      cache.put(key + '_a', chunkData.substring(0, half), 21600);
      cache.put(key + '_b', chunkData.substring(half), 21600);
      cache.put(key + '_split', '1', 21600);
      cache.put('payment_meta_' + id + '_count', String(chunkIndex + 1), 21600);
    }
    return { success: true, uploadId: id };
  } catch (e) {
    return { success: false, message: 'Upload Chunk Error: ' + e.message };
  }
}

function finalizeBookingPaymentUpload(uploadId, fileName, mimeType) {
  try {
    const cache = CacheService.getScriptCache();
    const countStr = cache.get('payment_meta_' + uploadId + '_count');
    if (!countStr) return { success: false, message: 'Sesi upload tidak ditemukan atau sudah kedaluwarsa' };
    
    const totalChunks = parseInt(countStr);
    let fullBase64 = '';
    for (let i = 0; i < totalChunks; i++) {
      const key = 'payment_chunk_' + uploadId + '_' + i;
      const isSplit = cache.get(key + '_split');
      if (isSplit) {
        fullBase64 += (cache.get(key + '_a') || '') + (cache.get(key + '_b') || '');
      } else {
        fullBase64 += cache.get(key) || '';
      }
    }
    
    if (!fullBase64) return { success: false, message: 'Data file kosong atau gagal direkonstruksi' };
    
    const blob = Utilities.newBlob(Utilities.base64Decode(fullBase64), mimeType || 'application/octet-stream', fileName);
    const folderId = CONFIG.BOOKING_PAYMENT_FOLDER_ID || CONFIG.DRIVE_FOLDER_ID;
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(blob);
    
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(err) {}
    
    return { success: true, url: file.getUrl() };
  } catch (e) {
    return { success: false, message: 'Finalize Upload Error: ' + e.message };
  }
}

// Fungsi Upload Instan untuk file kecil (< 10MB)
function uploadPaymentBookingInstant(base64, fileName, mimeType, bookingId) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, fileName);
    const folderId = CONFIG.BOOKING_PAYMENT_FOLDER_ID || CONFIG.DRIVE_FOLDER_ID;
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(blob);
    
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(err) {}
    
    const url = file.getUrl();
    
    // Update URL di spreadsheet secara otomatis
    const res = updateBuktiPembayaranBookingUrl(bookingId, url);
    if (!res.success) return { success: false, message: 'File terupload tapi gagal update data: ' + res.message };
    
    return { success: true, url: url };
  } catch (e) {
    return { success: false, message: 'Instant Upload Error: ' + e.message };
  }
}

// Ensure global scope visibility for doPost bridge
this.uploadChunkBookingPayment = uploadChunkBookingPayment;
this.finalizeBookingPaymentUpload = finalizeBookingPaymentUpload;
this.updateBuktiPembayaranBookingUrl = updateBuktiPembayaranBookingUrl;
this.uploadPaymentBookingInstant = uploadPaymentBookingInstant;





// ============================================================
// MODUL TUGAS CONSUMABLE


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

function getScheduledClockOut(nama, tanggal, divisi) {
  try {
    const ss = getSpreadsheet();
    const tz = Session.getScriptTimeZone();
    const settingsRes = getRosterSettings();
    const settings = settingsRes.success ? settingsRes.data : { pagiIn:"08:00", pagiOut:"17:00", malamIn:"20:00", malamOut:"05:00" };
    
    // 1. Cek Roster
    if (nama && tanggal) {
      const tglDate = new Date(tanggal + 'T12:00:00'); // avoid timezone offset issues
      const monthYear = Utilities.formatDate(tglDate, tz, "yyyy-MM");
      const dayNum = String(tglDate.getDate());
      
      const rosterSheet = ss.getSheetByName(CONFIG.SHEETS.JADWAL_ROSTER);
      if (rosterSheet) {
        const rData = rosterSheet.getDataRange().getValues();
        const rHeaders = rData[0];
        for (let i = 1; i < rData.length; i++) {
          let rBulan = rData[i][0];
          if (rBulan instanceof Date) rBulan = Utilities.formatDate(rBulan, tz, "yyyy-MM");
          
          if (String(rBulan).trim() === monthYear && String(rData[i][1]).trim() === nama) {
            const dCol = rHeaders.map(String).indexOf(dayNum);
            if (dCol !== -1) {
              const shiftVal = String(rData[i][dCol]).toUpperCase();
              if (shiftVal === 'PAGI' || shiftVal.includes('PAGI')) {
                return settings.pagiOut || "17:00";
              } else if (shiftVal === 'MALAM' || shiftVal.includes('MALAM')) {
                return settings.malamOut || "05:00";
              } else if (shiftVal === 'OFF') {
                return "OFF";
              }
            }
            break;
          }
        }
      }
    }

    // 2. Fallback: Jadwal Shift Divisi
    const sheet = ss.getSheetByName(CONFIG.SHEETS.JADWAL_SHIFT);
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getDataRange().getValues();
      const activeList = [];
      for (let i = 1; i < data.length; i++) {
        if (!data[i][0]) continue;
        if (divisi && data[i][2] !== divisi) continue;
        const aktif = data[i][7];
        if (String(aktif).toLowerCase() === 'ya' || aktif === true || String(aktif).toLowerCase() === 'true') {
          activeList.push(data[i]);
        }
      }
      if (activeList.length > 0) {
        let jamPulang = activeList[0][5];
        if (jamPulang instanceof Date) {
          return Utilities.formatDate(jamPulang, tz, 'HH:mm');
        }
        return String(jamPulang || "17:00");
      }
    }
    
    // 3. Fallback ke global roster pagi out
    return settings.pagiOut || "17:00";
  } catch(e) {
    return "17:00";
  }
}

function getAbsensiInOut(ss, tz, nama, tanggal, scheduledIn, scheduledOut) {
  const absSheet = ss.getSheetByName(CONFIG.SHEETS.ABSENSI_KARYAWAN);
  let inTime = '-';
  let outTime = '-';
  let outDateStr = tanggal;
  
  if (!absSheet) return { inTime, outTime, outDateStr };
  
  const inMnt = _parseTimeToMinutes(scheduledIn || "08:00");
  const outMnt = _parseTimeToMinutes(scheduledOut || "17:00");
  const isNightShift = (inMnt > outMnt);
  
  const tglDate = new Date(tanggal + 'T12:00:00');
  tglDate.setDate(tglDate.getDate() + 1);
  const nextDayStr = Utilities.formatDate(tglDate, tz, 'yyyy-MM-dd');
  
  const absData = absSheet.getDataRange().getValues();
  const nameLower = String(nama || '').trim().toLowerCase();
  
  for (let i = 1; i < absData.length; i++) {
    if (!absData[i][1] || !absData[i][4]) continue;
    if (String(absData[i][4]).trim().toLowerCase() !== nameLower) continue;
    
    let rowTgl = "";
    let rawTgl = absData[i][1];
    if (rawTgl instanceof Date) {
      rowTgl = Utilities.formatDate(rawTgl, tz, 'yyyy-MM-dd');
    } else {
      let s = String(rawTgl).trim().split('T')[0];
      if (s.includes('/')) {
        let parts = s.split('/');
        if (parts[0].length === 4) rowTgl = parts[0] + '-' + parts[1].padStart(2, '0') + '-' + parts[2].padStart(2, '0');
        else rowTgl = parts[2] + '-' + parts[1].padStart(2, '0') + '-' + parts[0].padStart(2, '0');
      } else {
        rowTgl = s;
      }
    }
    
    let rawJam = absData[i][2];
    let jamStr = "";
    if (rawJam instanceof Date) {
      jamStr = Utilities.formatDate(rawJam, tz, 'HH:mm');
    } else {
      let match = String(rawJam || '').trim().match(/^(\d{1,2}):(\d{2})/);
      if (match) {
        jamStr = match[1].padStart(2, '0') + ':' + match[2];
      } else {
        jamStr = String(rawJam || '').trim().substring(0, 5);
      }
    }
    
    const tipe = String(absData[i][7] || '').trim().toUpperCase();
    const isClockIn = (tipe === 'IN' || tipe.includes('MASUK') || tipe.includes('IN'));
    const isClockOut = (tipe === 'OUT' || tipe.includes('PULANG') || tipe.includes('OUT'));
    
    if (isClockIn && rowTgl === tanggal) {
      if (inTime === '-' || jamStr < inTime) {
        inTime = jamStr;
      }
    }
    
    if (isClockOut) {
      if (isNightShift) {
        if (rowTgl === nextDayStr) {
          if (outTime === '-' || jamStr > outTime) {
            outTime = jamStr;
            outDateStr = nextDayStr;
          }
        }
      } else {
        if (rowTgl === tanggal) {
          if (outTime === '-' || jamStr > outTime) {
            outTime = jamStr;
            outDateStr = tanggal;
          }
        }
      }
    }
  }
  
  return { inTime, outTime, outDateStr };
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
    const statusInfo = _hitungStatusAbsensi(divisi, tipe, jamStr, nama, tglStr, jabatan || '');
    const id = generateId();
    const row = [
      id, tglStr, jamStr, karyawanId, nama, divisi, jabatan || '',
      tipe, 'manual', '', statusInfo.status, keterangan || '', now.toISOString()
    ];
    sheet.appendRow(row);
    syncSheetRowToSupabase(CONFIG.SHEETS.ABSENSI_KARYAWAN, row);
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
      const statusInfo = _hitungStatusAbsensi(divisi, tipe, jamStr, nama, tglStr, jabatan);
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

function _hitungStatusAbsensi(divisi, tipe, jam, nama, tanggal, jabatan) {
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
               let shiftIn = "", shiftOut = "", shiftTol = parseInt(settings.toleransi) || 0;
               
               if (shiftVal === 'PAGI' || shiftVal.includes('PAGI')) {
                 shiftIn = settings.pagiIn; shiftOut = settings.pagiOut;
               } else if (shiftVal === 'MALAM' || shiftVal.includes('MALAM')) {
                 shiftIn = settings.malamIn; shiftOut = settings.malamOut;
               } else if (shiftVal === 'OFF') {
                 return { status: tipe === 'IN' ? 'Masuk (OFF)' : 'Pulang (OFF)' };
               }

               // Jika ada pengaturan jam per jabatan, override jam dari roster
               if (jabatan) {
                 const jabResInRoster = getJadwalShift(jabatan);
                 if (jabResInRoster.success && jabResInRoster.data.length) {
                   const jabAktif = jabResInRoster.data.filter(function(j) {
                     return String(j.aktif).toLowerCase() === 'ya' || j.aktif === true;
                   });
                   if (jabAktif.length) {
                     shiftIn  = jabAktif[0].jamMasuk  || shiftIn;
                     shiftOut = jabAktif[0].jamPulang || shiftOut;
                     shiftTol = parseInt(jabAktif[0].toleransiMenit) || shiftTol;
                   }
                 }
               }
               
               if (shiftIn && shiftOut) {
                 return _compareAttendanceTime(tipe, jam, shiftIn, shiftOut, shiftTol); 
               }
            }
            break;
         }
       }
    }

    // 2. Prioritas Kedua: Jadwal per Jabatan (tanpa roster)
    if (jabatan) {
      const jabRes = getJadwalShift(jabatan);
      if (jabRes.success && jabRes.data.length) {
        const jabAktif = jabRes.data.filter(function(j) {
          return String(j.aktif).toLowerCase() === 'ya' || j.aktif === true || String(j.aktif).toLowerCase() === 'true';
        });
        if (jabAktif.length) {
          return _compareAttendanceTime(tipe, jam, jabAktif[0].jamMasuk, jabAktif[0].jamPulang, parseInt(jabAktif[0].toleransiMenit) || 0);
        }
      }
    }

    // 3. Fallback: Jadwal Shift per Divisi
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
 * Pembantu untuk konversi berbagai format waktu ke total menit (0-1439)
 * Mendukung string "HH:mm" atau objek Date dari Google Sheets
 */
function _parseTimeToMinutes(timeInput) {
  if (!timeInput) return 0;
  
  if (timeInput instanceof Date) {
    // SpreadsheetApp sering mengembalikan waktu sebagai objek Date
    return (timeInput.getHours() * 60) + timeInput.getMinutes();
  }
  
  const parts = String(timeInput).trim().split(':');
  if (parts.length >= 2) {
    // ParseInt robust terhadap leading zero ("08")
    const hh = parseInt(parts[0], 10) || 0;
    const mm = parseInt(parts[1], 10) || 0;
    return (hh * 60) + mm;
  }
  
  return 0;
}

/**
 * Helper untuk membandingkan jam absen dengan jadwal
 */
function _compareAttendanceTime(tipe, jamAbsen, jamMasuk, jamPulang, toleransi) {
  const jamMnt     = _parseTimeToMinutes(jamAbsen);
  const masukMnt   = _parseTimeToMinutes(jamMasuk);
  const pulangMnt  = _parseTimeToMinutes(jamPulang);
  const tol        = parseInt(toleransi) || 0;

  if (tipe === 'IN') {
    // Hadir jika jam absen <= jam masuk + toleransi
    return { status: jamMnt <= (masukMnt + tol) ? 'Hadir' : 'Terlambat' };
  } else {
    // Normalisasi untuk shift malam (jika pulang jam 05:00 pagi besoknya)
    // Jika jam masuk > jam pulang (misal 20:00 -> 05:00), maka pulang dianggap hari berikutnya
    let effectivePulangMnt = pulangMnt;
    
    // Logika Pulang: Hadir jika jam absen >= jam pulang - toleransi (toleransi pulang biasanya 0/negatif tapi kita ikuti flow)
    // Jika user absen jam 18:25 untuk jadwal 17:00, jamMnt (1105) >= 1020, maka "Pulang"
    if (jamMnt >= (effectivePulangMnt - tol)) {
      return { status: 'Pulang' };
    } else {
      return { status: 'Pulang Awal' };
    }
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
      const row = [newId, namaJadwal, divisi, shiftType, jamMasuk, jamPulang, toleransiMenit, aktif, now, now];
      sheet.appendRow(row);
      syncSheetRowToSupabase(CONFIG.SHEETS.JADWAL_SHIFT, row);
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
        // Recalculate status berdasarkan pengaturan terbaru
        const tglStr = aData[j][1] instanceof Date ? Utilities.formatDate(aData[j][1], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(aData[j][1]);
        const jamStr = aData[j][2] instanceof Date ? Utilities.formatDate(aData[j][2], Session.getScriptTimeZone(), "HH:mm:ss") : String(aData[j][2]);
        const tipe   = String(aData[j][7]).toUpperCase();
        const currentStatus = String(aData[j][10]);

        const newStatusInfo = _hitungStatusAbsensi(match[1], tipe, jamStr, match[0], tglStr, match[2] || '');
        const newStatus = newStatusInfo.status;

        if (isPlaceholder || isNameDiff || isDivDiff || currentStatus !== newStatus) {
          // Update Nama (E), Divisi (F), Jabatan (G) -> Kolom 5, 6, 7
          absSheet.getRange(rowNum, 5, 1, 3).setValues([[match[0], match[1], match[2]]]);
          // Update Status (K) -> Kolom 11
          absSheet.getRange(rowNum, 11).setValue(newStatus);
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
  var body;
  try {
    body = JSON.parse(rawContent);
  } catch (parseErr) {
    // ===========================================================
    // JIKA BUKAN JSON (ERROR PARSE), INI PASTI DARI MESIN X900
    // ===========================================================
    try {
      var ss = getSpreadsheet();
      var logSheetName = "LogMesinX900";
      var logSheet = ss.getSheetByName(logSheetName);
      if (!logSheet) {
        logSheet = ss.insertSheet(logSheetName);
        logSheet.appendRow(["Waktu Terima", "Data Mentah (Raw Text) dari Mesin"]);
        logSheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#1a3a5c").setFontColor("#ffffff");
      }
      logSheet.appendRow([new Date(), rawContent]);
      var rows_raw = rawContent.split('\n');
      var records_to_sync = [];
      rows_raw.forEach(function(line) {
        var cleanLine = line.trim();
        if (!cleanLine) return;
        var parts = cleanLine.split(/[\s,]+/);
        if (parts.length >= 3) {
          records_to_sync.push({
            fingerprintId: parts[0],
            tanggal: parts[1],
            jam: parts[2],
            tipe: (parts[3] == "0" || (parts[3]||"").toLowerCase() == "in") ? "IN" : "OUT"
          });
        }
      });
      if (records_to_sync.length > 0) syncFingerprintData(records_to_sync);
      return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
    } catch (logErr) {
      return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
    }
  }

  // 3. Eksekusi fungsi jika data adalah JSON valid
  try {
    var func = body.func;
    var args = body.args || [];
    var result;
    const context = typeof globalThis !== 'undefined' ? globalThis : this;

    if (func && typeof context[func] === 'function') {
      result = context[func].apply(null, args);
    } else {
      result = { success: false, message: 'Fungsi tidak dikenal: ' + func };
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (execErr) {
    return ContentService.createTextOutput(JSON.stringify({ 
      success: false, 
      message: 'Server Execution Error: ' + execErr.message 
    })).setMimeType(ContentService.MimeType.JSON);
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
    
    // --- Strategi Sinkronisasi (Merge) ---
    // 1. Pisahkan data bulan lain dan data bulan ini
    const otherMonthsRows = [headers];
    const currentMonthMap = {}; // Key: Nama (lower), Value: Full Row Array
    const tz = ss.getSpreadsheetTimeZone();

    for (let i = 1; i < values.length; i++) {
      let rowBulan = values[i][0];
      let rowBulanStr = "";
      if (rowBulan instanceof Date) {
        rowBulanStr = Utilities.formatDate(rowBulan, tz, "yyyy-MM");
      } else {
        rowBulanStr = String(rowBulan).trim();
      }

      if (rowBulanStr === monthYear) {
        const nameKey = String(values[i][1]).trim().toLowerCase();
        currentMonthMap[nameKey] = values[i];
      } else {
        otherMonthsRows.push(values[i]);
      }
    }

    // 2. Gabungkan data baru ke data eksisting (atau tambah baru jika belum ada)
    rosterData.forEach(item => {
      const namaAsli = item.name || item.nama;
      const nameKey = String(namaAsli).trim().toLowerCase();
      
      let row;
      if (currentMonthMap[nameKey]) {
        // Update baris eksisting
        row = currentMonthMap[nameKey];
      } else {
        // Buat baris baru, isi default kosong untuk 31 hari
        row = [monthYear, namaAsli];
        for (let d = 1; d <= 31; d++) row.push("");
        currentMonthMap[nameKey] = row;
      }

      // --- Strategi Merge (Patch): Hanya update kolom tanggal yang ada di dalam item ---
      for (let d = 1; d <= 31; d++) {
        const dStr = d.toString();
        if (Object.prototype.hasOwnProperty.call(item, dStr)) {
          const val = item[dStr];
          // Hanya perbarui jika nilainya tidak kosong (agar tidak menimpa data lama dengan kosong)
          // Jika di Excel kosong, maka data di database tidak akan berubah (Preserved)
          if (val !== undefined && val !== null && String(val).trim() !== "") {
            // Indeks kolom: Bulan(0), Nama(1), Tgl 1(2), ..., Tgl 31(32)
            row[d + 1] = String(val).trim().toUpperCase();
          }
        }
      }
    });

    // 3. Gabungkan kembali semua baris dan tulis ke sheet
    const finalValues = [...otherMonthsRows, ...Object.values(currentMonthMap)];
    sheet.clear();
    sheet.getRange(1, 1, finalValues.length, headers.length).setValues(finalValues);

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

    // Parse special dates
    let specialDates = {};
    try {
      if (result.roster_special_dates) {
        specialDates = JSON.parse(result.roster_special_dates);
      }
    } catch (e) { specialDates = {}; }

    return { 
      success: true, 
      data: {
        pagiIn: result.roster_pagi_in || "08:00",
        pagiOut: result.roster_pagi_out || "17:00",
        malamIn: result.roster_malam_in || "20:00",
        malamOut: result.roster_malam_out || "05:00",
        toleransi: parseInt(result.roster_toleransi) || 0,
        specialDates: specialDates
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

// ============================================================
// STOCK OPNAME ASSET
// ============================================================

/**
 * Ambil semua sesi Stock Opname Asset (untuk history & approval list)
 */
function getAssetOpnameSessions() {
  try {
    const ss = getSpreadsheet();
    let sheet = setupSheet(ss, 'AssetOpnameSession', ['id','tanggal','divisi','status','totalAsset','terscan','createdBy','createdAt','approvedBy','approvedAt','history']);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      result.push({
        id:         data[i][0],
        tanggal:    data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]||''),
        divisi:     data[i][2] || 'Semua',
        status:     data[i][3] || 'Draft',
        totalAsset: parseInt(data[i][4]) || 0,
        terscan:    parseInt(data[i][5]) || 0,
        createdBy:  data[i][6],
        createdAt:  data[i][7] instanceof Date ? data[i][7].toISOString() : String(data[i][7]||''),
        approvedBy: data[i][8] || '',
        approvedAt: data[i][9] || '',
        history:    data[i][10] || '[]'
      });
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

/**
 * Mulai sesi Stock Opname Asset baru
 */
function createAssetOpnameSession(tanggal, divisi, createdBy) {
  try {
    // Cari semua asset sesuai filter divisi
    const assetSheet = getSheet(CONFIG.SHEETS.ASSET_WAREHOUSE);
    const assetData = assetSheet.getDataRange().getValues();
    let totalAsset = 0;
    for (let i = 1; i < assetData.length; i++) {
      if (assetData[i].join('').trim() === '') continue;
      const aDiv = String(assetData[i][4] || '');
      if (!divisi || divisi === 'Semua' || aDiv === divisi) totalAsset++;
    }

    const ss = getSpreadsheet();
    let sheet = setupSheet(ss, 'AssetOpnameSession', ['id','tanggal','divisi','status','totalAsset','terscan','createdBy','createdAt','approvedBy','approvedAt','history']);

    const id = 'SO-AST-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
    const now = new Date().toISOString();
    const historyInit = JSON.stringify([{ action: 'Sesi dibuat', by: createdBy, time: now }]);
    const sessionRow = [id, tanggal, divisi || 'Semua', 'Draft', totalAsset, 0, createdBy, now, '', '', historyInit];
    sheet.appendRow(sessionRow);
    // AssetOpnameSession is a dynamic sheet — sync to audit_reports as closest match
    try { syncSheetRowToSupabase('AssetOpnameSession', sessionRow); } catch(_) {}
    return { success: true, id: id, totalAsset: totalAsset };
  } catch(e) { return { success: false, message: e.message }; }
}

/**
 * Ambil detail sesi + semua log scan-nya
 */
function getAssetOpnameSessionDetail(sessionId) {
  try {
    // Ambil session
    const ss = getSpreadsheet();
    const sesSheet = ss.getSheetByName('AssetOpnameSession');
    if (!sesSheet) return { success: false, message: 'Sheet sesi tidak ditemukan' };
    const sesData = sesSheet.getDataRange().getValues();
    let session = null;
    for (let i = 1; i < sesData.length; i++) {
      if (String(sesData[i][0]) === String(sessionId)) {
        session = {
          id: sesData[i][0], tanggal: sesData[i][1], divisi: sesData[i][2],
          status: sesData[i][3], totalAsset: parseInt(sesData[i][4])||0,
          terscan: parseInt(sesData[i][5])||0, createdBy: sesData[i][6],
          createdAt: sesData[i][7], approvedBy: sesData[i][8], approvedAt: sesData[i][9],
          history: sesData[i][10] || '[]'
        };
        break;
      }
    }
    if (!session) return { success: false, message: 'Sesi tidak ditemukan' };

    // Ambil log
    let logSheet = ss.getSheetByName('AssetOpnameLog');
    if (!logSheet) return { success: true, session: session, logs: [] };
    const logData = logSheet.getDataRange().getValues();
    const logs = [];
    for (let i = 1; i < logData.length; i++) {
      if (String(logData[i][1]) !== String(sessionId)) continue;
      if (logData[i].join('').trim() === '') continue;
      logs.push({
        id:         logData[i][0],
        sessionId:  logData[i][1],
        assetId:    logData[i][2],
        assetCode:  logData[i][3],
        assetNama:  logData[i][4],
        divisi:     logData[i][5],
        qtyFisik:   parseFloat(logData[i][6]) || 1,
        qtySistem:  parseFloat(logData[i][7]) || 1,
        kondisi:    logData[i][8] || 'Baik',
        catatan:    logData[i][9] || '',
        scannedBy:  logData[i][10],
        scannedAt:  logData[i][11] instanceof Date ? logData[i][11].toISOString() : String(logData[i][11]||'')
      });
    }
    return { success: true, session: session, logs: logs };
  } catch(e) { return { success: false, message: e.message }; }
}

/**
 * Rekam satu scan asset dalam sesi
 */
function scanAssetForOpname(sessionId, assetCode, qtyFisik, kondisi, catatan, scannedBy) {
  try {
    // Validasi: cari asset berdasarkan kode
    const assetSheet = getSheet(CONFIG.SHEETS.ASSET_WAREHOUSE);
    const assetData = assetSheet.getDataRange().getValues();
    let foundAsset = null;
    for (let i = 1; i < assetData.length; i++) {
      if (String(assetData[i][2]).trim().toLowerCase() === String(assetCode).trim().toLowerCase()) {
        foundAsset = {
          id: assetData[i][0], code: assetData[i][2], nama: assetData[i][3],
          divisi: assetData[i][4], qty: parseFloat(assetData[i][10]) || 1
        };
        break;
      }
    }
    if (!foundAsset) return { success: false, message: 'Kode asset tidak ditemukan: ' + assetCode };

    // Cek sudah pernah scan dalam sesi ini?
    const ss = getSpreadsheet();
    let logSheet = setupSheet(ss, 'AssetOpnameLog', ['id','sessionId','assetId','assetCode','assetNama','divisi','qtyFisik','qtySistem','kondisi','catatan','scannedBy','scannedAt']);
    const logData = logSheet.getDataRange().getValues();
    for (let i = 1; i < logData.length; i++) {
      if (String(logData[i][1]) === String(sessionId) && String(logData[i][3]).toLowerCase() === String(assetCode).trim().toLowerCase()) {
        return { success: false, message: 'Asset ini sudah di-scan dalam sesi ini', alreadyScanned: true };
      }
    }

    // Tambah log
    const id = generateId();
    const now = new Date().toISOString();
    logSheet.appendRow([id, sessionId, foundAsset.id, foundAsset.code, foundAsset.nama, foundAsset.divisi,
      parseFloat(qtyFisik)||1, foundAsset.qty, kondisi||'Baik', catatan||'', scannedBy, now]);

    // Update counter terscan di session
    const sesSheet = ss.getSheetByName('AssetOpnameSession');
    if (sesSheet) {
      const sesData = sesSheet.getDataRange().getValues();
      for (let i = 1; i < sesData.length; i++) {
        if (String(sesData[i][0]) === String(sessionId)) {
          const current = parseInt(sesData[i][5]) || 0;
          sesSheet.getRange(i+1, 6).setValue(current + 1);
          break;
        }
      }
    }
    return { success: true, asset: foundAsset };
  } catch(e) { return { success: false, message: e.message }; }
}

/**
 * Submit sesi → status Pending Approval
 */
function submitAssetOpname(sessionId, createdBy) {
  try {
    const ss = getSpreadsheet();
    const sesSheet = ss.getSheetByName('AssetOpnameSession');
    if (!sesSheet) return { success: false, message: 'Sheet sesi tidak ditemukan' };
    const sesData = sesSheet.getDataRange().getValues();
    for (let i = 1; i < sesData.length; i++) {
      if (String(sesData[i][0]) === String(sessionId)) {
        if (String(sesData[i][3]) !== 'Draft') return { success: false, message: 'Sesi ini sudah disubmit atau selesai' };
        sesSheet.getRange(i+1, 4).setValue('Pending Approval');
        let hist = [];
        try { hist = JSON.parse(sesData[i][10] || '[]'); } catch(e) {}
        hist.push({ action: 'Diajukan untuk approval', by: createdBy, time: new Date().toISOString() });
        sesSheet.getRange(i+1, 11).setValue(JSON.stringify(hist));
        return { success: true };
      }
    }
    return { success: false, message: 'Sesi tidak ditemukan' };
  } catch(e) { return { success: false, message: e.message }; }
}

/**
 * Approve atau Reject sesi Stock Opname Asset
 * Jika Approve: update qty & status tiap asset di AssetWarehouse sesuai hasil scan
 */
function approveAssetOpname(sessionId, action, approverNama) {
  try {
    const ss = getSpreadsheet();
    const sesSheet = ss.getSheetByName('AssetOpnameSession');
    if (!sesSheet) return { success: false, message: 'Sheet sesi tidak ditemukan' };
    const sesData = sesSheet.getDataRange().getValues();
    let sessionRow = -1;

    for (let i = 1; i < sesData.length; i++) {
      if (String(sesData[i][0]) === String(sessionId)) {
        sessionRow = i;
        break;
      }
    }
    if (sessionRow === -1) return { success: false, message: 'Sesi tidak ditemukan' };

    const now = new Date();
    const nowStr = now.toISOString();
    const newStatus = (action === 'Approve') ? 'Approved' : 'Ditolak';

    // Update status sesi
    sesSheet.getRange(sessionRow+1, 4).setValue(newStatus);
    sesSheet.getRange(sessionRow+1, 9).setValue(approverNama);
    sesSheet.getRange(sessionRow+1, 10).setValue(nowStr);
    let hist = [];
    try { hist = JSON.parse(sesData[sessionRow][10] || '[]'); } catch(e) {}
    hist.push({ action: action === 'Approve' ? 'Disetujui' : 'Ditolak', by: approverNama, time: nowStr });
    sesSheet.getRange(sessionRow+1, 11).setValue(JSON.stringify(hist));

    // Update status di AuditReports (Laporan SO)
    try {
      const reportSheet = getSheet(CONFIG.SHEETS.AUDIT_REPORTS);
      if (reportSheet) {
        const reportData = reportSheet.getDataRange().getValues();
        for (let r = 1; r < reportData.length; r++) {
          if (String(reportData[r][0]) === String(sessionId)) {
            reportSheet.getRange(r+1, 7).setValue(newStatus);
            let repHist = [];
            try { repHist = JSON.parse(reportData[r][9] || '[]'); } catch(e) {}
            const reportNowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
            repHist.push({ action: action === 'Approve' ? 'Laporan disetujui' : 'Laporan ditolak', by: approverNama, time: reportNowStr });
            reportSheet.getRange(r+1, 10).setValue(JSON.stringify(repHist));
            break;
          }
        }
      }
    } catch(err) {
      console.error('Gagal mengupdate status di AuditReports:', err);
    }

    // Jika Approve → update AssetWarehouse
    if (action === 'Approve') {
      const logSheet = ss.getSheetByName('AssetOpnameLog');
      if (logSheet) {
        const logData = logSheet.getDataRange().getValues();
        const assetSheet = getSheet(CONFIG.SHEETS.ASSET_WAREHOUSE);
        const assetData = assetSheet.getDataRange().getValues();
        const nowLocale = now.toLocaleString('id-ID');
        const sessionDivisi = sesData[sessionRow][2] || 'Semua';

        // 1. Dapatkan set ID asset yang terscan pada sesi ini
        const scannedAssetIds = new Set();
        for (let li = 1; li < logData.length; li++) {
          if (String(logData[li][1]) === String(sessionId)) {
            scannedAssetIds.add(String(logData[li][2]));
          }
        }

        // 2. Loop semua asset di warehouse
        for (let ai = 1; ai < assetData.length; ai++) {
          if (assetData[ai].join('').trim() === '') continue;
          const assetId = String(assetData[ai][0]);
          const assetDivisi = String(assetData[ai][4] || '');

          // Filter divisi asset agar sesuai divisi sesi SO
          if (sessionDivisi && sessionDivisi !== 'Semua' && assetDivisi !== sessionDivisi) {
            continue;
          }

          if (scannedAssetIds.has(assetId)) {
            // Asset terscan: cari log scan dan update sesuai isinya
            const logEntry = logData.find(l => String(l[1]) === String(sessionId) && String(l[2]) === assetId);
            if (logEntry) {
              const qtyFisik = parseFloat(logEntry[6]) || 1;
              const kondisi = String(logEntry[8] || 'Baik');

              assetSheet.getRange(ai+1, 11).setValue(qtyFisik); // qty (kolom 11)
              
              let newAssetStatus = assetData[ai][6] || 'Aktif';
              if (kondisi === 'Rusak') newAssetStatus = 'Tidak Aktif';
              else if (kondisi === 'Hilang') newAssetStatus = 'Hilang';
              assetSheet.getRange(ai+1, 7).setValue(newAssetStatus); // status (kolom 7)

              const oldHist = assetData[ai][9] || '';
              const entry = `🔄 SO Asset disetujui oleh ${approverNama} pada ${nowLocale}. Qty Fisik: ${qtyFisik}, Kondisi: ${kondisi}`;
              assetSheet.getRange(ai+1, 10).setValue(oldHist ? oldHist + '\n' + entry : entry); // history (kolom 10)
            }
          } else {
            // Asset TIDAK terscan: berarti Hilang!
            assetSheet.getRange(ai+1, 11).setValue(0); // qty = 0 (kolom 11)
            assetSheet.getRange(ai+1, 7).setValue('Hilang'); // status = Hilang (kolom 7)

            const oldHist = assetData[ai][9] || '';
            const entry = `❌ SO Asset disetujui oleh ${approverNama} pada ${nowLocale}. Asset tidak terscan saat opname, otomatis diubah menjadi Hilang (Qty: 0)`;
            assetSheet.getRange(ai+1, 10).setValue(oldHist ? oldHist + '\n' + entry : entry); // history (kolom 10)
          }
        }
      }
    }
    return { success: true, newStatus: newStatus };
  } catch(e) { return { success: false, message: e.message }; }
}

/**
 * Hapus sesi Draft (belum disubmit)
 */
function deleteAssetOpnameSession(sessionId) {
  try {
    const ss = getSpreadsheet();
    const sesSheet = ss.getSheetByName('AssetOpnameSession');
    if (!sesSheet) return { success: false, message: 'Sheet tidak ditemukan' };
    const sesData = sesSheet.getDataRange().getValues();
    for (let i = 1; i < sesData.length; i++) {
      if (String(sesData[i][0]) === String(sessionId)) {
        if (String(sesData[i][3]) !== 'Draft') return { success: false, message: 'Hanya sesi Draft yang dapat dihapus' };
        sesSheet.deleteRow(i+1);
        // Hapus semua log terkait
        const logSheet = ss.getSheetByName('AssetOpnameLog');
        if (logSheet) {
          const logData = logSheet.getDataRange().getValues();
          for (let li = logData.length - 1; li >= 1; li--) {
            if (String(logData[li][1]) === String(sessionId)) logSheet.deleteRow(li+1);
          }
        }
        return { success: true };
      }
    }
    return { success: false, message: 'Sesi tidak ditemukan' };
  } catch(e) { return { success: false, message: e.message }; }
}

/**
 * Ambil sesi Pending Approval untuk panel approval global
 */
function getAssetOpnamePendingApprovals() {
  try {
    const res = getAssetOpnameSessions();
    if (!res.success) return { success: true, data: [] };
    return { success: true, data: (res.data || []).filter(s => s.status === 'Pending Approval') };
  } catch(e) { return { success: false, message: e.message }; }
}

// ============================================================
// INVENTORY CONTROL & MONITORING
// ============================================================
function getInventoryControl() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.INVENTORY_CONTROL);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      result.push({
        id: data[i][0],
        tanggalPengerjaan: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][1]),
        cycleCount: data[i][2],
        lokasi: data[i][3],
        sku: data[i][4],
        batch: data[i][5],
        exp: data[i][6] instanceof Date ? Utilities.formatDate(data[i][6], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][6]),
        stockTTX: data[i][7],
        stockMabang: data[i][8],
        stockFisik: data[i][9],
        selisihMabang: data[i][10],
        selisihTTX: data[i][11],
        action: data[i][12],
        keterangan: data[i][13],
        createdBy: data[i][14],
        createdAt: data[i][15]
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function saveInventoryControl(data) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.INVENTORY_CONTROL);
    const sheetData = sheet.getDataRange().getValues();
    const now = new Date().toISOString();
    
    const rowData = [
      data.id || generateId(),
      data.tanggalPengerjaan,
      data.cycleCount,
      data.lokasi,
      data.sku,
      data.batch,
      data.exp,
      data.stockTTX,
      data.stockMabang,
      data.stockFisik,
      data.selisihMabang,
      data.selisihTTX,
      data.action,
      data.keterangan,
      data.createdBy,
      now
    ];

    if (data.id) {
      for (let i = 1; i < sheetData.length; i++) {
        if (String(sheetData[i][0]) === String(data.id)) {
          sheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
          syncSheetRowToSupabase(CONFIG.SHEETS.INVENTORY_CONTROL, rowData);
          return { success: true, id: data.id };
        }
      }
    }
    
    sheet.appendRow(rowData);
    syncSheetRowToSupabase(CONFIG.SHEETS.INVENTORY_CONTROL, rowData);
    return { success: true, id: rowData[0] };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteInventoryControl(id) {
  return deleteRow(CONFIG.SHEETS.INVENTORY_CONTROL, id);
}

function getInventoryMonitoring() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.INVENTORY_MONITORING);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      result.push({
        id: data[i][0],
        areaPosisi: data[i][1],
        keterangan: data[i][2],
        updatedAt: data[i][3]
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateInventoryMonitoring(id, areaPosisi, keterangan) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.INVENTORY_MONITORING);
    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();
    
    if (id) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(id)) {
          sheet.getRange(i + 1, 2).setValue(areaPosisi);
          sheet.getRange(i + 1, 3).setValue(keterangan);
          sheet.getRange(i + 1, 4).setValue(now);
          return { success: true };
        }
      }
    }
    
    sheet.appendRow([generateId(), areaPosisi, keterangan, now]);
    syncSheetRowToSupabase(CONFIG.SHEETS.INVENTORY_MONITORING, [generateId(), areaPosisi, keterangan, now]);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

// ============================================================
// RETURN DISTRIBUTOR
// Data disimpan di Spreadsheet Distributor (DISTRIBUTOR_QUEUE_SPREADSHEET_ID)
// Settings (daftar SKU bermasalah) tetap di spreadsheet utama (Settings sheet)
// ============================================================

// ============================================================
// RETURN DISTRIBUTOR
// Arsitektur: Header (1 baris per transaksi) + Detail (N baris SKU per transaksi)
// Kedua sheet ada di Spreadsheet Distributor
// Settings (daftar SKU bermasalah) di spreadsheet utama (sheet Settings)
// ============================================================

const RETURN_DISTRIBUTOR_SHEET        = 'ReturnDistributor';
const RETURN_DISTRIBUTOR_DETAIL_SHEET = 'ReturnDistributorDetail';

const RETURN_DISTRIBUTOR_HEADERS = [
  'id', 'noReturn', 'tanggal', 'namaDistributor', 'jenisReturn',
  'picSales', 'noMabang', 'totalSKU', 'totalQty', 'hargaOngkir', 'noResi', 'keterangan', 'createdBy', 'createdAt'
];
const RETURN_DISTRIBUTOR_DETAIL_HEADERS = [
  'id', 'returnId', 'noReturn', 'sku', 'batch', 'qty', 'expDate',
  'kategoriReturn', 'keterangan'
];

const RETURN_DIST_SETTINGS_KEYS = {
  PENARIKAN: 'returnDist_penarikan', // tidak dipakai lagi, data kini di sheet SKU Bermasalah
  BUYBACK:   'returnDist_buyback'    // tidak dipakai lagi, data kini di sheet SKU Bermasalah
};

/**
 * Fungsi test untuk memverifikasi Return Distributor dapat dimuat
 * Panggil dari Apps Script Editor untuk debugging
 */
function testReturnDistributor() {
  console.log('=== TEST RETURN DISTRIBUTOR ===');
  
  try {
    // Test 1: Setup sheet
    console.log('Test 1: Setup sheet...');
    const setupResult = setupReturnDistributorSheet();
    console.log('Setup result:', setupResult);
    
    // Test 2: Get sheet
    console.log('Test 2: Get sheet...');
    const sheet = getReturnDistributorSheet();
    console.log('Sheet name:', sheet.getName());
    console.log('Sheet ID:', sheet.getSheetId());
    
    // Test 3: Check headers
    console.log('Test 3: Check headers...');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    console.log('Current headers:', headers);
    console.log('Expected headers:', RETURN_DISTRIBUTOR_HEADERS);
    
    // Test 4: Get data
    console.log('Test 4: Get data...');
    const result = getReturnDistributor();
    console.log('Result:', JSON.stringify(result));
    
    if (result.success) {
      console.log('✅ SUCCESS: Data berhasil dimuat, jumlah:', result.data.length);
      if (result.data.length > 0) {
        console.log('Sample data:', JSON.stringify(result.data[0]));
      }
    } else {
      console.log('❌ FAILED:', result.message);
    }
    
    return result;
  } catch (e) {
    console.error('❌ ERROR:', e.message, e.stack);
    return { success: false, message: e.message, stack: e.stack };
  }
}

/**
 * Fungsi untuk memperbaiki header sheet yang tidak sesuai
 * Panggil dari Apps Script Editor jika header tidak sesuai
 */
function fixReturnDistributorHeaders() {
  try {
    console.log('=== FIX RETURN DISTRIBUTOR HEADERS ===');
    
    const ss = getReturnDistributorSpreadsheet();
    const sheet = ss.getSheetByName(RETURN_DISTRIBUTOR_SHEET);
    
    if (!sheet) {
      return { success: false, message: 'Sheet tidak ditemukan' };
    }
    
    // Baca header saat ini
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    console.log('Current headers:', currentHeaders);
    console.log('Expected headers:', RETURN_DISTRIBUTOR_HEADERS);
    
    // Update header
    sheet.getRange(1, 1, 1, RETURN_DISTRIBUTOR_HEADERS.length).setValues([RETURN_DISTRIBUTOR_HEADERS]);
    sheet.getRange(1, 1, 1, RETURN_DISTRIBUTOR_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#1a3a5c')
      .setFontColor('#ffffff');
    
    console.log('✅ Headers berhasil diperbaiki');
    return { success: true, message: 'Headers berhasil diperbaiki' };
  } catch (e) {
    console.error('❌ ERROR:', e.message);
    return { success: false, message: e.message };
  }
}

// ---- Spreadsheet distributor ----
function getReturnDistributorSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.DISTRIBUTOR_QUEUE_SPREADSHEET_ID);
}

// ---- Setup kedua sheet ----
function setupReturnDistributorSheet() {
  try {
    const ss = getReturnDistributorSpreadsheet();
    console.log('setupReturnDistributorSheet: Membuat sheet di spreadsheet:', ss.getId());
    setupSheet(ss, RETURN_DISTRIBUTOR_SHEET,        RETURN_DISTRIBUTOR_HEADERS);
    setupSheet(ss, RETURN_DISTRIBUTOR_DETAIL_SHEET, RETURN_DISTRIBUTOR_DETAIL_HEADERS);
    console.log('setupReturnDistributorSheet: Sheet berhasil dibuat');
    return { success: true };
  } catch (e) {
    console.error('setupReturnDistributorSheet Error:', e.message);
    throw e;
  }
}

function getReturnDistributorSheet() {
  try {
    const ss = getReturnDistributorSpreadsheet();
    let sheet = ss.getSheetByName(RETURN_DISTRIBUTOR_SHEET);
    if (!sheet) { 
      console.log('getReturnDistributorSheet: Sheet tidak ditemukan, membuat sheet baru');
      setupReturnDistributorSheet(); 
      sheet = ss.getSheetByName(RETURN_DISTRIBUTOR_SHEET); 
    }
    if (!sheet) {
      console.error('getReturnDistributorSheet: Gagal membuat sheet');
      throw new Error('Gagal membuat sheet ReturnDistributor');
    }
    return sheet;
  } catch (e) {
    console.error('getReturnDistributorSheet Error:', e.message);
    throw e;
  }
}

function getReturnDistributorDetailSheet() {
  try {
    const ss = getReturnDistributorSpreadsheet();
    let sheet = ss.getSheetByName(RETURN_DISTRIBUTOR_DETAIL_SHEET);
    if (!sheet) { 
      console.log('getReturnDistributorDetailSheet: Sheet tidak ditemukan, membuat sheet baru');
      setupReturnDistributorSheet(); 
      sheet = ss.getSheetByName(RETURN_DISTRIBUTOR_DETAIL_SHEET); 
    }
    if (!sheet) {
      console.error('getReturnDistributorDetailSheet: Gagal membuat sheet');
      throw new Error('Gagal membuat sheet ReturnDistributorDetail');
    }
    return sheet;
  } catch (e) {
    console.error('getReturnDistributorDetailSheet Error:', e.message);
    throw e;
  }
}

// ---- Generate nomor return: RD-YYYYMMDD-XXX ----
function generateNoReturn() {
  const sheet = getReturnDistributorSheet();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const prefix = 'RD-' + today + '-';
  const data = sheet.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const no = String(data[i][1] || '');
    if (no.startsWith(prefix)) {
      const num = parseInt(no.replace(prefix, '')) || 0;
      if (num > max) max = num;
    }
  }
  return prefix + String(max + 1).padStart(3, '0');
}

// ---- Get all return headers ----
function getReturnDistributor() {
  try {
    const sheet = getReturnDistributorSheet();
    if (!sheet) {
      console.error('getReturnDistributor: Sheet tidak ditemukan');
      return { success: false, message: 'Sheet ReturnDistributor tidak ditemukan' };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 1) {
      console.log('getReturnDistributor: Sheet kosong, mengembalikan array kosong');
      return { success: true, data: [] };
    }

    // Baca header row untuk mapping kolom secara dinamis
    const headers = data[0].map(h => String(h||'').trim().toLowerCase());
    console.log('getReturnDistributor: Headers found:', headers);
    
    const col = (name) => headers.indexOf(name.toLowerCase());

    // Mapping kolom berdasarkan header yang ditemukan
    // Header order: id, noReturn, tanggal, namaDistributor, jenisReturn, picSales, noMabang, totalSKU, totalQty, hargaOngkir, noResi, keterangan, createdBy, createdAt
    const idx = {
      id:              col('id')              >= 0 ? col('id')              : 0,
      noReturn:        col('noreturn')        >= 0 ? col('noreturn')        : 1,
      tanggal:         col('tanggal')         >= 0 ? col('tanggal')         : 2,
      namaDistributor: col('namadistributor') >= 0 ? col('namadistributor') : 3,
      jenisReturn:     col('jenisreturn')     >= 0 ? col('jenisreturn')     : 4,
      picSales:        col('picsales')        >= 0 ? col('picsales')        : 5,
      noMabang:        col('nomabang')        >= 0 ? col('nomabang')        : 6,
      totalSKU:        col('totalsku')        >= 0 ? col('totalsku')        : 7,
      totalQty:        col('totalqty')        >= 0 ? col('totalqty')        : 8,
      hargaOngkir:     col('hargaongkir')     >= 0 ? col('hargaongkir')     : 9,
      noResi:          col('noresi')          >= 0 ? col('noresi')          : 10,
      keterangan:      col('keterangan')      >= 0 ? col('keterangan')      : 11,
      createdBy:       col('createdby')       >= 0 ? col('createdby')       : 12,
      createdAt:       col('createdat')       >= 0 ? col('createdat')       : 13,
    };

    console.log('getReturnDistributor: Column mapping:', JSON.stringify(idx));

    const getVal = (row, i, def) => (i >= 0 && i < row.length) ? (row[i] || def) : def;

    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      const row = data[i];
      const tanggalRaw = getVal(row, idx.tanggal, '');
      
      const item = {
        id:              String(getVal(row, idx.id, '')),
        noReturn:        String(getVal(row, idx.noReturn, '')),
        tanggal:         tanggalRaw instanceof Date
          ? Utilities.formatDate(tanggalRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : String(tanggalRaw),
        namaDistributor: String(getVal(row, idx.namaDistributor, '')),
        jenisReturn:     String(getVal(row, idx.jenisReturn, '')),
        picSales:        String(getVal(row, idx.picSales, '')),
        noMabang:        String(getVal(row, idx.noMabang, '')),
        totalSKU:        Number(getVal(row, idx.totalSKU, 0)) || 0,
        totalQty:        Number(getVal(row, idx.totalQty, 0)) || 0,
        hargaOngkir:     Number(getVal(row, idx.hargaOngkir, 0)) || 0,
        noResi:          String(getVal(row, idx.noResi, '')),
        keterangan:      String(getVal(row, idx.keterangan, '')),
        createdBy:       String(getVal(row, idx.createdBy, '')),
        createdAt:       String(getVal(row, idx.createdAt, ''))
      };
      
      result.push(item);
    }
    console.log('getReturnDistributor: Berhasil memuat ' + result.length + ' data');
    if (result.length > 0) {
      console.log('getReturnDistributor: Sample data:', JSON.stringify(result[0]));
    }
    return { success: true, data: result };
  } catch (e) { 
    console.error('getReturnDistributor Error:', e.message, e.stack);
    return { success: false, message: 'Error: ' + e.message }; 
  }
}

// ---- Get detail by returnId ----
function getReturnDistributorDetail(returnId) {
  try {
    const sheet = getReturnDistributorDetailSheet();
    if (!sheet) {
      console.error('getReturnDistributorDetail: Sheet tidak ditemukan');
      return { success: false, message: 'Sheet ReturnDistributorDetail tidak ditemukan' };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 1) {
      console.log('getReturnDistributorDetail: Sheet kosong untuk returnId:', returnId);
      return { success: true, data: [] };
    }

    // Mapping kolom dinamis
    const headers = data[0].map(h => String(h||'').trim().toLowerCase());
    const col = (name) => headers.indexOf(name.toLowerCase());
    const idx = {
      id:             col('id')             >= 0 ? col('id')             : 0,
      returnId:       col('returnid')       >= 0 ? col('returnid')       : 1,
      noReturn:       col('noreturn')       >= 0 ? col('noreturn')       : 2,
      sku:            col('sku')            >= 0 ? col('sku')            : 3,
      batch:          col('batch')          >= 0 ? col('batch')          : 4,
      qty:            col('qty')            >= 0 ? col('qty')            : 5,
      expDate:        col('expdate')        >= 0 ? col('expdate')        : 6,
      kategoriReturn: col('kategorireturn') >= 0 ? col('kategorireturn') : 7,
      keterangan:     col('keterangan')     >= 0 ? col('keterangan')     : 8,
    };
    const getVal = (row, i, def) => (i >= 0 && i < row.length) ? (row[i] !== undefined && row[i] !== null ? row[i] : def) : def;

    const result = [];
    const rid = String(returnId).trim();

    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      const row = data[i];
      if (String(getVal(row, idx.returnId, '')).trim() !== rid) continue;

      const expRaw = getVal(row, idx.expDate, '');
      result.push({
        id:             String(getVal(row, idx.id, '')),
        returnId:       String(getVal(row, idx.returnId, '')),
        noReturn:       getVal(row, idx.noReturn, ''),
        sku:            getVal(row, idx.sku, ''),
        batch:          getVal(row, idx.batch, ''),
        qty:            getVal(row, idx.qty, ''),
        expDate:        expRaw instanceof Date
          ? Utilities.formatDate(expRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : String(expRaw || ''),
        kategoriReturn: getVal(row, idx.kategoriReturn, 'Return Normal'),
        keterangan:     getVal(row, idx.keterangan, '')
      });
    }

    console.log('getReturnDistributorDetail: Berhasil memuat ' + result.length + ' detail untuk returnId:', returnId);
    return { success: true, data: result };
  } catch (e) { 
    console.error('getReturnDistributorDetail Error:', e.message, e.stack);
    return { success: false, message: 'Error: ' + e.message }; 
  }
}

// ---- Resolve kategori ----
function resolveReturnKategori(sku, batch, defaultKategori, penarikanList, buybackList, bpomList) {
  const s = String(sku || '').trim().toLowerCase();
  const b = String(batch || '').trim().toLowerCase();
  if ((penarikanList||[]).some(x => matchSKUBatch(x, s, b))) return 'Return Penarikan';
  if ((buybackList||[]).some(x   => matchSKUBatch(x, s, b))) return 'Return Buy Back';
  if ((bpomList||[]).some(x      => matchSKUBatch(x, s, b))) return 'Return BPOM';
  return defaultKategori || 'Return Normal';
}

// ---- Save transaksi return (header + detail sekaligus) ----
// rec = { id?, tanggal, namaDistributor, keterangan, createdBy, items: [{sku,batch,qty,expDate,kategoriReturn,keterangan}] }
function saveReturnDistributor(rec) {
  try {
    // Backward compat: jika array dikirim langsung → bulk lama
    if (Array.isArray(rec)) return saveReturnDistributorBulk(rec);

    const settingsRes   = getReturnDistributorSettings();
    const penarikanList = settingsRes.success ? (settingsRes.data.penarikan || []) : [];
    const buybackList   = settingsRes.success ? (settingsRes.data.buyback  || []) : [];
    const bpomList      = settingsRes.success ? (settingsRes.data.bpom     || []) : [];
    const now           = new Date().toISOString();
    const hSheet        = getReturnDistributorSheet();
    const dSheet        = getReturnDistributorDetailSheet();

    // ---- MODE EDIT header saja (tanpa items) ----
    if (rec.id && (!rec.items || !rec.items.length)) {
      const hData = hSheet.getDataRange().getValues();
      for (let i = 1; i < hData.length; i++) {
        if (String(hData[i][0]) === String(rec.id)) {
          hSheet.getRange(i+1, 3).setValue(rec.tanggal);
          hSheet.getRange(i+1, 4).setValue(rec.namaDistributor);
          hSheet.getRange(i+1, 5).setValue(rec.jenisReturn || '');
          hSheet.getRange(i+1, 6).setValue(rec.picSales || '');
          hSheet.getRange(i+1, 7).setValue(rec.noMabang || '');
          hSheet.getRange(i+1, 10).setValue(rec.hargaOngkir || 0);
          hSheet.getRange(i+1, 11).setValue(rec.noResi || '');
          hSheet.getRange(i+1, 12).setValue(rec.keterangan || '');
          return { success: true, id: rec.id };
        }
      }
      return { success: false, message: 'Header tidak ditemukan' };
    }

    const items = rec.items || [];
    const validItems = items.filter(x => x.sku && x.batch);
    if (!validItems.length) return { success: false, message: 'Minimal 1 item SKU+Batch harus diisi' };

    const totalQty = validItems.reduce((s, x) => s + (parseFloat(x.qty) || 0), 0);

    // ---- MODE TAMBAH BARU ----
    if (!rec.id) {
      const returnId = generateId();
      const noReturn = generateNoReturn();

      // Simpan header
      hSheet.appendRow([
        returnId, noReturn, rec.tanggal, rec.namaDistributor,
        rec.jenisReturn || '', rec.picSales || '', rec.noMabang || '',
        validItems.length, totalQty, rec.hargaOngkir || 0, rec.noResi || '', rec.keterangan || '', rec.createdBy || '', now
      ]);

      // Simpan detail (batch)
      const detailRows = validItems.map(item => {
        const kat = resolveReturnKategori(item.sku, item.batch, item.kategoriReturn, penarikanList, buybackList, bpomList);
        return [generateId(), returnId, noReturn, item.sku, item.batch,
                parseFloat(item.qty)||'', item.expDate||'', kat, item.keterangan||''];
      });
      if (detailRows.length > 0) {
        dSheet.getRange(dSheet.getLastRow()+1, 1, detailRows.length, RETURN_DISTRIBUTOR_DETAIL_HEADERS.length).setValues(detailRows);
      }
      return { success: true, id: returnId, noReturn: noReturn, saved: validItems.length };
    }

    // ---- MODE UPDATE (ganti semua detail) ----
    const hData = hSheet.getDataRange().getValues();
    let noReturn = '';
    for (let i = 1; i < hData.length; i++) {
      if (String(hData[i][0]) === String(rec.id)) {
        noReturn = hData[i][1];
        hSheet.getRange(i+1, 3).setValue(rec.tanggal);
        hSheet.getRange(i+1, 4).setValue(rec.namaDistributor);
        hSheet.getRange(i+1, 5).setValue(rec.jenisReturn || '');
        hSheet.getRange(i+1, 6).setValue(rec.picSales || '');
        hSheet.getRange(i+1, 7).setValue(rec.noMabang || '');
        hSheet.getRange(i+1, 8).setValue(validItems.length);
        hSheet.getRange(i+1, 9).setValue(totalQty);
        hSheet.getRange(i+1, 10).setValue(rec.hargaOngkir || 0);
        hSheet.getRange(i+1, 11).setValue(rec.noResi || '');
        hSheet.getRange(i+1, 12).setValue(rec.keterangan || '');
        break;
      }
    }
    // Hapus detail lama
    const dData = dSheet.getDataRange().getValues();
    for (let i = dData.length - 1; i >= 1; i--) {
      if (String(dData[i][1]) === String(rec.id)) dSheet.deleteRow(i+1);
    }
    // Tulis detail baru
    const detailRows = validItems.map(item => {
      const kat = resolveReturnKategori(item.sku, item.batch, item.kategoriReturn, penarikanList, buybackList, bpomList);
      return [generateId(), rec.id, noReturn, item.sku, item.batch,
              parseFloat(item.qty)||'', item.expDate||'', kat, item.keterangan||''];
    });
    if (detailRows.length > 0) {
      dSheet.getRange(dSheet.getLastRow()+1, 1, detailRows.length, RETURN_DISTRIBUTOR_DETAIL_HEADERS.length).setValues(detailRows);
    }
    return { success: true, id: rec.id, noReturn: noReturn, saved: validItems.length };

  } catch (e) { return { success: false, message: e.message }; }
}

// ---- Delete return (header + semua detail) ----
function deleteReturnDistributor(id) {
  try {
    const hSheet = getReturnDistributorSheet();
    const dSheet = getReturnDistributorDetailSheet();
    // Hapus header
    const hData = hSheet.getDataRange().getValues();
    for (let i = hData.length - 1; i >= 1; i--) {
      if (String(hData[i][0]) === String(id)) { hSheet.deleteRow(i+1); break; }
    }
    // Hapus semua detail
    const dData = dSheet.getDataRange().getValues();
    for (let i = dData.length - 1; i >= 1; i--) {
      if (String(dData[i][1]) === String(id)) dSheet.deleteRow(i+1);
    }
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

// ---- Backward compat: bulk lama (tidak dipakai lagi tapi tetap ada) ----
function saveReturnDistributorBulk(records) {
  if (!records || !records.length) return { success: false, message: 'Tidak ada data.' };
  // Kelompokkan per tanggal+distributor menjadi 1 transaksi
  const groups = {};
  records.forEach(r => {
    const key = (r.tanggal||'') + '||' + (r.namaDistributor||'');
    if (!groups[key]) groups[key] = { tanggal: r.tanggal, namaDistributor: r.namaDistributor, keterangan: r.keterangan||'', createdBy: r.createdBy||'', items: [] };
    groups[key].items.push(r);
  });
  let totalSaved = 0;
  const results = [];
  Object.values(groups).forEach(g => {
    const res = saveReturnDistributor(g);
    if (res.success) totalSaved += (res.saved || 0);
    results.push(res);
  });
  return { success: true, saved: totalSaved, total: records.length, groups: results.length };
}

// ---- Sheet SKU Bermasalah di spreadsheet distributor ----
const RETURN_SKU_BERMASALAH_SHEET   = 'SKU Bermasalah';
const RETURN_SKU_BERMASALAH_HEADERS = ['sku', 'batch', 'manufaktur', 'tipe', 'updatedAt'];
// batch: nilai spesifik ATAU 'ALL' (semua batch dari SKU tersebut)
// tipe: 'Penarikan' | 'Buy Back' | 'BPOM'

function getReturnSKUBermasalahSheet() {
  const ss = getReturnDistributorSpreadsheet();
  let sheet = ss.getSheetByName(RETURN_SKU_BERMASALAH_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(RETURN_SKU_BERMASALAH_SHEET);
    sheet.appendRow(RETURN_SKU_BERMASALAH_HEADERS);
    sheet.getRange(1, 1, 1, RETURN_SKU_BERMASALAH_HEADERS.length)
      .setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160); // SKU
    sheet.setColumnWidth(2, 120); // Batch
    sheet.setColumnWidth(3, 160); // Manufaktur
    sheet.setColumnWidth(4, 120); // Tipe
    sheet.setColumnWidth(5, 200); // UpdatedAt
  } else {
    // Migrasi: pastikan header sudah ada kolom manufaktur
    const hdr = sheet.getRange(1, 1, 1, 5).getValues()[0];
    if (String(hdr[2]||'').toLowerCase() !== 'manufaktur') {
      // Insert kolom manufaktur di posisi 3
      sheet.insertColumnAfter(2);
      sheet.getRange(1, 3).setValue('manufaktur').setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
      sheet.setColumnWidth(3, 160);
    }
  }
  return sheet;
}

// ---- Helper: cek apakah SKU+Batch cocok (support ALL batch) ----
function matchSKUBatch(item, sku, batch) {
  const itemSKU   = String(item.sku   || '').trim().toLowerCase();
  const itemBatch = String(item.batch || '').trim().toLowerCase();
  const s = sku.toLowerCase(), b = batch.toLowerCase();
  if (itemSKU !== s) return false;
  return itemBatch === 'all' || itemBatch === b;
}

// ---- Get Return Distributor Settings (baca dari sheet SKU Bermasalah) ----
function getReturnDistributorSettings() {
  try {
    const sheet = getReturnSKUBermasalahSheet();
    const data  = sheet.getDataRange().getValues();
    const penarikan = [], buyback = [], bpom = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      const sku        = String(data[i][0] || '').trim();
      const batch      = String(data[i][1] || '').trim();
      const manufaktur = String(data[i][2] || '').trim();
      const tipe       = String(data[i][3] || '').trim().toLowerCase();
      if (!sku || !batch) continue;
      const entry = { sku, batch, manufaktur };
      if (tipe === 'penarikan')    penarikan.push(entry);
      else if (tipe === 'buy back') buyback.push(entry);
      else if (tipe === 'bpom')    bpom.push(entry);
    }
    return { success: true, data: { penarikan, buyback, bpom } };
  } catch (e) { return { success: false, message: e.message }; }
}

// ---- Save Return Distributor Settings (tulis ulang sheet SKU Bermasalah) ----
function saveReturnDistributorSettings(penarikanList, buybackList, bpomList) {
  try {
    const sheet = getReturnSKUBermasalahSheet();
    const now   = new Date().toISOString();

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

    const rows = [];
    const addRows = (list, tipeLabel) => {
      (list || []).forEach(item => {
        if (!item.sku || !item.batch) return;
        rows.push([item.sku, item.batch, item.manufaktur || '', tipeLabel, now]);
      });
    };
    addRows(penarikanList, 'Penarikan');
    addRows(buybackList,   'Buy Back');
    addRows(bpomList,      'BPOM');

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, RETURN_SKU_BERMASALAH_HEADERS.length).setValues(rows);
      rows.forEach((row, i) => {
        const rowNum = i + 2;
        const color  = row[3] === 'Penarikan' ? '#fce8e6'
                     : row[3] === 'Buy Back'  ? '#fef9e7'
                     : row[3] === 'BPOM'      ? '#f3e8ff'
                     : '#ffffff';
        sheet.getRange(rowNum, 1, 1, RETURN_SKU_BERMASALAH_HEADERS.length).setBackground(color);
        // Bold batch ALL
        if (String(row[1]).toUpperCase() === 'ALL') {
          sheet.getRange(rowNum, 2).setFontWeight('bold').setFontColor('#ef4444');
        }
      });
    }

    return { success: true, message: `${rows.length} SKU berhasil disimpan ke sheet "${RETURN_SKU_BERMASALAH_SHEET}".` };
  } catch (e) { return { success: false, message: e.message }; }
}

// ---- Check if SKU+Batch is flagged ----
function checkReturnDistributorSKU(sku, batch) {
  try {
    const res = getReturnDistributorSettings();
    if (!res.success) return { success: true, kategori: 'Return Normal', flagged: false };
    const s = String(sku || '').trim().toLowerCase();
    const b = String(batch || '').trim().toLowerCase();
    const inPenarikan = (res.data.penarikan || []).some(item => matchSKUBatch(item, s, b));
    const inBuyback   = (res.data.buyback   || []).some(item => matchSKUBatch(item, s, b));
    const inBPOM      = (res.data.bpom      || []).some(item => matchSKUBatch(item, s, b));
    if (inPenarikan) return { success: true, kategori: 'Return Penarikan', flagged: true, type: 'penarikan' };
    if (inBuyback)   return { success: true, kategori: 'Return Buy Back',  flagged: true, type: 'buyback'   };
    if (inBPOM)      return { success: true, kategori: 'Return BPOM',      flagged: true, type: 'bpom'      };
    return { success: true, kategori: 'Return Normal', flagged: false };
  } catch (e) { return { success: false, message: e.message }; }
}

// ============================================================
// PIC SALES RETURN DISTRIBUTOR (disimpan di Settings spreadsheet utama)
// ============================================================
const RD_PIC_SALES_KEY = 'returnDist_picSalesList';

function getReturnDistributorPICSales() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SETTINGS);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]||'').trim() === RD_PIC_SALES_KEY) {
        let list = [];
        try { list = JSON.parse(String(data[i][1]||'[]')); } catch(e) { list = []; }
        return { success: true, data: list };
      }
    }
    return { success: true, data: [] };
  } catch (e) { return { success: false, message: e.message }; }
}

function saveReturnDistributorPICSales(picList) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SETTINGS);
    const data  = sheet.getDataRange().getValues();
    const now   = new Date().toISOString();
    const val   = JSON.stringify((picList || []).filter(x => x && x.trim()));
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]||'').trim() === RD_PIC_SALES_KEY) {
        sheet.getRange(i+1, 2).setValue(val);
        sheet.getRange(i+1, 3).setValue(now);
        return { success: true };
      }
    }
    sheet.appendRow([RD_PIC_SALES_KEY, val, now]);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

// ============================================================
// PETTY CASH
// ============================================================

function setupPettyCashSheets() {
  const ss = getSpreadsheet();
  setupSheet(ss, CONFIG.SHEETS.PETTY_CASH_PERIOD, [
    'id', 'nama', 'tanggalMulai', 'tanggalSelesai', 'saldoAwal',
    'keterangan', 'status', 'createdBy', 'createdAt'
  ]);
  setupSheet(ss, CONFIG.SHEETS.PETTY_CASH, [
    'id', 'periodId', 'tanggal', 'kategori', 'keterangan',
    'nominal', 'tipe', 'buktiUrl', 'createdBy', 'createdAt', 'statusBayar'
  ]);
  return { success: true };
}

// --- PERIOD ---
function getPettyCashPeriods() {
  try {
    setupPettyCashSheets(); // Pastikan header selalu ada
    const sheet = getSheet(CONFIG.SHEETS.PETTY_CASH_PERIOD);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      result.push({
        id: data[i][0],
        nama: data[i][1],
        tanggalMulai: data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][2]),
        tanggalSelesai: data[i][3] instanceof Date ? Utilities.formatDate(data[i][3], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][3]),
        saldoAwal: parseFloat(data[i][4]) || 0,
        keterangan: data[i][5] || '',
        status: data[i][6] || 'Aktif',
        createdBy: data[i][7],
        createdAt: data[i][8]
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function addPettyCashPeriod(nama, tanggalMulai, tanggalSelesai, saldoAwal, keterangan, createdBy) {
  try {
    const id = generateId();
    const row = [
      id, nama, tanggalMulai, tanggalSelesai,
      parseFloat(saldoAwal) || 0, keterangan || '', 'Aktif',
      createdBy, new Date().toISOString()
    ];
    getSheet(CONFIG.SHEETS.PETTY_CASH_PERIOD).appendRow(row);
    syncSheetRowToSupabase(CONFIG.SHEETS.PETTY_CASH_PERIOD, row);
    return { success: true, id: id };
  } catch (e) { return { success: false, message: e.message }; }
}

function updatePettyCashPeriodStatus(id, status) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PETTY_CASH_PERIOD);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i + 1, 7).setValue(status);
        return { success: true };
      }
    }
    return { success: false, message: 'Periode tidak ditemukan' };
  } catch (e) { return { success: false, message: e.message }; }
}

function deletePettyCashPeriod(id) {
  return deleteRow(CONFIG.SHEETS.PETTY_CASH_PERIOD, id);
}

// --- TRANSACTIONS ---
function getPettyCash(periodId) {
  try {
    setupPettyCashSheets(); // Pastikan header selalu ada
    const sheet = getSheet(CONFIG.SHEETS.PETTY_CASH);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      if (periodId && String(data[i][1]) !== String(periodId)) continue;
      result.push({
        id: data[i][0],
        periodId: data[i][1],
        tanggal: data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][2]),
        kategori: data[i][3],
        keterangan: data[i][4],
        nominal: parseFloat(data[i][5]) || 0,
        tipe: data[i][6] || 'OUT',
        buktiUrl: data[i][7] || '',
        createdBy: data[i][8],
        createdAt: data[i][9],
        statusBayar: data[i][10] || 'Belum Bayar'
      });
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function addPettyCash(periodId, tanggal, kategori, keterangan, nominal, tipe, buktiUrl, createdBy) {
  try {
    const id = generateId();
    const row = [
      id, periodId, tanggal, kategori, keterangan,
      parseFloat(nominal) || 0, tipe || 'OUT',
      buktiUrl || '', createdBy, new Date().toISOString(), 'Belum Bayar'
    ];
    getSheet(CONFIG.SHEETS.PETTY_CASH).appendRow(row);
    syncSheetRowToSupabase(CONFIG.SHEETS.PETTY_CASH, row);
    return { success: true, id: id };
  } catch (e) { return { success: false, message: e.message }; }
}

function deletePettyCash(id) {
  return deleteRow(CONFIG.SHEETS.PETTY_CASH, id);
}

function updatePettyCashStatus(id, status) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PETTY_CASH);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i + 1, 11).setValue(status);
        return { success: true };
      }
    }
    return { success: false, message: 'Transaksi tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Ambil atau buat subfolder per periode di dalam folder Petty Cash Drive.
 * Struktur: PettyCash Drive Root → [Nama Periode]
 */
function getPettyCashPeriodFolder(periodNama) {
  try {
    const root = DriveApp.getFolderById(CONFIG.PETTY_CASH_FOLDER_ID);
    const safeName = (periodNama || 'Umum').replace(/[\/\\:*?"<>|]/g, '-').trim();
    const iter = root.getFoldersByName(safeName);
    return iter.hasNext() ? iter.next() : root.createFolder(safeName);
  } catch (e) {
    // Fallback ke folder Drive utama jika folder Petty Cash tidak bisa diakses
    Logger.log('Error accessing PETTY_CASH_FOLDER_ID, using DRIVE_FOLDER_ID: ' + e.message);
    const root = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    const safeName = 'PettyCash_' + (periodNama || 'Umum').replace(/[\/\\:*?"<>|]/g, '-').trim();
    const iter = root.getFoldersByName(safeName);
    return iter.hasNext() ? iter.next() : root.createFolder(safeName);
  }
}

/**
 * Upload atau ganti bukti untuk transaksi petty cash yang sudah ada.
 * File disimpan ke subfolder periode di Google Drive Petty Cash.
 * 
 * @param {string} txId - ID transaksi
 * @param {string} base64Data - Base64 data file (kosong jika hanya update URL)
 * @param {string} fileName - Nama file
 * @param {string} mimeType - MIME type
 * @param {string} existingUrl - URL yang sudah ada (jika hanya update tanpa upload baru)
 */
function uploadPettyCashBukti(txId, base64Data, fileName, mimeType, existingUrl) {
  let file = null;
  let fileUrl = existingUrl || '';
  let folderUrl = '';
  
  try {
    // Cari nama periode dari transaksi
    const sheet = getSheet(CONFIG.SHEETS.PETTY_CASH);
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let periodId = '';
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(txId)) {
        rowIndex = i;
        periodId = String(data[i][1]);
        break;
      }
    }
    if (rowIndex === -1) return { success: false, message: 'Transaksi tidak ditemukan' };

    // Jika existingUrl diberikan, langsung update database tanpa upload
    if (existingUrl && existingUrl.trim() !== '') {
      sheet.getRange(rowIndex + 1, 8).setValue(existingUrl);
      return { success: true, url: existingUrl, folderUrl: '', fileName: fileName || 'bukti' };
    }

    // Jika tidak ada base64Data, return error
    if (!base64Data || base64Data.trim() === '') {
      return { success: false, message: 'Tidak ada file untuk diupload' };
    }

    // Ambil nama periode
    let periodNama = 'Umum';
    try {
      const pSheet = getSheet(CONFIG.SHEETS.PETTY_CASH_PERIOD);
      const pData = pSheet.getDataRange().getValues();
      for (let i = 1; i < pData.length; i++) {
        if (String(pData[i][0]) === periodId) { periodNama = String(pData[i][1]); break; }
      }
    } catch(e) {
      Logger.log('Error getting period name: ' + e.message);
    }

    // STEP 1: Upload file ke Drive dulu (ke root My Drive untuk menghindari error akses folder)
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
    file = DriveApp.createFile(blob);
    
    // STEP 2: Coba pindahkan ke subfolder periode (jika berhasil)
    try {
      const folder = getPettyCashPeriodFolder(periodNama);
      file.moveTo(folder);
      folderUrl = folder.getUrl();
      Logger.log('File berhasil dipindahkan ke folder: ' + periodNama);
    } catch(e) {
      // Jika gagal pindah folder, biarkan di root - file tetap terupload
      Logger.log('File tetap di root Drive, tidak bisa pindah ke folder: ' + e.message);
      folderUrl = 'https://drive.google.com/drive/folders/' + CONFIG.DRIVE_FOLDER_ID;
    }
    
    // STEP 3: Set sharing permission dengan metode yang lebih reliable
    const fileId = file.getId();
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      Logger.log('File sharing permission set to ANYONE_WITH_LINK');
      
      // Tambahan: Set sharing secara eksplisit dengan Drive API
      Drive.Permissions.insert(
        {
          'type': 'anyone',
          'role': 'reader',
          'withLink': true
        },
        fileId,
        {
          'supportsAllDrives': true
        }
      );
      Logger.log('Drive API permission set successfully');
    } catch(e) {
      Logger.log('Warning: Tidak bisa set sharing permission: ' + e.message);
    }
    
    // Tunggu sebentar agar permission propagate
    Utilities.sleep(500);
    
    // STEP 4: Generate URL
    fileUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;

    // STEP 5: Update kolom buktiUrl (kolom 8) di sheet
    sheet.getRange(rowIndex + 1, 8).setValue(fileUrl);
    
    return { success: true, url: fileUrl, folderUrl: folderUrl, fileName: fileName };
    
  } catch (e) {
    Logger.log('Error in uploadPettyCashBukti: ' + e.message);
    
    // Jika file sudah dibuat tapi ada error di langkah lain, tetap return success dengan URL file
    if (file && file.getId()) {
      try {
        fileUrl = 'https://drive.google.com/uc?export=view&id=' + file.getId();
        const sheet = getSheet(CONFIG.SHEETS.PETTY_CASH);
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][0]) === String(txId)) {
            sheet.getRange(i + 1, 8).setValue(fileUrl);
            break;
          }
        }
        return { success: true, url: fileUrl, folderUrl: 'https://drive.google.com/drive/my-drive', fileName: fileName };
      } catch(e2) {
        Logger.log('Error saving file URL: ' + e2.message);
      }
    }
    
    return { success: false, message: 'Upload gagal: ' + e.message };
  }
}

/**
 * Export data Petty Cash periode ke file Excel (.xlsx) dan simpan ke Google Drive.
 * Mengembalikan URL file di Drive.
 */
function exportPettyCashToGDrive(periodId) {
  try {
    // Ambil data periode
    const pSheet = getSheet(CONFIG.SHEETS.PETTY_CASH_PERIOD);
    const pData = pSheet.getDataRange().getValues();
    let period = null;
    for (let i = 1; i < pData.length; i++) {
      if (String(pData[i][0]) === String(periodId)) {
        period = {
          id: pData[i][0],
          nama: pData[i][1],
          tanggalMulai: pData[i][2] instanceof Date ? Utilities.formatDate(pData[i][2], Session.getScriptTimeZone(), 'dd/MM/yyyy') : String(pData[i][2]),
          tanggalSelesai: pData[i][3] instanceof Date ? Utilities.formatDate(pData[i][3], Session.getScriptTimeZone(), 'dd/MM/yyyy') : String(pData[i][3]),
          saldoAwal: parseFloat(pData[i][4]) || 0
        };
        break;
      }
    }
    if (!period) return { success: false, message: 'Periode tidak ditemukan' };

    // Ambil transaksi periode ini
    const txSheet = getSheet(CONFIG.SHEETS.PETTY_CASH);
    const txData = txSheet.getDataRange().getValues();
    const txs = [];
    for (let i = 1; i < txData.length; i++) {
      if (txData[i].join('').trim() === '') continue;
      if (String(txData[i][1]) !== String(periodId)) continue;
      txs.push({
        tanggal: txData[i][2] instanceof Date ? Utilities.formatDate(txData[i][2], Session.getScriptTimeZone(), 'dd/MM/yyyy') : String(txData[i][2]),
        kategori: txData[i][3],
        keterangan: txData[i][4],
        nominal: parseFloat(txData[i][5]) || 0,
        tipe: txData[i][6] || 'OUT',
        buktiUrl: txData[i][7] || '',
        createdBy: txData[i][8]
      });
    }
    // Sort by tanggal
    txs.sort((a, b) => new Date(a.tanggal.split('/').reverse().join('-')) - new Date(b.tanggal.split('/').reverse().join('-')));

    // Build HTML table untuk Excel
    let running = period.saldoAwal;
    let totalOut = 0, totalIn = 0;

    const formatRpGs = (n) => 'Rp ' + Number(n).toLocaleString('id-ID');

    let rows = '';
    txs.forEach((d, i) => {
      if (d.tipe === 'IN') { running += d.nominal; totalIn += d.nominal; }
      else { running -= d.nominal; totalOut += d.nominal; }
      rows += `<tr>
        <td>${i + 1}</td>
        <td>${d.tanggal}</td>
        <td>${d.tipe}</td>
        <td>${d.kategori}</td>
        <td>${d.keterangan}</td>
        <td style="mso-number-format:'#,##0';">${d.tipe === 'OUT' ? -d.nominal : d.nominal}</td>
        <td style="mso-number-format:'#,##0';">${running}</td>
        <td>${d.buktiUrl ? d.buktiUrl : '-'}</td>
        <td>${d.createdBy}</td>
      </tr>`;
    });

    const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:x="urn:schemas-microsoft-com:office:excel"
      xmlns="http://www.w3.org/TR/REC-html40">
    <head><meta charset="UTF-8">
    <!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets>
    <x:ExcelWorksheet><x:Name>Petty Cash</x:Name>
    <x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>
    </x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
    <style>
      th { background:#1e3a5f; color:#fff; font-weight:bold; padding:6px; }
      td { padding:5px; }
      .footer { font-weight:bold; background:#f3f4f6; }
    </style>
    </head><body>
    <table>
      <tr><td colspan="9" style="font-size:16pt;font-weight:bold;color:#1e3a5f;">LAPORAN PETTY CASH - ${period.nama}</td></tr>
      <tr><td colspan="9">Periode: ${period.tanggalMulai} s/d ${period.tanggalSelesai}</td></tr>
      <tr><td colspan="9">Saldo Awal: ${formatRpGs(period.saldoAwal)}</td></tr>
      <tr></tr>
      <tr>
        <th>No</th><th>Tanggal</th><th>Tipe</th><th>Kategori</th>
        <th>Keterangan</th><th>Nominal</th><th>Saldo Berjalan</th>
        <th>Link Bukti</th><th>Input By</th>
      </tr>
      ${rows}
      <tr class="footer">
        <td colspan="5" style="text-align:right;">TOTAL PENGELUARAN</td>
        <td style="mso-number-format:'#,##0';">${-totalOut}</td>
        <td colspan="3"></td>
      </tr>
      <tr class="footer">
        <td colspan="5" style="text-align:right;">TOTAL PEMASUKAN</td>
        <td style="mso-number-format:'#,##0';">${totalIn}</td>
        <td colspan="3"></td>
      </tr>
      <tr class="footer">
        <td colspan="5" style="text-align:right;">SISA SALDO AKHIR</td>
        <td style="mso-number-format:'#,##0';">${period.saldoAwal + totalIn - totalOut}</td>
        <td colspan="3"></td>
      </tr>
    </table>
    </body></html>`;

    // Simpan ke subfolder periode di Drive
    const folder = getPettyCashPeriodFolder(period.nama);
    const fileName = 'PettyCash_' + period.nama.replace(/[\/\\:*?"<>|]/g, '-') + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm') + '.xls';
    const blob = Utilities.newBlob(html, 'application/vnd.ms-excel', fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return {
      success: true,
      fileUrl: file.getUrl(),
      folderUrl: folder.getUrl(),
      fileName: fileName,
      periodNama: period.nama
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getPettyCashFull() {
  try {
    const periods = getPettyCashPeriods();
    const txAll = getPettyCash(null);
    if (!periods.success || !txAll.success) return { success: false, message: 'Gagal memuat data' };
    return { success: true, periods: periods.data, transactions: txAll.data };
  } catch (e) { return { success: false, message: e.message }; }
}

// ============================================================
// FETCH MULTIPLE DRIVE FILES AS BASE64 (untuk Print Bukti Foto)
// ============================================================
/**
 * Mengambil beberapa file dari Google Drive dan mengembalikannya
 * sebagai data URI base64 agar bisa di-embed langsung di HTML print
 * tanpa perlu autentikasi ulang.
 *
 * @param {string[]} fileIds - Array of Google Drive file IDs
 * @returns {Object} { success, images: { [fileId]: { dataUri, mimeType } | null } }
 */
function getFilesAsBase64(fileIds) {
  try {
    const result = {};
    (fileIds || []).forEach(function(fileId) {
      if (!fileId) return;
      try {
        const file = DriveApp.getFileById(fileId);
        const blob = file.getBlob();
        const mimeType = blob.getContentType() || 'image/jpeg';
        const b64 = Utilities.base64Encode(blob.getBytes());
        result[fileId] = { dataUri: 'data:' + mimeType + ';base64,' + b64, mimeType: mimeType };
      } catch (e) {
        // File tidak bisa diakses (permission, tidak ada, dll)
        result[fileId] = null;
      }
    });
    return { success: true, images: result };
  } catch (e) {
    return { success: false, message: e.message, images: {} };
  }
}


// ============================================================
// TOKO GUDANG
// ============================================================

/**
 * Get all Payment Gudang records
 */
function getPaymentGudang() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PAYMENT_GUDANG);
    const data = sheet.getDataRange().getValues();
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({
        id: String(data[i][0]),
        nama: String(data[i][1] || ''),
        deskripsi: String(data[i][2] || ''),
        hargaPerOrang: Number(data[i][3] || 0),
        deadline: data[i][4] instanceof Date ? Utilities.formatDate(data[i][4], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][4] || ''),
        status: String(data[i][5] || 'Aktif'),
        createdBy: String(data[i][6] || ''),
        createdAt: data[i][7],
        midtransOrderId: String(data[i][8] || ''),
        midtransStatus: String(data[i][9] || ''),
        tipe: String(data[i][10] || 'Reguler')
      });
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Add new Payment Gudang
 */
function addPaymentGudang(nama, deskripsi, hargaPerOrang, deadline, createdBy, tipe) {
  try {
    const id = generateId();
    const sheet = getSheet(CONFIG.SHEETS.PAYMENT_GUDANG);
    const row = [
      id, nama, deskripsi,
      Number(hargaPerOrang) || 0,
      deadline, 'Aktif', createdBy,
      new Date().toISOString(),
      '', '', tipe || 'Reguler'
    ];
    sheet.appendRow(row);
    syncSheetRowToSupabase(CONFIG.SHEETS.PAYMENT_GUDANG, row);
    return { success: true, id: id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Get or create participant for MISTINE, syncing with employee account if necessary
 */
function getOrCreateParticipantForUser(paymentId, karyawanNama) {
  try {
    // 1. Check if participant already exists
    const sheetPart = getSheet(CONFIG.SHEETS.PAYMENT_GUDANG_PARTICIPANTS);
    const dataPart = sheetPart.getDataRange().getValues();
    for (let i = 1; i < dataPart.length; i++) {
      if (String(dataPart[i][1]) === String(paymentId) && String(dataPart[i][3]).toLowerCase() === String(karyawanNama).toLowerCase()) {
        return {
          success: true,
          participantId: String(dataPart[i][0]),
          karyawanId: String(dataPart[i][2]),
          namaKaryawan: String(dataPart[i][3])
        };
      }
    }

    // 2. Not found, search in Karyawan sheet to sync details
    let karyawanId = 'K-' + generateId();
    const sheetKar = getSheet(CONFIG.SHEETS.KARYAWAN);
    const dataKar = sheetKar.getDataRange().getValues();
    for (let j = 1; j < dataKar.length; j++) {
      if (String(dataKar[j][1]).toLowerCase() === String(karyawanNama).toLowerCase()) {
        karyawanId = String(dataKar[j][0]); // Found match, sync it!
        break;
      }
    }

    // 3. Create new participant
    const id = generateId();
    const partRow = [id, paymentId, karyawanId, karyawanNama, 'Belum Bayar', '', '', '', ''];
    sheetPart.appendRow(partRow);
    syncSheetRowToSupabase(CONFIG.SHEETS.PAYMENT_GUDANG_PARTICIPANTS, partRow);

    return {
      success: true,
      participantId: id,
      karyawanId: karyawanId,
      namaKaryawan: karyawanNama
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Update Payment Gudang status
 */
function updatePaymentGudangStatus(id, status) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PAYMENT_GUDANG);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i + 1, 6).setValue(status);
        return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Update Payment Gudang product details
 */
function updatePaymentGudang(id, nama, deskripsi, hargaPerOrang, deadline, tipe, status) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PAYMENT_GUDANG);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i + 1, 2).setValue(nama);
        sheet.getRange(i + 1, 3).setValue(deskripsi || '');
        sheet.getRange(i + 1, 4).setValue(Number(hargaPerOrang) || 0);
        sheet.getRange(i + 1, 5).setValue(deadline || '');
        sheet.getRange(i + 1, 6).setValue(status || 'Aktif');
        sheet.getRange(i + 1, 11).setValue(tipe || 'Reguler');
        return { success: true };
      }
    }
    return { success: false, message: 'Product tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Delete Payment Gudang
 */
function deletePaymentGudang(id) {
  return deleteRow(CONFIG.SHEETS.PAYMENT_GUDANG, id);
}

/**
 * Get participants for a payment
 */
function getPaymentParticipants(paymentId) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PAYMENT_GUDANG_PARTICIPANTS);
    const data = sheet.getDataRange().getValues();
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (String(data[i][1]) === String(paymentId)) {
        result.push({
          id: String(data[i][0]),
          paymentId: String(data[i][1]),
          karyawanId: String(data[i][2] || ''),
          namaKaryawan: String(data[i][3] || ''),
          statusBayar: String(data[i][4] || 'Belum Bayar'),
          tanggalBayar: data[i][5] ? (data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : String(data[i][5])) : '',
          metodeBayar: String(data[i][6] || ''),
          buktiUrl: String(data[i][7] || ''),
          midtransTransactionId: String(data[i][8] || '')
        });
      }
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Add participant to payment
 */
function addPaymentParticipant(paymentId, karyawanId, namaKaryawan) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PAYMENT_GUDANG_PARTICIPANTS);
    const id = generateId();
    const row = [id, paymentId, karyawanId, namaKaryawan, 'Belum Bayar', '', '', '', ''];
    sheet.appendRow(row);
    syncSheetRowToSupabase(CONFIG.SHEETS.PAYMENT_GUDANG_PARTICIPANTS, row);
    return { success: true, id: id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Add multiple participants at once
 */
function addPaymentParticipantsBulk(paymentId, karyawanList) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PAYMENT_GUDANG_PARTICIPANTS);
    let count = 0;
    
    karyawanList.forEach(k => {
      const id = generateId();
      const row = [id, paymentId, k.id || '', k.nama || '', 'Belum Bayar', '', '', '', ''];
      sheet.appendRow(row);
      syncSheetRowToSupabase(CONFIG.SHEETS.PAYMENT_GUDANG_PARTICIPANTS, row);
      count++;
    });
    
    return { success: true, count: count };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Update participant payment status
 */
function updateParticipantPaymentStatus(participantId, statusBayar, metodeBayar, buktiUrl, midtransTransactionId) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PAYMENT_GUDANG_PARTICIPANTS);
    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(participantId)) {
        sheet.getRange(i + 1, 5).setValue(statusBayar);
        sheet.getRange(i + 1, 6).setValue(now);
        sheet.getRange(i + 1, 7).setValue(metodeBayar || '');
        sheet.getRange(i + 1, 8).setValue(buktiUrl || '');
        sheet.getRange(i + 1, 9).setValue(midtransTransactionId || '');
        return { success: true };
      }
    }
    return { success: false, message: 'Participant tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Delete participant
 */
function deletePaymentParticipant(id) {
  return deleteRow(CONFIG.SHEETS.PAYMENT_GUDANG_PARTICIPANTS, id);
}

/**
 * Get Payment Gudang with participants (full data)
 */
function getPaymentGudangFull() {
  try {
    const paymentsRes = getPaymentGudang();
    if (!paymentsRes.success) return paymentsRes;
    
    const payments = paymentsRes.data;
    
    // Get participants for each payment
    payments.forEach(payment => {
      const participantsRes = getPaymentParticipants(payment.id);
      payment.participants = participantsRes.success ? participantsRes.data : [];
      
      // Calculate statistics
      const total = payment.participants.length;
      const lunas = payment.participants.filter(p => p.statusBayar === 'Lunas').length;
      const belum = total - lunas;
      
      payment.stats = {
        total: total,
        lunas: lunas,
        belum: belum,
        totalNominal: total * payment.hargaPerOrang,
        terkumpul: lunas * payment.hargaPerOrang,
        sisa: belum * payment.hargaPerOrang
      };
    });
    
    return { success: true, data: payments };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// MIDTRANS CONFIGURATION
// ============================================================

/**
 * Get Midtrans configuration
 */
function getMidtransConfig() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.MIDTRANS_CONFIG);
    const data = sheet.getDataRange().getValues();
    
    if (data.length > 1 && data[1][0]) {
      return {
        success: true,
        config: {
          id: String(data[1][0]),
          serverKey: String(data[1][1] || ''),
          clientKey: String(data[1][2] || ''),
          isProduction: Boolean(data[1][3]),
          updatedBy: String(data[1][4] || ''),
          updatedAt: data[1][5]
        }
      };
    }
    
    // Return empty config if not set
    return {
      success: true,
      config: {
        id: '',
        serverKey: '',
        clientKey: '',
        isProduction: false,
        updatedBy: '',
        updatedAt: ''
      }
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Save Midtrans configuration
 */
function saveMidtransConfig(serverKey, clientKey, isProduction, updatedBy) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.MIDTRANS_CONFIG);
    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();
    
    if (data.length > 1 && data[1][0]) {
      // Update existing
      sheet.getRange(2, 2).setValue(serverKey);
      sheet.getRange(2, 3).setValue(clientKey);
      sheet.getRange(2, 4).setValue(isProduction);
      sheet.getRange(2, 5).setValue(updatedBy);
      sheet.getRange(2, 6).setValue(now);
    } else {
      // Create new
      const id = generateId();
      const mcRow = [id, serverKey, clientKey, isProduction, updatedBy, now];
      sheet.appendRow(mcRow);
      syncSheetRowToSupabase(CONFIG.SHEETS.MIDTRANS_CONFIG, mcRow);
    }
    
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// MIDTRANS INTEGRATION
// ============================================================

/**
 * Create Midtrans Snap transaction
 */
function createMidtransTransaction(paymentId, participantId, amount, customerName, customerEmail, paymentMethodType) {
  try {
    const configRes = getMidtransConfig();
    if (!configRes.success) return configRes;
    
    const config = configRes.config;
    if (!config.serverKey) {
      return { success: false, message: 'Midtrans belum dikonfigurasi. Silakan setting API Key terlebih dahulu.' };
    }
    
    // Generate order ID
    const orderId = 'PG-' + paymentId.substring(0, 8) + '-' + participantId.substring(0, 8) + '-' + Date.now();
    
    // Midtrans API endpoint
    const apiUrl = config.isProduction 
      ? 'https://app.midtrans.com/snap/v1/transactions'
      : 'https://app.sandbox.midtrans.com/snap/v1/transactions';
    
    // Prepare transaction data
    const transactionData = {
      transaction_details: {
        order_id: orderId,
        gross_amount: amount
      },
      customer_details: {
        first_name: customerName,
        email: customerEmail || 'noreply@gudangfcl.com'
      },
      item_details: [{
        id: paymentId,
        price: amount,
        quantity: 1,
        name: 'MISTINE'
      }]
    };

    if (paymentMethodType) {
      if (paymentMethodType === 'va') {
        transactionData.enabled_payments = ["bca_va", "bni_va", "bri_va", "other_va", "permata_va", "echannel"];
      } else if (paymentMethodType === 'qr') {
        transactionData.enabled_payments = ["gopay", "shopeepay", "qris"];
      } else if (paymentMethodType === 'gopay') {
        transactionData.enabled_payments = ["gopay", "qris"];
      } else if (paymentMethodType === 'shopeepay') {
        transactionData.enabled_payments = ["shopeepay", "qris"];
      } else if (paymentMethodType === 'ovo') {
        transactionData.enabled_payments = ["gopay", "shopeepay", "qris"];
      } else if (paymentMethodType === 'cstore') {
        transactionData.enabled_payments = ["cstore"];
      }
    }
    
    // Create authorization header
    const authString = Utilities.base64Encode(config.serverKey + ':');
    
    // Make API request
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Basic ' + authString,
        'Accept': 'application/json'
      },
      payload: JSON.stringify(transactionData),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();
    
    Logger.log('Midtrans Response Code: ' + responseCode);
    Logger.log('Midtrans Response Body: ' + responseBody);
    
    if (responseCode === 201) {
      const result = JSON.parse(responseBody);
      
      // Update payment record with Midtrans order ID
      updateParticipantMidtransId(participantId, orderId);
      
      return {
        success: true,
        token: result.token,
        redirectUrl: result.redirect_url,
        orderId: orderId
      };
    } else {
      const error = JSON.parse(responseBody);
      return {
        success: false,
        message: 'Midtrans Error: ' + (error.error_messages ? error.error_messages.join(', ') : 'Unknown error')
      };
    }
  } catch (e) {
    Logger.log('createMidtransTransaction error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * Update participant with Midtrans transaction ID
 */
function updateParticipantMidtransId(participantId, midtransTransactionId) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PAYMENT_GUDANG_PARTICIPANTS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(participantId)) {
        sheet.getRange(i + 1, 9).setValue(midtransTransactionId);
        return { success: true };
      }
    }
    return { success: false, message: 'Participant tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Handle Midtrans payment notification (webhook)
 */
function handleMidtransNotification(notificationData) {
  try {
    const orderId = notificationData.order_id;
    const transactionStatus = notificationData.transaction_status;
    const fraudStatus = notificationData.fraud_status;
    
    Logger.log('Midtrans Notification - Order ID: ' + orderId + ', Status: ' + transactionStatus);
    
    // Find participant by Midtrans transaction ID
    const sheet = getSheet(CONFIG.SHEETS.PAYMENT_GUDANG_PARTICIPANTS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][8]) === String(orderId)) {
        let statusBayar = 'Belum Bayar';
        let metodeBayar = notificationData.payment_type || '';
        
        // Determine payment status based on Midtrans status
        if (transactionStatus === 'capture') {
          if (fraudStatus === 'accept') {
            statusBayar = 'Lunas';
          }
        } else if (transactionStatus === 'settlement') {
          statusBayar = 'Lunas';
        } else if (transactionStatus === 'pending') {
          statusBayar = 'Pending';
        } else if (transactionStatus === 'deny' || transactionStatus === 'expire' || transactionStatus === 'cancel') {
          statusBayar = 'Gagal';
        }
        
        // Update participant status
        sheet.getRange(i + 1, 5).setValue(statusBayar);
        sheet.getRange(i + 1, 6).setValue(new Date().toISOString());
        sheet.getRange(i + 1, 7).setValue(metodeBayar);
        
        Logger.log('Updated participant payment status to: ' + statusBayar);
        
        return { success: true, status: statusBayar };
      }
    }
    
    return { success: false, message: 'Order ID tidak ditemukan: ' + orderId };
  } catch (e) {
    Logger.log('handleMidtransNotification error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * Check Midtrans transaction status
 */
function checkMidtransStatus(orderId) {
  try {
    const configRes = getMidtransConfig();
    if (!configRes.success) return configRes;
    
    const config = configRes.config;
    if (!config.serverKey) {
      return { success: false, message: 'Midtrans belum dikonfigurasi' };
    }
    
    const apiUrl = config.isProduction
      ? 'https://api.midtrans.com/v2/' + orderId + '/status'
      : 'https://api.sandbox.midtrans.com/v2/' + orderId + '/status';
    
    const authString = Utilities.base64Encode(config.serverKey + ':');
    
    const options = {
      method: 'get',
      headers: {
        'Authorization': 'Basic ' + authString,
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();
    
    if (responseCode === 200) {
      const result = JSON.parse(responseBody);
      return {
        success: true,
        status: result.transaction_status,
        data: result
      };
    } else {
      return { success: false, message: 'Failed to check status' };
    }
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// PUSAT APPROVAL (APPROVAL CENTER) FUNCTIONS
// ==========================================

function getApprovalCenterData(username) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.STOCK_CONTROL);
    const data = sheet.getDataRange().getValues();
    const result = [];
    let pending = 0, approved = 0, rejected = 0, total = 0;

    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      const row = data[i];
      const statusRaw = String(row[7] || '');
      
      if (statusRaw !== 'Menunggu Approval' && statusRaw !== 'Approved' && statusRaw !== 'Ditolak' && statusRaw !== 'Disetujui' && statusRaw !== 'Rejected') {
        continue;
      }
      
      let status = 'pending';
      if (statusRaw === 'Menunggu Approval') {
        status = 'pending';
        pending++;
      } else if (statusRaw === 'Approved' || statusRaw === 'Disetujui') {
        status = 'approved';
        approved++;
      } else if (statusRaw === 'Ditolak' || statusRaw === 'Rejected') {
        status = 'rejected';
        rejected++;
      }
      total++;

      result.push({
        id: row[0],
        tanggal: row[1] instanceof Date ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : row[1],
        pic: row[2],
        area: row[3],
        kategori: row[4],
        temuan: row[5],
        qty: '-',
        status: status
      });
    }

    return { 
      success: true, 
      approvals: result.reverse(),
      stats: { pending, approved, rejected, total }
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}
this['getApprovalCenterData'] = getApprovalCenterData;

function approveStockAdjustment(id, username) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.STOCK_CONTROL);
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let kategori = '';
    let existingLog = '';

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        rowIndex = i + 1;
        kategori = data[i][4];
        existingLog = data[i][10] || '';
        break;
      }
    }
    if (rowIndex === -1) throw new Error('Data tidak ditemukan');

    // Update status di Master
    sheet.getRange(rowIndex, 8).setValue('Approved');
    const nowLocal = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    const newLog = `✅ Approved by ${username} pada ${nowLocal}`;
    sheet.getRange(rowIndex, 11).setValue(newLog + (existingLog ? '\n' + existingLog : ''));

    // Update ke Stock Jika Kategori = Stock Opname
    if (kategori === 'Stock Opname') {
      const detailSheet = getSheet(CONFIG.SHEETS.STOCK_CONTROL_DETAIL);
      const detailData = detailSheet.getDataRange().getValues();
      const stockSheet = getSheet(CONFIG.SHEETS.STOCK);
      const stockData = stockSheet.getDataRange().getValues();

      for (let j = 1; j < detailData.length; j++) {
        if (String(detailData[j][1]) === String(id)) {
          const sku = String(detailData[j][3] || '');
          const aksi = String(detailData[j][11] || '');
          const f = parseFloat(detailData[j][8]) || 0;
          const fs = parseFloat(detailData[j][13]) || 0;
          const finalStock = (detailData[j][13] !== '' && detailData[j][13] !== undefined && detailData[j][13] !== null) ? fs : f;

          if (aksi === 'Adjust Stock' && sku) {
            // Find in STOCK and adjust
            for (let k = 1; k < stockData.length; k++) {
              if (String(stockData[k][1]) === sku || String(stockData[k][3]) === sku) {
                stockSheet.getRange(k + 1, 8).setValue(finalStock); // Kolom 8 = stok
                break;
              }
            }
          }
        }
      }
    }

    return { success: true, message: 'Berhasil di-approve' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}
this['approveStockAdjustment'] = approveStockAdjustment;

function rejectStockAdjustment(id, username, reason) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.STOCK_CONTROL);
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let existingLog = '';

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        rowIndex = i + 1;
        existingLog = data[i][10] || '';
        break;
      }
    }
    if (rowIndex === -1) throw new Error('Data tidak ditemukan');

    // Update status di Master
    sheet.getRange(rowIndex, 8).setValue('Ditolak');
    const nowLocal = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    const newLog = `❌ Ditolak by ${username} pada ${nowLocal}. Alasan: ${reason}`;
    sheet.getRange(rowIndex, 11).setValue(newLog + (existingLog ? '\n' + existingLog : ''));

    return { success: true, message: 'Berhasil ditolak' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}
this['rejectStockAdjustment'] = rejectStockAdjustment;

// ============================================================
// NOTIFIKASI SISTEM
// ============================================================

/**
 * Mengambil notifikasi dari Settings sheet key 'notifikasi_data'.
 * Notifikasi disimpan sebagai JSON array di kolom value.
 */
function getNotifikasi() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SETTINGS);
    if (!sheet) return { success: true, data: [] };

    const data = sheet.getDataRange().getValues();
    let notifRaw = '';
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'notifikasi_data') {
        notifRaw = data[i][1] || '[]';
        break;
      }
    }

    let notifArr = [];
    try { notifArr = JSON.parse(notifRaw); } catch (e) { notifArr = []; }

    // Filter: hanya tampilkan notifikasi yang belum kadaluarsa (7 hari terakhir)
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 7);
    notifArr = notifArr.filter(n => {
      if (!n.tanggal) return true;
      return new Date(n.tanggal) >= cutoff;
    });

    // Tambah notifikasi otomatis dari data sistem
    const autoNotifs = _generateAutoNotifications();
    notifArr = autoNotifs.concat(notifArr);

    return { success: true, data: notifArr };
  } catch (e) {
    return { success: false, message: e.message, data: [] };
  }
}

/**
 * Generate notifikasi otomatis berdasarkan data terkini.
 */
function _generateAutoNotifications() {
  const notifs = [];
  const now = new Date();
  const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  try {
    // 1. Notifikasi approval pending
    const ijinSheet = getSheet(CONFIG.SHEETS.IJIN);
    if (ijinSheet) {
      const ijinData = ijinSheet.getDataRange().getValues();
      let pendingIjin = 0;
      for (let i = 1; i < ijinData.length; i++) {
        const status = String(ijinData[i][6] || '');
        if (status.startsWith('Pending')) pendingIjin++;
      }
      if (pendingIjin > 0) {
        notifs.push({
          id: 'auto_ijin_pending',
          title: 'Approval Ijin Menunggu',
          text: pendingIjin + ' pengajuan ijin/cuti menunggu persetujuan Anda',
          type: 'warning',
          page: 'approval',
          targetAkses: 'approval',
          tanggal: now.toISOString()
        });
      }
    }
  } catch (e) { /* silent */ }

  try {
    // 2. Notifikasi stok rendah
    const stockSheet = getSheet(CONFIG.SHEETS.STOCK);
    if (stockSheet) {
      const stockData = stockSheet.getDataRange().getValues();
      let lowStock = 0;
      for (let i = 1; i < stockData.length; i++) {
        const stok = Number(stockData[i][7]) || 0;
        const stokMin = Number(stockData[i][8]) || 0;
        if (stokMin > 0 && stok <= stokMin) lowStock++;
      }
      if (lowStock > 0) {
        notifs.push({
          id: 'auto_stock_low',
          title: 'Stok Rendah',
          text: lowStock + ' item mencapai batas stok minimum',
          type: 'danger',
          page: 'stock',
          targetAkses: 'stock',
          tanggal: now.toISOString()
        });
      }
    }
  } catch (e) { /* silent */ }

  return notifs;
}

/**
 * Simpan notifikasi manual ke Settings sheet.
 */
function saveNotifikasi(notifJson, username) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.SETTINGS);
    if (!sheet) return { success: false, message: 'Sheet Settings tidak ditemukan' };

    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();
    let found = false;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'notifikasi_data') {
        sheet.getRange(i + 1, 2).setValue(notifJson);
        sheet.getRange(i + 1, 3).setValue(now);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow(['notifikasi_data', notifJson, now]);
    }

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// UPLOAD FILE INSTANT (untuk file kecil < 1MB)
// ============================================================

/**
 * Upload file langsung (tanpa chunking) untuk file kecil.
 * @param {string} base64Data - Data file dalam format base64
 * @param {string} fileName   - Nama file
 * @param {string} mimeType   - MIME type file
 * @param {string} folderName - Nama folder tujuan di Drive
 */
function uploadFileInstant(base64Data, fileName, mimeType, folderName) {
  try {
    if (!base64Data) return { success: false, message: 'Data file kosong' };

    const folder = getOrCreateBuktiFolder(folderName || 'FCL_Uploads');
    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, mimeType || 'application/octet-stream', fileName || 'file');
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return {
      success: true,
      url: 'https://drive.google.com/file/d/' + file.getId() + '/view',
      fileId: file.getId()
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// ROSTER SPECIAL DATES
// ============================================================

/**
 * Menyimpan jadwal khusus tanggal untuk Roster ke Settings sheet.
 * @param {string} jsonStr - JSON string dari objek specialDates
 * @param {string} username - Username yang menyimpan
 */
function saveRosterSpecialDates(jsonStr, username) {
  try {
    if (username && !checkPermission(username, 'jadwalShift')) {
      return { success: false, message: 'Akses ditolak. Anda tidak memiliki izin mengubah jadwal shift.' };
    }

    // Validasi JSON
    try { JSON.parse(jsonStr); } catch (e) {
      return { success: false, message: 'Format data tidak valid: ' + e.message };
    }

    const sheet = getSheet(CONFIG.SHEETS.SETTINGS);
    if (!sheet) return { success: false, message: 'Sheet Settings tidak ditemukan' };

    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();
    let found = false;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'roster_special_dates') {
        sheet.getRange(i + 1, 2).setValue(jsonStr);
        sheet.getRange(i + 1, 3).setValue(now);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow(['roster_special_dates', jsonStr, now]);
    }

    return { success: true, message: 'Jadwal khusus berhasil disimpan' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// REKAP ONGKIR MISTINE
// ============================================================

// Helper untuk membersihkan nilai angka dari sheet (menangani format Rp, titik, koma)
function parseSheetNumber(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  
  let str = String(val).trim();
  if (str === '' || str === '-' || str === 'null') return 0;
  
  // Ambil hanya angka, titik, dan koma
  let clean = str.replace(/[^0-9,.]/g, '');
  
  // Logika cerdas untuk membedakan pemisah ribuan dan desimal
  if (clean.includes('.') && clean.includes(',')) {
    // Jika ada keduanya, biasanya format ID (titik=ribu, koma=desimal) atau US (kebalikannya)
    if (clean.indexOf('.') < clean.indexOf(',')) {
      // ID Style: 1.234.567,89
      clean = clean.replace(/\./g, '').replace(',', '.');
    } else {
      // US Style: 1,234,567.89
      clean = clean.replace(/,/g, '');
    }
  } else if (clean.includes(',')) {
    // Hanya ada koma
    let parts = clean.split(',');
    if (parts[parts.length - 1].length === 3) {
      // Kemungkinan besar pemisah ribuan: 95,000
      clean = clean.replace(/,/g, '');
    } else {
      // Kemungkinan besar desimal: 95,5
      clean = clean.replace(',', '.');
    }
  } else if (clean.includes('.')) {
    // Hanya ada titik
    let parts = clean.split('.');
    if (parts[parts.length - 1].length === 3) {
      // Kemungkinan besar pemisah ribuan: 95.000
      clean = clean.replace(/\./g, '');
    }
    // Jika bukan 3 digit, biarkan titik sebagai desimal
  }
  
  const num = parseFloat(clean);
  return isNaN(num) ? 0 : num;
}

function getRekapOngkirMistineData(bulanFilter) {
  try {
    const ss = getDistributorQueueSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.DISTRIBUTOR_QUEUE_MISTINE);
    if (!sheet) return { success: false, message: 'Sheet ANTRIAN MISTINE tidak ditemukan' };

    const data = sheet.getDataRange().getValues();
    const result = [];

    // Parse bulan filter (format: YYYY-MM)
    let filterYear = null, filterMonth = null;
    if (bulanFilter && bulanFilter.includes('-')) {
      const parts = bulanFilter.split('-');
      filterYear = parseInt(parts[0]);
      filterMonth = parseInt(parts[1]);
    }

    // Header index mapping (berdasarkan DISTRIBUTOR_QUEUE_HEADERS)
    // Index 22: nomor resi (kolom W), Index 24: Harga Ongkir (kolom Y), Index 25: Harga Ongkir Ekspedisi (kolom Z)
    
    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;
      
      // Filter berdasarkan bulan jika ada
      if (filterYear && filterMonth) {
        const orderDate = data[i][1];
        if (orderDate) {
          let dateObj;
          if (orderDate instanceof Date) {
            dateObj = orderDate;
          } else {
            dateObj = new Date(orderDate);
          }
          
          if (dateObj && !isNaN(dateObj.getTime())) {
            const rowYear = dateObj.getFullYear();
            const rowMonth = dateObj.getMonth() + 1;
            if (rowYear !== filterYear || rowMonth !== filterMonth) continue;
          }
        }
      }
      
      result.push({
        rowNumber: i + 1,
        orderQueueTime: formatDistributorQueueDate(data[i][1], false),
        picSales: String(data[i][2] || ''),
        namaDistributor: String(data[i][3] || ''),
        noMabang: String(data[i][7] || ''),
        metodePengiriman: String(data[i][8] || ''),
        nomorResi: String(data[i][22] || ''),
        hargaOngkir: parseSheetNumber(data[i][24]), // Kolom Y
        hargaOngkirEkspedisi: parseSheetNumber(data[i][25]), // Kolom Z
        poNumber: String(data[i][6] || '')
      });
    }

    return { success: true, data: result.reverse() }; // Terbaru di atas
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function saveRekapOngkirMistine(rowNumber, nomorResi, hargaOngkir, hargaOngkirEkspedisi, updatedBy) {
  try {
    const ss = getDistributorQueueSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.DISTRIBUTOR_QUEUE_MISTINE);
    if (!sheet) return { success: false, message: 'Sheet ANTRIAN MISTINE tidak ditemukan' };

    const rNum = parseInt(rowNumber);
    if (isNaN(rNum) || rNum < 2) return { success: false, message: 'Row number tidak valid' };

    // Kolom 23 = W (nomor resi), Kolom 25 = Y (Harga Ongkir), Kolom 26 = Z (Harga Ongkir Ekspedisi)
    sheet.getRange(rNum, 23).setValue(nomorResi || ''); // Kolom W
    const rangeOngkir = sheet.getRange(rNum, 25);
    const rangeEkspedisi = sheet.getRange(rNum, 26);
    
    rangeOngkir.setValue(parseFloat(hargaOngkir) || 0);
    rangeEkspedisi.setValue(parseFloat(hargaOngkirEkspedisi) || 0);
    
    // Set format Rupiah
    rangeOngkir.setNumberFormat('"Rp"#,##0');
    rangeEkspedisi.setNumberFormat('"Rp"#,##0');

    SpreadsheetApp.flush();
    return { success: true, message: 'Data ongkir berhasil diperbarui' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// PAYMENT KOL INSTANT (MISTINE)
// ============================================================

function getPaymentKOLInstantData(bulanFilter) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PAYMENT_KOL_INSTANT);
    if (!sheet) return { success: false, message: 'Sheet Payment KOL Instant tidak ditemukan' };

    const data = sheet.getDataRange().getValues();
    const result = [];

    // Parse bulan filter (format: YYYY-MM)
    let filterYear = null, filterMonth = null;
    if (bulanFilter && bulanFilter.includes('-')) {
      const parts = bulanFilter.split('-');
      filterYear = parseInt(parts[0]);
      filterMonth = parseInt(parts[1]);
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i].join('').trim() === '') continue;

      const tgl = data[i][1];
      if (filterYear && filterMonth && tgl) {
        const dateObj = (tgl instanceof Date) ? tgl : new Date(tgl);
        if (dateObj && !isNaN(dateObj.getTime())) {
          if (dateObj.getFullYear() !== filterYear || (dateObj.getMonth() + 1) !== filterMonth) continue;
        }
      }

      result.push({
        id: data[i][0],
        tanggal: tgl,
        noOrder: String(data[i][2] || ''),
        noResi: String(data[i][3] || ''),
        harga: parseSheetNumber(data[i][4]),
        createdBy: data[i][5],
        createdAt: data[i][6]
      });
    }

    return { success: true, data: result.reverse() };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function savePaymentKOLInstant(id, noOrder, noResi, harga, createdBy) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PAYMENT_KOL_INSTANT);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };

    const now = new Date();
    const tgl = now.toISOString().split('T')[0]; // Default tanggal hari ini

    if (id) {
      // Edit
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(id)) {
          sheet.getRange(i + 1, 3).setValue(noOrder || '');
          sheet.getRange(i + 1, 4).setValue(noResi || '');
          const rangeHarga = sheet.getRange(i + 1, 5);
          rangeHarga.setValue(parseFloat(harga) || 0);
          rangeHarga.setNumberFormat('"Rp"#,##0');
          SpreadsheetApp.flush();
          return { success: true, message: 'Data berhasil diperbarui' };
        }
      }
      return { success: false, message: 'ID tidak ditemukan' };
    } else {
      // Add
      const newId = generateId();
      const lastRow = sheet.getLastRow();
      const newRow = [
        newId, tgl, noOrder || '', noResi || '',
        parseFloat(harga) || 0, createdBy || '', now.toISOString()
      ];
      sheet.appendRow(newRow);
      syncSheetRowToSupabase(CONFIG.SHEETS.PAYMENT_KOL_INSTANT, newRow);
      sheet.getRange(lastRow + 1, 5).setNumberFormat('"Rp"#,##0');
      return { success: true, message: 'Data berhasil ditambahkan' };
    }
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deletePaymentKOLInstant(id) {
  return deleteRow(CONFIG.SHEETS.PAYMENT_KOL_INSTANT, id);
}