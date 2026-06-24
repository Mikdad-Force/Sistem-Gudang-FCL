document.getElementById('ubkDropZone').addEventListener('click', () => document.getElementById('ubkFileInput').click());
              document.getElementById('ubkDropZone').addEventListener('dragover', e => { e.preventDefault(); e.target.closest('#ubkDropZone').style.borderColor = 'var(--teal)'; });
              document.getElementById('ubkDropZone').addEventListener('dragleave', e => { e.preventDefault(); e.target.closest('#ubkDropZone').style.borderColor = 'var(--border-light)'; });
              document.getElementById('ubkDropZone').addEventListener('drop', e => {
                e.preventDefault();
                e.target.closest('#ubkDropZone').style.borderColor = 'var(--border-light)';
                if (e.dataTransfer.files.length) handleFileSelectBookingPayment({ files: e.dataTransfer.files });
              });

// ==========================================
    // VERCEL / EXTERNAL HOSTING POLYFILL PROXY
    // ==========================================
    // Jika aplikasi ini dijalankan di luar Google Apps Script (misal di Vercel),
    // objek `google` tidak akan ada. Skrip ini akan membuat jembatan ajaib (Proxy)
    // agar semua fungsi google.script.run otomatis dirubah menjadi request HTTP POST.
    if (typeof google === 'undefined' || typeof google.script === 'undefined') {
      console.warn("🚀 Menjalankan mode External Hosting (Vercel). Mengaktifkan Proxy API.");

      // ⚠️ GANTI URL DI BAWAH INI DENGAN URL WEB APP DEPLOYMENT ANDA ⚠️
      const GAS_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbyZK-mfXbHaPq3xuQBqta40Z_1Kx6V2Q2crXujHFmLOsJBlSyHc5Ix28se0Bkvk4sgpcA/exec";

      window.google = window.google || {};
      window.google.script = window.google.script || {};

      // Mocking Host API (biasanya digunakan untuk scrollTo top)
      window.google.script.host = {
        close: function () { console.log("Host close requested"); },
        scrollTo: function (x, y) { window.scrollTo(x, y); }
      };

      // Proxy interceptor untuk google.script.run
      window.google.script.run = new Proxy({}, {
        get: function (target, prop) {
          // Fungsi utama untuk memanggil Web App via fetch()
          const callApi = function (funcName, args, successCb, failureCb) {
            if (GAS_WEBAPP_URL === "ISI_URL_WEB_APP_ANDA_DISINI") {
              console.error("URL WEB APP BELUM DIISI! Silakan masukkan URL Web App Google Apps Script ke dalam variabel GAS_WEBAPP_URL.");
              alert("Sistem belum terhubung ke database. Cek console.");
              if (failureCb) failureCb("URL_NOT_SET");
              return;
            }

            fetch(GAS_WEBAPP_URL, {
              method: 'POST',
              mode: 'cors',
              redirect: 'follow',
              body: JSON.stringify({ func: funcName, args: args }),
              headers: { 'Content-Type': 'text/plain' } // text/plain digunakan untuk bypass preflight CORS
            })
              .then(res => {
                if (!res.ok) throw new Error("HTTP " + res.status);
                return res.json();
              })
              .then(data => {
                if (successCb) successCb(data);
              })
              .catch(err => {
                console.error("❌ Vercel GAS API Error:", err);
                if (failureCb) failureCb(err);
              });
          };

          // Handle chain syntax: google.script.run.withSuccessHandler(...)
          if (prop === 'withSuccessHandler') {
            return function (successCb) {
              return new Proxy({}, {
                get: function (t, funcName) {
                  if (funcName === 'withFailureHandler') {
                    return function (failureCb) {
                      return new Proxy({}, {
                        get: function (t2, finalFunc) {
                          return function (...args) { callApi(finalFunc, args, successCb, failureCb); };
                        }
                      });
                    };
                  }
                  return function (...args) { callApi(funcName, args, successCb, null); };
                }
              });
            };
          }

          // Handle chain syntax: google.script.run.withFailureHandler(...)
          if (prop === 'withFailureHandler') {
            return function (failureCb) {
              return new Proxy({}, {
                get: function (t, funcName) {
                  if (funcName === 'withSuccessHandler') {
                    return function (successCb) {
                      return new Proxy({}, {
                        get: function (t2, finalFunc) { return function (...args) { callApi(finalFunc, args, successCb, failureCb); }; }
                      });
                    };
                  }
                  return function (...args) { callApi(funcName, args, null, failureCb); };
                }
              });
            };
          }

          // Panggilan langsung (direct call) tanpa handler
          return function (...args) {
            callApi(prop, args, null, null);
          };
        }
      });
    }
    // ==========================================

    const SUPABASE_URL = 'https://hnofmrmwkropijhpexpx.supabase.co';
    const SUPABASE_KEY = 'sb_publishable_nu_0fpE0B_dnH3Z6d6dkIA_rWE9xGV8';
    let supabaseClient = null;

    function initSupabase() {
      if (!supabaseClient && window.supabase) {
        try {
          supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY);
        } catch (err) {
          console.warn('Supabase init error:', err);
        }
      }
    }
    initSupabase();

    function callGASPromise(funcName, ...args) {
      return new Promise((resolve, reject) => {
        if (!window.google || !window.google.script || !window.google.script.run) {
          reject(new Error('Google Apps Script tidak tersedia.'));
          return;
        }
        try {
          window.google.script.run
            .withSuccessHandler(resolve)
            .withFailureHandler(err => reject(err))[funcName](...args);
        } catch (err) {
          reject(err);
        }
      });
    }

    // ============================================================
    // SUPABASE HELPER FUNCTIONS
    // Nama tabel selalu di-lowercase sesuai getSupabaseTableName() di Kode.gs
    // Contoh: 'KasGudang' → 'kasgudang', 'TeamBuilding' → 'teambuilding'
    // ============================================================

    function toSupabaseTableName(name) {
      if (!name) return '';
      return String(name).trim().replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_]/g, '').toLowerCase();
    }

    async function supabaseSelect(table, columns = '*', filters = {}, order = null, limit = null) {
      if (!supabaseClient) initSupabase();
      if (!supabaseClient) return { success: false, message: 'Supabase client belum siap.' };
      try {
        const tbl = toSupabaseTableName(table);
        let query = supabaseClient.from(tbl).select(columns);
        Object.entries(filters || {}).forEach(([key, value]) => {
          if (value === null || value === undefined) return;
          query = query.eq(key, value);
        });
        if (order && order.column) {
          query = query.order(order.column, { ascending: order.ascending !== false });
        }
        if (limit != null) {
          query = query.limit(limit);
        }
        const { data, error } = await query;
        if (error) return { success: false, message: error.message };
        return { success: true, data };
      } catch (err) {
        return { success: false, message: err.message || String(err) };
      }
    }

    // Upsert (insert or update) satu baris ke Supabase langsung dari Web App.
    // Dipakai setelah GAS berhasil simpan agar Supabase langsung up-to-date tanpa menunggu onEdit trigger.
    async function supabaseUpsert(table, rowObject) {
      if (!supabaseClient) initSupabase();
      if (!supabaseClient) return { success: false, message: 'Supabase client belum siap.' };
      try {
        const tbl = toSupabaseTableName(table);
        const { error } = await supabaseClient.from(tbl).upsert(rowObject, { onConflict: 'id' });
        if (error) return { success: false, message: error.message };
        return { success: true };
      } catch (err) {
        return { success: false, message: err.message || String(err) };
      }
    }

    // Hapus satu baris dari Supabase langsung dari Web App berdasarkan id.
    async function supabaseDelete(table, id) {
      if (!supabaseClient) initSupabase();
      if (!supabaseClient) return { success: false, message: 'Supabase client belum siap.' };
      try {
        const tbl = toSupabaseTableName(table);
        const { error } = await supabaseClient.from(tbl).delete().eq('id', id);
        if (error) return { success: false, message: error.message };
        return { success: true };
      } catch (err) {
        return { success: false, message: err.message || String(err) };
      }
    }

    // Update kolom tertentu di Supabase berdasarkan filter key=value.
    async function supabaseUpdate(table, id, updates) {
      if (!supabaseClient) initSupabase();
      if (!supabaseClient) return { success: false, message: 'Supabase client belum siap.' };
      try {
        const tbl = toSupabaseTableName(table);
        const { error } = await supabaseClient.from(tbl).update(updates).eq('id', id);
        if (error) return { success: false, message: error.message };
        return { success: true };
      } catch (err) {
        return { success: false, message: err.message || String(err) };
      }
    }

    // Ambil data dengan filter custom (range, ilike, dll).
    async function supabaseQuery(table, builderFn) {
      if (!supabaseClient) initSupabase();
      if (!supabaseClient) return { success: false, message: 'Supabase client belum siap.' };
      try {
        const tbl = toSupabaseTableName(table);
        let query = supabaseClient.from(tbl).select('*');
        if (typeof builderFn === 'function') query = builderFn(query);
        const { data, error } = await query;
        if (error) return { success: false, message: error.message };
        return { success: true, data };
      } catch (err) {
        return { success: false, message: err.message || String(err) };
      }
    }

    async function fetchWithFallback(supabaseTask, gasFuncName, gasArgs = []) {
      let supaResult = { success: false };
      try {
        supaResult = await supabaseTask();
      } catch (err) {
        supaResult = { success: false, message: err.message || String(err) };
      }
      if (supaResult && supaResult.success) {
        return supaResult;
      }
      console.warn('Supabase fallback ke GAS:', supaResult.message || supaResult.error || 'Tidak ada respon Supabase.');
      try {
        const gasRes = await callGASPromise(gasFuncName, ...gasArgs);
        return gasRes;
      } catch (err) {
        return { success: false, message: 'GAS fallback gagal: ' + (err.message || String(err)) };
      }
    }

    // Resilience: Global Error Logger
    window.onerror = function (msg, url, line, col, error) {
      console.error('GLOBAL_ERROR:', msg, 'at', line + ':' + col);
      const el = document.getElementById('sbAttStatus');
      if (el && el.textContent.includes('Memuat')) {
        el.innerHTML = '<span style="color:#ef4444;font-size:9px;">System Error</span>';
      }
      return false;
    };

    // Auto-Trigger for Dashboard Ready
    window.addEventListener('DOMContentLoaded', () => {
      setTimeout(() => {
        if (typeof loadMyAttendanceSummary === 'function' && document.getElementById('sbAttStatus')?.textContent?.includes('Memuat')) {
          console.log('DOMContentLoaded Trigger: Retrying attendance load...');
          loadMyAttendanceSummary();
        }
      }, 5000);
    });
    // Resilience: Global Error Logger
    window.onerror = function (msg, url, line, col, error) {
      console.error('GLOBAL_ERROR:', msg, 'at', line + ':' + col);
      const el = document.getElementById('sbAttStatus');
      if (el && el.textContent.includes('Memuat')) {
        el.innerHTML = '<span style="color:#ef4444;font-size:9px;">System Error</span>';
      }
      return false;
    };

    // Auto-Trigger for Dashboard Ready
    window.addEventListener('DOMContentLoaded', () => {
      setTimeout(() => {
        if (typeof loadMyAttendanceSummary === 'function' && document.getElementById('sbAttStatus')?.textContent?.includes('Memuat')) {
          console.log('DOMContentLoaded Trigger: Retrying attendance load...');
          loadMyAttendanceSummary();
        }
      }, 5000);
    });
    // Resilience: Global Error Logger
    window.onerror = function (msg, url, line, col, error) {
      console.error('GLOBAL_ERROR:', msg, 'at', line + ':' + col);
      const el = document.getElementById('sbAttStatus');
      if (el && el.textContent.includes('Memuat')) {
        el.innerHTML = '<span style="color:#ef4444;font-size:9px;">System Error</span>';
      }
      return false;
    };

    // Auto-Trigger for Dashboard Ready
    window.addEventListener('DOMContentLoaded', () => {
      setTimeout(() => {
        if (typeof loadMyAttendanceSummary === 'function' && document.getElementById('sbAttStatus')?.textContent?.includes('Memuat')) {
          console.log('DOMContentLoaded Trigger: Retrying attendance load...');
          loadMyAttendanceSummary();
        }
      }, 5000);
    });
    // Resilience: Global Error Logger
    window.onerror = function (msg, url, line, col, error) {
      console.error('GLOBAL_ERROR:', msg, 'at', line + ':' + col);
      const el = document.getElementById('sbAttStatus');
      if (el && el.textContent.includes('Memuat')) {
        el.innerHTML = '<span style="color:#ef4444;font-size:9px;">System Error</span>';
      }
      return false;
    };

    // Auto-Trigger for Dashboard Ready
    window.addEventListener('DOMContentLoaded', () => {
      setTimeout(() => {
        if (typeof loadMyAttendanceSummary === 'function' && document.getElementById('sbAttStatus')?.textContent?.includes('Memuat')) {
          console.log('DOMContentLoaded Trigger: Retrying attendance load...');
          loadMyAttendanceSummary();
        }
      }, 5000);
    });
    const i18nData = {
      id: {
        dashboard: "Dashboard",
        kasGudang: "Kas Gudang",
        teamBuilding: "Team Building",
        expense: "Expense List",
        paymentGudang: "Payment Gudang MISTINE",
        pgAddTransaction: "Tambah Transaksi MISTINE",
        pgSettingPayment: "Setting Payment MISTINE",
        pgToggleStatus: "Close/Buka Product MISTINE",
        pgEditProduct: "Edit Product MISTINE",
        pgUpdateStockHabis: "Update Stock Habis MISTINE",
        pgMarkLunas: "Mark Lunas MISTINE",
        laporanKerja: "Laporan Kerja",
        grafikLaporan: "Grafik Laporan",
        handover: "Stock Control",
        klaim: "Klaim Paket",
        tugasProject: "Tugas Project",
        karyawan: "Karyawan",
        ijin: "Ijin/Cuti",
        lembur: "Lembur",
        pengajuanAsset: "Pengajuan Asset",
        organisasi: "Organisasi",
        gantiPass: "Ganti Password",
        sop: "SOP Gudang",
        packingList: "Dokumen INBOUND",
        stock: "Stock Inventory",
        inbound: "Inbound Barang",
        outbound: "Potong Stok",
        retur: "Return Barang",
        order: "Orderan",
        stockOpname: "Stock Opname",
        analisis: "Analisis Stok",
        admin: "Admin",
        manajemenUser: "Manajemen User",
        keluar: "Keluar",
        saldoKasGudang: "Saldo Kas Gudang",
        saldoTeamBuilding: "Saldo Team Building",
        totalKasMasuk: "Total Kas Masuk",
        totalKasKeluar: "Total Kas Keluar",
        notifTitle: "Notifikasi",
        notifMarkRead: "✓ Tandai Semua Dibaca",
        notifAll: "Semua",
        notifUnread: "Belum Dibaca",
        notifSeeAll: "Lihat Semua Aktivitas",
        success: "Berhasil",
        error: "Gagal",
        confirmDelete: "Hapus data ini?",
        menyimpan: "Menyimpan...",
        simpan: "Simpan",
        batal: "Batal",
        tutup: "Tutup",
        pilihBahasa: "Pilih Bahasa",
        mandarin: "Mandarin",
        inggeris: "Inggris",
        indonesia: "Indonesia",
        pencarian: "Cari...",
        tambah: "Tambah",
        hapus: "Hapus",
        edit: "Edit",
        lihat: "Lihat",
        detail: "Detail",
        status: "Status",
        tanggal: "Tanggal",
        noOrder: "No. Order",
        pelanggan: "Pelanggan",
        alamat: "Alamat",
        qty: "Qty",
        aksi: "Aksi",
        utama: "Utama",
        keuangan: "Keuangan",
        operasionalKerja: "Operasional Kerja",
        sdm: "SDM",
        dokumentasi: "Dokumentasi",
        inventory: "Inventory",
        menungguApproval: "🔔 Menunggu Approval Saya",
        tipePengajuan: "Tipe Pengajuan",
        namaStaff: "Nama Staff",
        statusSaatIni: "Status Saat Ini",
        operasionalOrderan: "👥 Operasional & Orderan (Hari Ini)",
        divisiOperasional: "Divisi Operasional",
        pekerja: "Pekerja",
        pengeluaranKasGudang: "📉 Pengeluaran Kas Gudang",
        pengeluaranTeamBuilding: "🤝 Pengeluaran Team Building",
        historyTransaksi: "📋 History Transaksi Terbaru",
        refresh: "Refresh",
        kategori: "Kategori",
        tipe: "Tipe",
        nominal: "Nominal",
        keterangan: "Keterangan",
        loginTitle: "GUDANG FCL GROUP",
        loginSub: "Sistem Pengelola Gudang",
        username: "Username",
        password: "Password",
        loginUserPlaceholder: "Masukkan username",
        loginPassPlaceholder: "Masukkan password",
        masuk: "Masuk",
        statusTodo: "Belum Dimulai",
        statusInProgress: "Sedang Dikerjakan",
        statusDone: "Selesai",
        searchPlaceholderStock: "Cari SKU atau Nama Barang...",
        searchPlaceholderOrder: "Cari Customer atau No. Order...",
        pilihStatus: "-- Pilih Status --",
        semuaStatus: "Semua Status",
        bukti: "Bukti",
        inputBy: "Input By",
        tambahTransaksi: "Tambah Transaksi",
        pilihBulan: "-- Pilih Bulan --",
        printLaporan: "Print Laporan",
        dashboardExpense: "Dashboard Expense",
        rincianPerusahaan: "Rincian per Perusahaan",
        riwayatLaporanHarian: "Riwayat Laporan Kerja (Harian)",
        buatLaporan: "Buat Laporan",
        ringkasanPersonil: "Ringkasan Personil Lembur",
        tugasProjectWarehouse: "Tugas Project Warehouse",
        papanKanban: "Papan Kanban",
        tabelList: "Tabel List",
        karyawanAktif: "Karyawan Aktif",
        riwayatResign: "Riwayat Resign",
        suratPeringatan: "Surat Peringatan",
        ekspor: "Ekspor",
        pengajuanIjinCuti: "Pengajuan Ijin & Cuti",
        ajukanIjin: "Ajukan Ijin",
        pengajuanAssetTitle: "Pengajuan Asset",
        ajukanAsset: "Ajukan Asset",
        dashboardValidasiLembur: "Dashboard Validasi Lembur",
        namaKaryawan: "Nama Karyawan",
        jabatan: "Jabatan",
        telepon: "Telepon",
        email: "Email",
        tglMasuk: "Tanggal Masuk",
        selesaiKontrak: "Selesai Kontrak",
        sisaCuti: "Sisa Cuti",
        prioritas: "Prioritas",
        deskripsi: "Deskripsi",
        alasan: "Alasan",
        masaBerlaku: "Masa Berlaku",
        kadaluarsa: "Kadaluarsa",
        estimasiHarga: "Estimasi Harga",
        shift: "Shift",
        kendala: "Kendala",
        pilihUser: "-- Pilih User --",
        targetSelesai: "Target Selesai",
        durasi: "Durasi",
        catatan: "Catatan",
        uploadFile: "Upload File",
        urlLink: "URL Link",
        dragDrop: "Klik atau drag & drop file di sini",
        maksSize: "maks. 20MB",
        semuaDivisi: "Semua Divisi",
        rekapLembur: "Lihat Rekap Seluruh Tanggal",
        lemburDiLaporan: "Lembur di Laporan",
        sudahPengajuan: "Sudah Pengajuan",
        belumPengajuan: "Belum Pengajuan",
        rekapKetidaksinkronan: "Rekap Ketidaksinkronan",
        inputKasGudang: "Input Kas Gudang",
        tipeTransaksi: "Tipe Transaksi",
        pengeluaran: "Pengeluaran",
        pemasukan: "Pemasukan",
        keteranganKegiatan: "Keterangan Kegiatan",
        buktiInvoice: "Bukti Surat Jalan",
        tambahExpense: "Tambah Expense",
        perusahaan: "Perusahaan",
        bankRekening: "Bank & Rekening",
        gantiPassword: "Ganti Password",
        passwordLama: "Password Lama",
        passwordBaru: "Password Baru",
        konfirmasiPassword: "Konfirmasi Password",
        inputLaporanOrderan: "Input Laporan Orderan",
        jamKerjaPhl: "Jam Kerja PHL",
        orangBantu: "Orang Perbantuan",
        orangKurang: "Orang Pengurangan",
        pekerjaLembur: "Pekerja Lembur",
        lamaLembur: "Lama Lembur",
        kpi: "KPI (Total Output / Jam)",
        totalKeseluruhanJam: "Total Keseluruhan Jam",
        inputHandover: "Input Stock Control",
        inputKlaim: "Input Klaim Paket",
        hargaPaket: "Harga Paket",
        buatTugasBaru: "Buat Tugas Baru",
        judulTugas: "Judul Tugas",
        assignee: "Assignee",
        areaKategori: "Area / Kategori",
        tambahKaryawan: "Tambah Karyawan",
        cabangLokasi: "Cabang / Lokasi",
        formResign: "Form Resign",
        alasanUtama: "Alasan Utama",
        buatSP: "Buat Surat Peringatan",
        jenisSP: "Jenis SP",
        tglTerbit: "Tgl Terbit",
        masaBerlakuHari: "Masa Berlaku (Hari)",
        tglKadaluarsaSP: "Tgl Kadaluarsa SP",
        formIjinCuti: "Form Ijin / Cuti",
        formLembur: "Form Pengajuan Lembur",

        // New Inventory Keys
        formInboundTitle: "📥 Form Inbound Barang",
        formOutboundTitle: "✂️ Form Pemotongan Stok Manual",
        formReturTitle: "↩️ Form Retur Barang",
        formOrderTitle: "🛒 Buat Orderan Baru",
        formStockOpnameTitle: "⚖️ Pengajuan Stock Opname",
        formPackingListTitle: "📋 Tambah Dokumen INBOUND",
        tglInbound: "Tanggal Inbound",
        supplier: "Supplier",
        tglPotong: "Tanggal Pemotongan",
        alasanTujuan: "Alasan (Tujuan)",
        tglRetur: "Tanggal Retur",
        sumberRetur: "Sumber Retur",
        detailSJ: "Detail Surat Jalan",
        tglOrder: "Tanggal Order",
        pilihBarangAudit: "📋 Pilih Barang untuk di-Audit",
        mulaiSesiOpname: "🏁 Mulai Sesi Opname",
        ajukanOpname: "💾 Ajukan Opname",
        konfirmasiRetur: "↩️ Konfirmasi Retur & Tambah Stok",
        konfirmasiPotong: "✂️ Konfirmasi Potong Stok",
        simpanTambahStok: "💾 Simpan & Tambah Stok",
        scanBarcode: "🔍 SCAN BARCODE / SKU",
        stokSistem: "Stok Sistem",
        jumlahFisik: "Jumlah Fisik (QTY)",
        selisih: "Selisih",
        lewatkan: "Lewati",
        simpanLanjut: "Simpan & Lanjut",
        importCsv: "Import CSV",
        tambahBaris: "Tambah Baris",
        uploadBukti: "Upload Bukti Packing",
        klikPilihFile: "Klik untuk Pilih File Foto / PDF",
        mulaiUpload: "🚀 Mulai Upload",
        // Global UI
        memuatSistem: "Memuat sistem...",
        pengelolaGudangSub: "Sistem Pengelola Gudang",
        menungguApprovalSaya: "🔔 Menunggu Approval Saya",
        operasionalOrderanHariIni: "👥 Operasional & Orderan (Hari Ini)",
        divisiOperasional: "Divisi Operasional",
        pekerja: "Pekerja",
        pengeluaranKasGudang: "📉 Pengeluaran Kas Gudang",
        pengeluaranTeamBuilding: "🤝 Pengeluaran Team Building",
        inputBy: "Input By",
        bukti: "Bukti",
        tambahTransaksi: "+ Tambah Transaksi",
        jan: "Januari", feb: "Februari", mar: "Maret", apr: "April", mei: "Mei", jun: "Juni",
        jul: "Juli", agu: "Agustus", sep: "September", okt: "Oktober", nov: "November", des: "Desember",
        semuaBulan: "-- Semua Bulan --",
        // Additional UI keys
        totalPengeluaranTahun: "Total Pengeluaran (Tahun Terpilih)",
        totalPengeluaranBulan: "Total Pengeluaran (Bulan Terpilih)",
        rincianPerPerusahaan: "🏢 Rincian per Perusahaan",
        ringkasanPersonilLembur: "👥 Ringkasan Personil Lembur (Laporan Terbaru)",
        riwayatLaporanKerjaHarian: "📝 Riwayat Laporan Kerja (Harian)",
        totalGarisLaporan: "Total Laporan",
        rataRataPekerja: "Rata-rata Pekerja Masuk",
        totalLembur: "Total Jam Lembur",
        totalJamKerja: "Total Keseluruhan Jam Kerja",
        totalOrderan: "Total Orderan",
        kpiOrdJam: "KPI (Ord/Jam)",
        laporanHandover: "📦 Laporan Stock Control",
        laporanKlaimPaket: "⚠️ Laporan Klaim Paket",
        totalTagihanPending: "Total Tagihan (Pending)",
        totalSelesaiDibayar: "Total Selesai Dibayar",
        tugasProjectWarehouse: "📋 Tugas Project Warehouse",
        totalTugas: "Total Tugas",
        dalamProses: "Dalam Proses",
        melewatiTarget: "Melewati Target",
        buatTugas: "+ Buat Tugas",
        semuaDivisi: "Semua Divisi",
        semuaStatus: "Semua Status",
        semuaAssignee: "Semua Assignee",
        semuaPrioritas: "Semua Prioritas",
        // HR Labels
        karyawanAktif: "Karyawan Aktif",
        riwayatResign: "Riwayat Resign",
        spAktif: "SP Aktif",
        tetap: "Tetap",
        kontrak: "Kontrak",
        ijinHariIni: "Ijin Hari Ini",
        cutiHariIni: "Cuti Hari Ini",
        sakitHariIni: "Sakit Hari Ini",
        ekspor: "Ekspor",
        tambahKaryawan: "+ Tambah Karyawan",
        jabatan: "Jabatan",
        tglMasuk: "Tgl Masuk",
        selesaiKontrak: "Selesai Kontrak",
        sisaCuti: "Sisa Cuti",
        tglResign: "Tgl Resign",
        jenisSP: "Jenis SP",
        tglTerbit: "Tgl Terbit",
        masaBerlaku: "Masa Berlaku",
        ajukanIjin: "+ Ajukan Ijin",
        jenisIjinCuti: "Jenis Ijin / Cuti",
        statusApproval: "Status Approval",
        // Asset, Lembur, Org, SOP, Users, Stock
        totalPengajuan: "Total Pengajuan",
        menungguApproval: "Menunggu Approval",
        disetujui: "Disetujui",
        ditolak: "Ditolak",
        riwayatPengajuanAsset: "Riwayat Pengajuan Asset",
        ajukanAsset: "+ Ajukan Asset",
        dashboardValidasiLembur: "📊 Dashboard Validasi Lembur",
        lihatRekapSeluruhTanggal: "Lihat Rekap Seluruh Tanggal",
        lemburDiLaporan: "Lembur di Laporan",
        sudahPengajuan: "Sudah Pengajuan",
        belumPengajuan: "Belum Pengajuan",
        ketidaksinkronanLembur: "🔴 Karyawan Belum Melakukan Pengajuan Lembur:",
        rekapKetidaksinkronan: "📑 Rekap Ketidaksinkronan Seluruh Tanggal:",
        strukturOrganisasi: "🏗️ Struktur Organisasi",
        expand: "Expand",
        collapse: "Collapse",
        sopGudang: "📋 SOP Gudang",
        eksporGoogleDocs: "📤 Ekspor ke Google Docs",
        manajemenUser: "⚙️ Manajemen User",
        tambahUser: "+ Tambah User",
        passwordLama: "Password Lama",
        passwordBaru: "Password Baru",
        perbaruiPassword: "🔄 Perbarui Password",
        totalJenisBarang: "Total Jenis Barang",
        totalStokUnit: "Total Stok (Unit)",
        stokRendah: "Stok Rendah",
        hampirExp: "Hampir Exp (30hr)",
        daftarStockBarang: "📦 Daftar Stock Barang",
        exportCSV: "📤 Export CSV",
        // Modal & Form Labels
        inputLaporanOrderan: "Input Laporan Pengerjaan Orderan",
        tanggalLaporan: "Tanggal Laporan",
        shift: "Shift",
        namaPelaporPic: "Nama Pelapor (PIC)",
        totalPekerjaUtama: "Total Pekerja Utama",
        totalAdmin: "Total Admin",
        totalPhl: "Total PHL",
        jamKerjaPhl: "Jam Kerja PHL (Per Orang)",
        orangPerbantuan: "Orang Perbantuan (+)",
        jamOrangBantu: "Jam / Orang (Bantu)",
        orangPengurangan: "Orang Pengurangan (-)",
        jamOrangKurang: "Jam / Orang (Kurang)",
        totalPekerjaLembur: "Total Pekerja Lembur",
        lamaLemburJam: "Lama Lembur (Jam)",
        pilihNamaPersonilLembur: "Pilih Nama Personil yang Lembur:",
        totalOrderDidapat: "Total Order Didapat",
        totalPo: "Total PO",
        totalQtyPcs: "Total Qty (Pcs)",
        totalInboundSj: "Total Inbound (SJ)",
        totalQtyInbound: "Total Qty (Inbound)",
        kpiOutputJam: "KPI (Total Output / Total Jam Kerja)",
        kendalaPekerjaan: "Kendala Pekerjaan",
        batal: "Batal",
        simpanLaporan: "Simpan Laporan",
        detailLaporanKerja: "Detail Laporan Kerja",
        tutup: "Tutup",
        cetakLaporan: "Cetak Laporan",
        inputLaporanHandover: "Input Laporan Stock Control",
        nomorResi: "Nomor Resi",
        pengerjaan: "Pengerjaan",
        keteranganTambahan: "Keterangan Tambahan",
        inputLaporanKlaim: "Input Laporan Klaim Paket",
        hargaPaketRp: "Harga Paket (Rp)",
        keteranganDetail: "Keterangan Detail",
        buatTugasBaru: "Buat Tugas Baru",
        targetHariSelesai: "Target Hari Selesai",
        kategoriArea: "Kategori / Area",
        deskripsiCatatan: "Deskripsi / Catatan",
        simpanTugas: "Simpan Tugas",
        namaLengkap: "Nama Lengkap",
        cabangLokasi: "Cabang / Lokasi",
        telepon: "Telepon",
        email: "Email",
        sisaKuotaCuti: "Sisa Kuota Cuti (Hari)",
        formResign: "Form Resign / Non-Aktif",
        alasanUtama: "Alasan Utama",
        simpanNonAktif: "Simpan & Non-Aktifkan",
        buatSP: "Buat Surat Peringatan (SP)",
        pilihKaryawan: "Pilih Karyawan",
        masaBerlakuHari: "Masa Berlaku (Hari)",
        tglKadaluarsaSP: "Tanggal Kadaluarsa SP",
        alasanDeskripsiPelanggaran: "Alasan / Deskripsi Pelanggaran",
        formIjinCuti: "Form Pengajuan Ijin / Cuti",
        keteranganLengkap: "Keterangan Lengkap",
        suratDokterBukti: "Surat Dokter / Bukti Pendukung",
        uploadFile: "Upload File",
        urlLink: "URL Link",
        klikDragDrop: "Klik atau drag & drop file di sini",
        formLembur: "Form Pengajuan Lembur",
        jumlahJamLembur: "Jumlah Jam Lembur",
        keteranganTugasLembur: "Keterangan / Tugas Lembur",
        areaPosisi: "Area / Posisi",
        stkMabang: "Stock Mabang",
        stkTtx: "Stock TTX",
        stkFisik: "Stock Fisik",
        selisihMabang: "Selisih Mabang",
        selisihTtx: "Selisih TTX",
        aksiPerbaikan: "Aksi Perbaikan",

        formAsset: "Form Pengajuan Asset",
        estimasiHarga: "Estimasi Harga (Rp)",
        deskripsiKebutuhan: "Deskripsi Kebutuhan",
        alurPersetujuan: "Alur Persetujuan",
        riwayatApproval: "Riwayat Approval",
        tambahAnggotaOrg: "Tambah Anggota Organisasi",
        atasanLangsung: "Atasan Langsung",
        departemen: "Departemen",
        urlFoto: "URL Foto",
        urutanTampil: "Urutan Tampil",
        tambahSOP: "Tambah SOP",
        judulSOP: "Judul SOP",
        isiKontenSOP: "Isi / Konten SOP",
        hakAksesMenu: "Hak Akses Menu",
        lemburTanpaLaporan: "Lembur Tanpa Laporan",
        assetWarehouse: "Asset Warehouse",
        kpiKaryawan: "KPI Karyawan",
        tambahAsset: "+ Tambah Asset",
        pilihDivisi: "Pilih Divisi",
        moveAsset: "Move Asset",
        riwayatPerpindahan: "Riwayat Perpindahan",
        cetakLabel: "Cetak Label",
        namaAsset: "Nama Asset",
        kodeAsset: "Kode Asset",
        tanggalMasuk: "Tanggal Masuk",
        divisi: "Divisi",
        barcodeAsset: "Barcode Asset",
        bookingMobil: "Booking Mobil",
        tambahBooking: "+ Tambah Booking",
        tanggalBooking: "Tanggal Booking",
        picBooking: "PIC Booking",
        jamBerangkat: "Jam Berangkat",
        tujuan: "Tujuan",
        rute: "Rute",
        absensiKaryawan: "Absensi Karyawan",
        jadwalShift: "Jadwal Shift"
      },
      en: {
        dashboard: "Dashboard",
        kasGudang: "Warehouse Cash",
        teamBuilding: "Team Building",
        expense: "Expense List",
        paymentGudang: "Payment Gudang MISTINE",
        pgAddTransaction: "Add Transaction (MISTINE)",
        pgSettingPayment: "Setting Payment (MISTINE)",
        pgToggleStatus: "Close/Buka Product (MISTINE)",
        pgEditProduct: "Edit Product (MISTINE)",
        pgUpdateStockHabis: "Update Stock Habis (MISTINE)",
        pgMarkLunas: "Mark Lunas (MISTINE)",
        laporanKerja: "Daily Report",
        grafikLaporan: "Report Graph",
        handover: "Stock Control",
        klaim: "Parcel Claim",
        tugasProject: "Project Tasks",
        karyawan: "Employees",
        ijin: "Leave/Permission",
        lembur: "Overtime",
        pengajuanAsset: "Asset Request",
        organisasi: "Organization",
        gantiPass: "Change Password",
        sop: "SOP Warehouse",
        packingList: "Inbound Document",
        stock: "Stock Inventory",
        inbound: "Inbound Goods",
        outbound: "Stock Outbound",
        retur: "Returns",
        order: "Orders",
        stockOpname: "Stock Audit",
        analisis: "Stock Analysis",
        admin: "Admin",
        manajemenUser: "User Management",
        keluar: "Logout",
        saldoKasGudang: "Warehouse Balance",
        saldoTeamBuilding: "Team Building Balance",
        totalKasMasuk: "Total Cash In",
        totalKasKeluar: "Total Cash Out",
        notifTitle: "Notifications",
        notifMarkRead: "✓ Mark All Read",
        notifAll: "All",
        notifUnread: "Unread",
        notifSeeAll: "See All Activities",
        success: "Success",
        error: "Error",
        confirmDelete: "Delete this record?",
        menyimpan: "Saving...",
        simpan: "Save",
        batal: "Cancel",
        tutup: "Close",
        pilihBahasa: "Select Language",
        mandarin: "Mandarin",
        inggeris: "English",
        indonesia: "Indonesian",
        pencarian: "Search...",
        tambah: "Add",
        hapus: "Delete",
        edit: "Edit",
        lihat: "View",
        detail: "Detail",
        status: "Status",
        tanggal: "Date",
        noOrder: "Order No.",
        pelanggan: "Customer",
        alamat: "Address",
        qty: "Qty",
        aksi: "Action",
        utama: "Main",
        keuangan: "Finance",
        operasionalKerja: "Operational",
        sdm: "HR",
        dokumentasi: "Documentation",
        inventory: "Inventory",
        menungguApproval: "🔔 Pending My Approval",
        tipePengajuan: "Submission Type",
        namaStaff: "Staff Name",
        statusSaatIni: "Current Status",
        operasionalOrderan: "👥 Operational & Order (Today)",
        divisiOperasional: "Operational Division",
        pekerja: "Worker",
        pengeluaranKasGudang: "📉 Warehouse Cash Out",
        pengeluaranTeamBuilding: "🤝 Team Building Out",
        historyTransaksi: "📋 Latest Transactions",
        refresh: "Refresh",
        kategori: "Category",
        tipe: "Type",
        nominal: "Amount",
        keterangan: "Description",
        loginTitle: "GUDANG FCL GROUP",
        loginSub: "Warehouse Management System",
        username: "Username",
        password: "Password",
        loginUserPlaceholder: "Enter username",
        loginPassPlaceholder: "Enter password",
        masuk: "Login",
        statusTodo: "To Do",
        statusInProgress: "In Progress",
        statusDone: "Done",
        searchPlaceholderStock: "Search SKU or Product Name...",
        searchPlaceholderOrder: "Search Customer or Order No...",
        pilihStatus: "-- Select Status --",
        semuaStatus: "All Status",
        bukti: "Proof",
        inputBy: "Input By",
        tambahTransaksi: "Add Transaction",
        pilihBulan: "-- Select Month --",
        printLaporan: "Print Report",
        dashboardExpense: "Expense Dashboard",
        rincianPerusahaan: "Details per Company",
        riwayatLaporanHarian: "Daily Work Report History",
        buatLaporan: "Create Report",
        ringkasanPersonil: "Overtime Personnel Summary",
        tugasProjectWarehouse: "Warehouse Project Tasks",
        papanKanban: "Kanban Board",
        tabelList: "List Table",
        karyawanAktif: "Active Employees",
        riwayatResign: "Resign History",
        suratPeringatan: "Warning Letter",
        ekspor: "Export",
        pengajuanIjinCuti: "Leave/Permission Request",
        ajukanIjin: "Request Leave",
        pengajuanAssetTitle: "Asset Request",
        ajukanAsset: "Request Asset",
        dashboardValidasiLembur: "Overtime Validation Dashboard",
        namaKaryawan: "Employee Name",
        jabatan: "Position",
        telepon: "Phone",
        email: "Email",
        tglMasuk: "Entry Date",
        selesaiKontrak: "End of Contract",
        sisaCuti: "Remaining Leave",
        prioritas: "Priority",
        deskripsi: "Description",
        alasan: "Reason",
        masaBerlaku: "Validity Period",
        kadaluarsa: "Expired",
        estimasiHarga: "Estimated Price",
        shift: "Shift",
        kendala: "Constraint",
        pilihUser: "-- Select User --",
        targetSelesai: "Target Done",
        durasi: "Duration",
        absensiKaryawan: "Employee Attendance",
        jadwalShift: "Shift Schedule",
        catatan: "Notes",
        uploadFile: "Upload File",
        urlLink: "URL Link",
        dragDrop: "Click or drag & drop file here",
        maksSize: "max. 20MB",
        semuaDivisi: "All Divisions",
        rekapLembur: "View All Dates Summary",
        lemburDiLaporan: "Overtime in Report",
        sudahPengajuan: "Requested",
        belumPengajuan: "Not Requested",
        rekapKetidaksinkronan: "Mismatch Summary",
        inputKasGudang: "Input Warehouse Cash",
        tipeTransaksi: "Transaction Type",
        pengeluaran: "Expense",
        pemasukan: "Income",
        keteranganKegiatan: "Activity Description",
        buktiInvoice: "Invoice Proof",
        tambahExpense: "Add Expense",
        perusahaan: "Company",
        bankRekening: "Bank & Account",
        gantiPassword: "Change Password",
        passwordLama: "Old Password",
        passwordBaru: "New Password",
        konfirmasiPassword: "Confirm Password",
        inputLaporanOrderan: "Input Order Report",
        jamKerjaPhl: "PHL Work Hours",
        orangBantu: "Support Personnel",
        orangKurang: "Personnel Reduction",
        pekerjaLembur: "Overtime Workers",
        lamaLembur: "Overtime Duration",
        kpi: "KPI (Output/Hour)",
        totalKeseluruhanJam: "Total Total Hours",
        inputHandover: "Input Stock Control",
        inputKlaim: "Input Claim Report",
        hargaPaket: "Package Price",
        buatTugasBaru: "Create New Task",
        judulTugas: "Task Title",
        assignee: "Assignee",
        areaKategori: "Area / Category",
        tambahKaryawan: "Add Employee",
        cabangLokasi: "Branch / Location",
        formResign: "Resign Form",
        alasanUtama: "Main Reason",
        buatSP: "Create Warning Letter",
        jenisSP: "Warning Type",
        tglTerbit: "Issue Date",
        masaBerlakuHari: "Validity (Days)",
        tglKadaluarsaSP: "Expiry Date",
        formIjinCuti: "Leave/Permission Form",
        formLembur: "Overtime Request Form",
        areaPosisi: "Area / Position",
        stkMabang: "Mabang Stock",
        stkTtx: "TTX Stock",
        stkFisik: "Physical Stock",
        selisihMabang: "Mabang Diff",
        selisihTtx: "TTX Diff",
        aksiPerbaikan: "Corrective Action",
        formLemburKhusus: "Special Overtime Form",
        // New Inventory Keys
        formInboundTitle: "📥 Inbound Goods Form",
        formOutboundTitle: "✂️ Manual Stock Outbound Form",
        formReturTitle: "↩️ Goods Return Form",
        formOrderTitle: "🛒 Create New Order",
        formStockOpnameTitle: "⚖️ Stock Audit Submission",
        formPackingListTitle: "📋 Add Inbound Document",
        tglInbound: "Inbound Date",
        supplier: "Supplier",
        tglPotong: "Outbound Date",
        alasanTujuan: "Reason (Destination)",
        tglRetur: "Return Date",
        sumberRetur: "Return Source",
        detailSJ: "Delivery Note Detail",
        tglOrder: "Order Date",
        pilihBarangAudit: "📋 Select Items for Audit",
        mulaiSesiOpname: "🏁 Start Audit Session",
        ajukanOpname: "💾 Submit Audit",
        konfirmasiRetur: "↩️ Confirm Return & Add Stock",
        konfirmasiPotong: "✂️ Confirm Stock Outbound",
        simpanTambahStok: "💾 Save & Add Stock",
        scanBarcode: "🔍 SCAN BARCODE / SKU",
        stokSistem: "System Stock",
        jumlahFisik: "Physical Qty",
        selisih: "Difference",
        lewatkan: "Skip",
        simpanLanjut: "Save & Continue",
        importCsv: "Import CSV",
        tambahBaris: "Add Row",
        uploadBukti: "Upload Packing Proof",
        klikPilihFile: "Click to Select Photo / PDF File",
        mulaiUpload: "🚀 Start Upload",
        // Global UI
        memuatSistem: "Loading system...",
        pengelolaGudangSub: "Warehouse Management System",
        menungguApprovalSaya: "🔔 Pending My Approval",
        operasionalOrderanHariIni: "👥 Operations & Orders (Today)",
        divisiOperasional: "Operations Division",
        pekerja: "Worker",
        pengeluaranKasGudang: "📉 Warehouse Cash Expense",
        pengeluaranTeamBuilding: "🤝 Team Building Expense",
        inputBy: "Input By",
        bukti: "Proof",
        tambahTransaksi: "+ Add Transaction",
        jan: "January", feb: "February", mar: "March", apr: "April", mei: "May", jun: "June",
        jul: "July", agu: "August", sep: "September", okt: "October", nov: "November", des: "December",
        semuaBulan: "-- All Months --",
        // Additional UI keys
        totalPengeluaranTahun: "Total Expense (Selected Year)",
        totalPengeluaranBulan: "Total Expense (Selected Month)",
        rincianPerPerusahaan: "🏢 Breakdown by Company",
        ringkasanPersonilLembur: "👥 Overtime Personnel Summary (Latest)",
        riwayatLaporanKerjaHarian: "📝 Daily Work Report History",
        totalGarisLaporan: "Total Reports",
        rataRataPekerja: "Avg Workers Present",
        totalLembur: "Total Overtime Hours",
        totalJamKerja: "Total Working Hours",
        totalOrderan: "Total Orders",
        kpiOrdJam: "KPI (Ord/Hr)",
        laporanHandover: "📦 Stock Control Report",
        laporanKlaimPaket: "⚠️ Package Claim Report",
        totalTagihanPending: "Total Pending Bill",
        totalSelesaiDibayar: "Total Paid Out",
        tugasProjectWarehouse: "📋 Warehouse Project Tasks",
        totalTugas: "Total Tasks",
        dalamProses: "In Progress",
        melewatiTarget: "Overdue",
        buatTugas: "+ Create Task",
        semuaDivisi: "All Divisions",
        semuaStatus: "All Status",
        semuaAssignee: "All Assignee",
        semuaPrioritas: "All Priorities",
        // HR Labels
        karyawanAktif: "Active Employees",
        riwayatResign: "Resign History",
        spAktif: "Active Warning (SP)",
        tetap: "Permanent",
        kontrak: "Contract",
        ijinHariIni: "Permission Today",
        cutiHariIni: "Leave Today",
        sakitHariIni: "Sick Today",
        ekspor: "Export",
        tambahKaryawan: "+ Add Employee",
        jabatan: "Position",
        tglMasuk: "Join Date",
        selesaiKontrak: "End Contract",
        sisaCuti: "Leave Balance",
        tglResign: "Resign Date",
        jenisSP: "Warning Type",
        tglTerbit: "Issue Date",
        masaBerlaku: "Validity",
        ajukanIjin: "+ Request Permission",
        jenisIjinCuti: "Leave/Permission Type",
        statusApproval: "Approval Status",
        // Asset, Lembur, Org, SOP, Users, Stock
        totalPengajuan: "Total Submission",
        menungguApproval: "Pending Approval",
        disetujui: "Approved",
        ditolak: "Rejected",
        riwayatPengajuanAsset: "Asset Submission History",
        ajukanAsset: "+ Request Asset",
        dashboardValidasiLembur: "📊 Overtime Validation Dashboard",
        lihatRekapSeluruhTanggal: "See All Dates Summary",
        lemburDiLaporan: "Overtime in Report",
        sudahPengajuan: "Already Requested",
        belumPengajuan: "Missing Request",
        ketidaksinkronanLembur: "🔴 Employees Missing Overtime Request:",
        rekapKetidaksinkronan: "📑 All Dates Inconsistency Summary:",
        strukturOrganisasi: "🏗️ Organization Structure",
        expand: "Expand",
        collapse: "Collapse",
        sopGudang: "📋 Warehouse SOP",
        eksporGoogleDocs: "📤 Export to Google Docs",
        manajemenUser: "⚙️ User Management",
        tambahUser: "+ Add User",
        passwordLama: "Old Password",
        passwordBaru: "New Password",
        perbaruiPassword: "🔄 Update Password",
        totalJenisBarang: "Total SKU Count",
        totalStokUnit: "Total Stock (Units)",
        stokRendah: "Low Stock",
        hampirExp: "Expiring Soon (30d)",
        daftarStockBarang: "📦 Stock Inventory List",
        exportCSV: "📤 Export CSV",
        // Modal & Form Labels
        inputLaporanOrderan: "Work Order Report Input",
        tanggalLaporan: "Report Date",
        shift: "Shift",
        namaPelaporPic: "Reporter Name (PIC)",
        totalPekerjaUtama: "Total Main Workers",
        totalAdmin: "Total Admin",
        totalPhl: "Total PHL (Freelance)",
        jamKerjaPhl: "PHL Working Hours (Per Person)",
        orangPerbantuan: "Support Personnel (+)",
        jamOrangBantu: "Hrs / Person (Support)",
        orangPengurangan: "Personnel Deduction (-)",
        jamOrangKurang: "Hrs / Person (Deduction)",
        totalPekerjaLembur: "Total Overtime Workers",
        lamaLemburJam: "Overtime Duration (Hrs)",
        pilihNamaPersonilLembur: "Select Overtime Personnel:",
        totalOrderDidapat: "Total Orders Received",
        totalPo: "Total PO",
        totalQtyPcs: "Total Qty (Pcs)",
        totalInboundSj: "Total Inbound (Waybill)",
        totalQtyInbound: "Total Qty (Inbound)",
        kpiOutputJam: "KPI (Total Output / Total Working Hours)",
        kendalaPekerjaan: "Work Obstacles",
        batal: "Cancel",
        simpanLaporan: "Save Report",
        detailLaporanKerja: "Work Report Detail",
        tutup: "Close",
        cetakLaporan: "Print Report",
        inputLaporanHandover: "Stock Control Report Input",
        nomorResi: "Tracking Number",
        pengerjaan: "Work Details",
        keteranganTambahan: "Additional Notes",
        inputLaporanKlaim: "Package Claim Report Input",
        hargaPaketRp: "Package Price (Rp)",
        keteranganDetail: "Detailed Notes",
        buatTugasBaru: "Create New Task",
        judulTugas: "Task Title",
        targetHariSelesai: "Target Days to Complete",
        kategoriArea: "Category / Area",
        deskripsiCatatan: "Description / Notes",
        simpanTugas: "Save Task",
        namaLengkap: "Full Name",
        cabangLokasi: "Branch / Location",
        telepon: "Phone",
        email: "Email",
        sisaKuotaCuti: "Remaining Leave (Days)",
        formResign: "Resignation / Inactive Form",
        alasanUtama: "Main Reason",
        simpanNonAktif: "Save & Deactivate",
        buatSP: "Create Warning Letter (SP)",
        pilihKaryawan: "Select Employee",
        masaBerlakuHari: "Validity Period (Days)",
        tglKadaluarsaSP: "SP Expiration Date",
        alasanDeskripsiPelanggaran: "Reason / Violation Details",
        formIjinCuti: "Leave / Permission Request Form",
        keteranganLengkap: "Full Explanation",
        suratDokterBukti: "Doctor's Note / Supporting Evidence",
        uploadFile: "Upload File",
        urlLink: "URL Link",
        klikDragDrop: "Click or drag & drop file here",
        formLembur: "Overtime Request Form",
        jumlahJamLembur: "Total Overtime Hours",
        keteranganTugasLembur: "Overtime Notes / Task",
        formLemburKhusus: "Special Overtime Request Form",
        formAsset: "Asset Request Form",
        estimasiHarga: "Estimated Price (Rp)",
        deskripsiKebutuhan: "Requirement Description",
        alurPersetujuan: "Approval Workflow",
        riwayatApproval: "Approval History",
        tambahAnggotaOrg: "Add Organization Member",
        atasanLangsung: "Direct Supervisor",
        departemen: "Department",
        urlFoto: "Photo URL",
        urutanTampil: "Display Order",
        tambahSOP: "Add SOP",
        judulSOP: "SOP Title",
        isiKontenSOP: "SOP Content",
        hakAksesMenu: "Menu Access Permissions",
        lemburTanpaLaporan: "Overtime Without Report",
        assetWarehouse: "Asset Warehouse",
        kpiKaryawan: "Employee KPI",
        tambahAsset: "+ Add Asset",
        pilihDivisi: "Select Division",
        moveAsset: "Move Asset",
        riwayatPerpindahan: "Transfer History",
        cetakLabel: "Print Label",
        namaAsset: "Asset Name",
        kodeAsset: "Asset Code",
        tanggalMasuk: "Entry Date",
        divisi: "Division",
        barcodeAsset: "Asset Barcode",
        bookingMobil: "Car Booking",
        tambahBooking: "+ Add Booking",
        tanggalBooking: "Booking Date",
        picBooking: "PIC Booking",
        jamBerangkat: "Departure Time",
        tujuan: "Destination",
        rute: "Route"
      },
      zh: {
        dashboard: "仪表板",
        kasGudang: "仓库现金",
        teamBuilding: "团队建设",
        expense: "费用列表",
        paymentGudang: "Payment Gudang MISTINE",
        pgAddTransaction: "添加交易 (MISTINE)",
        pgSettingPayment: "支付设置 (MISTINE)",
        pgToggleStatus: "关闭/打开产品 (MISTINE)",
        pgEditProduct: "编辑产品 (MISTINE)",
        pgUpdateStockHabis: "更新无库存 (MISTINE)",
        pgMarkLunas: "标记为已付 (MISTINE)",
        laporanKerja: "工作报告",
        grafikLaporan: "报告图表",
        handover: "库存控制",
        klaim: "件索赔",
        tugasProject: "项目任务",
        karyawan: "员工数据",
        ijin: "请假申请",
        lembur: "加班申请",
        pengajuanAsset: "资产申请",
        organisasi: "组织结构",
        gantiPass: "修改密码",
        sop: "SOP 仓库",
        packingList: "入库文档",
        stock: "库存管理",
        inbound: "入库详情",
        outbound: "库存扣除",
        retur: "退货管理",
        order: "订单管理",
        stockOpname: "盘点管理",
        analisis: "库存分析",
        admin: "管理员",
        manajemenUser: "用户管理",
        lemburTanpaLaporan: "无需报告加班",
        assetWarehouse: "资产仓库",
        keluar: "登出",
        saldoKasGudang: "仓库现金余额",
        saldoTeamBuilding: "团建现金余额",
        totalKasMasuk: "总现金收入",
        totalKasKeluar: "总现金支出",
        notifTitle: "通知",
        notifMarkRead: "✓ 全部读过",
        notifAll: "全部",
        notifUnread: "未读",
        notifSeeAll: "查看所有活动",
        success: "成功",
        error: "错误",
        confirmDelete: "删除此记录吗？",
        menyimpan: "保存中...",
        simpan: "保存",
        batal: "取消",
        tutup: "关闭",
        pilihBahasa: "选择语言",
        mandarin: "中文",
        inggeris: "英语",
        indonesia: "印尼语",
        pencarian: "搜索...",
        tambah: "添加",
        hapus: "删除",
        edit: "编辑",
        lihat: "视图",
        detail: "详情",
        status: "状态",
        tanggal: "日期",
        noOrder: "订单编号",
        pelanggan: "客户",
        alamat: "地址",
        qty: "数量",
        aksi: "操作",
        utama: "主要",
        keuangan: "财务",
        operasionalKerja: "运营",
        sdm: "人力资源",
        dokumentasi: "文档",
        inventory: "库存",
        menungguApproval: "🔔 待我审批",
        tipePengajuan: "提交类型",
        namaStaff: "员工姓名",
        statusSaatIni: "当前状态",
        operasionalOrderan: "👥 运营与订单 (今天)",
        divisiOperasional: "运营部门",
        pekerja: "工人",
        pengeluaranKasGudang: "📉 仓库现金支出",
        pengeluaranTeamBuilding: "🤝 团队建设支出",
        historyTransaksi: "📋 最新交易历史",
        refresh: "刷新",
        kategori: "类别",
        tipe: "类型",
        nominal: "金额",
        keterangan: "备注",
        loginTitle: "仓库 FCL",
        loginSub: "仓库管理系统",
        username: "用户名",
        password: "密码",
        loginUserPlaceholder: "输入用户名",
        loginPassPlaceholder: "输入密码",
        masuk: "登录",
        statusTodo: "待办",
        statusInProgress: "进行中",
        statusDone: "已完成",
        searchPlaceholderStock: "查询商品编码或名称...",
        searchPlaceholderOrder: "查询客户或订单号...",
        pilihStatus: "-- 选择状态 --",
        semuaStatus: "全部状态",
        bukti: "凭证",
        inputBy: "录入人",
        tambahTransaksi: "添加交易",
        pilihBulan: "-- 选择月份 --",
        printLaporan: "打印报告",
        dashboardExpense: "费用仪表板",
        rincianPerusahaan: "公司明细",
        riwayatLaporanHarian: "每日工作报告历史",
        buatLaporan: "创建报告",
        ringkasanPersonil: "加班人员摘要",
        tugasProjectWarehouse: "仓库项目任务",
        papanKanban: "看板",
        tabelList: "列表视图",
        karyawanAktif: "在职员工",
        riwayatResign: "离职历史",
        suratPeringatan: "警告信",
        ekspor: "导出",
        pengajuanIjinCuti: "请假申请",
        ajukanIjin: "申请请假",
        pengajuanAssetTitle: "资产申请",
        ajukanAsset: "申请资产",
        dashboardValidasiLembur: "加班校验仪表板",
        namaKaryawan: "员工姓名",
        jabatan: "职位",
        telepon: "电话",
        email: "电子邮件",
        tglMasuk: "入职日期",
        selesaiKontrak: "合同到期",
        sisaCuti: "剩余假期",
        prioritas: "优先级",
        deskripsi: "描述",
        alasan: "原因",
        masaBerlaku: "有效期",
        kadaluarsa: "过期",
        estimasiHarga: "预估价格",
        shift: "班次",
        kendala: "障碍/问题",
        pilihUser: "-- 选择用户 --",
        targetSelesai: "目标完成日期",
        durasi: "时长",
        catatan: "备注",
        uploadFile: "上传文件",
        urlLink: "URL 链接",
        dragDrop: "点击或拖拽文件到这里",
        maksSize: "最大 20MB",
        semuaDivisi: "所有部门",
        rekapLembur: "查看所有日期摘要",
        lemburDiLaporan: "报告中的加班",
        sudahPengajuan: "已申请",
        belumPengajuan: "未申请",
        rekapKetidaksinkronan: "不匹配摘要",
        inputKasGudang: "输入仓库现金",
        tipeTransaksi: "交易类型",
        pengeluaran: "支出",
        pemasukan: "收入",
        keteranganKegiatan: "活动描述",
        buktiInvoice: "发票凭证",
        tambahExpense: "添加费用",
        perusahaan: "公司",
        bankRekening: "银行与账号",
        gantiPassword: "修改密码",
        passwordLama: "旧密码",
        passwordBaru: "新密码",
        konfirmasiPassword: "确认密码",
        inputLaporanOrderan: "输入订单报告",
        jamKerjaPhl: "临时工工时",
        orangBantu: "支援人员",
        orangKurang: "人员减少",
        pekerjaLembur: "加班人员",
        lamaLembur: "加班时长",
        kpi: "KPI (产出/小时)",
        totalKeseluruhanJam: "总时数",
        inputHandover: "输入库存控制报告",
        inputKlaim: "输入索赔报告",
        hargaPaket: "包裹价格",
        buatTugasBaru: "创建新任务",
        judulTugas: "任务标题",
        assignee: "受托人",
        areaKategori: "区域 / 类别",
        tambahKaryawan: "添加员工",
        cabangLokasi: "分支 / 地点",
        formResign: "离职表单",
        alasanUtama: "主要原因",
        buatSP: "创建警告信",
        jenisSP: "警告类型",
        tglTerbit: "发布日期",
        masaBerlakuHari: "有效期 (天)",
        tglKadaluarsaSP: "过期日期",
        formIjinCuti: "请假表单",
        formLembur: "加班申请表",
        areaPosisi: "区域 / 位置",
        stkMabang: "Mabang 库存",
        stkTtx: "TTX 库存",
        stkFisik: "实物库存",
        selisihMabang: "Mabang 差异",
        selisihTtx: "TTX 差异",
        aksiPerbaikan: "改进措施",
        formLemburKhusus: "特殊加班表",
        // New Inventory Keys
        formInboundTitle: "📥 入库表单",
        formOutboundTitle: "✂️ 手动扣库存表单",
        formReturTitle: "↩️ 退货表单",
        formOrderTitle: "🛒 创建新订单",
        formStockOpnameTitle: "⚖️ 提交库存盘点",
        formPackingListTitle: "📋 添加入库文档",
        tglInbound: "入库日期",
        supplier: "供应商",
        tglPotong: "扣减日期",
        alasanTujuan: "原因（目的）",
        tglRetur: "退货日期",
        sumberRetur: "退货来源",
        detailSJ: "送货单详情",
        tglOrder: "订单日期",
        pilihBarangAudit: "📋 选择要审计的商品",
        mulaiSesiOpname: "🏁 开始盘点期间",
        ajukanOpname: "💾 提交盘点",
        konfirmasiRetur: "↩️ 确认退货并增加库存",
        konfirmasiPotong: "✂️ 确认扣除库存",
        simpanTambahStok: "💾 保存并增加库存",
        scanBarcode: "🔍 扫描条形码 / SKU",
        stokSistem: "系统库存",
        jumlahFisik: "实物数量",
        selisih: "差异",
        lewatkan: "跳过",
        simpanLanjut: "保存并继续",
        importCsv: "导入 CSV",
        tambahBaris: "添加行",
        uploadBukti: "上传打包凭证",
        klikPilihFile: "点击选择照片 / PDF 文件",
        mulaiUpload: "🚀 开始上传",
        // Global UI
        memuatSistem: "正在加载系统...",
        pengelolaGudangSub: "仓库管理系统",
        menungguApprovalSaya: "🔔 等待我的批准",
        operasionalOrderanHariIni: "👥 运营与订单（今天）",
        divisiOperasional: "运营部门",
        pekerja: "工作人员",
        pengeluaranKasGudang: "📉 仓库现金支出",
        pengeluaranTeamBuilding: "🤝 团建支出",
        inputBy: "输入者",
        bukti: "凭证",
        tambahTransaksi: "+ 添加交易",
        jan: "一月", feb: "二月", mar: "三月", apr: "四月", mei: "五月", jun: "六月",
        jul: "七月", agu: "八月", sep: "九月", okt: "十月", nov: "十一月", des: "十二月",
        semuaBulan: "-- 所有月份 --",
        // Additional UI keys
        totalPengeluaranTahun: "总支出（所选年份）",
        totalPengeluaranBulan: "总支出（所选月份）",
        rincianPerPerusahaan: "🏢 按公司划分的详细信息",
        ringkasanPersonilLembur: "👥 加班人员摘要（最新报告）",
        riwayatLaporanKerjaHarian: "📝 每日工作报告历史",
        totalGarisLaporan: "报告总数",
        rataRataPekerja: "平均入职人数",
        totalLembur: "总加班小时数",
        totalJamKerja: "总工作小时数",
        totalOrderan: "总订单数",
        kpiOrdJam: "绩效（订单/小时）",
        laporanHandover: "📦 库存控制报告",
        laporanKlaimPaket: "⚠️ 包裹索赔报告",
        totalTagihanPending: "待处理账单总额",
        totalSelesaiDibayar: "已支付总额",
        tugasProjectWarehouse: "📋 仓库项目任务",
        totalTugas: "任务总数",
        dalamProses: "进行中",
        melewatiTarget: "超过目标",
        buatTugas: "+ 创建任务",
        semuaDivisi: "所有部门",
        semuaStatus: "所有状态",
        semuaAssignee: "所有负责人",
        semuaPrioritas: "所有优先级",
        // HR Labels
        karyawanAktif: "在职员工",
        riwayatResign: "离职历史",
        spAktif: "有效警告 (SP)",
        tetap: "正式工",
        kontrak: "合同工",
        ijinHariIni: "今日请假",
        cutiHariIni: "今日年假",
        sakitHariIni: "今日病假",
        ekspor: "导出",
        tambahKaryawan: "+ 添加员工",
        jabatan: "职位",
        tglMasuk: "入职日期",
        selesaiKontrak: "合同结束",
        sisaCuti: "剩余年假",
        tglResign: "离职日期",
        jenisSP: "警告类型",
        tglTerbit: "发布日期",
        masaBerlaku: "有效期",
        ajukanIjin: "+ 提交申请",
        jenisIjinCuti: "请假/年假类型",
        statusApproval: "审批状态",
        // Asset, Lembur, Org, SOP, Users, Stock
        totalPengajuan: "总提交",
        menungguApproval: "待审批",
        disetujui: "已批准",
        ditolak: "已拒绝",
        riwayatPengajuanAsset: "资产申请历史",
        ajukanAsset: "+ 申请资产",
        dashboardValidasiLembur: "📊 加班验证看板",
        lihatRekapSeluruhTanggal: "查看全日期摘要",
        lemburDiLaporan: "报表中的加班",
        sudahPengajuan: "已申请",
        belumPengajuan: "未申请",
        ketidaksinkronanLembur: "🔴 未提交加班申请的员工：",
        rekapKetidaksinkronan: "📑 全日期不一致摘要：",
        strukturOrganisasi: "🏗️ 组织架构",
        expand: "展开",
        collapse: "折叠",
        sopGudang: "📋 仓库 SOP",
        eksporGoogleDocs: "📤 导出到 Google Docs",
        manajemenUser: "⚙️ 用户管理",
        tambahUser: "+ 添加用户",
        passwordLama: "旧密码",
        passwordBaru: "新密码",
        perbaruiPassword: "🔄 更新密码",
        totalJenisBarang: "总 SKU 数量",
        totalStokUnit: "总库存（单位）",
        stokRendah: "库存不足",
        hampirExp: "即将过期（30天）",
        daftarStockBarang: "📦 库存清单",
        exportCSV: "📤 导出 CSV",
        // Modal & Form Labels
        inputLaporanOrderan: "输入工作订单报告",
        tanggalLaporan: "报告日期",
        shift: "班次",
        namaPelaporPic: "报告人姓名 (PIC)",
        totalPekerjaUtama: "主要工人总数",
        totalAdmin: "行政人员总数",
        totalPhl: "临时工总数",
        jamKerjaPhl: "临时工工作时间（每人）",
        orangPerbantuan: "支援人员 (+)",
        jamOrangBantu: "小时/每人（支援）",
        orangPengurangan: "人员减少 (-)",
        jamOrangKurang: "小时/每人（减少）",
        totalPekerjaLembur: "加班人员总数",
        lamaLemburJam: "加班时间（小时）",
        pilihNamaPersonilLembur: "选择加班人员：",
        totalOrderDidapat: "收到的订单总数",
        totalPo: "订单总数 (PO)",
        totalQtyPcs: "总数量 (Pcs)",
        totalInboundSj: "入库总数 (SJ)",
        totalQtyInbound: "入库总数量",
        kpiOutputJam: "绩效 (总产出 / 总工作时间)",
        kendalaPekerjaan: "工作障碍",
        batal: "取消",
        simpanLaporan: "保存报告",
        detailLaporanKerja: "工作报告详情",
        tutup: "关闭",
        cetakLaporan: "打印报告",
        inputLaporanHandover: "输入库存控制报告",
        nomorResi: "快递单号",
        pengerjaan: "工作详情",
        keteranganTambahan: "附加备注",
        inputLaporanKlaim: "输入包裹索赔报告",
        hargaPaketRp: "包裹价格 (Rp)",
        keteranganDetail: "详细备注",
        buatTugasBaru: "创建新任务",
        judulTugas: "任务标题",
        targetHariSelesai: "预计完成天数",
        kategoriArea: "类别 / 区域",
        deskripsiCatatan: "描述 / 备注",
        simpanTugas: "保存任务",
        namaLengkap: "全名",
        cabangLokasi: "分支机构 / 地点",
        telepon: "电话",
        email: "电子邮件",
        sisaKuotaCuti: "剩余年假（天）",
        formResign: "离职 / 停职表单",
        alasanUtama: "主要原因",
        simpanNonAktif: "保存并停职",
        buatSP: "创建警告信 (SP)",
        pilihKaryawan: "选择员工",
        masaBerlakuHari: "有效期（天）",
        tglKadaluarsaSP: "SP 到期日期",
        alasanDeskripsiPelanggaran: "原因 / 违规细节",
        formIjinCuti: "请假 / 批准申请表",
        keteranganLengkap: "详细说明",
        suratDokterBukti: "医生证明 / 佐证材料",
        uploadFile: "上传文件",
        urlLink: "URL 链接",
        klikDragDrop: "点击或拖拽文件到此处",
        formLembur: "加班申请表",
        jumlahJamLembur: "总加班小时数",
        keteranganTugasLembur: "加班备注 / 任务",
        formLemburKhusus: "特别加班申请表",
        formAsset: "资产申请表",
        estimasiHarga: "预计价格 (Rp)",
        deskripsiKebutuhan: "需求描述",
        alurPersetujuan: "审批流程",
        riwayatApproval: "审批历史",
        tambahAnggotaOrg: "添加组织成员",
        atasanLangsung: "直接上级",
        departemen: "部门",
        urlFoto: "照片 URL",
        urutanTampil: "显示顺序",
        tambahSOP: "添加 SOP",
        judulSOP: "SOP 标题",
        isiKontenSOP: "SOP 内容",
        hakAksesMenu: "菜单访问权限",
        assetWarehouse: "资产仓库",
        kpiKaryawan: "员工 KPI",
        tambahAsset: "+ 添加资产",
        pilihDivisi: "选择部门",
        moveAsset: "移动资产",
        riwayatPerpindahan: "移动历史",
        cetakLabel: "打印标签",
        namaAsset: "资产名称",
        kodeAsset: "资产编号",
        tanggalMasuk: "入库日期",
        divisi: "部门",
        barcodeAsset: "资产条形码",
        absensiKaryawan: "员工考勤",
        jadwalShift: "轮班安排"
      }
    };

    let currentLang = localStorage.getItem('fcl_lang') || 'id';

    function setLanguage(lang) {
      currentLang = lang;
      localStorage.setItem('fcl_lang', lang);
      applyLanguage();
      toggleLangDropdown();
      updateLangUI();
    }

    function applyLanguage() {
      const data = i18nData[currentLang];
      document.querySelectorAll('[data-i18n]').forEach(el => {
        const key = el.getAttribute('data-i18n');
        const text = data[key];
        if (text) {
          // Check for placeholders in input/textarea
          if ((el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') && el.getAttribute('placeholder')) {
            el.setAttribute('placeholder', text);
          } else {
            // Handle elements with icons (e.g., sidebar)
            const icon = el.querySelector('i, .icon');
            if (icon) {
              const iconHtml = icon.outerHTML;
              // Special case: if the text had an emoji prefix, we might want to keep the emoji or replace it
              // For now, let's just append the translated text after the icon
              el.innerHTML = iconHtml + ' ' + text;
            } else {
              // Standard text replacement
              // Check if the element has other children we should preserve
              if (el.children.length === 0) {
                el.textContent = text;
              } else {
                // If it has children but no identified icon, we attempt to only replace the text node
                // This is safer for complex buttons or labels
                Array.from(el.childNodes).forEach(node => {
                  if (node.nodeType === Node.TEXT_NODE && node.textContent.trim().length > 0) {
                    node.textContent = ' ' + text + ' ';
                  }
                });
              }
            }
          }
        }
      });
      // Also update the current language label in topbar
      const currentLangLabel = document.getElementById('currentLangLabel');
      if (currentLangLabel) {
        currentLangLabel.textContent = currentLang === 'id' ? '🇮🇩 ID' : currentLang === 'en' ? '🇺🇸 EN' : '🇨🇳 ZH';
      }

      // Special cases
      if (document.getElementById('pageTitle')) {
        const pageKey = document.getElementById('pageTitle').getAttribute('data-page-key');
        if (pageKey && data[pageKey]) {
          const icon = document.querySelector(`.nav-item[onclick*="${pageKey}"] .icon`)?.textContent || '📊';
          document.getElementById('pageTitle').textContent = icon + ' ' + data[pageKey];
        }
      }
    }

    function updateLangUI() {
      const langNames = { id: '🇮🇩 ID', en: '🇺🇸 EN', zh: '🇨🇳 ZH' };
      document.getElementById('currentLangLabel').textContent = langNames[currentLang];
      document.querySelectorAll('.lang-item').forEach(item => {
        item.classList.remove('active');
        if (item.getAttribute('onclick').includes(currentLang)) item.classList.add('active');
      });
    }

    function toggleLangDropdown() {
      const dropdown = document.getElementById('langDropdown');
      dropdown.classList.toggle('show');
    }

    function getLangText(key) {
      return i18nData[currentLang][key] || key;
    }

    // Initialize on start
    window.addEventListener('DOMContentLoaded', () => {
      applyLanguage();
      updateLangUI();
      // Closing dropdown on outside click
      window.addEventListener('click', (e) => {
        const wrap = document.querySelector('.lang-wrapper');
        const drop = document.getElementById('langDropdown');
        if (wrap && !wrap.contains(e.target) && drop.classList.contains('show')) {
          drop.classList.remove('show');
        }
      });
    });

    const savedTheme = localStorage.getItem('fcl_theme') || 'dark';
    let stockOpnameDataAll = [];
    let opnameQueue = []; let currentOpnameIndex = 0;
    document.documentElement.setAttribute('data-theme', savedTheme);
    document.documentElement.setAttribute('data-bs-theme', savedTheme);

    function toggleTheme() {
      const current = document.documentElement.getAttribute('data-theme');
      const next = current === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', next);
      document.documentElement.setAttribute('data-bs-theme', next);
      localStorage.setItem('fcl_theme', next);
      document.getElementById('themeIcon').className = next === 'dark' ? 'bi bi-moon-stars' : 'bi bi-brightness-high-fill';
    }

    window.addEventListener('DOMContentLoaded', () => {
      if (document.getElementById('themeIcon')) {
        document.getElementById('themeIcon').className = savedTheme === 'dark' ? 'bi bi-moon-stars' : 'bi bi-brightness-high-fill';
      }
    });

    let currentUser = null;
    let karyawanData = [];
    // Variabel Payment Gudang (Toko Produk) - DINONAKTIFKAN karena fitur dihapus
    // let paymentGudangData = [];
    // let paymentGudangEmployeeData = [];
    // let paymentGudangSelectedEmployees = new Set();
    // let paymentParticipantData = [];
    // let paymentSelectedId = '';
    // let pgActiveProductId = '';
    // let pgActiveProductFilter = 'Aktif';
    // let pgEditingProductId = '';
    // let pgFilteredTransactions = [];
    // let midtransClientKey = '';

    // === GLOBAL LOADING SYSTEM ===
    function showLoading(msg = 'Memproses data...') {
      let loader = document.getElementById('globalLoader');
      if (!loader) {
        loader = document.createElement('div');
        loader.id = 'globalLoader';
        loader.innerHTML = `
          <div class="loader-content">
            <div class="spinner-border text-teal" role="status"></div>
            <div class="loader-text mt-3">${msg}</div>
          </div>`;
        document.body.appendChild(loader);

        // Add CSS for loader if not exists
        if (!document.getElementById('loaderStyles')) {
          const style = document.createElement('style');
          style.id = 'loaderStyles';
          style.textContent = `
            #globalLoader {
              position: fixed; top: 0; left: 0; width: 100%; height: 100%;
              background: rgba(10, 22, 40, 0.8); backdrop-filter: blur(4px);
              display: flex; align-items: center; justify-content: center;
              z-index: 9999; color: white; flex-direction: column;
            }
            .loader-content { text-align: center; }
            .loader-text { font-weight: 600; letter-spacing: 0.5px; }
          `;
          document.head.appendChild(style);
        }
      } else {
        loader.querySelector('.loader-text').textContent = msg;
        loader.style.display = 'flex';
      }
    }

    function hideLoading() {
      const loader = document.getElementById('globalLoader');
      if (loader) loader.style.display = 'none';
    }

    let orderDataList = []; let stockData = []; let laporanKerjaData = []; let klaimDataList = [];
    let chartInstances = { mp: null, dist: null, ret: null };
    let riwayatData = []; let spData = [];
    let unreadNotifCount = 0;
    let pendingApprovalCount = 0;

    function refreshGlobalBadge() {
      const total = unreadNotifCount + pendingApprovalCount;
      const badge = document.getElementById('notifBadge');
      if (badge) {
        if (total > 0) {
          badge.style.display = 'flex';
          badge.textContent = total > 99 ? '99+' : total;
          badge.classList.add('pulse-notif');
        } else {
          badge.style.display = 'none';
        }
      }
    }

    // === NOTIFICATION SYSTEM (Facebook-style floating) ===
    let notifications = [];
    let notifPanelOpen = false;

    // Inject floating panel into body on first toggle
    function ensureNotifPanel() {
      if (document.getElementById('notifPanel')) return;
      const backdrop = document.createElement('div');
      backdrop.id = 'notifBackdrop';
      backdrop.onclick = closeNotifPanel;
      document.body.appendChild(backdrop);

      const panel = document.createElement('div');
      panel.id = 'notifPanel';
      panel.innerHTML = `
        <div class="notif-panel-header">
          <div class="notif-panel-title">
            🔔 Notifikasi
            <span class="notif-unread-count" id="notifUnreadCount" style="display:none">0</span>
          </div>
          <button class="notif-mark-all" onclick="markAllNotifRead()">✓ Baca Semua</button>
        </div>
        <div class="notif-tabs">
          <div class="notif-tab active" onclick="filterNotif('all', this)">Semua</div>
          <div class="notif-tab" onclick="filterNotif('unread', this)">Belum Dibaca</div>
        </div>
        <div class="notif-scroll-list" id="notifList"></div>
        <div class="notif-panel-footer">
          <button onclick="showPage('dashboard'); closeNotifPanel();">Lihat Semua Aktivitas</button>
        </div>`;
      document.body.appendChild(panel);
    }

    let _notifFilter = 'all';
    function filterNotif(filter, tab) {
      _notifFilter = filter;
      document.querySelectorAll('.notif-tab').forEach(t => t.classList.remove('active'));
      if (tab) tab.classList.add('active');
      renderNotifications();
    }

    function toggleNotifPanel(event) {
      if (event) event.stopPropagation();
      ensureNotifPanel();
      if (notifPanelOpen) { closeNotifPanel(); return; }

      // Position the panel below the bell button
      const btn = document.getElementById('notifBtn');
      const panel = document.getElementById('notifPanel');
      const backdrop = document.getElementById('notifBackdrop');
      const rect = btn.getBoundingClientRect();
      const panelW = 360;
      let left = rect.right - panelW;
      if (left < 8) left = 8;
      panel.style.top = (rect.bottom + 8) + 'px';
      panel.style.left = left + 'px';

      // Adjust arrow position
      const arrowRight = rect.right - left - 7;
      panel.style.setProperty('--arrow-right', arrowRight + 'px');

      renderNotifications();
      panel.classList.add('open');
      backdrop.classList.add('open');
      btn.classList.add('active');
      notifPanelOpen = true;
    }

    function closeNotifPanel() {
      const panel = document.getElementById('notifPanel');
      const backdrop = document.getElementById('notifBackdrop');
      const btn = document.getElementById('notifBtn');
      if (panel) panel.classList.remove('open');
      if (backdrop) backdrop.classList.remove('open');
      if (btn) btn.classList.remove('active');
      notifPanelOpen = false;
    }

    // Keep backward compat
    function toggleNotifDropdown(event) { toggleNotifPanel(event); }

    function addNotification(title, text, type = 'info', page = 'dashboard', id = null) {
      const exists = notifications.find(n => n.title === title && n.text === text);
      if (exists) return;
      const notif = { id: id || Date.now(), title, text, type, page, time: new Date(), unread: true };
      notifications.unshift(notif);
      if (notifications.length > 50) notifications.pop();
      updateNotifBadge();
      // Animate bell
      const btn = document.getElementById('notifBtn');
      if (btn) { btn.style.animation = 'none'; btn.offsetHeight; btn.style.animation = 'notifBell 0.5s ease'; }
    }

    function updateNotifBadge() {
      unreadNotifCount = notifications.filter(n => n.unread).length;
      const countEl = document.getElementById('notifUnreadCount');
      if (countEl) {
        if (unreadNotifCount > 0) { countEl.style.display = 'inline'; countEl.textContent = unreadNotifCount; }
        else { countEl.style.display = 'none'; }
      }
      refreshGlobalBadge();
    }

    function renderNotifications() {
      const list = document.getElementById('notifList');
      if (!list) return;
      const filtered = _notifFilter === 'unread' ? notifications.filter(n => n.unread) : notifications;
      if (filtered.length === 0) {
        list.innerHTML = `<div class="notif-empty"><span class="notif-empty-icon">🔔</span><p>${_notifFilter === 'unread' ? 'Semua sudah dibaca!' : 'Belum ada notifikasi'}</p><small>${_notifFilter === 'unread' ? 'Tidak ada notifikasi baru.' : 'Notifikasi akan muncul di sini.'}</small></div>`;
        return;
      }
      const icons = { info: { bg: '#0ea5e922', icon: '💡', badge_bg: '#0ea5e9' }, warning: { bg: '#f59e0b22', icon: '⚠️', badge_bg: '#f59e0b' }, danger: { bg: '#ef444422', icon: '🚨', badge_bg: '#ef4444' }, success: { bg: '#10b98122', icon: '✅', badge_bg: '#10b981' } };
      list.innerHTML = filtered.map(n => {
        const ic = icons[n.type] || icons.info;
        return `<div class="notif-item ${n.unread ? 'unread' : ''}" onclick="clickNotif(event,'${n.id}','${n.page}')">
          <div class="notif-item-avatar" style="background:${ic.bg};">${ic.icon}</div>
          <div class="notif-item-body">
            <div class="notif-item-title">${n.title}</div>
            <div class="notif-item-text">${n.text}</div>
            <div class="notif-item-time">${timeAgo(n.time)}</div>
          </div>
        </div>`;
      }).join('');
    }

    function clickNotif(event, id, page) {
      if (event) event.stopPropagation();
      const notif = notifications.find(n => String(n.id) === String(id));
      if (notif) notif.unread = false;
      let readNotifs = JSON.parse(localStorage.getItem('fcl_read_notifs') || '[]');
      if (!readNotifs.includes(String(id))) { readNotifs.push(String(id)); localStorage.setItem('fcl_read_notifs', JSON.stringify(readNotifs)); }
      updateNotifBadge();
      renderNotifications();
      showPage(page);
      setTimeout(closeNotifPanel, 200);
    }

    function markAllNotifRead() {
      let readNotifs = JSON.parse(localStorage.getItem('fcl_read_notifs') || '[]');
      notifications.forEach(n => { n.unread = false; if (!readNotifs.includes(String(n.id))) readNotifs.push(String(n.id)); });
      localStorage.setItem('fcl_read_notifs', JSON.stringify(readNotifs));
      updateNotifBadge();
      renderNotifications();
    }

    function timeAgo(date) {
      const s = Math.floor((new Date() - date) / 1000);
      if (s < 60) return '🟢 Baru saja';
      const m = Math.floor(s / 60); if (m < 60) return m + ' menit lalu';
      const h = Math.floor(m / 60); if (h < 24) return h + ' jam lalu';
      return Math.floor(h / 24) + ' hari lalu';
    }

    function loadNotifications() {
      google.script.run.withSuccessHandler(res => {
        if (!res.success) {
          console.error('Gagal mengambil notifikasi:', res.message);
          return;
        }

        console.log('Notifikasi Diterima:', res.data.length, 'item');

        const role = currentUser.role || '';
        const isPimpinan = (role === 'admin' || role === 'Super Admin' || role === 'SPV' || role === 'Supervisor' || role === 'Vice SPV' || role === 'Vice Supervisor' || role === 'HR' || role === 'Team Leader' || role === 'TL' || role.includes('Team Leader') || role.includes('Supervisor') || role.includes('SPV') || role.includes('HR'));

        let allowedMenus = [];
        try { allowedMenus = JSON.parse(currentUser.permissions || '[]'); } catch (e) { }

        const isHR = (role === 'HR' || allowedMenus.includes('aksesHr'));
        const canManageUsers = role === 'admin' || isHR || allowedMenus.includes('kelolaUser');

        let readNotifs = JSON.parse(localStorage.getItem('fcl_read_notifs') || '[]');
        notifications = [];

        res.data.forEach(d => {
          let hasAccess = false;
          const target = (d.targetAkses || '').toLowerCase();

          if (role === 'admin') {
            hasAccess = true;
          } else if (target === 'all') {
            hasAccess = true;
          } else if (target === 'approval' && isPimpinan) {
            hasAccess = true;
          } else if (target === 'users' && canManageUsers) {
            hasAccess = true;
          } else if (target === 'tugas' || target === 'tugasproject') {
            if (allowedMenus.includes('tugasProject')) hasAccess = true;
          } else if (allowedMenus.includes(d.targetAkses)) {
            // Jika targetAkses ada di dalam daftar menu yang diizinkan user
            hasAccess = true;
          }

          if (hasAccess) {
            notifications.push({
              id: d.id,
              title: d.title,
              text: d.text,
              type: d.type || 'info',
              page: d.page || 'dashboard',
              time: new Date(d.tanggal),
              unread: !readNotifs.includes(String(d.id))
            });
          }
        });

        console.log('Notifikasi Ditampilkan (Setelah Filter):', notifications.length, 'item');
        updateNotifBadge();
        if (notifPanelOpen) renderNotifications();
      }).getNotifikasi();
    }

    // === UTILS ===
    const v = id => document.getElementById(id)?.value || '';
    const setVal = (id, val) => {
      const el = document.getElementById(id);
      if (el) {
        if (el.tagName === 'INPUT' || el.tagName === 'SELECT' || el.tagName === 'TEXTAREA') el.value = val;
        else el.textContent = val;
      }
    };
    const openModal = id => document.getElementById(id).classList.add('show');
    const closeModal = id => document.getElementById(id).classList.remove('show');
    const formatRpInput = el => { let raw = el.value.replace(/[^\d]/g, ''); if (!raw) { el.value = ''; return; } el.value = parseInt(raw, 10).toLocaleString('id-ID'); };
    const getRpValue = id => { const raw = v(id); return parseFloat(raw.replace(/[^\d]/g, '')) || 0; };
    const formatRp = n => 'Rp ' + (parseFloat(n) || 0).toLocaleString('id-ID');
    const formatDate = d => {
      if (!d) return '-';
      try {
        let dt;
        if (typeof d === 'string' && d.includes('-') && d.length === 10) {
          const parts = d.split('-');
          if (parts.length === 3) {
            dt = new Date(parts[0], parts[1] - 1, parts[2]);
          } else {
            dt = new Date(d);
          }
        } else {
          dt = new Date(d);
        }
        return isNaN(dt.getTime()) ? d : dt.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' });
      } catch (e) {
        return d;
      }
    };
    const normalizeDateKey = value => {
      if (!value && value !== 0) return '';
      const dateStr = String(value).trim();
      if (!dateStr) return '';
      if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return dateStr;
      if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(dateStr)) {
        const [d, m, y] = dateStr.split('/');
        return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
      }
      if (/^\d{1,2}-\d{1,2}-\d{4}$/.test(dateStr)) {
        const [d, m, y] = dateStr.split('-');
        return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
      }
      const parsed = new Date(dateStr);
      return isNaN(parsed.getTime()) ? dateStr : parsed.toISOString().split('T')[0];
    };
    let toastTimer;
    const toast = (msg, type = 'success', callback = null) => {
      const el = document.getElementById('toast');
      if (!el) return;
      const translatedMsg = getLangText(msg);
      el.textContent = (type === 'success' ? '✅ ' : '❌ ') + translatedMsg;
      el.className = 'show ' + type;
      if (callback) {
        el.classList.add('clickable');
        el.style.cursor = 'pointer';
        el.onclick = () => { callback(); el.className = ''; };
      } else {
        el.classList.remove('clickable');
        el.style.cursor = 'default';
        el.onclick = null;
      }
      clearTimeout(toastTimer);
      toastTimer = setTimeout(() => el.className = '', callback ? 6000 : 3000);
    };
    const showToast = toast;

    const resetForm = ids => ids.forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
    function escHtml(str) { if (!str) return ''; return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

    function ensureDistributorQueueSheet(silent = true) {
      google.script.run
        .withSuccessHandler(res => {
          if (!res || !res.success) {
            const msg = res ? res.message : 'Gagal menyiapkan sheet Antrian Distributor';
            console.error(msg);
            if (!silent && typeof toast === 'function') toast(msg, 'error');
          }
        })
        .withFailureHandler(err => {
          console.error('Setup Antrian Distributor gagal:', err);
          if (!silent && typeof toast === 'function') toast('Setup Antrian Distributor gagal: ' + err, 'error');
        })
        .setupDistributorQueueDatabase();
    }

    function getCurrentUserPermissions() {
      if (!currentUser) return [];
      try { return JSON.parse(currentUser.permissions || '[]'); } catch (e) { return []; }
    }

    function hasPermission(permissionKey) {
      if (!currentUser) return false;
      if (currentUser.role === 'admin' || currentUser.role === 'Super Admin') return true;
      const perms = getCurrentUserPermissions();
      return perms.includes(permissionKey);
    }

    function canViewDistributorQueue() {
      if (!currentUser) return false;
      if (currentUser.role === 'admin') return true;
      const perms = getCurrentUserPermissions();
      return perms.includes('lihatAntrianDistributor') || perms.includes('editAntrianDistributor') || perms.includes('lihatDashboardLateShipment') || perms.includes('antrianDistributor') || perms.includes('rekapOngkirMistine');
    }

    function canEditDistributorQueue() {
      if (!currentUser) return false;
      if (currentUser.role === 'admin') return true;
      const perms = getCurrentUserPermissions();
      return perms.includes('editAntrianDistributor') || perms.includes('antrianDistributor');
    }

    function canViewLateShipmentDashboard() {
      if (!currentUser) return false;
      if (currentUser.role === 'admin') return true;
      const perms = getCurrentUserPermissions();
      return perms.includes('lihatDashboardLateShipment') || perms.includes('antrianDistributor');
    }

    function canUpdateLateNote() {
      if (!currentUser) return false;
      if (currentUser.role === 'admin') return true;
      const perms = getCurrentUserPermissions();
      return perms.includes('editSLARule') || perms.includes('lihatDashboardLateShipment') || perms.includes('editAntrianDistributor');
    }

    function canEditSLARule() {
      if (!currentUser) return false;
      if (currentUser.role === 'admin') return true;
      const perms = getCurrentUserPermissions();
      return perms.includes('editSLARule');
    }

    function applyDistributorQueuePermissions() {
      const formCard = document.getElementById('dqFormCard');
      const lateCard = document.getElementById('dqLateCard');
      const headerActions = document.getElementById('dqHeaderActions');
      const actionHeader = document.getElementById('dqActionHeader');
      const canEdit = canEditDistributorQueue();
      const canViewLate = canViewLateShipmentDashboard();

      if (formCard) formCard.style.display = canEdit ? 'block' : 'none';
      if (lateCard) lateCard.style.display = canViewLate ? 'block' : 'none';
      if (headerActions) headerActions.style.display = canEdit ? 'flex' : 'none';
      if (actionHeader) actionHeader.textContent = canEdit ? 'Aksi' : 'Info';
    }

    window.onload = function () {
      ensureDistributorQueueSheet(true);
      const ls = document.getElementById('loadingScreen');
      const lp = document.getElementById('loginPage');
      const ap = document.getElementById('app');

      // ── FAILSAFE LAYER 1: 3s – jika loading masih tampil, paksa muncul login ──
      const failsafe1 = setTimeout(function () {
        try {
          const lsEl = document.getElementById('loadingScreen');
          if (lsEl && lsEl.style.display !== 'none') {
            console.warn('[FAILSAFE-1] Loading stuck, forcing login page...');
            lsEl.style.display = 'none';
            const apEl = document.getElementById('app');
            if (apEl) apEl.style.display = 'none';
            const lpEl = document.getElementById('loginPage');
            if (lpEl) { lpEl.style.display = 'flex'; lpEl.style.zIndex = '9999'; lpEl.style.opacity = '1'; }
          }
        } catch (e) { console.error('Failsafe-1 error:', e); }
      }, 3000);

      // ── FAILSAFE LAYER 2: 6s – paksa mutlak tanpa kondisi ──
      const failsafe2 = setTimeout(function () {
        try {
          const lsEl = document.getElementById('loadingScreen');
          if (lsEl) lsEl.style.setProperty('display', 'none', 'important');
          const apEl = document.getElementById('app');
          if (apEl && apEl.style.display !== 'block') apEl.style.display = 'none';
          const lpEl = document.getElementById('loginPage');
          const appIsShown = apEl && apEl.style.display === 'block';
          if (!appIsShown && lpEl) { lpEl.style.setProperty('display', 'flex', 'important'); lpEl.style.zIndex = '9999'; }
        } catch (e) { console.error('Failsafe-2 error:', e); }
      }, 6000);

      try {
        let saved = sessionStorage.getItem('fcl_user');
        if (!saved) {
          saved = localStorage.getItem('fcl_user');
          if (saved) {
            sessionStorage.setItem('fcl_user', saved);
          }
        }

        if (saved) {
          currentUser = JSON.parse(saved);
          clearTimeout(failsafe1);
          clearTimeout(failsafe2);
          initApp();
        } else {
          clearTimeout(failsafe1);
          clearTimeout(failsafe2);
          if (ls) ls.style.display = 'none';
          if (ap) ap.style.display = 'none';
          if (lp) { lp.style.display = 'flex'; lp.style.zIndex = '9999'; }
        }
      } catch (e) {
        clearTimeout(failsafe1);
        clearTimeout(failsafe2);
        console.error("Init Error:", e);
        if (ls) ls.style.display = 'none';
        if (lp) lp.style.display = 'flex';
      }
    };


    function initApp() {
      ensureDistributorQueueSheet(true);
      // Remove redundant attendance loading from startup - causes app to hang
      // These will be loaded after menu is displayed

      // Emergency Layout Fix: Force-hide all overlays
      const ls = document.getElementById('loadingScreen');
      const lp = document.getElementById('loginPage');
      if (ls) { ls.style.setProperty('display', 'none', 'important'); }
      if (lp) { lp.style.setProperty('display', 'none', 'important'); }

      const role = currentUser.role || '';
      const isPimpinan = (role === 'admin' || role === 'Super Admin' || role === 'SPV' || role === 'Supervisor' || role === 'Vice SPV' || role === 'Vice Supervisor' || role === 'Vice VPV' || role === 'HR' || role === 'Team Leader' || role === 'TL' || role.includes('Team Leader') || role.includes('Supervisor') || role.includes('SPV') || role.includes('HR'));

      document.getElementById('app').style.display = 'block';
      document.getElementById('userName').textContent = currentUser.nama;
      document.getElementById('userRole').textContent = role === 'admin' ? '👑 Admin' : '👤 ' + role;
      document.getElementById('userAvatar').textContent = currentUser.nama.charAt(0).toUpperCase();

      if (role === 'admin' || role === 'SPV' || role === 'Supervisor' || role === 'Vice SPV' || role === 'Vice Supervisor') {
        const btnExp = document.getElementById('btnExportLembur');
        if (btnExp) btnExp.style.display = 'inline-block';
      }

      let allowedMenus = []; if (currentUser.permissions) { try { allowedMenus = JSON.parse(currentUser.permissions); } catch (e) { } }
      const isHR = (role === 'HR');
      const isViceSPV = (role === 'Vice SPV' || role === 'Vice Supervisor');
      const canManageUsers = role === 'admin' || isHR || isViceSPV || allowedMenus.includes('kelolaUser');

      if (currentUser.role !== 'admin') {
        let firstAllowedPage = null;
        document.querySelectorAll('.nav-item').forEach(el => {
          const match = (el.getAttribute('onclick') || '').match(/showPage\('([^']+)'\)/);
          if (match) {
            const pName = match[1];
            // Visibility Check
            let isVisible = allowedMenus.includes(pName);
            if (pName === 'antrianDistributor') isVisible = canViewDistributorQueue();
            if (pName === 'users' && canManageUsers) isVisible = true;
            if (pName === 'returnDistributor' && (role === 'admin' || allowedMenus.includes('returnDistributor'))) isVisible = true;
            if (pName === 'approval') isVisible = isPimpinan; // Strictly role-based
            if (pName === 'approvalDashboard') isVisible = role === 'admin' || isPimpinan || allowedMenus.includes('approvalDashboard');

            if (isVisible) {
              el.style.display = 'flex';
              if (!firstAllowedPage) firstAllowedPage = pName;
            } else {
              el.style.display = 'none';
            }
          }
        });
        document.querySelectorAll('.nav-section').forEach(sec => {
          let nextEl = sec.nextElementSibling; let hasVis = false;
          while (nextEl && !nextEl.classList.contains('nav-section')) { if (nextEl.style.display !== 'none') { hasVis = true; break; } nextEl = nextEl.nextElementSibling; }
          if (sec.id === 'navAdmin') { if (canManageUsers || role === 'admin' || allowedMenus.includes('returnDistributor')) sec.style.display = 'block'; else sec.style.display = 'none'; }
          else if (!hasVis) sec.style.display = 'none';
          else sec.style.display = 'block';
        });
        if (canManageUsers) {
          document.getElementById('navAdmin').style.display = 'block';
          document.getElementById('navUsers').style.display = 'flex';
        }
        const canViewReturnDist = role === 'admin' || allowedMenus.includes('returnDistributor');
        if (canViewReturnDist) {
          document.getElementById('navAdmin').style.display = 'block';
          document.getElementById('navReturnDistributor').style.display = 'flex';
        }
        if (isPimpinan) {
          document.getElementById('navCentralApproval').style.display = 'flex';
        }
        const dashboardNav = Array.from(document.querySelectorAll('.nav-item')).find(el => el.getAttribute('onclick')?.includes("showPage('dashboard')") && el.style.display !== 'none');
        if (dashboardNav) {
          showPage('dashboard');
        } else if (firstAllowedPage) {
          showPage(firstAllowedPage);
        } else {
          document.querySelector('.content').innerHTML = '<div style="text-align:center;padding:50px;color:var(--gray)"><div style="font-size:40px; margin-bottom:10px;">⛔</div>Anda belum memiliki akses menu.<br>Silakan hubungi Administrator.</div>';
        }
      } else {
        document.getElementById('navAdmin').style.display = 'block';
        document.getElementById('navUsers').style.display = 'flex';
        document.getElementById('navReturnDistributor').style.display = 'flex';
        document.getElementById('navCentralApproval').style.display = 'flex';
        const navApprDash = document.getElementById('navApprovalDashboard');
        if (navApprDash) navApprDash.style.display = 'flex';
        showPage('dashboard');
      }
      setToday();
      updateLemburStaffSelect();

      // Visibilitas tombol khusus pimpinan
      const isPimpinanLembur = (role === 'admin' || role === 'SPV' || role === 'Supervisor' || role === 'Vice SPV' || role === 'Vice Supervisor' || role === 'Team Leader' || role === 'TL' || role.includes('Supervisor') || role.includes('SPV'));
      const btnK = document.getElementById('btnLemburKhusus');
      if (btnK) btnK.style.display = isPimpinanLembur ? 'block' : 'none';
      const btnE = document.getElementById('btnExportLembur');
      if (btnE) btnE.style.display = (isPimpinanLembur || role === 'HR') ? 'block' : 'none';

      // Granular Overtime Controls
      const btnTglMerah = document.getElementById('btnTglMerahSetting');
      if (btnTglMerah) btnTglMerah.style.display = (role === 'admin' || allowedMenus.includes('pengaturanTglMerah')) ? 'block' : 'none';

      // Show/hide Opname Report button for management and stock opname users
      const btnOpname = document.getElementById('btnOpnameReports');
      const canViewOpnameReports = (role === 'admin' || isPimpinan || allowedMenus.includes('laporanSO') || allowedMenus.includes('stockOpname') || allowedMenus.includes('stockOpnameAsset') || canManageUsers);
      if (btnOpname) btnOpname.style.display = canViewOpnameReports ? 'inline-block' : 'none';

      if (typeof checkPendingApprovals === 'function') {
        try { checkPendingApprovals(); } catch (e) { console.warn('checkPendingApprovals error:', e); }
        setInterval(checkPendingApprovals, 2 * 60 * 1000); // 2 minutes
      }
      try { loadNotifications(); } catch (e) { console.warn('loadNotifications error:', e); }
      // loadMyAttendanceSummary (moved after menu display to prevent app hang)
      setInterval(() => { try { loadNotifications(); } catch (e) { } }, 1 * 60 * 1000); // 1 minute
      if (typeof loadUsersDropdown === 'function') {
        try { loadUsersDropdown(); } catch (e) { console.warn('loadUsersDropdown error:', e); }
      }

      // Auto-load attendance widget on login (dashboard load happens via showPage)
      setTimeout(() => {
        try { loadMyAttendanceSummary(); } catch (e) { console.warn('Attendance load error:', e); }
      }, 500);

      // Forced Password Change Check
      if (currentUser.isDefaultPassword) {
        forceChangePassword();
      }
    }
    // checkPendingApprovals original implementation is below

    function setToday() {
      const t = new Date().toISOString().split('T')[0];
      ['kgTanggal', 'tbTanggal', 'kTglMasuk', 'inbTanggal', 'outTanggal', 'ordTanggal', 'retTanggal', 'lapTanggal', 'hoTanggal', 'klTanggal', 'ijTanggal', 'lbTanggal', 'exTanggal', 'tpTanggalMulai', 'assetTanggal', 'soTanggal', 'plTanggal', 'awTanggal'].forEach(id => setVal(id, t));
      if (currentUser) {
        const role = currentUser.role || '';
        const isAdminOrHR = (role === 'admin' || role === 'HR' || role === 'SPV' || role === 'Supervisor' || role === 'Vice SPV' || role === 'Vice Supervisor' || role === 'Team Leader' || role === 'TL' || role.includes('Supervisor') || role.includes('SPV') || role.includes('HR'));
        ['ijNama', 'assetNama'].forEach(id => {
          const el = document.getElementById(id);
          if (el) { el.value = currentUser.nama; el.readOnly = true; el.style.background = 'var(--input-bg)'; el.style.opacity = '0.7'; }
        });
        const lbN = document.getElementById('lbNama');
        if (lbN && !isAdminOrHR) { lbN.disabled = true; lbN.style.opacity = '0.7'; }
      }
      setVal('filterBulanExpense', (new Date().getMonth() + 1).toString()); setVal('filterBulanGrafik', t.substring(0, 7));
      setVal('filterMonthOpname', t.substring(0, 7));
    }

    function doLogin() {
      const user = v('loginUser').trim(), pass = v('loginPass');
      if (!user || !pass) { document.getElementById('loginError').style.display = 'block'; document.getElementById('loginError').textContent = '⚠️ Isi username dan password'; return; }
      const btn = document.querySelector('.btn-login'); btn.textContent = '⏳ Memuat...';
      google.script.run.withSuccessHandler(res => {
        btn.textContent = '🔐 Masuk';
        if (res.success) {
          currentUser = res.user;
          sessionStorage.setItem('fcl_user', JSON.stringify(currentUser));
          localStorage.setItem('fcl_user', JSON.stringify(currentUser));
          ensureDistributorQueueSheet(true);
          initApp();
        }
        else { document.getElementById('loginError').style.display = 'block'; document.getElementById('loginError').textContent = '⚠️ ' + res.message; }
      }).login(user, pass);
    }

    function forceChangePassword() {
      document.getElementById('changePassTitle').innerHTML = '<span style="color:var(--red)">⚠️ Wajib Ganti Password</span><br><small style="font-size:11px;font-weight:400;color:var(--gray)">Silakan ganti password default "1" Anda untuk melanjutkan.</small>';
      document.getElementById('btnExitPassModal').style.display = 'none';
      document.getElementById('btnCancelPassModal').style.display = 'none';
      openModal('modalChangePassword');
    }
    document.addEventListener('keydown', e => { if (e.key === 'Enter' && document.getElementById('loginPage').style.display !== 'none') doLogin(); });
    function submitChangePassword() {
      const o = v('oldPassModal'), n = v('newPassModal'), c = v('confirmPassModal');
      if (!o || !n || !c) return toast('Lengkapi semua field', 'error');
      if (n !== c) return toast('Password konfirmasi tidak cocok', 'error');
      if (n.length < 5) return toast('Password minimal 5 karakter', 'error');

      const btn = document.querySelector('#modalChangePassword .btn-primary');
      const oldTxt = btn.textContent;
      btn.disabled = true; btn.textContent = '⏳ Memproses...';

      google.script.run.withSuccessHandler(res => {
        btn.disabled = false; btn.textContent = oldTxt;
        if (res.success) {
          toast('Password berhasil diperbarui ✅');
          currentUser.isDefaultPassword = false;
          sessionStorage.setItem('fcl_user', JSON.stringify(currentUser));
          document.getElementById('btnExitPassModal').style.display = 'flex';
          document.getElementById('btnCancelPassModal').style.display = 'inline-flex';
          document.getElementById('changePassTitle').textContent = '🔑 Ganti Password';
          closeModal('modalChangePassword');
          resetForm(['oldPassModal', 'newPassModal', 'confirmPassModal']);
        } else {
          toast(res.message, 'error');
        }
      }).changePassword(currentUser.username, o, n);
    }

    function doLogout() {
      sessionStorage.clear();
      localStorage.removeItem('fcl_user');
      currentUser = null;
      document.getElementById('app').style.display = 'none';
      document.getElementById('loginPage').style.display = 'flex';
      document.getElementById('loginPage').style.opacity = '1';
      document.getElementById('loginUser').value = '';
      document.getElementById('loginPass').value = '';
    }

    const pageTitles = {
      dashboard: '📊 Dashboard', kasGudang: '💰 Kas Gudang', teamBuilding: '🤝 Team Building', expense: '💸 Expense List', pettyCash: '🪙 Petty Cash', paymentGudang: '💄 MISTINE',
      rekapOngkirMistine: '💰 Rekap Ongkir MISTINE',
      karyawan: '👥 Data Karyawan', laporanKerja: '📝 Input Laporan Kerja', grafikLaporan: '📈 Dashboard Laporan Kerja',
      handover: '📦 Laporan Stock Control', klaim: '⚠️ Laporan Klaim Paket', tugasProject: '📋 Tugas Project',
      ijin: '✉️ Pengajuan Ijin & Cuti', lembur: '⏱️ Pengajuan Lembur', pengajuanAsset: '📦 Pengajuan Asset',
      organisasi: '🏗️ Struktur Organisasi', sop: '📋 SOP Gudang',
      users: '⚙️ Manajemen User', stock: '📦 Stock Barang', stockOpname: '⚖️ Stock Opname', packingList: '📋 Dokumen INBOUND',
      inbound: '📥 Inbound Barang', outbound: '✂️ Pemotongan Stok',
      retur: '↩️ Retur Barang', order: '🛒 Orderan', antrianDistributor: '🚚 Antrian Distributor', analisis: '📈 Analisis Pemakaian Stok',
      assetWarehouse: '🏪 Asset Warehouse',
      bookingMobil: '🚗 Booking Mobil',
      approvalDashboard: '⚖️ Dashboard Approval',
      approval: '✅ Pusat Approval'
    };
    const pageLoaders = {
      dashboard: loadDashboard, kasGudang: loadKasGudang, teamBuilding: loadTB, expense: loadExpense, pettyCash: loadPettyCash, paymentGudang: loadPaymentGudang,
      karyawan: loadKaryawan, laporanKerja: loadLaporanKerja, grafikLaporan: loadGrafikLaporan,
      handover: loadStockControl, klaim: loadKlaim, tugasProject: loadTugasProject, ijin: loadIjin, lembur: loadLembur,
      pengajuanAsset: loadAsset,
      organisasi: loadOrg, sop: loadSOP, kpiKaryawan: loadKpiKaryawan, users: loadUsers, stock: loadStock, stockOpname: loadStockOpname, packingList: loadPackingList, inbound: loadInbound, outbound: loadOutbound, retur: loadRetur, order: loadOrder, antrianDistributor: loadDistributorQueue, analisis: loadAnalisis, assetWarehouse: loadAssetWarehouse, bookingMobil: loadBookingMobil,
      returnDistributor: loadReturnDistributor,
      approvalDashboard: loadApprovalDashboard,
      approval: loadCentralizedApprovalsOptimized, approvalCenter: renderApprovalCenter
    };

    function getKpiStorageKey() {
      return currentUser && currentUser.username ? `kpi_karyawan_${currentUser.username}` : 'kpi_karyawan_guest';
    }

    let kpiQuestionBanks = {
      Warehouse: [
        { question: 'Langkah pertama yang benar saat menerima barang masuk?', options: [
            { label: 'Periksa dokumen, jumlah, dan kondisi barang', point: 25 },
            { label: 'Langsung simpan barang tanpa pemeriksaan', point: 0 },
            { label: 'Tunggu supervisor memeriksa terlebih dahulu', point: 10 },
            { label: 'Tarik pallet tanpa membuka segel', point: 5 }
          ] },
        { question: 'Bagaimana cara menjaga akurasi stok gudang?', options: [
            { label: 'Melakukan pengecekan berkala dan pencatatan rapi', point: 25 },
            { label: 'Hanya mengandalkan ingatan saja', point: 0 },
            { label: 'Memindahkan barang tanpa update sistem', point: 5 },
            { label: 'Menunggu audit tahunan', point: 10 }
          ] },
        { question: 'Apa yang harus dilakukan saat menemukan kerusakan barang?', options: [
            { label: 'Catat segera dan laporkan ke supervisor', point: 25 },
            { label: 'Biarkan saja dan lanjutkan pekerjaan', point: 0 },
            { label: 'Rubutkan ke rekan kerja tanpa dokumentasi', point: 5 },
            { label: 'Buang barang tanpa izin', point: 0 }
          ] },
        { question: 'Sikap terbaik saat melayani tim operasional lain?', options: [
            { label: 'Responsif, sopan, dan bantu selesaikan masalah', point: 25 },
            { label: 'Mengabaikan jika bukan tugas saya', point: 0 },
            { label: 'Memberi jawaban samar-samar', point: 10 },
            { label: 'Menunggu perintah atasan baru bertindak', point: 5 }
          ] }
      ],
      Inbound: [
        { question: 'Dokumen apa yang wajib diverifikasi sebelum menerima inbound?', options: [
            { label: 'Surat jalan, invoice, dan daftar muatan', point: 25 },
            { label: 'Hanya melihat paket secara kasat mata', point: 0 },
            { label: 'Cek jumlah di akhir shift', point: 5 },
            { label: 'Tunggu petugas lain yang memeriksa', point: 10 }
          ] },
        { question: 'Bagaimana menangani barang inbound yang tidak sesuai?', options: [
            { label: 'Lapor dan pisahkan untuk tindak lanjut', point: 25 },
            { label: 'Simpan seperti biasa', point: 0 },
            { label: 'Buang agar tidak merepotkan', point: 0 },
            { label: 'Ubah label agar cocok', point: 0 }
          ] },
        { question: 'Sikap terbaik saat melakukan pengecekan kualitas?', options: [
            { label: 'Teliti dan terus jaga konsistensi', point: 25 },
            { label: 'Cukup lihat sepintas', point: 0 },
            { label: 'Tunda sampai ada waktu longgar', point: 10 },
            { label: 'Abaikan barang kecil yang rusak', point: 5 }
          ] },
        { question: 'Kapan Anda harus melakukan input data inbound ke sistem?', options: [
            { label: 'Segera setelah barang diterima dan dicek', point: 25 },
            { label: 'Besok pagi saja', point: 0 },
            { label: 'Setelah semua barang selesai diurut', point: 10 },
            { label: 'Hanya saat diminta supervisor', point: 5 }
          ] }
      ],
      Distributor: [
        { question: 'Apa prioritas utama saat menyiapkan order distributor?', options: [
            { label: 'Kecepatan, akurasi, dan kondisi barang baik', point: 25 },
            { label: 'Hanya cepat tanpa cek ulang', point: 0 },
            { label: 'Pilih barang secara acak', point: 0 },
            { label: 'Tunda hingga semua pesanan selesai', point: 10 }
          ] },
        { question: 'Bagaimana cara memastikan qty order benar?', options: [
            { label: 'Hitung ulang sebelum packing', point: 25 },
            { label: 'Percaya tanpa menghitung', point: 0 },
            { label: 'Tanya rekan kerja saja', point: 10 },
            { label: 'Hanya hitung sebagian', point: 5 }
          ] },
        { question: 'Apa yang harus dilakukan jika ada item yang rusak?', options: [
            { label: 'Pisahkan, laporkan, dan ajukan retur', point: 25 },
            { label: 'Tetap kirim agar tidak telat', point: 0 },
            { label: 'Tukar dengan item lain tanpa catatan', point: 0 },
            { label: 'Biarkan di gudang', point: 5 }
          ] },
        { question: 'Sikap terbaik saat bekerja dalam tim distributor?', options: [
            { label: 'Komunikasi jelas dan bantu rekan', point: 25 },
            { label: 'Kerja sendiri saja', point: 0 },
            { label: 'Menunggu instruksi terus', point: 10 },
            { label: 'Menyalahkan orang lain jika salah', point: 0 }
          ] }
      ],
      HR: [
        { question: 'Apa tujuan utama evaluasi karyawan rutin?', options: [
            { label: 'Meningkatkan kinerja dan kepuasan kerja', point: 25 },
            { label: 'Hanya mengikuti prosedur', point: 0 },
            { label: 'Untuk memberi nilai saja', point: 5 },
            { label: 'Agar terlihat sibuk', point: 0 }
          ] },
        { question: 'Bagaimana memastikan data karyawan tersimpan rapi?', options: [
            { label: 'Update sistem dan arsip secara konsisten', point: 25 },
            { label: 'Simpan hanya di catatan pribadi', point: 0 },
            { label: 'Tunggu audit tahunan', point: 10 },
            { label: 'Simpan di tempat acak', point: 0 }
          ] },
        { question: 'Sikap terbaik saat menerima keluhan staf?', options: [
            { label: 'Mendengarkan dengan empati dan memberi tindak lanjut', point: 25 },
            { label: 'Mendiscount dan tidak menindaklanjuti', point: 0 },
            { label: 'Menyarankan untuk lapor sendiri saja', point: 5 },
            { label: 'Menyalahkan staf atas masalahnya', point: 0 }
          ] },
        { question: 'Kapan Anda harus menindaklanjuti permintaan cuti?', options: [
            { label: 'Segera setelah menerima pengajuan lengkap', point: 25 },
            { label: 'Tunda hingga deadline', point: 0 },
            { label: 'Hanya setelah diingatkan', point: 5 },
            { label: 'Jangan menindaklanjuti sama sekali', point: 0 }
          ] }
      ],
      Operasional: [
        { question: 'Bagaimana mendeteksi risiko keselamatan di area kerja?', options: [
            { label: 'Patroli rutin dan laporkan setiap bahaya', point: 25 },
            { label: 'Hanya menunggu kecelakaan terjadi', point: 0 },
            { label: 'Tanyakan supervisor jika sempat', point: 10 },
            { label: 'Abaikan jika kecil', point: 0 }
          ] },
        { question: 'Apa yang harus dilakukan ketika tugas mendesak muncul?', options: [
            { label: 'Atur prioritas dan komunikasikan tim', point: 25 },
            { label: 'Kerjakan tanpa koordinasi', point: 0 },
            { label: 'Biarkan tim lain yang menangani', point: 5 },
            { label: 'Tunda sampai akhir shift', point: 0 }
          ] },
        { question: 'Bagaimana menjaga mutu output operasional?', options: [
            { label: 'Ikuti standar kerja dan lakukan pemeriksaan', point: 25 },
            { label: 'Selesaikan cepat tanpa cek ulang', point: 0 },
            { label: 'Serahkan ke tim lain', point: 10 },
            { label: 'Hanya lakukan saat diminta', point: 5 }
          ] },
        { question: 'Apa yang penting saat berkomunikasi dengan lini lain?', options: [
            { label: 'Jelas, sopan, dan tepat waktu', point: 25 },
            { label: 'Singkat tanpa detail', point: 0 },
            { label: 'Tunda sampai ada waktu', point: 10 },
            { label: 'Beri jawaban ambigu', point: 0 }
          ] }
      ],
      Admin: [
        { question: 'Bagaimana cara menjaga dokumen administrasi teratur?', options: [
            { label: 'Klasifikasi jelas dan arsip rutin', point: 25 },
            { label: 'Tumpuk di satu folder saja', point: 0 },
            { label: 'Simpan di komputer tanpa backup', point: 10 },
            { label: 'Biarkan rekan yang mengurusnya', point: 0 }
          ] },
        { question: 'Kapan sebaiknya memperbarui catatan biaya?', options: [
            { label: 'Segera setelah transaksi terjadi', point: 25 },
            { label: 'Akhir bulan saja', point: 5 },
            { label: 'Saat diminta audit', point: 0 },
            { label: 'Tidak perlu diperbarui', point: 0 }
          ] },
        { question: 'Apa fungsi utama data arsip?', options: [
            { label: 'Sebagai bukti dan referensi operasional', point: 25 },
            { label: 'Hanya untuk formalitas', point: 0 },
            { label: 'Untuk disimpan saja', point: 5 },
            { label: 'Tidak perlu diperhatikan', point: 0 }
          ] },
        { question: 'Bagaimana memastikan informasi tersampaikan dengan baik?', options: [
            { label: 'Gunakan bahasa jelas dan lengkap', point: 25 },
            { label: 'Kirim tanpa detail penting', point: 0 },
            { label: 'Sampaikan hanya sebagian', point: 10 },
            { label: 'Biarkan penerima menebak', point: 0 }
          ] }
      ]
    };

    function loadKpiQuestions(callback) {
      if (!isGoogleScriptAvailable()) {
        if (callback) callback();
        return;
      }
      const divisi = document.getElementById('kpiDivisi')?.value || '';
      google.script.run.withSuccessHandler(res => {
        if (res && res.success && Array.isArray(res.data)) {
          const bank = {};
          res.data.forEach(item => {
            const division = String(item.divisi || 'General').trim() || 'General';
            if (!bank[division]) bank[division] = [];
            bank[division].push({
              question: item.question,
              options: Array.isArray(item.options) ? item.options : []
            });
          });
          kpiQuestionBanks = Object.assign({}, kpiQuestionBanks, bank);
        }
        if (callback) callback();
      }).withFailureHandler(err => {
        console.error('Failed to load KPI questions:', err);
        if (callback) callback();
      }).getKpiQuestions(divisi);
    }

    function downloadKpiQuestionsTemplate() {
      const data = [
        ['ID (Kosongkan jika baru)', 'Divisi', 'Pertanyaan', 'Pilihan 1 Label', 'Pilihan 1 Poin', 'Pilihan 2 Label', 'Pilihan 2 Poin', 'Pilihan 3 Label', 'Pilihan 3 Poin', 'Pilihan 4 Label', 'Pilihan 4 Poin'],
        ['', 'Warehouse', 'Apa langkah pertama saat menerima barang masuk?', 'Periksa dokumen dan kondisi barang', 25, 'Langsung simpan tanpa cek', 0, 'Tunggu supervisor memeriksa', 10, 'Tarik pallet tanpa membuka segel', 5],
        ['', 'HR', 'Bagaimana cara menjaga kedisiplinan tim?', 'Reminder kehadiran dan monitoring', 25, 'Biarkan setiap orang mengatur sendiri', 0, 'Hanya catat masalah bila terjadi', 10, 'Serahkan sepenuhnya pada presensi', 5]
      ];
      const ws = XLSX.utils.aoa_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Template KPI Questions');
      XLSX.writeFile(wb, 'Template_KPI_Questions.xlsx');
    }

    function handleImportKpiQuestions(input) {
      if (!input.files.length) return;
      const file = input.files[0];
      const reader = new FileReader();
      reader.onload = e => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

        if (!rows.length) return toast('File kosong', 'error');

        const items = rows.map(r => ({
          id: String(r['ID (Kosongkan jika baru)'] || '').trim(),
          divisi: String(r['Divisi'] || 'General').trim() || 'General',
          question: String(r['Pertanyaan'] || '').trim(),
          opt1Label: String(r['Pilihan 1 Label'] || '').trim(),
          opt1Point: parseInt(r['Pilihan 1 Poin'] || 0, 10) || 0,
          opt2Label: String(r['Pilihan 2 Label'] || '').trim(),
          opt2Point: parseInt(r['Pilihan 2 Poin'] || 0, 10) || 0,
          opt3Label: String(r['Pilihan 3 Label'] || '').trim(),
          opt3Point: parseInt(r['Pilihan 3 Poin'] || 0, 10) || 0,
          opt4Label: String(r['Pilihan 4 Label'] || '').trim(),
          opt4Point: parseInt(r['Pilihan 4 Poin'] || 0, 10) || 0
        })).filter(x => x.question);

        if (!items.length) return toast('Tidak ada pertanyaan KPI valid', 'error');

        google.script.run.withSuccessHandler(res => {
          input.value = '';
          if (res && res.success) {
            toast('Import KPI Questions berhasil');
            loadKpiQuestions(() => renderKpiQuestions());
          } else {
            toast(res?.message ? 'Gagal impor: ' + res.message : 'Gagal impor KPI Questions.', 'error');
          }
        }).withFailureHandler(err => {
          console.error('Import KPI Questions error:', err);
          toast('Gagal mengimpor KPI Questions.', 'error');
          input.value = '';
        }).addBulkKpiQuestions(items);
      };
      reader.readAsArrayBuffer(file);
    }

    function renderKpiQuestions() {
      const division = document.getElementById('kpiDivisi')?.value || 'General';
      const questions = kpiQuestionBanks[division] || kpiQuestionBanks.General || [];
      const container = document.getElementById('kpiQuestionsContainer');
      if (!container) return;
      if (!questions.length) {
        container.innerHTML = '<div class="card"><div class="card-body">Tidak ada pertanyaan untuk divisi ini.</div></div>';
        return;
      }
      container.innerHTML = questions.map((q, idx) => `
        <div class="kpi-question-card">
          <div class="kpi-question-title"><strong>Soal ${idx + 1}.</strong> ${q.question}</div>
          ${q.options.map(opt => `
            <div class="form-check">
              <input class="form-check-input" type="radio" id="kpiQ${idx}-${opt.point}" name="kpiQ${idx}" value="${opt.point}" data-answer="${opt.label}" onchange="calcKpiScore()">
              <label class="form-check-label" for="kpiQ${idx}-${opt.point}">
                ${opt.label} <span style="color:var(--gray);">(${opt.point} poin)</span>
              </label>
            </div>
          `).join('')}
        </div>
      `).join('');
      calcKpiScore();
    }

    function calcKpiScore() {
      const container = document.getElementById('kpiQuestionsContainer');
      if (!container) return;
      let total = 0;
      const cards = container.querySelectorAll('.kpi-question-card');
      cards.forEach((card, idx) => {
        const checked = card.querySelector(`input[name="kpiQ${idx}"]:checked`);
        if (checked) total += parseInt(checked.value, 10) || 0;
      });
      const totalField = document.getElementById('kpiTotalPoint');
      const gradeField = document.getElementById('kpiGrade');
      if (totalField) totalField.value = total;
      if (gradeField) gradeField.textContent = getKpiGrade(total);
    }

    function getKpiGrade(total) {
      if (total >= 90) return 'A - Sangat Baik';
      if (total >= 75) return 'B - Baik';
      if (total >= 60) return 'C - Cukup';
      if (total > 0) return 'D - Perlu Peningkatan';
      return 'Belum Dinilai';
    }

    function isGoogleScriptAvailable() {
      return typeof google !== 'undefined' && google.script && typeof google.script.run !== 'undefined';
    }

    function renderKpiHistory(records) {
      const tbody = document.getElementById('kpiHistoryBody');
      if (!tbody) return;
      if (!records || records.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;color:var(--gray);">Belum ada catatan KPI.</td></tr>';
        return;
      }
      tbody.innerHTML = records.slice().reverse().map(record => `
        <tr>
          <td>${record.tanggal || '-'}</td>
          <td>${record.divisi || '-'}</td>
          <td>${record.totalPoints}</td>
          <td>${record.grade}</td>
          <td>${record.submittedAt || '-'}</td>
        </tr>
      `).join('');
    }

    function loadKpiHistory() {
      const localHistory = JSON.parse(localStorage.getItem(getKpiStorageKey()) || '[]');
      if (!isGoogleScriptAvailable() || !currentUser || !currentUser.username) {
        return renderKpiHistory(localHistory);
      }

      google.script.run.withSuccessHandler(res => {
        if (res && res.success && Array.isArray(res.data)) {
          localStorage.setItem(getKpiStorageKey(), JSON.stringify(res.data));
          renderKpiHistory(res.data);
        } else {
          renderKpiHistory(localHistory);
        }
      }).withFailureHandler(err => {
        console.error('KPI load error:', err);
        renderKpiHistory(localHistory);
      }).getKpiKaryawan(currentUser.username);
    }

    function submitKpiForm() {
      const division = document.getElementById('kpiDivisi')?.value || '';
      const tanggal = document.getElementById('kpiTanggal')?.value || new Date().toISOString().slice(0, 10);
      const container = document.getElementById('kpiQuestionsContainer');
      if (!container) return toast('Form KPI tidak ditemukan.', 'error');
      const cards = container.querySelectorAll('.kpi-question-card');
      const answers = [];
      let valid = true;
      cards.forEach((card, idx) => {
        const checked = card.querySelector(`input[name="kpiQ${idx}"]:checked`);
        answers.push({
          question: card.querySelector('.kpi-question-title')?.textContent || '',
          answer: checked?.dataset.answer || '',
          point: checked ? parseInt(checked.value, 10) || 0 : 0
        });
        if (!checked) valid = false;
      });
      if (!valid) return toast('Jawab semua pertanyaan sebelum menyimpan.', 'error');
      const total = parseInt(document.getElementById('kpiTotalPoint')?.value, 10) || 0;
      const record = {
        tanggal,
        divisi: division,
        totalPoints: total,
        grade: getKpiGrade(total),
        submittedAt: new Date().toLocaleString(),
        answers
      };

      const saveLocally = () => {
        const history = JSON.parse(localStorage.getItem(getKpiStorageKey()) || '[]');
        history.push(record);
        localStorage.setItem(getKpiStorageKey(), JSON.stringify(history));
        loadKpiHistory();
        toast('KPI Karyawan berhasil disimpan.');
      };

      if (!isGoogleScriptAvailable() || !currentUser || !currentUser.username) {
        return saveLocally();
      }

      google.script.run.withSuccessHandler(res => {
        if (res && res.success) {
          saveLocally();
        } else {
          toast(res?.message ? 'Gagal menyimpan KPI: ' + res.message : 'Gagal menyimpan KPI.', 'error');
        }
      }).withFailureHandler(err => {
        console.error('KPI submit error:', err);
        saveLocally();
        toast('Koneksi backend gagal. KPI disimpan secara lokal.', 'warning');
      }).addKpiKaryawan(currentUser.username, currentUser.nama || currentUser.username, division, tanggal, total, getKpiGrade(total), answers);
    }

    function loadKpiKaryawan() {
      const nameEl = document.getElementById('kpiUserName');
      const divisiEl = document.getElementById('kpiDivisi');
      const tanggalEl = document.getElementById('kpiTanggal');
      if (nameEl) nameEl.value = currentUser?.nama || currentUser?.username || '';
      const defaultDivisi = (currentUser?.divisi || currentUser?.role || 'Warehouse').trim();
      const divisions = ['Warehouse', 'Inbound', 'Distributor', 'HR', 'Operasional', 'Admin'];
      const finalDivisions = [...new Set([defaultDivisi, ...divisions].filter(Boolean))];
      if (divisiEl) {
        divisiEl.innerHTML = finalDivisions.map(d => `<option value="${d}">${d}</option>`).join('');
        if (defaultDivisi) divisiEl.value = finalDivisions.includes(defaultDivisi) ? defaultDivisi : finalDivisions[0];
      }
      if (tanggalEl) tanggalEl.value = new Date().toISOString().slice(0, 10);
      loadKpiQuestions(() => {
        renderKpiQuestions();
        loadKpiHistory();
      });
    }

    // Track last page to prevent redundant loads
    let _lastLoadedPage = null;
    let _pageLoadTimestamps = {};

    function showPage(name) {
      // Hard reset scroll with delay to ensure rendering is complete
      setTimeout(function () {
        window.scrollTo(0, 0);
        if (document.documentElement) document.documentElement.scrollTop = 0;
        if (typeof google !== 'undefined' && google.script && google.script.host) {
          google.script.host.scrollTo(0, 0);
        }
      }, 50);

      if (name === 'users') {
        const role = currentUser.role || '';
        const isHR = (role === 'HR');
        const isViceSPV = (role === 'Vice SPV' || role === 'Vice Supervisor');
        let allowedMenus = []; if (currentUser.permissions) { try { allowedMenus = JSON.parse(currentUser.permissions); } catch (e) { } }
        if (role !== 'admin' && !isHR && !isViceSPV && !allowedMenus.includes('kelolaUser')) return toast('Akses dilarang', 'error');
      }
      if (name === 'kpiKaryawan' && !hasPermission('kpiKaryawan')) {
        return toast('Akses KPI Karyawan dilarang', 'error');
      }
      if (name === 'paymentGudang' && !hasPermission('paymentGudang')) {
        return toast('Akses MISTINE dilarang', 'error');
      }
      if (name === 'rekapOngkirMistine' && !hasPermission('rekapOngkirMistine')) {
        return toast('Akses Rekap Ongkir dilarang', 'error');
      }
      if (name === 'antrianDistributor' && !canViewDistributorQueue()) {
        return toast('Akses Antrian Distributor dilarang', 'error');
      }
      document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
      document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
      const pg = document.getElementById('page-' + name);
      if (pg) pg.classList.add('active');
      document.querySelectorAll('.nav-item').forEach(n => { if (n.getAttribute('onclick') && n.getAttribute('onclick').includes("'" + name + "'")) n.classList.add('active'); });

      // Auto-hide sidebar on click (All devices)
      const sb = document.getElementById('sidebar');
      if (sb) sb.classList.remove('show');

      const titleEl = document.getElementById('pageTitle');
      if (titleEl) {
        titleEl.setAttribute('data-page-key', name);
        // Page icon mapping
        const pageIcons = {
          dashboard: '📊', kasGudang: '💰', teamBuilding: '🤝', expense: '💸', pettyCash: '🪙', paymentGudang: '💄', rekapOngkirMistine: '💰', karyawan: '👥',
          laporanKerja: '📝', grafikLaporan: '📈', handover: '📦', klaim: '⚠️', tugasProject: '📋', kpiKaryawan: '📝',
          ijin: '✉️', lembur: '⏱️', pengajuanAsset: '📦', organisasi: '🏗️', sop: '📋',
          users: '⚙️', stock: '📦', stockOpname: '⚖️', packingList: '📋', inbound: '📥',
          outbound: '✂️', retur: '↩️', order: '🛒', antrianDistributor: '🚚', analisis: '📈', bookingMobil: '🚗'
        };
        const icon = pageIcons[name] || '📄';
        const titleText = name === 'antrianDistributor' ? 'Antrian Distributor' : getLangText(name);
        titleEl.textContent = icon + ' ' + titleText;
      }

      if (pageLoaders[name]) pageLoaders[name]();

      // Auto-hide sidebar after selection
      const sidebar = document.getElementById('sidebar');
      if (window.innerWidth >= 768) {
        sidebar.classList.add('collapsed-desktop');
      } else {
        if (typeof bootstrap !== 'undefined') {
          const bsOffcanvas = bootstrap.Offcanvas.getInstance(sidebar);
          if (bsOffcanvas) bsOffcanvas.hide();
        } else {
          sidebar.classList.remove('show');
        }
      }
    }

    function toggleSidebar() {
      const sb = document.getElementById('sidebar');
      if (sb) sb.classList.toggle('show');
    }

    // === APPROVAL ===
    function canApprove(status, role, category) {
      console.log('canApprove called:', { status, role, category });
      if (role === 'admin' || role === 'Super Admin') {
        console.log('  -> Admin role, returning true');
        return true;
      }
      let userPerms = []; try { userPerms = JSON.parse(currentUser.permissions || '[]'); } catch (e) { }

      if (category === 'Asset Opname') return userPerms.includes('aksesApprovalAssetOpname');
      if (category === 'Stock Control') return userPerms.includes('aksesApprovalStockControl');

      if (userPerms.includes('aksesApproval')) {
        console.log('  -> Has aksesApproval permission, returning true');
        return true;
      }

      const isTL = (role === 'Team Leader' || role === 'TL' || role.includes('Team Leader'));
      const isVice = (role === 'Vice Supervisor' || role === 'Vice SPV' || role === 'Vice VPV' || role.includes('Vice'));
      const isSPV = (role === 'Supervisor' || role === 'SPV' || role === 'Supervisor HR' || (role.includes('Supervisor') && !role.includes('Vice')));
      const isHR = (role === 'HR' || role === 'Supervisor HR' || role.includes('HR'));

      console.log('  -> Role checks:', { isTL, isVice, isSPV, isHR });

      // Safety and final status check (Approved, Disetujui, Tolak, Rejected are final)
      if (!status) {
        console.log('  -> No status, returning false');
        return false;
      }
      const finalStatuses = ['Disetujui', 'Tolak', 'Disetujui Admin', 'Approved', 'Rejected', 'Approved Admin', 'Ditolak'];
      if (finalStatuses.includes(status)) {
        console.log('  -> Final status, returning false');
        return false;
      }

      // Enforce hierarchical approval: TL -> Vice -> SPV -> HR
      if (status === 'Pending Team Leader') {
        const result = isTL;
        console.log('  -> Pending Team Leader, isTL=' + result);
        return result;
      }
      if (status === 'Pending Vice Supervisor') {
        const result = isVice;
        console.log('  -> Pending Vice Supervisor, isVice=' + result);
        return result;
      }
      if (status === 'Pending Supervisor') {
        const result = isSPV;
        console.log('  -> Pending Supervisor, isSPV=' + result);
        return result;
      }
      if (status === 'Pending HR') {
        const result = isHR;
        console.log('  -> Pending HR, isHR=' + result);
        return result;
      }

      // Fallback untuk status 'Pending' (Record lama atau Default)
      if (status === 'Pending') {
        const result = isTL;
        console.log('  -> Pending (legacy), isTL=' + result);
        return result;
      }

      console.log('  -> Unknown status, returning false');
      return false;
    }
    // Helper: render dashApprovalList dari cache (tidak fetch server)
    function _renderDashApprovalFromCache() {
      if (!_approvalCache) return;
      const res = _approvalCache;
      const role = currentUser ? (currentUser.role || '') : '';
      const FINAL = ['Disetujui', 'Tolak', 'Disetujui Admin', 'Approved', 'Rejected', 'Approved Admin', 'Ditolak'];
      const PAGE_MAP = {
        'Ijin/Cuti': 'ijin', 'Lembur': 'lembur', 'Asset': 'pengajuanAsset',
        'Asset Opname': 'assetWarehouse', 'Stock Control': 'handover', 'Stock Opname': 'stockOpname'
      };

      let total = 0;
      const rows = [];

      const collect = (list, type) => {
        if (!list || !Array.isArray(list)) return;
        list.forEach(d => {
          if (!FINAL.includes(d.status) && canApprove(d.status, role, type)) {
            total++;
            const page = PAGE_MAP[type] || 'approval';
            const staff = d.nama || d.karyawan || d.sku || '-';
            rows.push(`<tr>
              <td>${formatDate(d.tanggal)}</td>
              <td><span class="badge-tb">${type}</span></td>
              <td><strong>${staff}</strong></td>
              <td><span class="badge-pending">⏳ ${d.status}</span></td>
              <td><button class="btn btn-ghost btn-sm" onclick="showPage('${page}')">Lihat</button></td>
            </tr>`);
          }
        });
      };

      collect(res.ijin, 'Ijin/Cuti');
      collect(res.lembur, 'Lembur');
      collect(res.asset, 'Asset');
      collect(res.stockOpname, 'Stock Opname');
      collect(res.assetOpname, 'Asset Opname');
      collect(res.stockControl, 'Stock Control');

      pendingApprovalCount = total;
      const dashList = document.getElementById('dashApprovalList');
      const dashWrap = document.getElementById('dashApprovalWrap');
      const centralBadge = document.getElementById('badgeTotalApproval');

      if (total > 0) {
        if (dashList) dashList.innerHTML = rows.join('');
        if (dashWrap) dashWrap.style.display = 'block';
        if (centralBadge) { centralBadge.textContent = total; centralBadge.style.display = 'inline-block'; }
      } else {
        if (dashList) dashList.innerHTML = '<tr><td colspan="5" style="text-align:center; color:var(--gray);">Tidak ada antrean approval.</td></tr>';
        if (dashWrap) dashWrap.style.display = 'none';
        if (centralBadge) centralBadge.style.display = 'none';
      }
      refreshGlobalBadge();
    }

    function checkPendingApprovals() {
      if (!currentUser) return;

      // Jika cache masih segar (< 30 detik), render langsung dari cache — instan
      const age = Date.now() - _approvalCacheTime;
      if (_approvalCache && age < 30000) {
        _renderDashApprovalFromCache();
        return;
      }

      google.script.run.withSuccessHandler(res => {
        if (!res) return;

        // Simpan ke cache global
        _approvalCache = res;
        _approvalCacheTime = Date.now();

        // Render dashboard approval list dari cache (pakai helper terpusat)
        _renderDashApprovalFromCache();

        // Jika halaman approval sedang terbuka, langsung render ulang dari cache
        if (document.getElementById('page-approval') && document.getElementById('page-approval').classList.contains('active')) {
          _renderApprovalFromCache();
        }
      }).getPendingApprovals();
    }
    function processApproval(tipe, id, action, nama, tanggal) {
      let reason = '';
      if (action === 'Reject') {
        reason = prompt('Alasan Penolakan:');
        if (reason === null) return;
      }

      const btn = event ? event.target : null;
      const oldTxt = btn ? btn.textContent : '';
      if (btn && btn.tagName === 'BUTTON') {
        btn.classList.add('btn-appr-animate');
        btn.disabled = true;
        btn.innerHTML = '⏳';
      }

      toast('⏳ Memproses approval...', 'success');

      google.script.run.withSuccessHandler(res => {
        if (btn && btn.tagName === 'BUTTON') {
          btn.classList.remove('btn-appr-animate');
          if (res.success) {
            if (action === 'Approve') {
              btn.classList.add('btn-appr-success');
            } else {
              btn.classList.add('btn-appr-reject');
            }
          }
        }
        if (res.success) {
          toast('Approval diproses ✅');
          const t = tipe.toLowerCase();
          if (t === 'ijin' || t === 'lembur' || t === 'asset' || t.includes('opname') || t === 'stockcontrol') {
            if (t === 'ijin') loadIjin();
            else if (t === 'lembur') loadLembur();
            else if (t === 'asset') loadAsset();
            else if (t === 'assetopname') loadAssetWarehouse();
            else if (t.includes('opname')) loadStockOpname();
            else if (t === 'stockcontrol') { if (typeof loadStockControl === 'function') loadStockControl(); }
          }

          if (document.getElementById('page-approval').classList.contains('active')) {
            // Invalidate cache agar fetch ulang data terbaru
            _approvalCache = null; _approvalCacheTime = 0;
            loadCentralizedApprovals();
          }
          checkPendingApprovals();
        } else {
          if (btn && btn.tagName === 'BUTTON') {
            btn.disabled = false;
            btn.textContent = oldTxt;
          }
          toast(res.message, 'error');
        }
      }).processApprovalStatus(tipe.toLowerCase(), id, action, currentUser.nama, currentUser.role, reason, nama, tanggal);
    }

    // ── Cache approval global ──────────────────────────────────────────────
    let _approvalCache = null;  // hasil getPendingApprovals()
    let _approvalCacheTime = 0;     // timestamp cache terakhir
    let _orgMapCache = null;  // orgMap untuk bulk enrichment
    const APPROVAL_CACHE_TTL = 3 * 60 * 1000; // 3 menit, sinkron dengan interval checkPending
    // ──────────────────────────────────────────────────────────────────────

    /**
     * Render halaman approval dari cache (instan, tanpa server call).
     * Dipanggil saat showPage('approval') atau saat cache diperbarui.
     */
    function _renderApprovalFromCache() {
      const tb = document.getElementById('tableCentralApproval');
      if (!tb) return;
      if (!_approvalCache) return; // belum ada cache, biarkan loadCentralizedApprovals tangani

      renderCentralizedApprovals(_approvalCache);

      // Bulk panel
      const role = currentUser.role || '';
      let userPerms = []; try { userPerms = JSON.parse(currentUser.permissions || '[]'); } catch (e) { }
      const isAdmin = (role === 'admin' || role === 'Super Admin');
      const isTL = (role === 'Team Leader' || role === 'TL' || role.includes('Team Leader'));
      const isVice = role.includes('Vice');
      const isSPV = (role === 'Supervisor' || role === 'SPV' || (role.includes('Supervisor') && !role.includes('Vice')));
      const isHR = (role === 'HR' || role === 'Supervisor HR' || role.includes('HR'));
      const hasCustom = userPerms.includes('aksesApproval');
      const canUseBulk = isAdmin || isTL || isVice || isSPV || isHR || hasCustom;

      const panel = document.getElementById('bulkApprovalPanel');
      if (canUseBulk) {
        panel.style.display = 'block';
        renderBulkActionButtons(role, userPerms);
        _buildBulkFromCache();
      } else {
        panel.style.display = 'none';
      }
    }

    /**
     * Bangun rawBulkData dari _approvalCache (tanpa ke server).
     * Enrichment _divisi dilakukan di sini dengan orgMapCache.
     */
    function _buildBulkFromCache() {
      if (!_approvalCache) return;
      const res = _approvalCache;

      const enrichList = (list, type, module) => (list || []).map(item => ({
        ...item,
        _type: type,
        _module: module,
        _divisi: item.divisi || (_orgMapCache && _orgMapCache[(item.nama || item.karyawan || '').trim()]) || 'Lainnya'
      }));

      rawBulkData = [
        ...enrichList(res.ijin, 'Ijin/Cuti', 'ijin'),
        ...enrichList(res.lembur, 'Lembur', 'lembur'),
        ...enrichList(res.asset, 'Asset', 'asset'),
        ...enrichList(res.stockOpname, 'Stock Opname', 'opname')
      ];
      filterBulkData();
    }

    function loadCentralizedApprovals() {
      const tb = document.getElementById('tableCentralApproval');
      if (!tb) return;

      // ── Tampilkan dari cache dulu (instan) ─────────────────────────────
      if (_approvalCache) {
        _renderApprovalFromCache();
      } else {
        // Tampilkan skeleton loading yang ringan saat pertama kali
        tb.innerHTML = `
          <tr>
            <td colspan="6" style="text-align: center; padding: 20px; color: var(--text-muted);">
              <div style="font-size: 14px; font-weight: 600;">⚡ Memuat data...</div>
            </td>
          </tr>`;
      }

      // ── Jika cache masih segar (< TTL), tidak perlu fetch lagi ────────
      const age = Date.now() - _approvalCacheTime;
      if (_approvalCache && age < APPROVAL_CACHE_TTL) return;

      // ── Fetch terbaru di background, update cache, lalu re-render ─────
      google.script.run.withSuccessHandler(res => {
        if (!res) return;
        _approvalCache = res;
        _approvalCacheTime = Date.now();
        _renderApprovalFromCache();
      }).getPendingApprovals();
    }

    /**
     * Optimized loader untuk approval - mencegah multiple fetch
     */
    function loadCentralizedApprovalsOptimized() {
      const now = Date.now();
      const lastLoad = _pageLoadTimestamps['approval'] || 0;
      const timeSinceLastLoad = now - lastLoad;

      // Jika baru saja di-load (< 2 detik), skip fetch
      if (timeSinceLastLoad < 2000) {
        console.log('Approval: Skipping fetch (too soon)', timeSinceLastLoad + 'ms');
        // Hanya render dari cache
        if (_approvalCache) {
          _renderApprovalFromCache();
        }
        return;
      }

      // Update timestamp
      _pageLoadTimestamps['approval'] = now;

      // Load normal
      loadCentralizedApprovals();
    }

    /**
     * Filter tabel approval berdasarkan pencarian dan status
     */
    function filterApprovalTable() {
      const searchTerm = (document.getElementById('searchApproval')?.value || '').toLowerCase();
      const statusFilter = document.getElementById('filterApprovalStatus')?.value || '';
      const rows = document.querySelectorAll('#tableCentralApproval tr');

      let visibleCount = 0;
      rows.forEach(row => {
        if (row.querySelector('.empty-state') || row.cells.length < 6) return;

        const nama = (row.cells[2]?.textContent || '').toLowerCase();
        const kategori = (row.cells[1]?.textContent || '').toLowerCase();
        const status = row.cells[4]?.textContent || '';

        const matchSearch = !searchTerm || nama.includes(searchTerm) || kategori.includes(searchTerm);
        const matchStatus = !statusFilter || status.includes(statusFilter);

        if (matchSearch && matchStatus) {
          row.style.display = '';
          visibleCount++;
        } else {
          row.style.display = 'none';
        }
      });
    }

    /**
     * Render tombol aksi di panel bulk sesuai jabatan user.
     * Admin mendapat semua tombol. Role lain hanya mendapat tombol stage miliknya.
     */
    function renderBulkActionButtons(role, userPerms) {
      const wrap = document.getElementById('bulkActionButtons');
      const badge = document.getElementById('bulkRoleBadge');
      if (!wrap) return;

      const isAdmin = (role === 'admin' || role === 'Super Admin');
      const isTL = (role === 'Team Leader' || role === 'TL' || role.includes('Team Leader'));
      const isVice = role.includes('Vice');
      const isSPV = (role === 'Supervisor' || role === 'SPV' || (role.includes('Supervisor') && !role.includes('Vice')));
      const isHR = (role === 'HR' || role === 'Supervisor HR' || role.includes('HR'));
      const hasCustom = (userPerms || []).includes('aksesApproval');

      // Konfigurasi tombol: [stage_status, label, warna, icon]
      const stages = [
        { status: 'Pending Team Leader', label: 'Setujui sebagai Team Leader', color: '#3b82f6', icon: '👤', show: isAdmin || isTL || hasCustom },
        { status: 'Pending Vice Supervisor', label: 'Setujui sebagai Vice SPV', color: '#8b5cf6', icon: '👥', show: isAdmin || isVice || hasCustom },
        { status: 'Pending Supervisor', label: 'Setujui sebagai Supervisor', color: '#f59e0b', icon: '🎖️', show: isAdmin || isSPV || hasCustom },
        { status: 'Pending HR', label: 'Setujui sebagai HR', color: '#10b981', icon: '🏢', show: isAdmin || isHR || hasCustom },
      ];

      const activeStages = stages.filter(s => s.show);

      // Badge jabatan
      let roleName = role;
      if (isAdmin) roleName = 'Admin (semua tahap)';
      else if (isTL) roleName = 'Team Leader';
      else if (isVice) roleName = 'Vice Supervisor';
      else if (isSPV) roleName = 'Supervisor';
      else if (isHR) roleName = 'HR';
      if (badge) badge.textContent = '🔑 Jabatan Anda: ' + roleName;

      // Render tombol
      let html = '<span style="font-size:12px; color:var(--text-muted); font-weight:700; flex-shrink:0;">AKSI:</span>';

      activeStages.forEach(s => {
        html += `
          <button class="btn btn-bulk-stage" id="btnBulk_${s.status.replace(/ /g, '_')}"
            onclick="processBulkApprovalByStage('${s.status}')"
            disabled
            style="background:${s.color}22; color:${s.color}; border:1px solid ${s.color}44;
                   padding:8px 18px; border-radius:8px; font-size:13px; font-weight:700;
                   cursor:pointer; transition:all 0.2s; display:flex; align-items:center; gap:7px; opacity:0.5;">
            ${s.icon} ${s.label} (<span class="bulk-stage-count" data-stage="${s.status}">0</span>)
          </button>`;
      });

      // Tombol tolak (hanya tampil saat ada item terpilih, awalnya hidden)
      html += `
        <button class="btn" id="btnBulkRejectAll"
          onclick="processBulkApprovalByStage('', 'Reject')"
          disabled
          style="background:#ef444422; color:var(--red); border:1px solid #ef444444;
                 padding:8px 18px; border-radius:8px; font-size:13px; font-weight:700;
                 cursor:pointer; display:flex; align-items:center; gap:7px; opacity:0.5;">
          ❌ Tolak Terpilih (<span id="bulkRejectCount">0</span>)
        </button>`;

      wrap.innerHTML = html;
    }

    let rawBulkData = []; // Cache for filtering

    function loadBulkApprovalData() {
      // Jika cache approval sudah ada, build langsung dari sana (instan)
      if (_approvalCache) {
        // Pastikan orgMap sudah ada; jika belum fetch sekali saja
        if (_orgMapCache) {
          _buildBulkFromCache();
        } else {
          // Fetch orgMap sekali, simpan ke cache, lalu build bulk
          google.script.run.withSuccessHandler(orgRes => {
            _orgMapCache = {};
            if (orgRes && orgRes.success && orgRes.data) {
              orgRes.data.forEach(o => { if (o.nama) _orgMapCache[o.nama.trim()] = o.departemen || 'Lainnya'; });
            }
            _buildBulkFromCache();
          }).getOrganisasi();
        }
        return;
      }

      // Fallback: tidak ada cache sama sekali → fetch dari server
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          rawBulkData = res.data || [];
          filterBulkData();
        } else {
          const tb = document.getElementById('tableBulkApproval');
          if (tb) tb.innerHTML = '<tr><td colspan="6" class="text-center text-muted">Gagal memuat data</td></tr>';
        }
      }).getBulkApprovalData();
    }

    function renderBulkTable(data) {
      const tb = document.getElementById('tableBulkApproval');
      tb.innerHTML = '';

      const visibleCount = document.getElementById('bulkVisibleCount');
      if (visibleCount) visibleCount.textContent = data ? data.length : 0;

      if (!data || data.length === 0) {
        tb.innerHTML = '<tr><td colspan="6" class="text-center text-muted p-4">Tidak ada pengajuan yang perlu disetujui</td></tr>';
        updateBulkCount();
        return;
      }

      const role = currentUser.role || '';
      let userPerms = []; try { userPerms = JSON.parse(currentUser.permissions || '[]'); } catch (e) { }
      const isAdmin = (role === 'admin' || role === 'Super Admin');
      const isTL = (role === 'Team Leader' || role === 'TL' || role.includes('Team Leader'));
      const isVice = role.includes('Vice');
      const isSPV = (role === 'Supervisor' || role === 'SPV' || (role.includes('Supervisor') && !role.includes('Vice')));
      const isHR = (role === 'HR' || role === 'Supervisor HR' || role.includes('HR'));
      const hasCustom = userPerms.includes('aksesApproval');

      // Status badge colors
      const stageColors = {
        'Pending Team Leader': { bg: 'rgba(59,130,246,0.15)', color: '#3b82f6' },
        'Pending Vice Supervisor': { bg: 'rgba(139,92,246,0.15)', color: '#8b5cf6' },
        'Pending Supervisor': { bg: 'rgba(245,158,11,0.15)', color: '#f59e0b' },
        'Pending HR': { bg: 'rgba(16,185,129,0.15)', color: '#10b981' },
        'Pending': { bg: 'rgba(59,130,246,0.15)', color: '#3b82f6' },
      };

      data.forEach(d => {
        const tgl = formatDate(d.tanggal || d.createdAt);
        const rawStatus = d.status || 'Pending';
        // Legacy 'Pending' → tampilkan sebagai Pending Team Leader
        const displayStatus = rawStatus === 'Pending' ? 'Pending Team Leader' : rawStatus;
        const sc = stageColors[rawStatus] || { bg: 'rgba(148,163,184,0.15)', color: '#94a3b8' };

        // Apakah user bisa approve item ini?
        const canAct = isAdmin || hasCustom
          || (isTL && (rawStatus === 'Pending Team Leader' || rawStatus === 'Pending'))
          || (isVice && rawStatus === 'Pending Vice Supervisor')
          || (isSPV && rawStatus === 'Pending Supervisor')
          || (isHR && rawStatus === 'Pending HR');

        // Row samar kalau tidak bisa di-act
        const rowStyle = canAct ? '' : 'opacity:0.45;';

        tb.innerHTML += `
          <tr style="${rowStyle}">
            <td style="text-align:center;">
              ${canAct
            ? `<input type="checkbox" class="bulk-item-check"
                      data-id="${d.id}" data-type="${d._module}"
                      data-nama="${(d.nama || d.karyawan || '').replace(/'/g, "&#39;")}"
                      data-tanggal="${d.tanggal || d.createdAt}"
                      data-stage="${rawStatus}"
                      onchange="updateBulkCount()">`
            : `<span title="Bukan tahap Anda" style="color:var(--text-muted);">🔒</span>`}
            </td>
            <td style="white-space:nowrap;">${tgl}</td>
            <td><span class="badge-pending-stage">${d._type}</span></td>
            <td><strong>${d.nama || d.karyawan || '-'}</strong></td>
            <td><span style="font-size:11px; background:rgba(255,255,255,0.05); padding:2px 8px; border-radius:6px;">${d._divisi || 'Lainnya'}</span></td>
            <td>
              <span style="font-size:11px; font-weight:700; padding:3px 10px; border-radius:20px;
                           background:${sc.bg}; color:${sc.color}; border:1px solid ${sc.color}44;">
                ${displayStatus}
              </span>
            </td>
          </tr>`;
      });

      updateBulkCount();
    }

    function filterBulkData() {
      const fDate = document.getElementById('bulkFilterDate').value;
      const fDiv = document.getElementById('bulkFilterDivision').value;
      const fKat = document.getElementById('bulkFilterKategori') ? document.getElementById('bulkFilterKategori').value : '';
      const fSearch = document.getElementById('bulkFilterSearch') ? document.getElementById('bulkFilterSearch').value.toLowerCase() : '';

      const filtered = rawBulkData.filter(d => {
        const itemDate = (d.tanggal || d.createdAt || '').split('T')[0];
        const nameLower = (d.nama || d.karyawan || '').toLowerCase();
        const matchDate = !fDate || itemDate === fDate;
        const matchDiv = !fDiv || d._divisi === fDiv;
        const matchKat = !fKat || d._module === fKat;
        const matchSearch = !fSearch || nameLower.includes(fSearch);
        return matchDate && matchDiv && matchKat && matchSearch;
      });

      renderBulkTable(filtered);
    }

    function toggleAllBulk(master) {
      const checks = document.querySelectorAll('.bulk-item-check');
      checks.forEach(c => c.checked = master.checked);
      updateBulkCount();
    }

    function updateBulkCount() {
      const checked = document.querySelectorAll('.bulk-item-check:checked');
      const total = checked.length;

      // Update info bar
      const countEl = document.getElementById('bulkCheckCount');
      if (countEl) countEl.textContent = total;

      // Hitung per stage
      const stageCounts = {};
      checked.forEach(c => {
        const stage = c.getAttribute('data-stage') || 'Pending';
        const key = stage === 'Pending' ? 'Pending Team Leader' : stage;
        stageCounts[key] = (stageCounts[key] || 0) + 1;
      });

      // Update label tiap tombol stage
      document.querySelectorAll('.bulk-stage-count').forEach(el => {
        const stage = el.getAttribute('data-stage');
        const cnt = stageCounts[stage] || 0;
        el.textContent = cnt;
        const btn = el.closest('button');
        if (btn) {
          const hasCnt = cnt > 0;
          btn.disabled = !hasCnt;
          btn.style.opacity = hasCnt ? '1' : '0.5';
          btn.style.cursor = hasCnt ? 'pointer' : 'not-allowed';
        }
      });

      // Update tombol tolak
      const btnReject = document.getElementById('btnBulkRejectAll');
      const rejectCnt = document.getElementById('bulkRejectCount');
      if (rejectCnt) rejectCnt.textContent = total;
      if (btnReject) {
        btnReject.disabled = total === 0;
        btnReject.style.opacity = total > 0 ? '1' : '0.5';
        btnReject.style.cursor = total > 0 ? 'pointer' : 'not-allowed';
      }

      // Sinkron master checkbox
      const allChecks = document.querySelectorAll('.bulk-item-check');
      const masterCb = document.getElementById('bulkCheckAll');
      if (masterCb && allChecks.length > 0) {
        masterCb.checked = total === allChecks.length;
        masterCb.indeterminate = total > 0 && total < allChecks.length;
      }
    }

    /**
     * Proses bulk approval/rejection.
     * @param {string} targetStage - Status stage yang diproses (kosong = semua yang terpilih)
     * @param {string} action      - 'Approve' (default) atau 'Reject'
     */
    function processBulkApprovalByStage(targetStage, action) {
      action = action || 'Approve';

      let reason = '';
      if (action === 'Reject') {
        reason = prompt('Alasan penolakan massal:');
        if (reason === null) return;
      }

      // Ambil semua checkbox yang dicentang
      const checked = document.querySelectorAll('.bulk-item-check:checked');
      if (checked.length === 0) { toast('Pilih minimal 1 item terlebih dahulu', 'error'); return; }

      // Filter sesuai stage jika targetStage diisi
      let items = Array.from(checked).map(c => ({
        id: c.getAttribute('data-id'),
        tipe: c.getAttribute('data-type'),
        nama: c.getAttribute('data-nama'),
        tanggal: c.getAttribute('data-tanggal'),
        stage: c.getAttribute('data-stage'),
        action: action,
        reason: reason
      }));

      if (targetStage) {
        items = items.filter(i => i.stage === targetStage || (targetStage === 'Pending Team Leader' && i.stage === 'Pending'));
      }

      if (items.length === 0) {
        toast('Tidak ada item pada tahap ' + targetStage + ' yang terpilih', 'error');
        return;
      }

      // Konfirmasi
      const stageLabel = targetStage || 'semua tahap';
      const actionLabel = action === 'Approve' ? 'menyetujui' : 'menolak';
      const confirmMsg = `${actionLabel.charAt(0).toUpperCase() + actionLabel.slice(1)} ${items.length} pengajuan (${stageLabel})?`;
      if (!confirm(confirmMsg)) return;

      // Disable semua tombol selama proses
      document.querySelectorAll('#bulkActionButtons button').forEach(b => { b.disabled = true; b.style.opacity = '0.6'; });
      showLoading(`${actionLabel.charAt(0).toUpperCase() + actionLabel.slice(1)} ${items.length} pengajuan...`);

      google.script.run
        .withSuccessHandler(res => {
          hideLoading();
          document.querySelectorAll('#bulkActionButtons button').forEach(b => { b.disabled = false; });

          if (res.success) {
            const msg = `⚡ Berhasil ${actionLabel} ${res.processed} item.`;
            if (res.failed > 0) {
              alert(`${msg}\n⚠️ ${res.failed} item gagal:\n${(res.errors || []).join('\n')}`);
            } else {
              if (typeof toast === 'function') toast(msg, 'success');
              else alert(msg);
            }
            // Invalidate cache agar fetch ulang data terbaru
            _approvalCache = null; _approvalCacheTime = 0;
            checkPendingApprovals();
            loadCentralizedApprovals();
          } else {
            alert('Gagal memproses approval massal: ' + res.message);
            updateBulkCount();
          }
        })
        .withFailureHandler(err => {
          hideLoading();
          document.querySelectorAll('#bulkActionButtons button').forEach(b => { b.disabled = false; });
          alert('Kesalahan Sistem: ' + err.message);
          updateBulkCount();
        })
        .processBatchApproval(items, currentUser.nama, currentUser.role);
    }

    // Legacy wrapper agar tidak error jika ada panggilan lama
    function processBulkApproval() { processBulkApprovalByStage(''); }

    function isAuthorizedToAct(status, role, category) {
      if (role === 'admin' || role === 'Super Admin') return true;
      let userPerms = []; try { userPerms = JSON.parse(currentUser.permissions || '[]'); } catch (e) { }

      if (category === 'Asset Opname') return userPerms.includes('aksesApprovalAssetOpname');
      if (category === 'Stock Control') return userPerms.includes('aksesApprovalStockControl');

      if (userPerms.includes('aksesApproval')) return true;

      const isTL = (role === 'Team Leader' || role === 'TL' || role.includes('Team Leader'));
      const isVice = (role === 'Vice Supervisor' || role === 'Vice SPV' || role === 'Vice VPV' || role.includes('Vice'));
      const isSPV = (role === 'Supervisor' || role === 'SPV' || role === 'Supervisor HR' || (role.includes('Supervisor') && !role.includes('Vice')));
      const isHR = (role === 'HR' || role === 'Supervisor HR' || role.includes('HR'));

      if (category === 'Stock Opname' && status === 'Pending') return (isSPV || role === 'admin' || role === 'Super Admin');

      // Alur Berjenjang (TL -> Vice -> SPV -> HR)
      if (status === 'Pending Team Leader') return isTL;
      if (status === 'Pending Vice Supervisor') return isVice;
      if (status === 'Pending Supervisor') return isSPV;
      if (status === 'Pending HR') return isHR;

      // Fallback untuk status 'Pending' (Record lama atau Default)
      if (status === 'Pending') return isTL || isAdmin;

      return false;
    }

    function renderCentralizedApprovals(res) {
      const tb = document.getElementById('tableCentralApproval');
      if (!tb) return;

      const role = currentUser.role || '';
      const items = [];

      const filterItems = (list, type, module) => {
        if (!list || !Array.isArray(list)) return [];
        return list.filter(d => {
          const status = d.status;
          return (status !== 'Disetujui' && status !== 'Tolak' && status !== 'Disetujui Admin' && status !== 'Approved' && status !== 'Rejected' && status !== 'Ditolak') && canApprove(status, role, type);
        }).map(d => ({ ...d, _type: type, _module: module }));
      };

      const ijin = filterItems(res.ijin, 'Ijin/Cuti', 'ijin');
      const lembur = filterItems(res.lembur, 'Lembur', 'lembur');
      const asset = filterItems(res.asset, 'Asset', 'asset');
      const opname = filterItems(res.stockOpname, 'Stock Opname', 'opname');
      const assetOpname = filterItems(res.assetOpname, 'Asset Opname', 'assetOpname');
      const stockControl = filterItems(res.stockControl, 'Stock Control', 'stockcontrol');

      items.push(...ijin, ...lembur, ...asset, ...opname, ...assetOpname, ...stockControl);
      items.sort((a, b) => new Date(b.tanggal || b.createdAt) - new Date(a.tanggal || a.createdAt));

      // Update statistik dengan animasi
      const stats = [
        { id: 'statApprTotal', value: items.length },
        { id: 'statApprIjin', value: ijin.length },
        { id: 'statApprLembur', value: lembur.length },
        { id: 'statApprAsset', value: asset.length },
        { id: 'statApprOpname', value: opname.length + assetOpname.length + stockControl.length }
      ];

      stats.forEach(stat => {
        const el = document.getElementById(stat.id);
        if (el) {
          el.style.transition = 'all 0.3s ease';
          el.textContent = stat.value;
        }
      });

      if (items.length === 0) {
        tb.innerHTML = `
          <tr>
            <td colspan="6" class="empty-state" style="text-align: center; padding: 60px 20px;">
              <div style="font-size: 64px; margin-bottom: 16px; opacity: 0.3;">🎉</div>
              <div style="font-size: 18px; font-weight: 700; color: var(--text-main); margin-bottom: 8px;">
                Tidak Ada Approval Pending
              </div>
              <div style="font-size: 13px; color: var(--text-muted);">
                Semua pengajuan telah diproses atau tidak ada pengajuan baru
              </div>
            </td>
          </tr>`;
        return;
      }

      // Gunakan DocumentFragment untuk performa rendering yang lebih baik
      const fragment = document.createDocumentFragment();
      const tempDiv = document.createElement('div');

      items.forEach(d => {
        const staff = d.nama || d.karyawan || d.sku || '-';
        let detail = d.keterangan || d.alasan || d.detailTugas || d.nama || '-';

        // Highlight info penting untuk lembur
        if (d._module === 'lembur' && d.jumlahJam) {
          detail = `<div style="display: flex; align-items: center; gap: 8px; margin-bottom: 4px;">
                      <span style="background: linear-gradient(135deg, rgba(245, 158, 11, 0.15), rgba(245, 158, 11, 0.08)); 
                                   color: var(--accent); font-weight: 800; padding: 4px 10px; border-radius: 6px; 
                                   font-size: 12px; border: 1px solid rgba(245, 158, 11, 0.3);">
                        ⏰ ${d.jumlahJam} Jam
                      </span>
                    </div>
                    <div style="font-size: 13px; color: var(--text-muted); margin-top: 4px;">${detail}</div>`;
        }

        // Info Jam Absen & Verifikasi untuk Lembur
        let absenHtml = '';
        if (d._module === 'lembur' && d.jamAbsen && d.jamAbsen !== '-') {
          const isLolos = d.statusVerifikasi === 'Lolos';
          const verifyBadge = isLolos
            ? `<span style="font-size: 10px; padding: 3px 8px; border-radius: 6px; font-weight: 700; 
                          background: rgba(16,185,129,0.15); color: var(--green); border: 1px solid rgba(16,185,129,0.3); 
                          display: inline-flex; align-items: center; gap: 4px;">
                 ✓ Lolos Verifikasi
               </span>`
            : `<span style="font-size: 10px; padding: 3px 8px; border-radius: 6px; font-weight: 700; 
                          background: rgba(239,68,68,0.15); color: var(--red); border: 1px solid rgba(239,68,68,0.3);
                          display: inline-flex; align-items: center; gap: 4px;">
                 ✗ Tidak Lolos
               </span>`;
          absenHtml = `<div style="margin-top: 8px; padding: 6px 10px; background: rgba(255,255,255,0.02); 
                                  border-radius: 6px; border: 1px solid var(--border-color); font-size: 11px; 
                                  display: inline-flex; align-items: center; gap: 8px;">
                         <span style="color: var(--text-muted);">🕒 ${d.jamAbsen}</span>
                         ${verifyBadge}
                       </div>`;
        }

        const tgl = formatDate(d.tanggal || d.createdAt);
        const status = d.status === 'Pending' ? 'Pending HR' : d.status;
        const staffNama = d.nama || d.karyawan || '';
        const isoTgl = d.tanggal || d.createdAt;

        // Badge kategori dengan warna berbeda
        const categoryColors = {
          'Ijin/Cuti': { bg: 'rgba(59, 130, 246, 0.12)', color: '#3b82f6', icon: '🏖️' },
          'Lembur': { bg: 'rgba(245, 158, 11, 0.12)', color: '#f59e0b', icon: '⏰' },
          'Asset': { bg: 'rgba(14, 165, 233, 0.12)', color: '#0ea5e9', icon: '🏢' },
          'Stock Opname': { bg: 'rgba(239, 68, 68, 0.12)', color: '#ef4444', icon: '📦' },
          'Asset Opname': { bg: 'rgba(139, 92, 246, 0.12)', color: '#8b5cf6', icon: '📋' },
          'Stock Control': { bg: 'rgba(16, 185, 129, 0.12)', color: '#10b981', icon: '⚖️' }
        };
        const catStyle = categoryColors[d._type] || { bg: 'rgba(148, 163, 184, 0.12)', color: '#94a3b8', icon: '📄' };

        // Status badge dengan styling lebih menarik
        const statusColors = {
          'Pending Team Leader': { bg: 'rgba(59, 130, 246, 0.12)', color: '#3b82f6' },
          'Pending Vice Supervisor': { bg: 'rgba(139, 92, 246, 0.12)', color: '#8b5cf6' },
          'Pending Supervisor': { bg: 'rgba(245, 158, 11, 0.12)', color: '#f59e0b' },
          'Pending HR': { bg: 'rgba(16, 185, 129, 0.12)', color: '#10b981' },
          'Pending': { bg: 'rgba(148, 163, 184, 0.12)', color: '#94a3b8' }
        };
        const statusStyle = statusColors[status] || { bg: 'rgba(148, 163, 184, 0.12)', color: '#94a3b8' };

        tempDiv.innerHTML = `
          <tr style="transition: all 0.2s ease; cursor: pointer;" 
              onmouseenter="this.style.background='rgba(245, 158, 11, 0.04)'" 
              onmouseleave="this.style.background=''">
            <td style="white-space: nowrap;">
              <span style="font-size: 13px; color: var(--text-muted); font-weight: 500;">${tgl}</span>
            </td>
            <td>
              <span style="font-size: 11px; font-weight: 700; padding: 5px 12px; border-radius: 8px; 
                          background: ${catStyle.bg}; color: ${catStyle.color}; 
                          border: 1px solid ${catStyle.color}44; display: inline-flex; align-items: center; gap: 5px;
                          letter-spacing: 0.3px;">
                ${catStyle.icon} ${d._type}
              </span>
            </td>
            <td>
              <strong style="font-size: 14px; color: var(--text-main);">${staff}</strong>
            </td>
            <td>
              <div style="max-width: 400px; font-size: 13px; line-height: 1.5;">${detail}</div>
              ${absenHtml}
            </td>
            <td>
              <span style="font-size: 11px; font-weight: 700; padding: 5px 12px; border-radius: 20px; 
                          background: ${statusStyle.bg}; color: ${statusStyle.color}; 
                          border: 1px solid ${statusStyle.color}44; white-space: nowrap; letter-spacing: 0.3px;">
                ${status}
              </span>
            </td>
            <td style="text-align: center;">
              ${isAuthorizedToAct(d.status, role, d._type) ? `
              <div style="display: flex; gap: 8px; justify-content: center;">
                <button class="btn btn-sm" 
                  onclick="processApproval('${d._module}', '${d.id}', 'Approve', '${staffNama}', '${isoTgl}')"
                  style="background: linear-gradient(135deg, rgba(16, 185, 129, 0.15), rgba(16, 185, 129, 0.08)); 
                         color: var(--green); border: 1px solid rgba(16, 185, 129, 0.3); font-weight: 700; 
                         padding: 6px 16px; transition: all 0.2s;"
                  onmouseenter="this.style.background='var(--green)'; this.style.color='#fff'"
                  onmouseleave="this.style.background='linear-gradient(135deg, rgba(16, 185, 129, 0.15), rgba(16, 185, 129, 0.08))'; this.style.color='var(--green)'">
                  ✓ Setuju
                </button>
                <button class="btn btn-sm" 
                  onclick="processApproval('${d._module}', '${d.id}', 'Reject', '${staffNama}', '${isoTgl}')"
                  style="background: linear-gradient(135deg, rgba(239, 68, 68, 0.15), rgba(239, 68, 68, 0.08)); 
                         color: var(--red); border: 1px solid rgba(239, 68, 68, 0.3); font-weight: 700; 
                         padding: 6px 16px; transition: all 0.2s;"
                  onmouseenter="this.style.background='var(--red)'; this.style.color='#fff'"
                  onmouseleave="this.style.background='linear-gradient(135deg, rgba(239, 68, 68, 0.15), rgba(239, 68, 68, 0.08))'; this.style.color='var(--red)'">
                  ✗ Tolak
                </button>
              </div>` : `
              <span style="font-size: 11px; padding: 5px 12px; border-radius: 8px; 
                          background: rgba(148, 163, 184, 0.08); color: var(--text-muted); 
                          border: 1px solid rgba(148, 163, 184, 0.2); font-weight: 600; 
                          display: inline-flex; align-items: center; gap: 5px;"
                    title="Menunggu approval dari role lain">
                🔒 Terkunci
              </span>
              `}
            </td>
          </tr>
        `;

        fragment.appendChild(tempDiv.firstElementChild);
      });

      // Clear and append in one operation (faster than multiple innerHTML +=)
      tb.innerHTML = '';
      tb.appendChild(fragment);
    }
    function showHistoryModal(str) {
      const content = document.getElementById('riwayatContent'); content.innerHTML = '';
      try { str = decodeURIComponent(str); } catch (e) { }
      if (!str || str === '[]') content.innerHTML = '<div style="color:var(--gray);text-align:center;">Belum ada riwayat.</div>';
      else {
        try {
          JSON.parse(str).forEach(h => {
            let hClass = h.action === 'Diajukan' ? 'pending' : (h.action === 'Approve' ? 'approved' : 'rejected');
            content.innerHTML += `<div class="timeline-item ${hClass}"><div class="timeline-date">${formatDate(h.date)}</div><div class="timeline-title">${h.status}</div><div class="timeline-desc">Oleh: <strong>${h.by}</strong> (${h.role})</div>${h.reason ? `<div class="timeline-reason">Alasan: ${h.reason}</div>` : ''}</div>`;
          });
        } catch (e) { content.innerHTML = '<div style="color:red">Error parsing</div>'; }
      } openModal('modalRiwayatApproval');
    }

    // === UPLOAD HELPERS ===
    function switchTab(prefix, tab, btn) {
      document.querySelectorAll(`#${prefix}-panel-upload, #${prefix}-panel-url`).forEach(e => e.classList.remove('active'));
      document.querySelectorAll(`.file-tab`).forEach(e => { if (e.parentElement.parentElement.id === btn.parentElement.parentElement.id) e.classList.remove('active'); });
      document.getElementById(`${prefix}-panel-${tab}`).classList.add('active'); btn.classList.add('active');
    }
    function handleFileSelect(prefix, input) {
      const file = input.files[0]; if (!file) return;
      if (file.size > 20 * 1024 * 1024) { toast('File maks 20MB', 'error'); input.value = ''; return; }
      document.getElementById(`${prefix}-fname`).textContent = file.name;
      document.getElementById(`${prefix}-fsize`).textContent = (file.size / 1024 / 1024).toFixed(2) + ' MB';
      document.getElementById(`${prefix}-dropzone`).style.display = 'none'; document.getElementById(`${prefix}-preview`).classList.add('show');
    }
    function removeFile(prefix) {
      const input = document.getElementById(`${prefix}File`); if (input) input.value = ''; window['_droppedFile_' + prefix] = null;
      document.getElementById(`${prefix}-dropzone`).style.display = 'block'; document.getElementById(`${prefix}-preview`).classList.remove('show');
      document.getElementById(`${prefix}-pbar`).style.width = '0%'; document.getElementById(`${prefix}-progress`).classList.remove('show');
    }
    function uploadFileAndProceed(prefix, file, folder, callback, btn) {
      document.getElementById(`${prefix}-progress`).classList.add('show');
      const reader = new FileReader();
      reader.onload = e => {
        const b64 = e.target.result.split(',')[1];
        const chunkSize = 90000; const chunks = [];
        for (let i = 0; i < b64.length; i += chunkSize) chunks.push(b64.substring(i, i + chunkSize));
        let cIdx = 0; let uId = '';
        const sendChunk = () => {
          document.getElementById(`${prefix}-pbar`).style.width = ((cIdx / chunks.length) * 100) + '%';
          if (cIdx < chunks.length) {
            google.script.run.withSuccessHandler(res => { if (res.success) { uId = res.uploadId; cIdx++; sendChunk(); } else { toast(res.message, 'error'); if (btn) { btn.disabled = false; btn.textContent = btn.getAttribute('data-orig') || '💾 Simpan'; } } }).uploadChunk(chunks[cIdx], cIdx, uId);
          } else {
            google.script.run.withSuccessHandler(res => { if (res.success) callback(res.url); else { toast(res.message, 'error'); if (btn) { btn.disabled = false; btn.textContent = btn.getAttribute('data-orig') || '💾 Simpan'; } } }).finalizeChunkedUpload(uId, file.name, file.type, folder);
          }
        };
        sendChunk();
      }; reader.readAsDataURL(file);
    }
    // Drag Drop global setup
    ['kg', 'tb', 'ij', 'ast', 'pl'].forEach(p => {
      const zone = document.getElementById(`${p}-dropzone`); if (!zone) return;
      zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
      zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
      zone.addEventListener('drop', e => { e.preventDefault(); zone.classList.remove('dragover'); const files = e.dataTransfer.files; if (files.length) { window['_droppedFile_' + p] = files[0]; const dt = new DataTransfer(); dt.items.add(files[0]); document.getElementById(`${p}File`).files = dt.files; handleFileSelect(p, document.getElementById(`${p}File`)); } });
    });


    // === DASHBOARD & KEUANGAN ===
    function loadDashboard() {
      const g = id => document.getElementById(id);

      google.script.run
        .withFailureHandler(err => {
          toast('Gagal memuat dashboard: ' + err.message, 'error');
          ['statSaldoGudang', 'statSaldoTB', 'statKasIn', 'statKasOut'].forEach(id => { if (g(id)) g(id).textContent = 'Error'; });
        })
        .withSuccessHandler(res => {
          console.log('Dashboard Response:', res);
          if (!res || !res.success) {
            toast(res ? res.message : 'Respon server tidak valid atau kosong. Periksa koneksi atau database.', 'error');
            return;
          }

          if (g('statSaldoGudang')) g('statSaldoGudang').textContent = formatRp(res.saldoGudang);
          if (g('statSaldoTB')) g('statSaldoTB').textContent = formatRp(res.saldoTB);
          if (g('statKasIn')) g('statKasIn').textContent = formatRp(res.totalKasIn);
          if (g('statKasOut')) g('statKasOut').textContent = formatRp(res.totalKasOut);

          // Gunakan tanggal lokal (WIB) untuk filtering di frontend
          const now = new Date();
          const t = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + String(now.getDate()).padStart(2, '0');

          const opBody = g('dashOperasional');
          if (opBody) {
            opBody.innerHTML = '';
            const lapHariIni = (res.laporanData || []).filter(d => d && d.tanggal && String(d.tanggal).startsWith(t));
            if (!lapHariIni.length) {
              opBody.innerHTML = '<tr><td colspan="3" class="text-center text-muted">Belum ada laporan kerja hari ini</td></tr>';
            } else {
              const divSummary = {};
              lapHariIni.forEach(l => {
                const div = l.divisi || 'Lainnya';
                if (!divSummary[div]) divSummary[div] = { p: 0, o: 0 };
                const staffMasuk = (parseInt(l.totalOrang) || 0) + (parseInt(l.totalAdmin) || 0);
                const totalMasuk = staffMasuk + (parseInt(l.totalPHL) || 0);
                divSummary[div].p += totalMasuk;
                if (div === 'Distributor' || div === 'Distributor SBY' || div === 'Distributor Surabaya') {
                  divSummary[div].o += (parseInt(l.totalQty) || 0);
                } else {
                  divSummary[div].o += (parseInt(l.totalOrder) || 0);
                }
              });
              Object.keys(divSummary).forEach(div => {
                const unit = (div === 'Distributor' || div === 'Distributor SBY' || div === 'Distributor Surabaya') ? 'Qty' : (div === 'Consumable' ? 'Bubble' : 'Order');
                const icon = (div === 'Marketplace' || div === 'Market Place SBY' || div === 'Marketplace Surabaya') ? '🛒' : (div === 'Distributor' || div === 'Distributor SBY' || div === 'Distributor Surabaya') ? '📦' : div === 'Return' ? '↩️' : div === 'KOL' ? '🎥' : div === 'Inbound' ? '📥' : '🫧';
                opBody.innerHTML += `<tr><td><strong>${icon} ${div}</strong></td><td><strong>${divSummary[div].p}</strong> Orang Hadir</td><td style="color:var(--teal)"><strong>${divSummary[div].o.toLocaleString('id-ID')}</strong> ${unit}</td></tr>`;
              });
            }
          }

          const kasOutBody = g('dashKasOut');
          const kasOutList = (res.kasData || []).filter(k => k && k.tipe === 'OUT');
          if (g('badgeKasOut')) g('badgeKasOut').textContent = formatRp(kasOutList.reduce((s, k) => s + (k.nominal || 0), 0));
          if (kasOutBody) kasOutBody.innerHTML = !kasOutList.length ? '<tr><td colspan="3" class="text-center text-muted">Kosong</td></tr>' : kasOutList.map(k => `<tr><td>${formatDate(k.tanggal)}</td><td>${k.keterangan || '-'}</td><td class="negative rupiah">${formatRp(k.nominal)}</td></tr>`).join('');

          const tbBody = g('dashTBOut');
          const tbList = (res.tbData || []).filter(x => x && (!x.tipe || x.tipe === 'Pengeluaran'));
          if (g('badgeTBOut')) g('badgeTBOut').textContent = formatRp(tbList.reduce((s, x) => s + (x.nominal || 0), 0));
          if (tbBody) tbBody.innerHTML = !tbList.length ? '<tr><td colspan="3" class="text-center text-muted">Kosong</td></tr>' : tbList.map(t => `<tr><td>${formatDate(t.tanggal)}</td><td>${t.keterangan || '-'}</td><td class="negative rupiah">${formatRp(t.nominal)}</td></tr>`).join('');

          const histBody = g('dashHistory');
          if (histBody) {
            histBody.innerHTML = '';
            if (!res.history || !res.history.length) {
              histBody.innerHTML = '<tr><td colspan="5" class="empty-state"><div class="emoji">📭</div>Belum ada transaksi</td></tr>';
            } else {
              res.history.forEach(h => {
                histBody.innerHTML += `<tr><td>${formatDate(h.tanggal)}</td><td><span class="badge-tb">${h.kategori || '-'}</span></td><td><span class="${h.tipe === 'Kas Masuk' ? 'badge-in' : h.tipe === 'Kas Keluar' ? 'badge-out' : 'badge-tb'}">${h.tipe || '-'}</span></td><td>${h.keterangan || '-'}</td><td class="rupiah">${formatRp(h.nominal)}</td></tr>`;
              });
            }
          }

        })
        .getDashboardData();
    }


    async function loadKasGudang() {
      const tb = document.getElementById('tableKasGudang');
      tb.innerHTML = '';
      const res = await fetchWithFallback(
        () => supabaseSelect('kasgudang', '*', {}, { column: 'createdAt', ascending: false }),
        'getKasGudang'
      );
      if (!res || !res.success || !Array.isArray(res.data)) {
        tb.innerHTML = '<tr><td colspan="8" class="empty-state">Kosong</td></tr>';
        return;
      }
      if (!res.data.length) {
        tb.innerHTML = '<tr><td colspan="8" class="empty-state">Kosong</td></tr>';
        return;
      }
      res.data.forEach(d => tb.innerHTML += `<tr><td>${formatDate(d.tanggal)}</td><td><span class="${d.tipe === 'IN' ? 'badge-in' : 'badge-out'}">${d.tipe === 'IN' ? '📈 Masuk' : '📉 Keluar'}</span></td><td>${d.keterangan}</td><td class="${d.tipe === 'IN' ? 'positive' : 'negative'} rupiah">${formatRp(d.nominal)}</td><td>${d.buktiUrl ? `<a href="${d.buktiUrl}" target="_blank">Lihat</a>` : '-'}</td><td>${d.createdBy}</td><td><button class="btn btn-danger btn-sm" onclick="delKasGudang('${d.id}')">🗑️</button></td></tr>`);
    }
    function submitKasGudang() {
      const t = v('kgTanggal'), tp = v('kgTipe'), k = v('kgKeterangan'), n = getRpValue('kgNominal');
      if (!t || !k || n <= 0) return toast('Lengkapi form', 'error');
      const btn = document.getElementById('btnSaveKg'); btn.disabled = true; btn.textContent = '⏳...';
      const proceed = url => google.script.run.withSuccessHandler(res => { btn.disabled = false; btn.textContent = '💾 Simpan'; if (res.success) { toast('Berhasil'); closeModal('modalKasGudang'); loadKasGudang(); loadDashboard(); resetForm(['kgKeterangan']); setVal('kgNominal', ''); removeFile('kg'); } else toast(res.message, 'error'); }).addKasGudang(t, tp, k, n, url, currentUser.username);
      const f = document.getElementById('kgFile').files[0] || window['_droppedFile_kg'];
      if (f) { window['_droppedFile_kg'] = null; uploadFileAndProceed('kg', f, 'Kas Gudang', proceed, btn); } else proceed(v('kgBukti'));
    }
    function delKasGudang(id) { if (confirm('Hapus?')) google.script.run.withSuccessHandler(res => { if (res.success) { toast('Dihapus'); loadKasGudang(); loadDashboard(); } else toast(res.message, 'error'); }).deleteKasGudang(id); }

    async function loadTB() {
      const tb = document.getElementById('tableTB');
      tb.innerHTML = '';
      const res = await fetchWithFallback(
        () => supabaseSelect('teambuilding', '*', {}, { column: 'createdAt', ascending: false }),
        'getTeamBuilding'
      );
      if (!res || !res.success || !Array.isArray(res.data)) {
        tb.innerHTML = '<tr><td colspan="8" class="empty-state">Kosong</td></tr>';
        return;
      }
      if (!res.data.length) {
        tb.innerHTML = '<tr><td colspan="8" class="empty-state">Kosong</td></tr>';
        return;
      }
      res.data.forEach(d => tb.innerHTML += `<tr><td>${formatDate(d.tanggal)}</td><td><span class="${d.tipe === 'Pemasukan' ? 'badge-in' : 'badge-out'}">${d.tipe || 'Pengeluaran'}</span></td><td>${d.keterangan}</td><td class="rupiah">${formatRp(d.nominal)}</td><td>${d.buktiUrl ? `<a href="${d.buktiUrl}" target="_blank">Lihat</a>` : '-'}</td><td>${d.createdBy}</td><td><button class="btn btn-danger btn-sm" onclick="delTB('${d.id}')">🗑️</button></td></tr>`);
    }
    function submitTB() {
      const t = v('tbTanggal'), tp = v('tbTipe'), k = v('tbKeterangan'), n = getRpValue('tbNominal');
      if (!t || !k || n <= 0) return toast('Lengkapi form', 'error');
      const btn = document.getElementById('btnSaveTb'); btn.disabled = true; btn.textContent = '⏳...';
      const proceed = url => google.script.run.withSuccessHandler(res => { btn.disabled = false; btn.textContent = '💾 Simpan'; if (res.success) { toast('Berhasil'); closeModal('modalTB'); loadTB(); loadDashboard(); resetForm(['tbKeterangan']); setVal('tbNominal', ''); removeFile('tb'); } else toast(res.message, 'error'); }).addTeamBuilding(t, k, n, url, currentUser.username, tp);
      const f = document.getElementById('tbFile').files[0] || window['_droppedFile_tb'];
      if (f) { window['_droppedFile_tb'] = null; uploadFileAndProceed('tb', f, 'Team Building', proceed, btn); } else proceed(v('tbBukti'));
    }
    function delTB(id) { if (confirm('Hapus?')) google.script.run.withSuccessHandler(res => { if (res.success) { toast('Dihapus'); loadTB(); loadDashboard(); } else toast(res.message, 'error'); }).deleteTeamBuilding(id); }

    async function loadExpense() {
      const res = await fetchWithFallback(
        () => supabaseSelect('expense', '*', {}, { column: 'createdAt', ascending: false }),
        'getExpense'
      );
      if (res.success) {
        expenseData = res.data;
        populateExpenseYears();
        filterExpense();
      } else {
        toast(res.message || 'Gagal memuat data expense.', 'error');
      }
    }
    function populateExpenseYears() { const sel = document.getElementById('filterTahunExpense'), ys = new Set(expenseData.map(d => new Date(d.tanggal).getFullYear()).filter(y => !isNaN(y))); ys.add(new Date().getFullYear()); sel.innerHTML = Array.from(ys).sort((a, b) => b - a).map(y => `<option value="${y}">${y}</option>`).join(''); if (!sel.value) sel.value = new Date().getFullYear(); }
    function toggleCustomPerusahaan() { const s = v('exPerusahaanSelect'); document.getElementById('exPerusahaanCustom').style.display = s === 'Lainnya' ? 'block' : 'none'; }
    function toggleCustomKategori() { const s = v('exKategoriSelect'); document.getElementById('exKategoriCustom').style.display = s === 'Lainnya' ? 'block' : 'none'; }
    function filterExpense() {
      const y = v('filterTahunExpense') || new Date().getFullYear().toString(), m = v('filterBulanExpense'); let tY = 0, tM = 0; const cStats = {};
      const filtered = expenseData.filter(d => {
        const dY = new Date(d.tanggal).getFullYear().toString(), dM = (new Date(d.tanggal).getMonth() + 1).toString();
        if (dY !== y) return false;
        tY += d.nominal; if (!cStats[d.perusahaan]) cStats[d.perusahaan] = { b: 0, t: 0 }; cStats[d.perusahaan].t += d.nominal;
        if (m === 'all' || dM === m) { tM += d.nominal; cStats[d.perusahaan].b += d.nominal; return true; } return false;
      });
      const tb = document.getElementById('tableExpense'); tb.innerHTML = '';
      if (!filtered.length) tb.innerHTML = '<tr><td colspan="8" class="empty-state">Kosong</td></tr>';
      else filtered.sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal)).forEach(d => tb.innerHTML += `<tr><td>${formatDate(d.tanggal)}</td><td><strong>${d.perusahaan}</strong></td><td><span class="badge-tb">${d.kategori}</span></td><td>${d.keterangan || '-'}</td><td class="negative rupiah">${formatRp(d.nominal)}</td><td>${d.bank ? `<strong>${d.bank}</strong><br><small>${d.rekening}</small>` : '-'}</td><td>${d.createdBy}</td><td><button class="btn btn-danger btn-sm" onclick="delExpense('${d.id}')">🗑️</button></td></tr>`);
      document.getElementById('statExpenseTahun').textContent = formatRp(tY); document.getElementById('statExpenseBulan').textContent = formatRp(tM);
      const csObj = document.getElementById('companyExpenseStats'); csObj.innerHTML = '';
      Object.keys(cStats).sort().forEach(c => csObj.innerHTML += `<div style="background:var(--navy2);border:1px solid #ffffff15;border-radius:12px;padding:16px"><div style="font-weight:700;color:var(--accent);margin-bottom:12px">🏢 ${c}</div><div style="display:flex;justify-content:space-between"><small class="text-muted">Tahun Ini</small><strong style="color:var(--red)">${formatRp(cStats[c].t)}</strong></div><div style="display:flex;justify-content:space-between"><small class="text-muted">${m === 'all' ? 'Semua Bulan' : 'Bulan Terpilih'}</small><strong style="color:var(--amber)">${formatRp(cStats[c].b)}</strong></div></div>`);
    }
    function submitExpense() {
      const sP = v('exPerusahaanSelect'), p = sP === 'Lainnya' ? v('exPerusahaanCustom') : sP, sK = v('exKategoriSelect'), kat = sK === 'Lainnya' ? v('exKategoriCustom') : sK;
      const t = v('exTanggal'), ket = v('exKeterangan'), n = getRpValue('exNominal'), b = v('exBank'), r = v('exRekening');
      if (!t || !p || !kat || n <= 0) return toast('Lengkapi data', 'error');
      const btn = document.querySelector('#modalExpense .btn-primary'); btn.disabled = true; btn.textContent = '⏳...';
      google.script.run.withSuccessHandler(res => { btn.disabled = false; btn.textContent = '💾 Simpan'; if (res.success) { toast('Berhasil'); closeModal('modalExpense'); loadExpense(); resetForm(['exKeterangan']); setVal('exNominal', ''); } else toast(res.message, 'error'); }).addExpense(t, p, kat, ket, n, b, r, currentUser.username);
    }
    function delExpense(id) { if (confirm('Hapus?')) google.script.run.withSuccessHandler(res => { if (res.success) { toast('Dihapus'); loadExpense(); } else toast(res.message, 'error'); }).deleteExpense(id); }
    function printExpense() { window.print(); }

    // ============================================================
    // PETTY CASH PERMISSION HELPER
    // ============================================================
    function hasPCPermission(key) {
      if (!currentUser) return false;
      if (currentUser.role === 'admin') return true;
      try {
        const perms = JSON.parse(currentUser.permissions || '[]');
        return perms.includes(key);
      } catch (e) { return false; }
    }

    function applyPettyCashPermissions() {
      const canSeeStats = hasPCPermission('lihatStatsPettyCash');
      const canCetak = hasPCPermission('cetakDokumenPettyCash');

      // Stats grid tetap tampil, tapi nilai disembunyikan jika tidak punya akses
      // Tidak perlu hide statsGrid lagi

      // Tombol Cetak Dokumen Bukti & Cetak Bukti Foto
      const btnCetak = document.getElementById('btnCetakDokumenPC');
      const btnCetakFoto = document.getElementById('btnCetakFotoPC');
      if (btnCetak) btnCetak.style.display = canCetak ? '' : 'none';
      if (btnCetakFoto) btnCetakFoto.style.display = canCetak ? '' : 'none';
    }

    // ============================================================
    // PETTY CASH
    // ============================================================
    let pcData = { periods: [], transactions: [] };
    let pcActivePeriodId = null;
    let pcPendingPreviewTxId = null; // Track transaksi yang baru diupload untuk auto-preview

    function loadPettyCash(autoPreview = false) {
      console.log('=== loadPettyCash called ===');

      // Terapkan hak akses tampilan
      applyPettyCashPermissions();

      google.script.run
        .withSuccessHandler(res => {
          console.log('loadPettyCash response:', res);

          if (!res.success) {
            console.error('Failed to load Petty Cash:', res.message);
            toast('Gagal memuat Petty Cash: ' + (res.message || ''), 'error');
            return;
          }

          pcData = res;
          console.log('pcData updated:', pcData.periods.length, 'periods,', pcData.transactions.length, 'transactions');

          renderPCPeriods();

          // Auto-select periode aktif pertama
          const aktif = pcData.periods.find(p => p.status === 'Aktif');
          console.log('Active period:', aktif);

          if (aktif) {
            selectPCPeriod(aktif.id);
          } else {
            console.warn('No active period found');
            renderPCTable([]);
          }

          populatePCPeriodSelect();

          // Auto-preview jika ada pending upload
          if (autoPreview && pcPendingPreviewTxId) {
            const tx = pcData.transactions.find(t => t.id === pcPendingPreviewTxId);
            if (tx && tx.buktiUrl) {
              toast('✅ Bukti berhasil diupload');
              setTimeout(() => {
                viewPCBukti(tx.buktiUrl, tx.keterangan, formatDate(tx.tanggal), formatRp(tx.nominal));
              }, 500);
            }
            pcPendingPreviewTxId = null;
          }
        })
        .withFailureHandler(err => {
          console.error('loadPettyCash error:', err);
          toast('Error memuat Petty Cash: ' + err.message, 'error');
        })
        .getPettyCashFull();
    }

    function loadPaymentGudang() {
      // Load kedua data untuk statistik gabungan
      const bulanFilter = document.getElementById('mistineFilterBulan');

      if (bulanFilter && !bulanFilter.value) {
        const now = new Date();
        bulanFilter.value = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
      }

      const bulan = bulanFilter ? bulanFilter.value : '';

      // Load data distributor dan KOL secara paralel untuk stats
      Promise.all([
        new Promise(resolve => google.script.run.withSuccessHandler(resolve).getRekapOngkirMistineData(bulan)),
        new Promise(resolve => google.script.run.withSuccessHandler(resolve).getPaymentKOLInstantData(bulan))
      ]).then(([resDist, resKOL]) => {
        mistineRekapData = (resDist && resDist.success) ? resDist.data : [];
        kolPaymentData = (resKOL && resKOL.success) ? resKOL.data : [];

        // Render sesuai tab aktif
        const isKOL = document.getElementById('mistineTabKOL').style.display !== 'none';
        if (isKOL) {
          renderPaymentKOL(kolPaymentData);
        } else {
          renderRekapOngkirMistine(mistineRekapData);
        }

        // Update stats gabungan
        updateCombinedMistineStats(mistineRekapData, kolPaymentData);
      });
    }

    function updateCombinedMistineStats(distData, kolData) {
      const totalDist = distData.length;
      const sudahResi = distData.filter(d => d.nomorResi && d.nomorResi !== '-' && d.nomorResi !== 'null').length;
      const belumResi = totalDist - sudahResi;

      const totalOngkirDist = distData.reduce((sum, d) => sum + (parseFloat(d.hargaOngkir) || 0) + (parseFloat(d.hargaOngkirEkspedisi) || 0), 0);
      const totalOngkirKOL = kolData.reduce((sum, d) => sum + (parseFloat(d.harga) || 0), 0);
      const totalGlobal = totalOngkirDist + totalOngkirKOL;

      const elTotal = document.getElementById('mistineStatTotal');
      const elResi = document.getElementById('mistineStatResi');
      const elBelum = document.getElementById('mistineStatBelum');
      const elOngkir = document.getElementById('mistineStatOngkir');

      if (elTotal) elTotal.textContent = totalDist;
      if (elResi) elResi.textContent = sudahResi;
      if (elBelum) elBelum.textContent = belumResi;
      if (elOngkir) elOngkir.textContent = formatRp(totalGlobal);

      // Update info di masing-masing tab
      const elInfoDist = document.getElementById('mistineTableInfo');
      if (elInfoDist) elInfoDist.textContent = `${totalDist} data | Total Tab: ${formatRp(totalOngkirDist)}`;

      const elInfoKOL = document.getElementById('kolTableInfo');
      if (elInfoKOL) elInfoKOL.textContent = `${kolData.length} data | Total Tab: ${formatRp(totalOngkirKOL)}`;

      // Update stats di tab KOL juga jika ada
      const elStatKOLTotal = document.getElementById('kolStatTotal');
      const elStatKOLHarga = document.getElementById('kolStatHarga');
      if (elStatKOLTotal) elStatKOLTotal.textContent = kolData.length;
      if (elStatKOLHarga) elStatKOLHarga.textContent = formatRp(totalOngkirKOL);
    }

    function switchMistineTab(tab) {
      const distTab = document.getElementById('mistineTabDistributor');
      const kolTab = document.getElementById('mistineTabKOL');
      const distBtn = document.getElementById('btnTabMistineDistributor');
      const kolBtn = document.getElementById('btnTabMistineKOL');

      if (tab === 'kol') {
        distTab.style.display = 'none';
        kolTab.style.display = 'block';
        distBtn.classList.remove('active');
        kolBtn.classList.add('active');
        renderPaymentKOL(kolPaymentData);
      } else {
        distTab.style.display = 'block';
        kolTab.style.display = 'none';
        distBtn.classList.add('active');
        kolBtn.classList.remove('active');
        renderRekapOngkirMistine(mistineRekapData);
      }
    }

    // ============================================================
    // REKAP ONGKIR MISTINE
    // ============================================================

    let mistineRekapData = []; // Cache data mistine

    function loadRekapOngkirMistine() {
      loadPaymentGudang(); // Gunakan load global agar stats sinkron
    }

    function updateMistineStats(data) {
      // Fungsi ini sekarang digantikan oleh updateCombinedMistineStats
      // Namun kita tetap biarkan untuk kompatibilitas jika dipanggil manual
      if (typeof kolPaymentData !== 'undefined') {
        updateCombinedMistineStats(data, kolPaymentData);
      }
    }

    function renderRekapOngkirMistine(data) {
      const body = document.getElementById('tableRekapOngkirMistine');
      const footer = document.getElementById('mistineFooterTotal');
      const footerEkspedisi = document.getElementById('mistineFooterTotalEkspedisi');
      const footerBayar = document.getElementById('mistineFooterTotalBayar');
      if (!body) return;

      if (!data.length) {
        body.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:60px;color:var(--text-muted);"><div style="font-size:64px;margin-bottom:16px;opacity:0.3;">📦</div><div style="font-size:18px;font-weight:700;margin-bottom:8px;">Tidak Ada Data</div><div style="font-size:13px;">Belum ada antrian distributor MISTINE untuk bulan ini</div></td></tr>';
        if (footer) footer.textContent = 'Rp 0';
        if (footerEkspedisi) footerEkspedisi.textContent = 'Rp 0';
        if (footerBayar) footerBayar.textContent = 'Rp 0';
        return;
      }

      let totalOngkir = 0;
      let totalEkspedisi = 0;
      body.innerHTML = data.map((item, idx) => {
        const ongkir = parseFloat(item.hargaOngkir) || 0;
        const ekspedisi = parseFloat(item.hargaOngkirEkspedisi) || 0;
        const totalBayar = ongkir + ekspedisi;

        totalOngkir += ongkir;
        totalEkspedisi += ekspedisi;

        const hasResi = item.nomorResi && item.nomorResi !== '-' && item.nomorResi !== 'null' && item.nomorResi !== '';
        const resiText = hasResi ? item.nomorResi : 'Belum Input';
        const resiClass = hasResi ? 'badge-mistine-success' : 'badge-mistine-warning';

        return `
        <tr>
          <td style="text-align:center;font-weight:700;color:var(--text-muted);">${idx + 1}</td>
          <td style="font-size:12px;color:var(--text-muted);line-height:1.4;">${item.orderQueueTime || '-'}</td>
          <td>
            <div style="font-weight:700;color:var(--text-main);font-size:14px;">${item.namaDistributor || '-'}</div>
            <div style="font-size:11px;color:var(--text-muted);margin-top:2px;">PIC: ${item.picSales || '-'}</div>
          </td>
          <td style="font-family:monospace;font-size:12px;color:var(--teal);font-weight:600;">${item.noMabang || '-'}</td>
          <td style="font-size:12px;">${item.metodePengiriman || '-'}</td>
          <td><span class="badge-mistine ${resiClass}">${resiText}</span></td>
          <td style="text-align:right;font-weight:700;font-family:'Outfit',sans-serif;color:var(--accent);">${formatRp(ongkir)}</td>
          <td style="text-align:right;font-weight:700;font-family:'Outfit',sans-serif;color:var(--teal);">${formatRp(ekspedisi)}</td>
          <td style="text-align:right;font-weight:800;font-family:'Outfit',sans-serif;color:var(--green);font-size:14px;">${formatRp(totalBayar)}</td>
          <td style="text-align:center;">
            <button class="btn btn-ghost btn-sm" onclick="openEditOngkirMistineModal(${item.rowNumber}, '${item.nomorResi || ''}', ${ongkir}, ${ekspedisi})" title="Edit Data">✏️</button>
          </td>
        </tr>
      `;
      }).join('');

      if (footer) footer.textContent = formatRp(totalOngkir);
      if (footerEkspedisi) footerEkspedisi.textContent = formatRp(totalEkspedisi);
      if (footerBayar) footerBayar.textContent = formatRp(totalOngkir + totalEkspedisi);
    }

    function filterMistineTable() {
      const statusFilter = document.getElementById('mistineFilterStatus').value;
      const searchTerm = (document.getElementById('mistineFilterSearch').value || '').toLowerCase();
      const rows = document.querySelectorAll('#tableRekapOngkirMistine tr');

      let visibleCount = 0;
      let visibleOngkir = 0;
      let visibleEkspedisi = 0;

      rows.forEach(row => {
        if (row.querySelector('.empty-state') || row.cells.length < 11) return;

        const resi = (row.cells[6]?.textContent || '').toLowerCase();
        const distributor = (row.cells[3]?.textContent || '').toLowerCase();
        const noMabang = (row.cells[4]?.textContent || '').toLowerCase();

        const ongkirText = row.cells[7]?.textContent || 'Rp 0';
        const ongkirVal = parseFloat(ongkirText.replace(/[^0-9]/g, '')) || 0;

        const ekspedisiText = row.cells[8]?.textContent || 'Rp 0';
        const ekspedisiVal = parseFloat(ekspedisiText.replace(/[^0-9]/g, '')) || 0;

        const hasResi = !resi.includes('belum');
        const matchStatus = !statusFilter || (statusFilter === 'sudah' && hasResi) || (statusFilter === 'belum' && !hasResi);
        const matchSearch = !searchTerm || distributor.includes(searchTerm) || noMabang.includes(searchTerm);

        if (matchStatus && matchSearch) {
          row.style.display = '';
          visibleCount++;
          visibleOngkir += ongkirVal;
          visibleEkspedisi += ekspedisiVal;
        } else {
          row.style.display = 'none';
        }
      });

      const footer = document.getElementById('mistineFooterTotal');
      if (footer) footer.textContent = formatRp(visibleOngkir);

      const footerEkspedisi = document.getElementById('mistineFooterTotalEkspedisi');
      if (footerEkspedisi) footerEkspedisi.textContent = formatRp(visibleEkspedisi);

      const footerBayar = document.getElementById('mistineFooterTotalBayar');
      if (footerBayar) footerBayar.textContent = formatRp(visibleOngkir + visibleEkspedisi);

      const info = document.getElementById('mistineTableInfo');
      if (info) info.textContent = `${visibleCount} data ditampilkan | Total: ${formatRp(visibleOngkir + visibleEkspedisi)}`;
    }

    function openEditOngkirMistineModal(rowNumber, resi, harga, hargaEkspedisi) {
      setVal('ongkirMistineRowNumber', rowNumber);
      setVal('ongkirMistineResi', (resi === 'null' || resi === 'undefined' || !resi) ? '' : resi);
      setVal('ongkirMistineHarga', harga || 0);
      setVal('ongkirMistineHargaEkspedisi', hargaEkspedisi || 0);
      openModal('modalEditOngkirMistine');
    }

    function submitRekapOngkirMistine() {
      const rowNumber = v('ongkirMistineRowNumber');
      const resi = v('ongkirMistineResi') || '';
      const harga = parseFloat(v('ongkirMistineHarga')) || 0;
      const hargaEkspedisi = parseFloat(v('ongkirMistineHargaEkspedisi')) || 0;

      if (!rowNumber) return toast('Row number tidak valid', 'error');

      const btn = document.querySelector('#modalEditOngkirMistine .btn-primary');
      const oldText = btn.textContent;
      btn.disabled = true;
      btn.textContent = '⏳ Menyimpan...';

      google.script.run.withSuccessHandler(res => {
        btn.disabled = false;
        btn.textContent = oldText;
        if (!res.success) {
          toast(res.message || 'Gagal menyimpan data', 'error');
          return;
        }
        toast('✅ Data ongkir berhasil disimpan', 'success');
        closeModal('modalEditOngkirMistine');
        loadRekapOngkirMistine();
      }).withFailureHandler(err => {
        btn.disabled = false;
        btn.textContent = oldText;
        toast('Kesalahan sistem: ' + err.message, 'error');
      }).saveRekapOngkirMistine(rowNumber, resi, harga, hargaEkspedisi, currentUser ? currentUser.username : 'System');
    }

    function exportRekapOngkirMistine() {
      if (!mistineRekapData || mistineRekapData.length === 0) {
        return toast('Tidak ada data untuk diekspor', 'info');
      }

      const headers = ["No", "Order Queue Time", "PIC Sales", "Nama Distributor", "No Mabang", "Metode Pengiriman", "No Resi", "Harga Ongkir", "Harga Ongkir Ekspedisi", "Total Bayar"];
      const rows = mistineRekapData.map((d, idx) => {
        const ongkir = d.hargaOngkir || 0;
        const ekspedisi = d.hargaOngkirEkspedisi || 0;
        return [
          idx + 1,
          d.orderQueueTime || '-',
          d.picSales || '-',
          d.namaDistributor || '-',
          d.noMabang || '-',
          d.metodePengiriman || '-',
          d.nomorResi || 'Belum Input',
          ongkir,
          ekspedisi,
          ongkir + ekspedisi
        ];
      });

      // Header dengan BOM untuk Excel UTF-8
      const csvContent = "\uFEFF" + [headers, ...rows].map(e => e.join(",")).join("\n");
      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const link = document.createElement("a");
      const url = URL.createObjectURL(blob);
      const bulan = document.getElementById('mistineFilterBulan').value || 'All';
      link.setAttribute("href", url);
      link.setAttribute("download", `Rekap_Ongkir_MISTINE_${bulan}.csv`);
      link.style.visibility = 'hidden';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      toast('📥 File berhasil didownload', 'success');
    }

    // ============================================================
    // PAYMENT KOL INSTANT
    // ============================================================

    let kolPaymentData = [];

    function loadPaymentKOL() {
      loadPaymentGudang(); // Gunakan load global agar stats sinkron
    }

    function updateKOLStats(data) {
      const total = data.length;
      const totalHarga = data.reduce((sum, d) => sum + (parseFloat(d.harga) || 0), 0);

      const elTotal = document.getElementById('kolStatTotal');
      const elHarga = document.getElementById('kolStatHarga');
      const elInfo = document.getElementById('kolTableInfo');

      if (elTotal) elTotal.textContent = total;
      if (elHarga) elHarga.textContent = formatRp(totalHarga);
      if (elInfo) elInfo.textContent = `${total} data | Total Pembayaran: ${formatRp(totalHarga)}`;
    }

    function renderPaymentKOL(data) {
      const body = document.getElementById('tablePaymentKOL');
      const footer = document.getElementById('kolFooterTotal');
      if (!body) return;

      if (!data.length) {
        body.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:60px;color:var(--text-muted);"><div style="font-size:64px;margin-bottom:16px;opacity:0.3;">🎥</div><div style="font-size:18px;font-weight:700;margin-bottom:8px;">Tidak Ada Data</div><div style="font-size:13px;">Belum ada data pembayaran KOL untuk bulan ini</div></td></tr>';
        if (footer) footer.textContent = 'Rp 0';
        return;
      }

      let totalHarga = 0;
      body.innerHTML = data.map((item, idx) => {
        const harga = parseFloat(item.harga) || 0;
        totalHarga += harga;
        return `
        <tr>
          <td style="text-align:center;font-weight:700;color:var(--text-muted);">${idx + 1}</td>
          <td style="font-size:12px;color:var(--text-muted);">${formatDate(item.tanggal)}</td>
          <td><strong style="font-size:14px;color:var(--text-main);">${item.noOrder || '-'}</strong></td>
          <td style="font-size:13px;font-family:monospace;color:var(--teal);">${item.noResi || '-'}</td>
          <td style="text-align:right;"><strong style="font-size:15px;font-family:'Outfit',sans-serif;color:var(--accent);">${formatRp(harga)}</strong></td>
          <td style="text-align:center;">
            <div style="display:flex;gap:6px;justify-content:center;">
              <button class="btn btn-ghost btn-sm" onclick="openEditKOLModal('${item.id}', '${item.noOrder}', '${item.noResi}', ${harga})" title="Edit">✏️</button>
              <button class="btn btn-ghost btn-sm" onclick="deleteKOLPayment('${item.id}')" style="color:var(--red);" title="Hapus">🗑️</button>
            </div>
          </td>
        </tr>
      `;
      }).join('');

      if (footer) footer.textContent = formatRp(totalHarga);
    }

    function filterKOLTable() {
      const searchTerm = (document.getElementById('kolFilterSearch').value || '').toLowerCase();
      const rows = document.querySelectorAll('#tablePaymentKOL tr');

      let visibleCount = 0;
      let totalVisibleHarga = 0;
      rows.forEach(row => {
        if (row.cells.length < 5) return;
        const text = row.textContent.toLowerCase();
        const hargaText = row.cells[4]?.textContent || 'Rp 0';
        const hargaVal = parseFloat(hargaText.replace(/[^0-9]/g, '')) || 0;

        if (text.includes(searchTerm)) {
          row.style.display = '';
          visibleCount++;
          totalVisibleHarga += hargaVal;
        } else {
          row.style.display = 'none';
        }
      });

      const elInfo = document.getElementById('kolTableInfo');
      if (elInfo) elInfo.textContent = `${visibleCount} data ditampilkan | Total: ${formatRp(totalVisibleHarga)}`;

      const footer = document.getElementById('kolFooterTotal');
      if (footer) footer.textContent = formatRp(totalVisibleHarga);
    }

    function openAddKOLModal() {
      Swal.fire({
        title: 'Tambah Pembayaran KOL',
        html: `
          <div style="text-align:left;">
            <label class="form-label" style="font-size:12px;color:var(--text-muted);">No Order</label>
            <input id="swalNoOrder" class="form-control mb-3" placeholder="Contoh: KOL-12345">
            <label class="form-label" style="font-size:12px;color:var(--text-muted);">No Resi</label>
            <input id="swalNoResi" class="form-control mb-3" placeholder="Contoh: JX123456789">
            <label class="form-label" style="font-size:12px;color:var(--text-muted);">Harga</label>
            <input id="swalHarga" type="number" class="form-control mb-3" placeholder="Masukkan Nominal">
          </div>
        `,
        background: 'var(--bg-panel)',
        color: 'var(--text-main)',
        showCancelButton: true,
        confirmButtonText: 'Simpan',
        cancelButtonText: 'Batal',
        confirmButtonColor: 'var(--accent)',
        preConfirm: () => {
          return {
            noOrder: document.getElementById('swalNoOrder').value,
            noResi: document.getElementById('swalNoResi').value,
            harga: document.getElementById('swalHarga').value
          }
        }
      }).then(result => {
        if (result.isConfirmed) {
          const { noOrder, noResi, harga } = result.value;
          if (!noOrder) return toast('No Order wajib diisi', 'error');

          showLoading('Menyimpan data...');
          google.script.run.withSuccessHandler(res => {
            Swal.close();
            if (res.success) {
              toast(res.message, 'success');
              loadPaymentKOL();
            } else {
              toast(res.message, 'error');
            }
          }).savePaymentKOLInstant(null, noOrder, noResi, harga, currentUser.username);
        }
      });
    }

    function openEditKOLModal(id, noOrder, noResi, harga) {
      Swal.fire({
        title: 'Edit Pembayaran KOL',
        html: `
          <div style="text-align:left;">
            <label class="form-label" style="font-size:12px;color:var(--text-muted);">No Order</label>
            <input id="swalNoOrder" class="form-control mb-3" value="${noOrder}" placeholder="Masukkan No Order">
            <label class="form-label" style="font-size:12px;color:var(--text-muted);">No Resi</label>
            <input id="swalNoResi" class="form-control mb-3" value="${noResi}" placeholder="Masukkan No Resi">
            <label class="form-label" style="font-size:12px;color:var(--text-muted);">Harga</label>
            <input id="swalHarga" type="number" class="form-control mb-3" value="${harga}" placeholder="Masukkan Harga">
          </div>
        `,
        background: 'var(--bg-panel)',
        color: 'var(--text-main)',
        showCancelButton: true,
        confirmButtonText: 'Update',
        cancelButtonText: 'Batal',
        confirmButtonColor: 'var(--accent)',
        preConfirm: () => {
          return {
            noOrder: document.getElementById('swalNoOrder').value,
            noResi: document.getElementById('swalNoResi').value,
            harga: document.getElementById('swalHarga').value
          }
        }
      }).then(result => {
        if (result.isConfirmed) {
          const { noOrder, noResi, harga } = result.value;
          showLoading('Memperbarui data...');
          google.script.run.withSuccessHandler(res => {
            Swal.close();
            if (res.success) {
              toast(res.message, 'success');
              loadPaymentKOL();
            } else {
              toast(res.message, 'error');
            }
          }).savePaymentKOLInstant(id, noOrder, noResi, harga, currentUser.username);
        }
      });
    }

    function deleteKOLPayment(id) {
      Swal.fire({
        title: 'Hapus Data?',
        text: "Data yang dihapus tidak dapat dikembalikan!",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#ef4444',
        confirmButtonText: 'Ya, Hapus!',
        background: 'var(--bg-panel)',
        color: 'var(--text-main)'
      }).then((result) => {
        if (result.isConfirmed) {
          showLoading('Menghapus data...');
          google.script.run.withSuccessHandler(res => {
            Swal.close();
            if (res.success) {
              toast('Data berhasil dihapus', 'success');
              loadPaymentKOL();
            } else {
              toast(res.message, 'error');
            }
          }).deletePaymentKOLInstant(id);
        }
      });
    }

    function renderPCPeriods() {
      const wrap = document.getElementById('pcPeriodList');
      if (!pcData.periods.length) {
        wrap.innerHTML = '<div style="color:var(--text-muted);font-size:13px;padding:8px;">Belum ada periode. Klik <strong>Buat Periode Baru</strong> untuk memulai.</div>';
        return;
      }

      // Apply active filter selection (default: Aktif prioritized)
      const filter = window.pcPeriodFilter || 'Aktif';

      // Sort periods: Aktif first, then Selesai, then others
      const sorted = [...pcData.periods].sort((a, b) => {
        const order = status => (status === 'Aktif' ? 0 : (status === 'Selesai' ? 1 : 2));
        const oa = order(a.status || '');
        const ob = order(b.status || '');
        if (oa !== ob) return oa - ob;
        // fallback: newer start date first
        return new Date(b.tanggalMulai) - new Date(a.tanggalMulai);
      });

      const toShow = sorted.filter(p => {
        if (filter === 'Semua') return true;
        return p.status === filter;
      });

      // If filter results empty and filter is Aktif, fallback to showing all
      const finalList = toShow.length ? toShow : sorted;

      wrap.innerHTML = finalList.map(p => {
        const isActive = p.id === pcActivePeriodId;
        const statusColor = p.status === 'Aktif' ? 'var(--green)' : 'var(--gray)';
        return `<div onclick="selectPCPeriod('${p.id}')" style="cursor:pointer;padding:10px 16px;border-radius:10px;border:2px solid ${isActive ? 'var(--accent)' : 'var(--border-color)'};background:${isActive ? 'rgba(245,158,11,0.1)' : 'var(--bg-panel-light)'};min-width:180px;transition:all 0.2s;">
          <div style="font-weight:700;font-size:13px;color:${isActive ? 'var(--accent)' : 'var(--text-main)'};">${p.nama}</div>
          <div style="font-size:11px;color:var(--text-muted);margin-top:3px;">${formatDate(p.tanggalMulai)} – ${formatDate(p.tanggalSelesai)}</div>
          <div style="display:flex;justify-content:space-between;align-items:center;margin-top:6px;">
            <span style="font-size:11px;font-weight:700;color:${statusColor};">● ${p.status}</span>
            <div style="display:flex;gap:4px;">
              ${p.status === 'Aktif' ? `<button class="btn btn-ghost btn-sm" style="padding:2px 8px;font-size:10px;" onclick="event.stopPropagation();closePCPeriod('${p.id}')">🔒 Tutup</button>` : ''}
              <button class="btn btn-danger btn-sm" style="padding:2px 8px;font-size:10px;" onclick="event.stopPropagation();delPCPeriod('${p.id}')">🗑️</button>
            </div>
          </div>
        </div>`;
      }).join('');

      // Update filter button styles
      const btnAll = document.getElementById('btnPcPeriodFilterAll');
      const btnAktif = document.getElementById('btnPcPeriodFilterAktif');
      const btnSelesai = document.getElementById('btnPcPeriodFilterSelesai');
      if (btnAll) btnAll.className = filter === 'Semua' ? 'btn btn-primary btn-xs' : 'btn btn-outline-secondary btn-xs';
      if (btnAktif) btnAktif.className = filter === 'Aktif' ? 'btn btn-primary btn-xs' : 'btn btn-outline-secondary btn-xs';
      if (btnSelesai) btnSelesai.className = filter === 'Selesai' ? 'btn btn-primary btn-xs' : 'btn btn-outline-secondary btn-xs';
    }

    function selectPCPeriod(id) {
      pcActivePeriodId = id;
      const period = pcData.periods.find(p => p.id === id);
      const txs = pcData.transactions.filter(t => t.periodId === id);
      renderPCPeriods();
      renderPCTable(txs, period);
      updatePCStats(txs, period);
    }

    function renderPCTable(txs, period) {
      console.log('=== renderPCTable called ===');
      console.log('Transactions:', txs.length);
      console.log('Period:', period);

      const canSeeStats = hasPCPermission('lihatStatsPettyCash');
      const hiddenValue = '**********';

      const tb = document.getElementById('tablePettyCash');
      const title = document.getElementById('pcTableTitle');
      const sub = document.getElementById('pcTableSubtitle');

      if (period) {
        title.textContent = `📋 Transaksi: ${period.nama}`;
        sub.textContent = `${formatDate(period.tanggalMulai)} s/d ${formatDate(period.tanggalSelesai)} | Saldo Awal: ${canSeeStats ? formatRp(period.saldoAwal) : hiddenValue}`;
      }

      if (!txs.length) {
        tb.innerHTML = '<tr><td colspan="10" class="empty-state">Belum ada transaksi di periode ini</td></tr>';
        document.getElementById('pcFooterTotal').textContent = 'Rp 0';
        console.log('No transactions to render');
        return;
      }

      // Sort by date ascending
      const sorted = [...txs].sort((a, b) => new Date(a.tanggal) - new Date(b.tanggal));
      let running = period ? period.saldoAwal : 0;
      let totalOut = 0, totalIn = 0;

      tb.innerHTML = sorted.map((d, i) => {
        if (d.tipe === 'IN') { running += d.nominal; totalIn += d.nominal; }
        else { running -= d.nominal; totalOut += d.nominal; }
        const runColor = running >= 0 ? 'var(--green)' : 'var(--red)';
        const tipeHtml = d.tipe === 'IN'
          ? `<span class="badge-in">📈 Masuk</span>`
          : `<span class="badge-out">📉 Keluar</span>`;

        // Validasi data untuk mencegah error null/undefined
        const safeKeterangan = String(d.keterangan || '');
        const safeKategori = String(d.kategori || 'Lainnya');
        const safeCreatedBy = String(d.createdBy || '-');
        const safeBuktiUrl = String(d.buktiUrl || '');
        const safeNominal = parseFloat(d.nominal) || 0;
        const safeTanggal = d.tanggal || '';
        const escapedKeterangan = safeKeterangan.replace(/'/g, "\\'");

        // Bukti cell: show link + upload/replace button + inline progress
        const buktiHtml = safeBuktiUrl
          ? `<div style="display:flex;flex-direction:column;gap:4px;align-items:flex-start;" id="pcBukti_${d.id}">
               <a href="javascript:void(0)" onclick="viewPCBukti('${safeBuktiUrl}','${escapedKeterangan}','${formatDate(safeTanggal)}','${formatRp(safeNominal)}')" style="color:var(--teal);font-size:12px;font-weight:700;">📎 Lihat Bukti</a>
               <button class="btn btn-ghost btn-sm" style="padding:2px 8px;font-size:10px;color:var(--accent);" onclick="triggerPCRowUpload('${d.id}')">🔄 Ganti</button>
             </div>`
          : `<div id="pcBukti_${d.id}">
               <button class="btn btn-ghost btn-sm" style="padding:3px 10px;font-size:11px;border:1px dashed var(--accent);color:var(--accent);" onclick="triggerPCRowUpload('${d.id}')">📎 Upload</button>
             </div>`;

        // Saldo Berjalan: tampilkan bintang jika tidak punya permission
        const saldoBerjalanDisplay = canSeeStats ? formatRp(running) : hiddenValue;

        const statusText = (d.statusBayar || 'Belum Bayar');
        const statusBadge = statusText === 'Lunas' ? '<span class="badge badge-in">✅ Lunas</span>' : '<span class="badge badge-out">⏳ Belum Bayar</span>';
        const markButton = statusText === 'Lunas' ? '' : `<button class="btn btn-teal btn-sm" onclick="markPCTxPaid('${d.id}')">✅ Selesai Bayar</button>`;

        return `<tr>
          <td style="color:var(--text-muted);font-size:12px;">${i + 1}</td>
          <td>${formatDate(safeTanggal)}</td>
          <td>${tipeHtml}</td>
          <td><span style="background:rgba(14,165,233,0.1);color:var(--teal);padding:2px 8px;border-radius:6px;font-size:11px;font-weight:700;">${safeKategori}</span></td>
          <td>${safeKeterangan}</td>
          <td class="${d.tipe === 'IN' ? 'positive' : 'negative'} rupiah" style="font-weight:700;">${d.tipe === 'OUT' ? '-' : '+'}${formatRp(safeNominal)}</td>
          <td style="font-weight:800;color:${runColor};">${saldoBerjalanDisplay}</td>
          <td style="font-size:12px;">${statusBadge}</td>
          <td style="min-width:130px;">${buktiHtml}</td>
          <td style="font-size:12px;color:var(--text-muted);">${safeCreatedBy}</td>
          <td style="white-space:nowrap;display:flex;gap:6px;">${markButton}<button class="btn btn-danger btn-sm" onclick="delPCTx('${d.id}')">🗑️</button></td>
        </tr>`;
      }).join('');

      document.getElementById('pcFooterTotal').textContent = `OUT: ${formatRp(totalOut)} | IN: ${formatRp(totalIn)}`;
      console.log('Table rendered successfully with', sorted.length, 'rows');
    }

    // Upload bukti per baris transaksi
    function triggerPCRowUpload(txId) {
      document.getElementById('pcRowUploadTargetId').value = txId;
      document.getElementById('pcRowFileInput').value = '';
      document.getElementById('pcRowFileInput').click();
    }

    // ============================================================
    // PETTY CASH ROW FILE UPLOAD
    // ============================================================
    function handlePCRowFileUpload(input) {
      const file = input.files[0];
      if (!file) return;

      // Validasi ukuran file (max 10MB untuk keamanan)
      const maxSize = 10 * 1024 * 1024; // 10MB
      if (file.size > maxSize) {
        toast('❌ File terlalu besar. Maksimal 10MB', 'error');
        input.value = '';
        return;
      }

      const txId = document.getElementById('pcRowUploadTargetId').value;
      if (!txId) return;

      const tx = pcData.transactions.find(t => t.id === txId);
      if (!tx) return;

      // Get the bukti cell container
      const buktiCell = document.getElementById('pcBukti_' + txId);
      if (!buktiCell) return;

      // Replace cell content with inline progress bar
      buktiCell.innerHTML = `
        <div style="display:flex;flex-direction:column;gap:6px;min-width:120px;">
          <div style="display:flex;align-items:center;gap:6px;">
            <span id="pcIcon_${txId}" style="font-size:16px;animation:spin 1s linear infinite;">📤</span>
            <span id="pcStatus_${txId}" style="font-size:11px;font-weight:700;color:var(--text-main);">Uploading...</span>
          </div>
          <div style="background:var(--border-color);border-radius:10px;height:6px;overflow:hidden;">
            <div id="pcProgress_${txId}" style="height:100%;background:linear-gradient(90deg,var(--teal),var(--accent));border-radius:10px;width:0%;transition:width 0.2s ease-out;"></div>
          </div>
          <span id="pcPct_${txId}" style="font-size:10px;font-weight:700;color:var(--accent);">0%</span>
        </div>
      `;

      const progressBar = document.getElementById('pcProgress_' + txId);
      const pctText = document.getElementById('pcPct_' + txId);
      const statusText = document.getElementById('pcStatus_' + txId);
      const iconEl = document.getElementById('pcIcon_' + txId);

      const reader = new FileReader();
      reader.onload = function (e) {
        const b64 = e.target.result.split(',')[1];
        // Ukuran chunk 30KB untuk stabilitas maksimal dengan file besar
        const chunkSize = 30000;
        const chunks = [];
        for (let i = 0; i < b64.length; i += chunkSize) chunks.push(b64.substring(i, i + chunkSize));
        let cIdx = 0, uId = '';

        console.log('File size:', file.size, 'bytes');
        console.log('Total chunks:', chunks.length);

        const sendChunk = () => {
          const progress = Math.round((cIdx / chunks.length) * 85);
          if (progressBar) {
            progressBar.style.width = progress + '%';
            progressBar.style.transition = 'width 0.2s ease-out';
          }
          if (pctText) pctText.textContent = progress + '%';

          if (cIdx < chunks.length) {
            let retryCount = 0;
            const maxRetries = 3;

            const attemptUpload = () => {
              google.script.run
                .withSuccessHandler(res => {
                  if (res.success) {
                    uId = res.uploadId;
                    cIdx++;
                    sendChunk();
                  } else {
                    console.error('Chunk', cIdx, 'failed:', res.message);
                    if (retryCount < maxRetries) {
                      retryCount++;
                      console.log('Retrying chunk', cIdx, '(attempt', retryCount, ')');
                      setTimeout(attemptUpload, 1000 * retryCount); // Exponential backoff
                    } else {
                      onUploadError('Chunk ' + cIdx + ' gagal setelah ' + maxRetries + ' percobaan');
                    }
                  }
                })
                .withFailureHandler(err => {
                  console.error('Chunk', cIdx, 'error:', err);
                  if (retryCount < maxRetries) {
                    retryCount++;
                    console.log('Retrying chunk', cIdx, '(attempt', retryCount, ')');
                    setTimeout(attemptUpload, 1000 * retryCount);
                  } else {
                    onUploadError(err.message);
                  }
                })
                .uploadChunk(chunks[cIdx], cIdx, uId);
            };

            attemptUpload();
          } else {
            // Finalizing
            if (progressBar) progressBar.style.width = '90%';
            if (pctText) pctText.textContent = '90%';
            if (statusText) statusText.textContent = 'Finalizing...';

            // Ambil nama periode untuk folder
            const period = pcData.periods.find(p => p.id === tx.periodId);
            const periodName = period ? period.nama : 'Umum';

            google.script.run
              .withSuccessHandler(res => {
                if (res.success) {
                  // Progress 95% saat menyimpan ke database
                  if (progressBar) progressBar.style.width = '95%';
                  if (pctText) pctText.textContent = '95%';
                  if (statusText) statusText.textContent = 'Saving...';

                  onUploadDone(res.url);
                } else {
                  onUploadError(res.message, true);
                }
              })
              .withFailureHandler(err => onUploadError(err.message, true))
              .finalizeChunkedUpload(uId, file.name, file.type, 'PettyCash', periodName);
          }
        };

        const onUploadDone = (url) => {
          // Progress 100% dengan animasi
          if (progressBar) {
            progressBar.style.width = '100%';
            progressBar.style.background = 'linear-gradient(90deg,var(--green),var(--teal))';
          }
          if (pctText) {
            pctText.textContent = '100%';
            pctText.style.color = 'var(--green)';
          }
          if (statusText) statusText.textContent = 'Complete!';
          if (iconEl) {
            iconEl.style.animation = 'none';
            iconEl.textContent = '✅';
          }

          console.log('Upload done, URL:', url);

          // STEP 1: Update buktiUrl ke database dulu
          google.script.run
            .withSuccessHandler(updateRes => {
              console.log('Database update success:', updateRes);

              // STEP 2: Tunggu sebentar agar user lihat "Complete!"
              setTimeout(() => {
                // STEP 3: Reload SEMUA data dari server (PASTI SINKRON)
                console.log('Reloading all data from server...');
                google.script.run
                  .withSuccessHandler(res => {
                    console.log('Data reloaded:', res);

                    if (res.success) {
                      // Update global data
                      pcData = res;

                      // Render ulang tabel dengan data FRESH dari server
                      const period = pcData.periods.find(p => p.id === pcActivePeriodId);
                      const txs = pcData.transactions.filter(t => t.periodId === pcActivePeriodId);

                      console.log('Rendering table with', txs.length, 'transactions');
                      renderPCTable(txs, period);
                      updatePCStats(txs, period);

                      toast('✅ Bukti berhasil diupload');
                    } else {
                      console.error('Failed to reload data:', res.message);
                      toast('⚠️ Upload berhasil, silakan refresh halaman', 'warning');
                    }
                  })
                  .withFailureHandler(err => {
                    console.error('Failed to reload data:', err);
                    toast('⚠️ Upload berhasil, silakan refresh halaman', 'warning');
                  })
                  .getPettyCashFull();
              }, 600);
            })
            .withFailureHandler(err => {
              console.error('Database update failed:', err);

              // Tetap coba reload data
              setTimeout(() => {
                google.script.run
                  .withSuccessHandler(res => {
                    if (res.success) {
                      pcData = res;
                      const period = pcData.periods.find(p => p.id === pcActivePeriodId);
                      const txs = pcData.transactions.filter(t => t.periodId === pcActivePeriodId);
                      renderPCTable(txs, period);
                      updatePCStats(txs, period);
                      toast('✅ Bukti berhasil diupload');
                    }
                  })
                  .withFailureHandler(() => {
                    toast('⚠️ Upload berhasil, silakan refresh halaman', 'warning');
                  })
                  .getPettyCashFull();
              }, 600);
            })
            .uploadPettyCashBukti(txId, '', '', '', url);
        };

        const onUploadError = (msg, maybeOk) => {
          if (buktiCell) {
            buktiCell.innerHTML = `
              <div style="display:flex;flex-direction:column;gap:4px;">
                <span style="font-size:11px;color:var(--red);font-weight:600;">❌ Upload gagal</span>
                <button class="btn btn-ghost btn-sm" style="padding:2px 8px;font-size:10px;" onclick="triggerPCRowUpload('${txId}')">🔄 Coba Lagi</button>
              </div>
            `;
          }
          if (maybeOk) {
            toast('⚠️ Memeriksa status upload...', 'warning');
            setTimeout(() => loadPettyCash(), 1500);
          } else {
            toast('❌ Upload gagal: ' + (msg || ''), 'error');
          }
        };

        sendChunk();
      };
      reader.readAsDataURL(file);
    }

    // Dipanggil saat tombol "Lihat Foto" di overlay diklik
    function pcUploadViewPhoto() {
      const overlay = document.getElementById('pcUploadOverlay');
      overlay.classList.remove('show');
      const url = window._pcUploadViewUrl;
      const tx = window._pcUploadViewTx;
      if (url && tx) {
        viewPCBukti(url, tx.keterangan, formatDate(tx.tanggal), formatRp(tx.nominal));
      }
    }

    // View bukti in modal
    function viewPCBukti(url, keterangan, tanggal, nominal) {
      console.log('viewPCBukti called with URL:', url);

      if (!url) {
        toast('❌ URL bukti tidak tersedia', 'error');
        return;
      }

      document.getElementById('pcBuktiTanggal').textContent = tanggal;
      document.getElementById('pcBuktiNominal').textContent = nominal;
      document.getElementById('pcBuktiKeterangan').textContent = keterangan;

      // Buat URL yang lebih baik untuk preview dan download
      const fileId = extractGDriveFileId(url);
      console.log('Extracted file ID:', fileId);

      // Gunakan multiple URL format untuk preview - prioritaskan yang paling reliable
      const embedUrl = fileId ? `https://drive.google.com/file/d/${fileId}/preview` : null;
      const thumbnailUrl = fileId ? `https://drive.google.com/thumbnail?id=${fileId}&sz=w1000` : null;
      const directUrl = fileId ? `https://lh3.googleusercontent.com/d/${fileId}` : null;
      const ucUrl = fileId ? `https://drive.google.com/uc?export=view&id=${fileId}` : url;
      const downloadUrl = fileId ? `https://drive.google.com/file/d/${fileId}/view` : url;

      console.log('Embed URL:', embedUrl);
      console.log('Thumbnail URL:', thumbnailUrl);
      console.log('Direct URL:', directUrl);
      console.log('UC URL:', ucUrl);
      console.log('Download URL:', downloadUrl);

      document.getElementById('pcBuktiDownloadBtn').href = downloadUrl;

      const preview = document.getElementById('pcBuktiPreview');
      preview.innerHTML = '<div style="text-align:center;padding:40px;color:var(--text-muted);font-size:13px;">⏳ Memuat bukti...</div>';

      // Buka modal dulu
      openModal('modalViewPCBukti');

      // Coba beberapa metode untuk menampilkan gambar
      const previewUrls = [
        { url: thumbnailUrl, type: 'img', name: 'Thumbnail API' },
        { url: directUrl, type: 'img', name: 'Google User Content' },
        { url: ucUrl, type: 'img', name: 'UC Export' },
        { url: embedUrl, type: 'iframe', name: 'Embed Preview' }
      ].filter(item => item.url); // Filter out null URLs

      let currentIndex = 0;

      const tryLoadImage = () => {
        if (currentIndex >= previewUrls.length) {
          // Semua metode gagal, tampilkan fallback
          console.log('All preview methods failed, showing fallback');
          showFallback();
          return;
        }

        const current = previewUrls[currentIndex];
        console.log(`Trying preview method ${currentIndex + 1} (${current.name}):`, current.url);

        if (current.type === 'img') {
          const img = new Image();
          img.onload = function () {
            console.log(`Preview method ${currentIndex + 1} (${current.name}) succeeded!`);
            preview.innerHTML = `<div style="text-align:center;">
              <img src="${current.url}" style="max-width:100%;max-height:500px;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.2);" alt="Bukti Transaksi">
            </div>`;
          };
          img.onerror = function (e) {
            console.error(`Preview method ${currentIndex + 1} (${current.name}) failed:`, e);
            currentIndex++;
            tryLoadImage();
          };
          img.crossOrigin = 'anonymous';
          img.src = current.url;
        } else if (current.type === 'iframe') {
          // Try iframe embed
          preview.innerHTML = `<div style="text-align:center;">
            <iframe src="${current.url}" style="width:100%;height:500px;border:none;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.2);" allow="autoplay"></iframe>
          </div>`;

          // Give iframe some time to load
          setTimeout(() => {
            console.log(`Preview method ${currentIndex + 1} (${current.name}) loaded (iframe)`);
          }, 1000);
        }
      };

      const showFallback = () => {
        // Jika semua gagal, tampilkan dengan iframe atau link
        const isPDF = /\.pdf(\?|$)/i.test(url) || /\.pdf$/i.test(downloadUrl);

        if (isPDF) {
          console.log('Detected as PDF, showing iframe');
          preview.innerHTML = `
            <div style="width:100%;height:500px;border-radius:8px;overflow:hidden;background:#f5f5f5;border:1px solid var(--border-color);">
              <iframe src="${ucUrl}" style="width:100%;height:100%;border:none;" frameborder="0"></iframe>
            </div>
            <div style="text-align:center;margin-top:12px;font-size:12px;color:var(--text-muted);">
              💡 Jika PDF tidak tampil, klik "Buka di Tab Baru"
            </div>
          `;
        } else {
          console.log('Showing download link fallback');
          preview.innerHTML = `<div style="text-align:center;padding:40px;color:var(--text-muted);">
            <div style="font-size:64px;margin-bottom:16px;">📷</div>
            <div style="font-size:16px;font-weight:700;color:var(--text-main);margin-bottom:8px;">Bukti Transaksi</div>
            <div style="font-size:13px;margin-bottom:20px;">Foto berhasil diupload ke Google Drive</div>
            <a href="${downloadUrl}" target="_blank" class="btn btn-teal" style="text-decoration:none;display:inline-block;">
              📥 Buka Foto di Google Drive
            </a>
          </div>`;
        }
      };

      // Mulai mencoba load gambar
      tryLoadImage();
    }

    // Helper function untuk extract file ID dari Google Drive URL
    function extractGDriveFileId(url) {
      if (!url) return null;
      // Pattern 1: https://drive.google.com/uc?export=view&id=FILE_ID
      let match = url.match(/[?&]id=([^&]+)/);
      if (match) return match[1];
      // Pattern 2: https://drive.google.com/file/d/FILE_ID/view
      match = url.match(/\/file\/d\/([^\/]+)/);
      if (match) return match[1];
      // Pattern 3: https://drive.google.com/open?id=FILE_ID
      match = url.match(/open\?id=([^&]+)/);
      if (match) return match[1];
      return null;
    }

    function filterPCTable() {
      if (!pcActivePeriodId) return;
      const q = document.getElementById('pcSearchInput').value.toLowerCase();
      const period = pcData.periods.find(p => p.id === pcActivePeriodId);
      let txs = pcData.transactions.filter(t => t.periodId === pcActivePeriodId);
      if (q) txs = txs.filter(t => t.keterangan.toLowerCase().includes(q) || t.kategori.toLowerCase().includes(q));
      // If filter toggle active, show only transactions that are not marked 'Lunas'
      if (window.pcFilterBelum) txs = txs.filter(t => (t.statusBayar || 'Belum Bayar') !== 'Lunas');
      renderPCTable(txs, period);
    }

    function updatePCStats(txs, period) {
      const canSeeStats = hasPCPermission('lihatStatsPettyCash');

      const saldoAwal = period ? period.saldoAwal : 0;
      let totalOut = 0, totalIn = 0;
      txs.forEach(t => { if (t.tipe === 'IN') totalIn += t.nominal; else totalOut += t.nominal; });
      const sisa = saldoAwal + totalIn - totalOut;

      // Jika tidak punya permission, tampilkan bintang
      const hiddenValue = '**********';
      document.getElementById('pcStatSaldoAwal').textContent = canSeeStats ? formatRp(saldoAwal) : hiddenValue;
      document.getElementById('pcStatOut').textContent = canSeeStats ? formatRp(totalOut) : hiddenValue;
      document.getElementById('pcStatIn').textContent = canSeeStats ? formatRp(totalIn) : hiddenValue;
      document.getElementById('pcStatSisa').textContent = canSeeStats ? formatRp(sisa) : hiddenValue;
    }

    function populatePCPeriodSelect() {
      const sel = document.getElementById('pcTxPeriodSelect');
      sel.innerHTML = '<option value="">-- Pilih Periode --</option>';
      pcData.periods.forEach(p => {
        const opt = document.createElement('option');
        opt.value = p.id;
        opt.textContent = `${p.nama} (${formatDate(p.tanggalMulai)} – ${formatDate(p.tanggalSelesai)})`;
        sel.appendChild(opt);
      });
      // Auto-select active period
      if (pcActivePeriodId) sel.value = pcActivePeriodId;
      // Ensure default filter for period list is Aktif
      if (!window.pcPeriodFilter) window.pcPeriodFilter = 'Aktif';
    }

    function setPCPeriodFilter(f) {
      window.pcPeriodFilter = f || 'Aktif';
      renderPCPeriods();
    }

    function togglePCKategoriCustom() {
      const v = document.getElementById('pcTxKategoriSelect').value;
      document.getElementById('pcKategoriCustomRow').style.display = v === 'Lainnya' ? 'flex' : 'none';
    }

    function submitPCPeriod() {
      const nama = document.getElementById('pcpNama').value.trim();
      const mulai = document.getElementById('pcpMulai').value;
      const selesai = document.getElementById('pcpSelesai').value;
      const saldo = getRpValue('pcpSaldoAwal');
      const ket = document.getElementById('pcpKeterangan').value.trim();
      if (!nama || !mulai || !selesai) return toast('Nama, tanggal mulai & selesai wajib diisi', 'error');
      if (new Date(selesai) < new Date(mulai)) return toast('Tanggal selesai harus setelah tanggal mulai', 'error');
      const btn = document.getElementById('btnSavePCP');
      btn.disabled = true; btn.textContent = '⏳...';
      google.script.run.withSuccessHandler(res => {
        btn.disabled = false; btn.textContent = '💾 Simpan Periode';
        if (res.success) {
          toast('Periode berhasil dibuat');
          closeModal('modalPCPeriod');
          resetForm(['pcpNama', 'pcpKeterangan']); setVal('pcpSaldoAwal', '');
          loadPettyCash();
        } else toast(res.message, 'error');
      }).addPettyCashPeriod(nama, mulai, selesai, saldo, ket, currentUser.username);
    }

    function submitPettyCash() {
      const periodId = document.getElementById('pcTxPeriodSelect').value;
      const tgl = document.getElementById('pcTxTanggal').value;
      const tipe = document.getElementById('pcTxTipe').value;
      const katSel = document.getElementById('pcTxKategoriSelect').value;
      const kat = katSel === 'Lainnya' ? document.getElementById('pcTxKategoriCustom').value.trim() : katSel;
      const ket = document.getElementById('pcTxKeterangan').value.trim();
      const nom = getRpValue('pcTxNominal');
      if (!periodId) return toast('Pilih periode terlebih dahulu', 'error');
      if (!tgl || !kat || nom <= 0) return toast('Lengkapi semua field', 'error');
      const btn = document.getElementById('btnSavePCTx');
      btn.disabled = true; btn.textContent = '⏳...';
      const hasFile = !!(document.getElementById('pcTxFile').files[0] || window['_droppedFile_pcTx']);
      const proceed = url => google.script.run.withSuccessHandler(res => {
        btn.disabled = false; btn.textContent = '💾 Simpan Transaksi';
        if (res.success) {
          toast('Transaksi berhasil disimpan');
          closeModal('modalPettyCash');
          resetForm(['pcTxKeterangan', 'pcTxKategoriCustom']); setVal('pcTxNominal', ''); removeFile('pcTx');
          // Jika ada bukti, reload lalu auto-buka Lihat Foto
          if (url) {
            google.script.run.withSuccessHandler(r => {
              if (!r.success) return;
              pcData = r;
              renderPCPeriods();
              const aktif = pcData.periods.find(p => p.id === periodId);
              if (aktif) selectPCPeriod(aktif.id);
              populatePCPeriodSelect();
              // Cari transaksi yang baru disimpan (url cocok)
              const newTx = pcData.transactions.find(t => t.buktiUrl === url);
              if (newTx) {
                setTimeout(() => {
                  viewPCBukti(url, newTx.keterangan, formatDate(newTx.tanggal), formatRp(newTx.nominal));
                }, 400);
              }
            }).getPettyCashFull();
          } else {
            loadPettyCash();
          }
        } else toast(res.message, 'error');
      }).addPettyCash(periodId, tgl, kat, ket, nom, tipe, url, currentUser.username);
      const f = document.getElementById('pcTxFile').files[0] || window['_droppedFile_pcTx'];
      if (f) { window['_droppedFile_pcTx'] = null; uploadFileAndProceed('pcTx', f, 'PettyCash', proceed, btn); }
      else proceed(document.getElementById('pcTxBukti').value);
    }

    function delPCTx(id) {
      if (!confirm('Hapus transaksi ini?')) return;
      google.script.run.withSuccessHandler(res => {
        if (res.success) { toast('Transaksi dihapus'); loadPettyCash(); }
        else toast(res.message, 'error');
      }).deletePettyCash(id);
    }

    // Toggle: show only 'Belum Bayar' transactions
    window.pcFilterBelum = false;
    function togglePCFilterBelum() {
      window.pcFilterBelum = !window.pcFilterBelum;
      const btn = document.getElementById('btnPcFilterBelum');
      if (btn) {
        btn.className = window.pcFilterBelum ? 'btn btn-primary btn-sm' : 'btn btn-outline-secondary btn-sm';
        btn.textContent = window.pcFilterBelum ? '🔎 Menampilkan: Belum Bayar' : '⏳ Belum Bayar';
      }
      // Re-filter current period
      if (pcActivePeriodId) {
        const period = pcData.periods.find(p => p.id === pcActivePeriodId);
        let txs = pcData.transactions.filter(t => t.periodId === pcActivePeriodId);
        const q = (document.getElementById('pcSearchInput')?.value || '').toLowerCase();
        if (q) txs = txs.filter(t => t.keterangan.toLowerCase().includes(q) || t.kategori.toLowerCase().includes(q));
        if (window.pcFilterBelum) txs = txs.filter(t => (t.statusBayar || 'Belum Bayar') !== 'Lunas');
        renderPCTable(txs, period);
      }
    }

    function markPCTxPaid(txId) {
      if (!confirm('Tandai transaksi ini sebagai Lunas (Selesai Bayar)?')) return;
      google.script.run.withSuccessHandler(res => {
        if (!res.success) return toast(res.message || 'Gagal menandai Lunas', 'error');
        toast('Transaksi ditandai Lunas');
        loadPettyCash();
      }).updatePettyCashStatus(txId, 'Lunas');
    }

    function delPCPeriod(id) {
      if (!confirm('Hapus periode ini beserta semua transaksinya?')) return;
      google.script.run.withSuccessHandler(res => {
        if (res.success) { toast('Periode dihapus'); if (pcActivePeriodId === id) pcActivePeriodId = null; loadPettyCash(); }
        else toast(res.message, 'error');
      }).deletePettyCashPeriod(id);
    }

    function closePCPeriod(id) {
      if (!confirm('Tutup periode ini? Status akan berubah menjadi Selesai.')) return;
      google.script.run.withSuccessHandler(res => {
        if (res.success) { toast('Periode ditutup'); loadPettyCash(); }
        else toast(res.message, 'error');
      }).updatePettyCashPeriodStatus(id, 'Selesai');
    }

    function openModal_PettyCash_withPeriod() {
      populatePCPeriodSelect();
      openModal('modalPettyCash');
    }

    function exportPCExcel() {
      if (!pcActivePeriodId) return toast('Pilih periode terlebih dahulu', 'warning');
      const period = pcData.periods.find(p => p.id === pcActivePeriodId);
      if (!period) return;

      // 1. Download lokal dulu (instan, tidak perlu server)
      const txs = pcData.transactions.filter(t => t.periodId === pcActivePeriodId)
        .sort((a, b) => new Date(a.tanggal) - new Date(b.tanggal));
      let running = period.saldoAwal;
      let totalOut = 0, totalIn = 0;
      let rows = '';
      txs.forEach((d, i) => {
        if (d.tipe === 'IN') { running += d.nominal; totalIn += d.nominal; }
        else { running -= d.nominal; totalOut += d.nominal; }
        rows += `<tr>
          <td>${i + 1}</td><td>${d.tanggal}</td><td>${d.tipe}</td><td>${d.kategori}</td>
          <td>${d.keterangan}</td>
          <td>${d.tipe === 'OUT' ? -d.nominal : d.nominal}</td>
          <td>${running}</td>
          <td>${d.buktiUrl || '-'}</td>
          <td>${d.createdBy}</td>
        </tr>`;
      });
      const html = `<table>
        <tr><td colspan="9"><b>LAPORAN PETTY CASH - ${period.nama}</b></td></tr>
        <tr><td colspan="9">Periode: ${formatDate(period.tanggalMulai)} s/d ${formatDate(period.tanggalSelesai)} | Saldo Awal: ${formatRp(period.saldoAwal)}</td></tr>
        <tr></tr>
        <thead><tr><th>No</th><th>Tanggal</th><th>Tipe</th><th>Kategori</th><th>Keterangan</th><th>Nominal</th><th>Saldo Berjalan</th><th>Link Bukti</th><th>Input By</th></tr></thead>
        <tbody>${rows}</tbody>
        <tfoot>
          <tr><td colspan="5"><b>TOTAL PENGELUARAN</b></td><td><b>${-totalOut}</b></td><td colspan="3"></td></tr>
          <tr><td colspan="5"><b>SISA SALDO AKHIR</b></td><td><b>${period.saldoAwal + totalIn - totalOut}</b></td><td colspan="3"></td></tr>
        </tfoot>
      </table>`;
      exportToExcel(html, `PettyCash_${period.nama.replace(/\s/g, '_')}.xls`);

      // 2. Upload ke Google Drive (background, tampilkan link setelah selesai)
      const btn = document.querySelector('[onclick="exportPCExcel()"]');
      if (btn) { btn.disabled = true; btn.textContent = '⏳ Upload Drive...'; }
      google.script.run.withSuccessHandler(res => {
        if (btn) { btn.disabled = false; btn.textContent = '📤 Export Excel'; }
        if (res.success) {
          // Tampilkan notifikasi dengan link ke file dan folder
          const msg = `✅ File tersimpan di Google Drive!\n📁 Folder: ${res.periodNama}`;
          toast(msg, 'success');
          // Tampilkan popup kecil dengan link
          showPCDriveLinks(res.fileUrl, res.folderUrl, res.fileName);
        } else {
          toast('⚠️ Download lokal berhasil, tapi upload Drive gagal: ' + res.message, 'warning');
        }
      }).exportPettyCashToGDrive(pcActivePeriodId);
    }

    function showPCDriveLinks(fileUrl, folderUrl, fileName) {
      // Buat toast/popup dengan link Drive
      const existing = document.getElementById('pcDriveToast');
      if (existing) existing.remove();
      const el = document.createElement('div');
      el.id = 'pcDriveToast';
      el.style.cssText = `
        position:fixed; bottom:80px; right:20px; z-index:9999;
        background:var(--bg-panel); border:1px solid var(--green);
        border-radius:12px; padding:16px 20px; max-width:340px;
        box-shadow:0 8px 32px rgba(0,0,0,0.4); animation:slideDown 0.3s ease;
      `;
      el.innerHTML = `
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
          <span style="font-weight:800;color:var(--green);font-size:13px;">✅ Tersimpan di Google Drive</span>
          <button onclick="document.getElementById('pcDriveToast').remove()" style="background:none;border:none;color:var(--text-muted);cursor:pointer;font-size:16px;">✕</button>
        </div>
        <div style="font-size:11px;color:var(--text-muted);margin-bottom:10px;word-break:break-all;">${fileName}</div>
        <div style="display:flex;flex-direction:column;gap:8px;">
          <a href="${fileUrl}" target="_blank" style="display:flex;align-items:center;gap:8px;padding:8px 12px;background:rgba(16,185,129,0.1);border:1px solid rgba(16,185,129,0.3);border-radius:8px;color:var(--green);font-size:12px;font-weight:700;text-decoration:none;">
            📄 Buka File Excel
          </a>
          <a href="${folderUrl}" target="_blank" style="display:flex;align-items:center;gap:8px;padding:8px 12px;background:rgba(14,165,233,0.1);border:1px solid rgba(14,165,233,0.3);border-radius:8px;color:var(--teal);font-size:12px;font-weight:700;text-decoration:none;">
            📁 Buka Folder Periode
          </a>
        </div>
      `;
      document.body.appendChild(el);
      // Auto-close setelah 15 detik
      setTimeout(() => { if (document.getElementById('pcDriveToast')) document.getElementById('pcDriveToast').remove(); }, 15000);
    }

    function cetakDokumenPettyCash() {
      if (!hasPCPermission('cetakDokumenPettyCash')) return toast('❌ Anda tidak memiliki akses untuk mencetak dokumen', 'error');
      if (!pcActivePeriodId) return toast('Pilih periode terlebih dahulu', 'warning');
      const period = pcData.periods.find(p => p.id === pcActivePeriodId);
      if (!period) return;
      const txs = pcData.transactions
        .filter(t => t.periodId === pcActivePeriodId)
        .sort((a, b) => new Date(a.tanggal) - new Date(b.tanggal));
      if (!txs.length) return toast('Belum ada transaksi di periode ini', 'warning');

      let running = period.saldoAwal;
      let totalOut = 0, totalIn = 0;

      // Build table rows
      const rows = txs.map((d, i) => {
        if (d.tipe === 'IN') { running += d.nominal; totalIn += d.nominal; }
        else { running -= d.nominal; totalOut += d.nominal; }
        const nomStr = d.tipe === 'OUT'
          ? `<span style="color:#dc2626;font-weight:700;">- Rp ${d.nominal.toLocaleString('id-ID')}</span>`
          : `<span style="color:#16a34a;font-weight:700;">+ Rp ${d.nominal.toLocaleString('id-ID')}</span>`;
        const runColor = running >= 0 ? '#16a34a' : '#dc2626';
        const buktiCell = d.buktiUrl
          ? `<a href="${d.buktiUrl}" target="_blank" style="color:#0369a1;font-size:11px;">📎 Lihat</a>`
          : `<span style="color:#9ca3af;font-size:11px;">-</span>`;
        return `<tr style="border-bottom:1px solid #e5e7eb;">
          <td style="padding:8px 6px;text-align:center;color:#6b7280;font-size:12px;">${i + 1}</td>
          <td style="padding:8px 6px;font-size:12px;">${new Date(d.tanggal).toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' })}</td>
          <td style="padding:8px 6px;text-align:center;">
            <span style="padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700;background:${d.tipe === 'IN' ? '#dcfce7' : '#fee2e2'};color:${d.tipe === 'IN' ? '#16a34a' : '#dc2626'};">${d.tipe}</span>
          </td>
          <td style="padding:8px 6px;font-size:11px;font-weight:700;color:#0369a1;">${d.kategori}</td>
          <td style="padding:8px 6px;font-size:12px;">${d.keterangan}</td>
          <td style="padding:8px 6px;text-align:right;">${nomStr}</td>
          <td style="padding:8px 6px;text-align:right;font-weight:800;color:${runColor};font-size:12px;">Rp ${running.toLocaleString('id-ID')}</td>
          <td style="padding:8px 6px;text-align:center;">${buktiCell}</td>
          <td style="padding:8px 6px;font-size:11px;color:#6b7280;">${d.createdBy}</td>
        </tr>`;
      }).join('');

      // Build proof images section — grouped by kategori
      const withBukti = txs.filter(d => d.buktiUrl);

      let proofSection = '';
      if (withBukti.length > 0) {
        // Group by kategori
        const grouped = {};
        withBukti.forEach(d => {
          const k = d.kategori || 'Lainnya';
          if (!grouped[k]) grouped[k] = [];
          grouped[k].push(d);
        });

        const categoryBlocks = Object.entries(grouped).map(([kat, items]) => {
          const cards = items.map(d => {
            const txNo = txs.indexOf(d) + 1;
            const isImg = /\.(jpg|jpeg|png|gif|webp)/i.test((d.buktiUrl || '').split('?')[0]);
            const fileId = extractGDriveFileId ? extractGDriveFileId(d.buktiUrl) : null;
            const imgSrc = fileId ? `https://drive.google.com/uc?export=view&id=${fileId}` : d.buktiUrl;
            const imgHtml = isImg || fileId
              ? `<img src="${imgSrc}" style="width:100%;max-height:420px;object-fit:contain;border-radius:6px;border:1px solid #e5e7eb;margin-top:8px;display:block;" onerror="this.style.display='none';this.nextElementSibling.style.display='flex';">
                 <div style="display:none;padding:20px;background:#f9fafb;border-radius:6px;text-align:center;color:#6b7280;font-size:12px;margin-top:8px;align-items:center;justify-content:center;flex-direction:column;gap:6px;">
                   <span style="font-size:24px;">🖼️</span><span>Gambar tidak dapat dimuat</span>
                   <a href="${d.buktiUrl}" target="_blank" style="color:#0369a1;font-size:11px;">Buka Link</a>
                 </div>`
              : `<div style="margin-top:8px;padding:12px;background:#f0f9ff;border-radius:6px;border:1px solid #bae6fd;">
                   <a href="${d.buktiUrl}" target="_blank" style="color:#0369a1;font-size:12px;font-weight:700;word-break:break-all;">📎 ${d.buktiUrl}</a>
                 </div>`;
            return `<div style="border:1px solid #e5e7eb;border-radius:8px;padding:12px;background:#fff;page-break-inside:avoid;">
              <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">
                <span style="font-size:11px;font-weight:800;color:#1e3a5f;background:#eff6ff;padding:2px 8px;border-radius:4px;">#${txNo}</span>
                <span style="font-size:10px;color:#6b7280;">${new Date(d.tanggal).toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' })}</span>
              </div>
              <div style="font-size:12px;font-weight:700;color:#111827;margin:4px 0 2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${d.keterangan}">${d.keterangan}</div>
              <div style="font-size:11px;color:#6b7280;"><strong style="color:${d.tipe === 'OUT' ? '#dc2626' : '#16a34a'};">${d.tipe === 'OUT' ? '- ' : '+'}Rp ${d.nominal.toLocaleString('id-ID')}</strong> &nbsp;·&nbsp; ${d.createdBy}</div>
              ${imgHtml}
            </div>`;
          }).join('');

          return `<div style="margin-bottom:32px;page-break-inside:avoid;">
            <div style="display:flex;align-items:center;gap:10px;margin-bottom:14px;padding:10px 16px;background:linear-gradient(135deg,#1e3a5f,#1d4ed8);border-radius:8px;">
              <span style="font-size:16px;">📂</span>
              <span style="font-family:'Outfit',sans-serif;font-size:14px;font-weight:800;color:#fff;letter-spacing:0.3px;">${kat}</span>
              <span style="margin-left:auto;background:rgba(255,255,255,0.2);color:#fff;font-size:11px;font-weight:700;padding:2px 10px;border-radius:20px;">${items.length} bukti</span>
            </div>
            <div style="display:grid;grid-template-columns:repeat(2,1fr);gap:20px;">
              ${cards}
            </div>
          </div>`;
        }).join('');

        proofSection = `
          <div style="page-break-before:always;">
            <h2 style="font-family:'Outfit',sans-serif;font-size:18px;font-weight:800;color:#1e3a5f;border-bottom:3px solid #f59e0b;padding-bottom:8px;margin-bottom:24px;">
              📎 LAMPIRAN BUKTI TRANSAKSI &nbsp;<span style="font-size:13px;font-weight:600;color:#6b7280;">(${withBukti.length} dari ${txs.length} transaksi)</span>
            </h2>
            ${categoryBlocks}
          </div>`;
      }

      const printWin = window.open('', '_blank', 'width=1000,height=800,scrollbars=yes');
      if (!printWin) return toast('Pop-up diblokir browser! Izinkan popup untuk website ini.', 'error');

      printWin.document.write(`<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <title>Dokumen Petty Cash - ${period.nama}</title>
  <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;600;700;800&display=swap" rel="stylesheet">
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Outfit', Arial, sans-serif; color: #111827; background: #fff; padding: 32px; font-size: 13px; }
    .doc-header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 28px; padding-bottom: 16px; border-bottom: 3px solid #f59e0b; }
    .doc-title { font-size: 22px; font-weight: 800; color: #1e3a5f; }
    .doc-subtitle { font-size: 13px; color: #6b7280; margin-top: 4px; }
    .doc-meta { text-align: right; font-size: 12px; color: #6b7280; }
    .doc-meta strong { color: #1e3a5f; display: block; font-size: 14px; }
    .summary-grid { display: grid; grid-template-columns: repeat(4,1fr); gap: 12px; margin-bottom: 24px; }
    .summary-box { border-radius: 8px; padding: 14px; text-align: center; }
    .summary-box .label { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px; }
    .summary-box .value { font-size: 16px; font-weight: 800; }
    .box-blue { background: #eff6ff; border: 1px solid #bfdbfe; }
    .box-blue .label { color: #1d4ed8; } .box-blue .value { color: #1d4ed8; }
    .box-red { background: #fef2f2; border: 1px solid #fecaca; }
    .box-red .label { color: #dc2626; } .box-red .value { color: #dc2626; }
    .box-green { background: #f0fdf4; border: 1px solid #bbf7d0; }
    .box-green .label { color: #16a34a; } .box-green .value { color: #16a34a; }
    .box-amber { background: #fffbeb; border: 1px solid #fde68a; }
    .box-amber .label { color: #d97706; } .box-amber .value { color: #d97706; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 24px; }
    thead tr { background: #1e3a5f; }
    thead th { padding: 10px 8px; color: #fff; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; text-align: left; }
    thead th:nth-child(1), thead th:nth-child(3), thead th:nth-child(8) { text-align: center; }
    thead th:nth-child(6), thead th:nth-child(7) { text-align: right; }
    tbody tr:nth-child(even) { background: #f9fafb; }
    .footer-row td { padding: 10px 8px; font-weight: 800; background: #f3f4f6; border-top: 2px solid #d1d5db; }
    .sign-section { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 40px; margin-top: 48px; }
    .sign-box { text-align: center; }
    .sign-line { border-top: 1px solid #374151; margin-top: 60px; padding-top: 6px; font-size: 12px; font-weight: 700; }
    .sign-role { font-size: 11px; color: #6b7280; margin-top: 2px; }
    @media print {
      body { padding: 16px; }
      .no-print { display: none !important; }
      @page { margin: 1cm; size: A4; }
    }
  </style>
</head>
<body>
  <div class="doc-header">
    <div>
      <div class="doc-title">🪙 LAPORAN PETTY CASH</div>
      <div class="doc-subtitle">${period.nama}</div>
      <div class="doc-subtitle" style="margin-top:4px;">Periode: ${new Date(period.tanggalMulai).toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' })} s/d ${new Date(period.tanggalSelesai).toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' })}</div>
    </div>
    <div class="doc-meta">
      <strong>GUDANG FCL GROUP</strong>
      Dicetak: ${new Date().toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' })}<br>
      Oleh: ${currentUser.nama}
    </div>
  </div>

  <div class="summary-grid">
    <div class="summary-box box-blue">
      <div class="label">Saldo Awal</div>
      <div class="value">Rp ${period.saldoAwal.toLocaleString('id-ID')}</div>
    </div>
    <div class="summary-box box-red">
      <div class="label">Total Pengeluaran</div>
      <div class="value">Rp ${totalOut.toLocaleString('id-ID')}</div>
    </div>
    <div class="summary-box box-green">
      <div class="label">Total Pemasukan</div>
      <div class="value">Rp ${totalIn.toLocaleString('id-ID')}</div>
    </div>
    <div class="summary-box box-amber">
      <div class="label">Sisa Saldo</div>
      <div class="value">Rp ${(period.saldoAwal + totalIn - totalOut).toLocaleString('id-ID')}</div>
    </div>
  </div>

  <h2 style="font-size:15px;font-weight:800;color:#1e3a5f;margin-bottom:12px;">📋 RINCIAN TRANSAKSI</h2>
  <table>
    <thead>
      <tr>
        <th style="width:36px;">No</th>
        <th>Tanggal</th>
        <th style="width:50px;">Tipe</th>
        <th>Kategori</th>
        <th>Keterangan</th>
        <th style="text-align:right;">Nominal</th>
        <th style="text-align:right;">Saldo</th>
        <th style="width:60px;text-align:center;">Bukti</th>
        <th>Input By</th>
      </tr>
    </thead>
    <tbody>${rows}</tbody>
    <tfoot>
      <tr class="footer-row">
        <td colspan="5" style="text-align:right;padding:10px 8px;color:#374151;">TOTAL PENGELUARAN</td>
        <td style="text-align:right;padding:10px 8px;color:#dc2626;">- Rp ${totalOut.toLocaleString('id-ID')}</td>
        <td colspan="3"></td>
      </tr>
      <tr class="footer-row">
        <td colspan="5" style="text-align:right;padding:10px 8px;color:#374151;">SISA SALDO AKHIR</td>
        <td style="text-align:right;padding:10px 8px;color:${(period.saldoAwal + totalIn - totalOut) >= 0 ? '#16a34a' : '#dc2626'};font-size:15px;">Rp ${(period.saldoAwal + totalIn - totalOut).toLocaleString('id-ID')}</td>
        <td colspan="3"></td>
      </tr>
    </tfoot>
  </table>

  ${proofSection}

  <div class="sign-section">
    <div class="sign-box">
      <div class="sign-line">Dibuat Oleh</div>
      <div class="sign-role">(${currentUser.nama})</div>
    </div>
    <div class="sign-box">
      <div class="sign-line">Diperiksa Oleh</div>
      <div class="sign-role">( _________________ )</div>
    </div>
    <div class="sign-box">
      <div class="sign-line">Disetujui Oleh</div>
      <div class="sign-role">( _________________ )</div>
    </div>
  </div>

  <div class="no-print" style="position:fixed;bottom:20px;right:20px;display:flex;gap:10px;">
    <button onclick="window.print()" style="padding:12px 28px;background:#f59e0b;color:#000;border:none;border-radius:8px;font-size:14px;font-weight:800;cursor:pointer;box-shadow:0 4px 12px rgba(245,158,11,0.4);">🖨️ Print / Save PDF</button>
    <button onclick="window.close()" style="padding:12px 20px;background:#6b7280;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;">✕ Tutup</button>
  </div>
</body>
</html>`);
      printWin.document.close();
    }

    // ============================================================
    // CETAK BUKTI FOTO PETTY CASH — Dikelompokkan Per Kategori
    // ============================================================

    // Helper: ekstrak Google Drive file ID dari berbagai format URL
    function extractDriveFileId(url) {
      if (!url) return null;
      const m1 = url.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
      if (m1) return m1[1];
      const m2 = url.match(/[?&]id=([a-zA-Z0-9_-]{20,})/);
      return m2 ? m2[1] : null;
    }

    function cetakBuktiFotoPettyCash() {
      if (!hasPCPermission('cetakDokumenPettyCash')) return toast('❌ Anda tidak memiliki akses untuk mencetak bukti foto', 'error');
      if (!pcActivePeriodId) return toast('Pilih periode terlebih dahulu', 'warning');
      const period = pcData.periods.find(p => p.id === pcActivePeriodId);
      if (!period) return;

      const txs = pcData.transactions
        .filter(t => t.periodId === pcActivePeriodId)
        .sort((a, b) => new Date(a.tanggal) - new Date(b.tanggal));

      const withBukti = txs.filter(d => d.buktiUrl);
      if (!withBukti.length) return toast('Tidak ada bukti foto di periode ini', 'warning');

      // Kumpulkan semua Drive file IDs yang perlu di-fetch
      const fileIdMap = {}; // fileId -> buktiUrl
      withBukti.forEach(d => {
        const fid = extractDriveFileId(d.buktiUrl);
        if (fid) fileIdMap[fid] = d.buktiUrl;
      });
      const fileIds = Object.keys(fileIdMap);

      // Tampilkan loading indicator
      const btn = document.querySelector('[onclick="cetakBuktiFotoPettyCash()"]');
      if (btn) { btn.disabled = true; btn.textContent = '⏳ Memuat foto...'; }
      toast('⏳ Mengambil foto dari Drive, mohon tunggu...', 'info');

      // Fetch semua gambar sebagai base64 dari server
      google.script.run
        .withSuccessHandler(function (res) {
          if (btn) { btn.disabled = false; btn.textContent = '📸 Cetak Bukti Foto'; }
          const imageMap = (res && res.images) ? res.images : {};
          _openPrintWindowBuktiFoto(period, txs, withBukti, imageMap);
        })
        .withFailureHandler(function (err) {
          if (btn) { btn.disabled = false; btn.textContent = '📸 Cetak Bukti Foto'; }
          toast('⚠️ Gagal memuat foto: ' + err.message, 'error');
        })
        .getFilesAsBase64(fileIds);
    }

    function _openPrintWindowBuktiFoto(period, txs, withBukti, imageMap) {
      // Group by kategori, sorted alphabetically
      const grouped = {};
      withBukti.forEach(d => {
        const k = d.kategori || 'Lainnya';
        if (!grouped[k]) grouped[k] = [];
        grouped[k].push(d);
      });
      const sortedKats = Object.keys(grouped).sort();

      // Category color palette (cycles)
      const catColors = [
        { bg: '#1e3a5f', light: '#eff6ff', border: '#bfdbfe', text: '#1d4ed8' },
        { bg: '#065f46', light: '#f0fdf4', border: '#bbf7d0', text: '#16a34a' },
        { bg: '#7c2d12', light: '#fff7ed', border: '#fed7aa', text: '#ea580c' },
        { bg: '#4c1d95', light: '#f5f3ff', border: '#ddd6fe', text: '#7c3aed' },
        { bg: '#831843', light: '#fdf2f8', border: '#f9a8d4', text: '#db2777' },
        { bg: '#1e40af', light: '#eff6ff', border: '#93c5fd', text: '#2563eb' },
        { bg: '#14532d', light: '#f0fdf4', border: '#86efac', text: '#15803d' },
        { bg: '#78350f', light: '#fffbeb', border: '#fde68a', text: '#d97706' },
      ];

      const categoryBlocks = sortedKats.map((kat, ki) => {
        const col = catColors[ki % catColors.length];
        const items = grouped[kat];

        const cards = items.map(d => {
          const txNo = txs.indexOf(d) + 1;
          const fileId = extractDriveFileId(d.buktiUrl);

          // Gunakan base64 data URI jika tersedia, fallback ke URL asli
          let mediaHtml;
          if (fileId && imageMap[fileId] && imageMap[fileId].dataUri) {
            // Gambar sudah di-embed sebagai base64 — pasti tampil saat print
            mediaHtml = `<div class="photo-wrap">
              <img src="${imageMap[fileId].dataUri}" class="photo-img" alt="Bukti #${txNo}">
            </div>`;
          } else if (d.buktiUrl) {
            // Fallback: tampilkan link jika tidak bisa di-fetch
            mediaHtml = `<div class="photo-link-wrap">
              <span style="font-size:28px;">📎</span>
              <span style="font-size:11px;color:#6b7280;">File tidak dapat dimuat otomatis</span>
              <a href="${d.buktiUrl}" target="_blank" class="photo-link">${d.buktiUrl}</a>
            </div>`;
          } else {
            mediaHtml = `<div class="photo-link-wrap"><span style="color:#9ca3af;font-size:12px;">Tidak ada bukti</span></div>`;
          }

          return `<div class="photo-card">
            <div class="photo-card-header">
              <div class="photo-card-num" style="background:${col.light};color:${col.text};border:1px solid ${col.border};">#${txNo}</div>
              <div class="photo-card-date">${new Date(d.tanggal).toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' })}</div>
            </div>
            <div class="photo-card-ket">${d.keterangan}</div>
            <div class="photo-card-nom" style="color:${d.tipe === 'OUT' ? '#dc2626' : '#16a34a'};">
              ${d.tipe === 'OUT' ? '−' : '+'}Rp ${d.nominal.toLocaleString('id-ID')}
              <span class="photo-card-by">· ${d.createdBy}</span>
            </div>
            ${mediaHtml}
          </div>`;
        }).join('');

        return `<section class="cat-section">
          <div class="cat-header" style="background:${col.bg};">
            <span class="cat-icon">📂</span>
            <span class="cat-name">${kat}</span>
            <span class="cat-count">${items.length} bukti</span>
          </div>
          <div class="photo-grid">
            ${cards}
          </div>
        </section>`;
      }).join('');

      // Summary per kategori
      const summaryRows = sortedKats.map((kat, ki) => {
        const col = catColors[ki % catColors.length];
        const items = grouped[kat];
        const total = items.reduce((s, d) => d.tipe === 'OUT' ? s + d.nominal : s - d.nominal, 0);
        return `<tr>
          <td style="padding:8px 12px;">
            <span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${col.bg};margin-right:6px;"></span>
            <strong>${kat}</strong>
          </td>
          <td style="padding:8px 12px;text-align:center;">${items.length}</td>
          <td style="padding:8px 12px;text-align:right;color:${total >= 0 ? '#dc2626' : '#16a34a'};font-weight:700;">
            Rp ${Math.abs(total).toLocaleString('id-ID')}
          </td>
        </tr>`;
      }).join('');

      const printWin = window.open('', '_blank', 'width=1100,height=900,scrollbars=yes');
      if (!printWin) return toast('Pop-up diblokir browser! Izinkan popup untuk website ini.', 'error');

      printWin.document.write(`<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <title>Bukti Foto Petty Cash — ${period.nama}</title>
  <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;600;700;800&display=swap" rel="stylesheet">
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Outfit', Arial, sans-serif;
      background: #f8fafc;
      color: #111827;
      padding: 32px;
      font-size: 13px;
    }

    /* ── HEADER ── */
    .doc-header {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      margin-bottom: 24px;
      padding-bottom: 16px;
      border-bottom: 4px solid #f59e0b;
    }
    .doc-title { font-size: 24px; font-weight: 800; color: #1e3a5f; }
    .doc-period { font-size: 13px; color: #6b7280; margin-top: 4px; }
    .doc-meta { text-align: right; font-size: 12px; color: #6b7280; line-height: 1.7; }
    .doc-meta strong { color: #1e3a5f; font-size: 15px; display: block; }

    /* ── SUMMARY TABLE ── */
    .summary-wrap {
      background: #fff;
      border: 1px solid #e5e7eb;
      border-radius: 10px;
      padding: 16px 20px;
      margin-bottom: 28px;
    }
    .summary-title {
      font-size: 13px;
      font-weight: 800;
      color: #1e3a5f;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      margin-bottom: 12px;
    }
    .summary-table { width: 100%; border-collapse: collapse; }
    .summary-table th {
      background: #f3f4f6;
      padding: 8px 12px;
      font-size: 11px;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      color: #6b7280;
      text-align: left;
    }
    .summary-table th:nth-child(2) { text-align: center; }
    .summary-table th:nth-child(3) { text-align: right; }
    .summary-table td { border-top: 1px solid #f3f4f6; font-size: 12px; }
    .summary-total td {
      border-top: 2px solid #d1d5db;
      font-weight: 800;
      background: #f9fafb;
      padding: 8px 12px;
    }

    /* ── CATEGORY SECTION ── */
    .cat-section { margin-bottom: 36px; }
    .cat-header {
      display: flex;
      align-items: center;
      gap: 10px;
      padding: 12px 18px;
      border-radius: 10px 10px 0 0;
      color: #fff;
    }
    .cat-icon { font-size: 18px; }
    .cat-name { font-size: 15px; font-weight: 800; letter-spacing: 0.3px; flex: 1; }
    .cat-count {
      background: rgba(255,255,255,0.22);
      padding: 3px 12px;
      border-radius: 20px;
      font-size: 12px;
      font-weight: 700;
    }

    /* ── PHOTO GRID ── */
    .photo-grid {
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 20px;
      padding: 16px;
      background: #fff;
      border: 1px solid #e5e7eb;
      border-top: none;
      border-radius: 0 0 10px 10px;
    }

    /* ── PHOTO CARD ── */
    .photo-card {
      border: 1px solid #e5e7eb;
      border-radius: 10px;
      padding: 14px;
      background: #fafafa;
      page-break-inside: avoid;
      break-inside: avoid;
    }
    .photo-card-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 6px;
    }
    .photo-card-num {
      font-size: 11px;
      font-weight: 800;
      padding: 3px 10px;
      border-radius: 4px;
    }
    .photo-card-date { font-size: 11px; color: #9ca3af; }
    .photo-card-ket {
      font-size: 13px;
      font-weight: 700;
      color: #111827;
      margin-bottom: 4px;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .photo-card-nom {
      font-size: 12px;
      font-weight: 700;
      margin-bottom: 10px;
    }
    .photo-card-by { font-size: 11px; color: #9ca3af; font-weight: 400; }

    /* ── PHOTO IMAGE ── */
    .photo-wrap { position: relative; }
    .photo-img {
      width: 100%;
      max-height: 420px;
      object-fit: contain;
      border-radius: 6px;
      border: 1px solid #e5e7eb;
      display: block;
      background: #f9fafb;
    }
    .photo-err {
      display: none;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      gap: 4px;
      padding: 20px 10px;
      background: #f9fafb;
      border-radius: 6px;
      border: 1px dashed #d1d5db;
      text-align: center;
      font-size: 11px;
      color: #9ca3af;
    }
    .photo-err a { color: #0369a1; font-size: 10px; }
    .photo-link-wrap {
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 6px;
      padding: 16px 8px;
      background: #f0f9ff;
      border-radius: 6px;
      border: 1px solid #bae6fd;
      text-align: center;
    }
    .photo-link {
      color: #0369a1;
      font-size: 10px;
      word-break: break-all;
      font-weight: 600;
    }

    /* ── FOOTER ── */
    .doc-footer {
      margin-top: 40px;
      padding-top: 16px;
      border-top: 2px solid #e5e7eb;
      display: flex;
      justify-content: space-between;
      align-items: center;
      font-size: 11px;
      color: #9ca3af;
    }

    /* ── PRINT BUTTON ── */
    .no-print {
      position: fixed;
      bottom: 20px;
      right: 20px;
      display: flex;
      gap: 10px;
      z-index: 999;
    }

    /* ── PRINT MEDIA ── */
    @media print {
      body { background: #fff; padding: 12px; }
      .no-print { display: none !important; }
      .photo-grid { grid-template-columns: repeat(2, 1fr); }
      @page { margin: 1cm; size: A4; }
    }
  </style>
</head>
<body>

  <!-- HEADER -->
  <div class="doc-header">
    <div>
      <div class="doc-title">📸 BUKTI FOTO PETTY CASH</div>
      <div class="doc-period">${period.nama}</div>
      <div class="doc-period">Periode: ${new Date(period.tanggalMulai).toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' })} s/d ${new Date(period.tanggalSelesai).toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' })}</div>
    </div>
    <div class="doc-meta">
      <strong>GUDANG FCL GROUP</strong>
      Dicetak: ${new Date().toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' })}<br>
      Oleh: ${currentUser.nama}<br>
      Total Bukti: <strong style="color:#1e3a5f;font-size:13px;">${withBukti.length} foto</strong>
    </div>
  </div>

  <!-- SUMMARY PER KATEGORI -->
  <div class="summary-wrap">
    <div class="summary-title">📊 Ringkasan Per Kategori</div>
    <table class="summary-table">
      <thead>
        <tr>
          <th>Kategori</th>
          <th style="text-align:center;">Jumlah Bukti</th>
          <th style="text-align:right;">Total Nominal</th>
        </tr>
      </thead>
      <tbody>
        ${summaryRows}
      </tbody>
      <tfoot>
        <tr class="summary-total">
          <td style="padding:8px 12px;">TOTAL</td>
          <td style="padding:8px 12px;text-align:center;">${withBukti.length}</td>
          <td style="padding:8px 12px;text-align:right;color:#dc2626;">
            Rp ${withBukti.filter(d => d.tipe === 'OUT').reduce((s, d) => s + d.nominal, 0).toLocaleString('id-ID')}
          </td>
        </tr>
      </tfoot>
    </table>
  </div>

  <!-- FOTO PER KATEGORI -->
  ${categoryBlocks}

  <!-- FOOTER -->
  <div class="doc-footer">
    <span>GUDANG FCL GROUP &nbsp;·&nbsp; Petty Cash: ${period.nama}</span>
    <span>Dicetak ${new Date().toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' })} oleh ${currentUser.nama}</span>
  </div>

  <!-- PRINT BUTTON -->
  <div class="no-print">
    <button onclick="window.print()" style="padding:12px 28px;background:#f59e0b;color:#000;border:none;border-radius:8px;font-size:14px;font-weight:800;cursor:pointer;box-shadow:0 4px 12px rgba(245,158,11,0.4);">🖨️ Print / Save PDF</button>
    <button onclick="window.close()" style="padding:12px 20px;background:#6b7280;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;">✕ Tutup</button>
  </div>

</body>
</html>`);
      printWin.document.close();
    } // end _openPrintWindowBuktiFoto

    function loadKaryawan() {
      // Menggunakan Batch Request untuk performa maksimal (1 Round Trip ke Server)
      google.script.run.withSuccessHandler(res => {
        if (!res.success) { toast('Gagal memuat data SDM', 'error'); return; }

        // 1. Data Ijin
        if (res.ijin && res.ijin.success) ijinDataAll = res.ijin.data;

        // 2. Data Riwayat Resign
        if (res.riwayat && res.riwayat.success) {
          riwayatData = res.riwayat.data;
          renderRiwayatKaryawan();
        }

        // 3. Data Surat Peringatan (SP)
        if (res.sp && res.sp.success) {
          spData = res.sp.data;
          renderSP();
        }

        // 4. Data Karyawan Aktif
        if (res.karyawan && res.karyawan.success) {
          karyawanData = res.karyawan.data;
          renderKaryawan(karyawanData);
          try {
            updateKaryawanStats();
            renderKaryawanWarehouseStats();
          } catch (e) {
            console.error('Stats update error:', e);
          }
          populateSPKaryawanSelect();
        }
      }).getKaryawanFullData();
    }
    function updateKaryawanStats() {
      if (!Array.isArray(karyawanData)) karyawanData = [];
      if (!Array.isArray(riwayatData)) riwayatData = [];
      if (!Array.isArray(ijinDataAll)) ijinDataAll = [];
      if (!Array.isArray(spData)) spData = [];

      try {
        setVal('statTotalKar', karyawanData.length);
        setVal('statTetap', karyawanData.filter(d => d.status === 'Tetap').length);
        setVal('statKontrak', karyawanData.filter(d => d.status === 'Kontrak').length);

        const now = new Date();
        const in40 = karyawanData.filter(d => {
          if (d.status !== 'Kontrak' || !d.tanggalSelesai) return false;
          const diff = (new Date(d.tanggalSelesai) - now) / 86400000;
          return diff >= 0 && diff <= 40;
        });
        setVal('statKontrakHabis', in40.length);

        const wrap = document.getElementById('kontrakWarningWrap');
        const list = document.getElementById('kontrakWarningList');
        if (wrap && list) {
          if (in40.length) {
            wrap.style.display = 'block';
            list.innerHTML = '';
            in40.forEach(d => list.innerHTML += `<div style="display:flex;justify-content:space-between;background:#ef444415;padding:10px;border-radius:8px"><div><strong>${d.nama}</strong> <small>(${d.jabatan})</small></div><div style="text-align:right"><small style="color:var(--red)">${formatDate(d.tanggalSelesai)}</small><br><strong style="color:var(--red)">${Math.ceil((new Date(d.tanggalSelesai) - now) / 86400000)} hr lagi</strong></div></div>`);
          } else wrap.style.display = 'none';
        }

        // STATS IJIN HARI INI
        const tToday = now.toISOString().split('T')[0];
        const iToday = ijinDataAll.filter(d => d.tanggal === tToday && (d.status || '') === 'Disetujui');
        setVal('statIjinToday', iToday.filter(d => {
          const j = (d.jenis || '').toLowerCase();
          return !j.includes('sakit') && !j.includes('cuti');
        }).length);
        setVal('statCutiToday', iToday.filter(d => (d.jenis || '').toLowerCase().includes('cuti')).length);
        setVal('statSakitToday', iToday.filter(d => (d.jenis || '').toLowerCase().includes('sakit')).length);

        const iW = document.getElementById('ijinTodayWrap'), iL = document.getElementById('ijinTodayList');
        if (iW && iL) {
          if (iToday.length) {
            iW.style.display = 'block'; iL.innerHTML = '';
            ['Sakit', 'Cuti', 'Ijin'].forEach(cat => {
              const filtered = iToday.filter(d => {
                const j = (d.jenis || '').toLowerCase();
                if (cat === 'Sakit') return j.includes('sakit');
                if (cat === 'Cuti') return j.includes('cuti');
                return !j.includes('sakit') && !j.includes('cuti');
              });
              if (filtered.length) {
                let color = cat === 'Sakit' ? 'var(--red)' : (cat === 'Cuti' ? 'var(--orange)' : 'var(--blue)');
                filtered.forEach(d => {
                  iL.innerHTML += `<div style="display:flex;align-items:center;gap:10px;background:#ffffff0a;padding:10px;border-radius:10px;border-left:4px solid ${color};">
                    <div style="flex:1;">
                      <div style="font-weight:700;font-size:14px;color:var(--white);">${d.nama}</div>
                      <div style="font-size:11px;color:var(--gray);">${d.keterangan || 'Tanpa keterangan'}</div>
                    </div>
                    <span class="badge-tb" style="background:${color}22; color:${color}; border-color:${color}44;">${cat}</span>
                  </div>`;
                });
              }
            });
          } else iW.style.display = 'none';
        }

        // Tab Counts
        setVal('countKarAktif', karyawanData.length);
        setVal('countKarRiwayat', riwayatData.length);
        setVal('countKarSP', spData.filter(d => (d.status || '') === 'Aktif').length);

        setVal('statTotalRiwayat', riwayatData.length);
        setVal('statTotalSP', spData.filter(d => (d.status || '') === 'Aktif').length);

        const spAktif = spData.filter(d => (d.status || '') === 'Aktif');
        setVal('statSPTotal', spAktif.length);
        setVal('statSP1', spAktif.filter(d => d.jenisSP === 'SP 1').length);
        setVal('statSP2', spAktif.filter(d => d.jenisSP === 'SP 2').length);
        setVal('statSP3', spAktif.filter(d => d.jenisSP === 'SP 3').length);

        renderKaryawanWarehouseStats();
      } catch (err) {
        console.error('Error in updateKaryawanStats:', err);
      }
    }
    function renderKaryawanWarehouseStats() {
      try {
        const cont = document.getElementById('karWarehouseStats');
        if (!cont) return;

        let kData = Array.isArray(karyawanData) ? karyawanData : [];
        let rData = Array.isArray(riwayatData) ? riwayatData : [];

        // Filter active only
        const activeData = kData.filter(d => !rData.some(r => String(r.id) === String(d.id)));

        const stats = {};
        activeData.forEach(d => {
          const loc = String(d.cabang || 'Lainnya').trim() || 'Lainnya';
          const div = String(d.jabatan || 'Staf').trim() || 'Staf';
          if (!stats[loc]) stats[loc] = { total: 0, divs: {} };
          stats[loc].total++;
          stats[loc].divs[div] = (stats[loc].divs[div] || 0) + 1;
        });

        if (Object.keys(stats).length === 0) {
          cont.innerHTML = `<div style="grid-column: 1/-1; text-align:center; padding:20px; border:1px dashed var(--border-color); border-radius:12px; color:var(--text-muted);">
            🏢 Belum ada data lokasi gudang (Surabaya / Jakarta). Isi kolom "Lokasi" pada data karyawan.
          </div>`;
          return;
        }

        let html = '';
        Object.keys(stats).sort().forEach(loc => {
          const s = stats[loc];
          const divsHtml = Object.entries(s.divs)
            .sort((a, b) => b[1] - a[1])
            .map(([d, c]) => `<div class="div-pill"><span>${d}</span><strong>${c}</strong></div>`)
            .join('');

          html += `<div class="warehouse-card">
            <div class="warehouse-header">
              <div class="warehouse-name">🏗️ ${loc}</div>
              <div class="warehouse-total">${s.total} Org</div>
            </div>
            <div class="warehouse-divisions">${divsHtml}</div>
          </div>`;
        });
        cont.innerHTML = html;
      } catch (err) {
        console.error('Render Warehouse Stats Fail:', err);
      }
    }

    function renderKaryawan(data) {
      const tb = document.getElementById('tableKaryawan'); tb.innerHTML = ''; if (!data.length) { tb.innerHTML = '<tr><td colspan="10" class="empty-state">Kosong</td></tr>'; return; }
      const now = new Date();
      let userPerms = []; try { userPerms = JSON.parse(currentUser.permissions || '[]'); } catch (e) { }
      const canEdit = currentUser.role === 'admin' || userPerms.includes('editKaryawan');
      const canDelete = currentUser.role === 'admin';

      data.forEach(d => {
        let sCell = '-'; if (d.status === 'Kontrak' && d.tanggalSelesai) { const diff = Math.ceil((new Date(d.tanggalSelesai) - now) / 86400000); sCell = `<span style="${diff <= 40 ? 'color:var(--red);font-weight:bold' : ''}">${formatDate(d.tanggalSelesai)} ${diff <= 40 && diff >= 0 ? `(${diff}hr)` : ''}</span>`; }
        const spAktif = spData.filter(s => s.karyawanId === d.id && s.status === 'Aktif').sort((a, b) => b.jenisSP.localeCompare(a.jenisSP))[0];
        let spBadge = '-';
        if (spAktif) {
          const cls = spAktif.jenisSP.replace(' ', '').toLowerCase();
          spBadge = `<span class="sp-badge ${cls}" style="font-size:10px; padding:2px 6px;">${spAktif.jenisSP}</span>`;
        }

        const editBtn = canEdit ? `<button class="btn btn-ghost btn-sm" onclick="editKaryawan('${d.id}')" title="Edit Data">✏️</button>` : '';
        const cardBtn = canEdit ? `<button class="btn btn-ghost btn-sm" onclick="printEmployeeCard('${d.id}')" title="Cetak ID Card QR" style="color:var(--teal)">📇</button>` : '';
        const spBtn = canEdit ? `<button class="btn btn-ghost btn-sm" onclick="openSPModal('${d.id}')" title="Beri SP" style="color:var(--amber)">⚠️</button>` : '';
        const resBtn = canEdit ? `<button class="btn btn-danger btn-sm" onclick="openResignModal('${d.id}')" title="Karyawan Resign" style="padding:2px 6px; font-size:10px;">👋 Resign</button>` : '';
        const delBtn = canDelete ? `<button class="btn btn-danger btn-sm" onclick="delKaryawan('${d.id}')" title="Hapus Permanen">🗑️</button>` : '';

        tb.innerHTML += `<tr>
          <td><strong>${d.nama}</strong></td>
          <td>${d.jabatan}</td>
          <td><span class="badge-tb">${d.cabang || '-'}</span></td>
          <td><span class="${d.status === 'Tetap' ? 'badge-aktif' : 'badge-nonaktif'}">${d.status}</span></td>
          <td><span class="badge-tb">${d.fingerprintId || '-'}</span></td>
          <td>${formatDate(d.tanggalMasuk)}</td>
          <td>${sCell}</td>
          <td><strong style="color:var(--teal)">${d.sisaCuti || 0} hr</strong></td>
          <td>${spBadge}</td>
          <td style="display:flex;gap:4px;">${editBtn} ${cardBtn} ${spBtn} ${resBtn} ${delBtn}</td>
        </tr>`;
      });
    }

    function switchKarTab(tab) {
      document.querySelectorAll('.kar-tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.kar-tab-panel').forEach(p => p.classList.remove('active'));

      const btn = document.querySelector(`.kar-tab[onclick*="'${tab}'"]`);
      const panel = document.getElementById(`panel-kar-${tab}`);
      if (btn) btn.classList.add('active');
      if (panel) panel.classList.add('active');
    }

    // RIWAYAT KARYAWAN
    function renderRiwayatKaryawan() {
      const tb = document.getElementById('tableRiwayatKaryawan'); tb.innerHTML = '';
      const q = v('riwayatSearch').toLowerCase();
      const filtered = riwayatData.filter(d => d.nama.toLowerCase().includes(q) || d.jabatan.toLowerCase().includes(q));

      if (!filtered.length) { tb.innerHTML = '<tr><td colspan="7" class="empty-state">Belum ada riwayat resign</td></tr>'; return; }
      filtered.sort((a, b) => new Date(b.tanggalResign) - new Date(a.tanggalResign)).forEach(d => {
        tb.innerHTML += `<tr>
          <td><strong>${d.nama}</strong></td>
          <td>${d.jabatan}</td>
          <td><span class="badge-tb">${d.cabang}</span></td>
          <td>${formatDate(d.tanggalMasuk)}</td>
          <td style="color:var(--red)">${formatDate(d.tanggalResign)}</td>
          <td><small>${d.alasanResign}</small></td>
          <td><button class="btn btn-danger btn-sm" onclick="delRiwayat('${d.id}')">🗑️</button></td>
        </tr>`;
      });
    }

    function openResignModal(id) {
      const d = karyawanData.find(x => x.id === id); if (!d) return;
      setVal('resId', d.id); setVal('resNama', d.nama); setVal('resJabatan', d.jabatan);
      setVal('resTanggal', new Date().toISOString().split('T')[0]);
      setVal('resAlasan', 'Mengundurkan Diri');
      setVal('resKeterangan', '');
      openModal('modalResign');
    }

    function submitResign() {
      const id = v('resId'), t = v('resTanggal'), a = v('resAlasan'), k = v('resKeterangan');
      const d = karyawanData.find(x => x.id === id);
      if (!t || !a) return toast('Lengkapi form', 'error');
      if (!confirm(`Yakin tandai ${d.nama} sebagai Resign? Akun akan dinonaktifkan.`)) return;

      const btn = document.querySelector('#modalResign .btn-danger'); btn.disabled = true; btn.textContent = '⏳...';
      google.script.run.withSuccessHandler(res => {
        btn.disabled = false; btn.textContent = '💾 Simpan & Non-Aktifkan';
        if (res.success) { toast('Berhasil diproses'); closeModal('modalResign'); loadKaryawan(); }
        else toast(res.message, 'error');
      }).addRiwayatKaryawan(id, d.nama, d.jabatan, d.cabang, d.telepon, d.tanggalMasuk, t, a, k, currentUser.username);
    }

    function delRiwayat(id) { if (confirm('Hapus dari riwayat?')) google.script.run.withSuccessHandler(res => { if (res.success) { toast('Dihapus'); loadKaryawan(); } }).deleteRiwayatKaryawan(id); }

    function downloadKaryawanTemplate() {
      const data = [
        ['ID (Kosongkan jika baru)', 'Nama Lengkap', 'Jabatan', 'Cabang/Lokasi', 'Telepon', 'Email', 'Tanggal Masuk (YYYY-MM-DD)', 'Status (Tetap/Kontrak)', 'Selesai Kontrak (YYYY-MM-DD)', 'Sisa Cuti', 'Fingerprint ID'],
        ['', 'Budi Santoso', 'Staff Gudang', 'Jakarta', '08123456789', 'budi@example.com', '2024-01-01', 'Tetap', '', '12', '001'],
        ['', 'Siti Aminah', 'Admin', 'Surabaya', '08129876543', 'siti@example.com', '2024-02-15', 'Kontrak', '2025-02-15', '12', '002']
      ];
      const ws = XLSX.utils.aoa_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Template Karyawan");
      XLSX.writeFile(wb, "Template_Data_Karyawan.xlsx");
    }

    function handleImportKaryawan(input) {
      if (!input.files.length) return;
      const file = input.files[0];
      const reader = new FileReader();
      reader.onload = e => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet);

        if (!rows.length) return toast('File kosong', 'error');

        const items = rows.map(r => ({
          id: r['ID (Kosongkan jika baru)'] || '',
          nama: r['Nama Lengkap'] || '',
          jabatan: r['Jabatan'] || '',
          cabang: r['Cabang/Lokasi'] || '',
          telepon: r['Telepon'] || '',
          email: r['Email'] || '',
          tanggalMasuk: r['Tanggal Masuk (YYYY-MM-DD)'] || '',
          status: r['Status (Tetap/Kontrak)'] || 'Tetap',
          tanggalSelesai: r['Selesai Kontrak (YYYY-MM-DD)'] || '',
          sisaCuti: parseInt(r['Sisa Cuti'] || 12),
          fingerprintId: r['Fingerprint ID'] || ''
        })).filter(x => x.nama);

        if (!items.length) return toast('Tidak ada data valid', 'error');

        google.script.run.withSuccessHandler(res => {
          if (res.success) {
            toast('Import berhasil');
            loadKaryawan();
          } else toast(res.message, 'error');
        }).addBulkKaryawan(items);
      };
      reader.readAsArrayBuffer(file);
    }

    // SURAT PERINGATAN (SP)
    function populateSPKaryawanSelect() {
      const sel = document.getElementById('spNamaSelect'); sel.innerHTML = '<option value="">-- Pilih Karyawan --</option>';
      karyawanData.sort((a, b) => a.nama.localeCompare(b.nama)).forEach(d => {
        sel.innerHTML += `<option value="${d.id}" data-nama="${d.nama}">${d.nama} (${d.jabatan})</option>`;
      });
    }

    function setSPMasaBerlaku(days, btn) {
      setVal('spMasaBerlaku', days);
      document.querySelectorAll('.mb-preset').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      calcSPKadaluarsa();
    }

    function calcSPKadaluarsa() {
      const tgl = v('spTanggal'), masa = parseInt(v('spMasaBerlaku')) || 0;
      if (!tgl || masa <= 0) return setVal('spTglKadaluarsa', '-');
      const d = new Date(tgl); d.setDate(d.getDate() + masa);
      setVal('spTglKadaluarsa', formatDate(d.toISOString().split('T')[0]));
    }

    function renderSP() {
      const tb = document.getElementById('tableSP'); tb.innerHTML = '';
      const q = v('spSearch').toLowerCase();
      const filtered = spData.filter(d => d.karyawanNama.toLowerCase().includes(q) || d.jenisSP.toLowerCase().includes(q));

      if (!filtered.length) { tb.innerHTML = '<tr><td colspan="8" class="empty-state">Belum ada SP yang diterbitkan</td></tr>'; return; }
      filtered.sort((a, b) => new Date(b.tanggalSP) - new Date(a.tanggalSP)).forEach(d => {
        const cls = d.jenisSP.replace(' ', '').toLowerCase();
        const pct = Math.max(0, Math.min(100, (d.sisaHari / d.masaBerlaku) * 100));
        const barColor = d.status === 'Kadaluarsa' ? '#94a3b8' : (pct < 20 ? 'var(--red)' : (pct < 50 ? 'var(--accent)' : 'var(--green)'));

        tb.innerHTML += `<tr>
          <td><strong>${d.karyawanNama}</strong></td>
          <td><span class="sp-badge ${cls}">${d.jenisSP}</span></td>
          <td><small style="display:block; max-width:200px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${d.alasan}">${d.alasan}</small></td>
          <td>${formatDate(d.tanggalSP)}</td>
          <td>${d.masaBerlaku} hari</td>
          <td style="color:${d.status === 'Kadaluarsa' ? 'var(--gray)' : 'var(--red)'}">${formatDate(d.tanggalKadaluarsa)}</td>
          <td>
            <div class="sp-progress-wrap">
              <span class="sp-badge ${d.status === 'Kadaluarsa' ? 'kadaluarsa' : 'aktif'}">${d.status}</span>
              ${d.status === 'Aktif' ? `
              <div class="sp-progress-bar-bg"><div class="sp-progress-bar" style="width:${pct}%; background:${barColor}"></div></div>
              <div class="sp-sisa-label" style="color:${barColor}">${d.sisaHari} hari lagi</div>
              ` : ''}
            </div>
          </td>
          <td><button class="btn btn-danger btn-sm" onclick="delSP('${d.id}')">🗑️</button></td>
        </tr>`;
      });
    }

    function openSPModal(karId) {
      // Pastikan list karyawan di dropdown sudah yang terbaru
      populateSPKaryawanSelect();

      setVal('spId', '');
      setVal('spNamaSelect', karId || '');
      setVal('spJenis', 'SP 1');
      setVal('spTanggal', new Date().toISOString().split('T')[0]);
      setVal('spMasaBerlaku', 180);
      setVal('spAlasan', '');

      document.querySelectorAll('.mb-preset').forEach(b => b.textContent === '6 Bulan' ? b.classList.add('active') : b.classList.remove('active'));
      calcSPKadaluarsa();

      if (karId) switchKarTab('sp');
      openModal('modalSP');
    }

    function submitSP() {
      const id = v('spNamaSelect'), j = v('spJenis'), t = v('spTanggal'), m = v('spMasaBerlaku'), a = v('spAlasan');
      if (!id || !t || !m || !a) return toast('Lengkapi data', 'error');
      const nama = document.querySelector(`#spNamaSelect option[value="${id}"]`).getAttribute('data-nama');

      const btn = document.querySelector('#modalSP .btn-primary'); btn.disabled = true; btn.textContent = '⏳...';
      google.script.run.withSuccessHandler(res => {
        btn.disabled = false; btn.textContent = '💾 Simpan & Terbitkan SP';
        if (res.success) { toast('SP Diterbitkan'); closeModal('modalSP'); loadKaryawan(); }
        else toast(res.message, 'error');
      }).addSuratPeringatan(nama, id, j, a, t, m, currentUser.username);
    }

    function delSP(id) { if (confirm('Hapus SP ini?')) google.script.run.withSuccessHandler(res => { if (res.success) { toast('Dihapus'); loadKaryawan(); } }).deleteSuratPeringatan(id); }

    function filterKaryawan() { const q = v('karyawanSearch').toLowerCase(); renderKaryawan(karyawanData.filter(d => d.nama.toLowerCase().includes(q) || (d.jabatan || '').toLowerCase().includes(q))); }
    function toggleTglSelesai() { document.getElementById('wrapTglSelesai').style.display = v('kStatus') === 'Kontrak' ? 'block' : 'none'; }
    function openKaryawanModal() { setVal('karyawanId', ''); document.getElementById('karyawanModalTitle').textContent = '👤 Tambah Karyawan'; resetForm(['kNama', 'kJabatan', 'kCabang', 'kTelepon', 'kEmail', 'kTglSelesai', 'kFingerprintId']); setVal('kSisaCuti', 12); setVal('kStatus', 'Kontrak'); toggleTglSelesai(); openModal('modalKaryawan'); }
    function editKaryawan(id) { const d = karyawanData.find(x => x.id === id); if (!d) return; setVal('karyawanId', d.id); document.getElementById('karyawanModalTitle').textContent = '✏️ Edit Karyawan';['Nama', 'Jabatan', 'Cabang', 'Telepon', 'Email', 'TglMasuk', 'Status', 'TglSelesai', 'FingerprintId'].forEach(f => setVal('k' + f, d[f.charAt(0).toLowerCase() + f.slice(1)] || '')); setVal('kSisaCuti', d.sisaCuti || 0); toggleTglSelesai(); openModal('modalKaryawan'); }
    function submitKaryawan() {
      const id = v('karyawanId'), n = v('kNama'), j = v('kJabatan'), cbg = v('kCabang'), tlp = v('kTelepon'), em = v('kEmail'), tm = v('kTglMasuk'), st = v('kStatus'), ts = v('kTglSelesai'), sc = v('kSisaCuti'), fp = v('kFingerprintId');
      if (!n || !j) return toast('Nama & Jabatan wajib', 'error');
      const btn = document.querySelector('#modalKaryawan .btn-primary'); btn.disabled = true; btn.textContent = '⏳...';
      const cb = res => { btn.disabled = false; btn.textContent = '💾 Simpan'; if (res.success) { toast('Berhasil'); closeModal('modalKaryawan'); loadKaryawan(); } else toast(res.message, 'error'); };
      if (id) google.script.run.withSuccessHandler(cb).updateKaryawan(id, n, j, cbg, tlp, em, tm, st, ts, sc, fp); else google.script.run.withSuccessHandler(cb).addKaryawan(n, j, cbg, tlp, em, tm, st, ts, sc, fp);
    }
    function delKaryawan(id) { if (confirm('Hapus?')) google.script.run.withSuccessHandler(res => { if (res.success) { toast('Dihapus'); loadKaryawan(); } else toast(res.message, 'error'); }).deleteKaryawan(id); }

    function loadIjin() {
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          ijinDataAll = res.data;
          const tb = document.getElementById('tableIjin'); tb.innerHTML = '';
          if (!res.data.length) { tb.innerHTML = '<tr><td colspan="8" class="empty-state">Kosong</td></tr>'; return; }

          const role = currentUser?.role || 'user';
          res.data.sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal)).forEach(d => {
            let stsCls = 'badge-pending';
            if (d.status === 'Disetujui') stsCls = 'badge-approved';
            else if (d.status === 'Ditolak') stsCls = 'badge-rejected';
            else if (d.status.includes('Pending')) stsCls = 'badge-pending-stage';

            let actionHtml = `<button class="btn btn-ghost btn-sm" onclick="showHistoryModal('${encodeURIComponent(d.history)}')">Riwayat</button>`;

            let userPerms = []; try { userPerms = JSON.parse(currentUser.permissions || '[]'); } catch (e) { }
            const escapedNamaIjin = (d.nama || '').replace(/'/g, "\\'");
            const tanggalStrIjin = d.tanggal || '';
            const canAppr = canApprove(d.status, role, 'Ijin/Cuti');

            if (canAppr) {
              actionHtml += `
                <button class="btn btn-teal btn-sm" style="padding:2px 8px; font-size:11px;" onclick="processApproval('Ijin','${d.id}','Approve','${escapedNamaIjin}','${tanggalStrIjin}')">Setuju</button>
                <button class="btn btn-danger btn-sm" style="padding:2px 8px; font-size:11px;" onclick="processApproval('Ijin','${d.id}','Reject','${escapedNamaIjin}','${tanggalStrIjin}')">Tolak</button>
              `;
            }

            if (currentUser.username === d.createdBy && d.status === 'Pending Team Leader') {
              actionHtml += `<button class="btn btn-danger btn-sm" onclick="delIjin('${d.id}')">Batal</button>`;
            }

            tb.innerHTML += `<tr>
              <td>${formatDate(d.tanggal)}</td>
              <td><strong>${d.nama}</strong></td>
              <td><span class="badge-tb">${d.jenis}</span></td>
              <td>${d.keterangan}</td>
              <td>${d.bukti ? `<a href="${d.bukti}" target="_blank">Lihat Bukti</a>` : '-'}</td>
              <td><span class="${stsCls}">${d.status}</span></td>
              <td style="display:flex; gap:5px; align-items:center;">${actionHtml}</td>
            </tr>`;
          });
        }
      }).getIjin();
    }
    function toggleBuktiIjin() { document.getElementById('ijBuktiWrap').style.display = v('ijJenis').includes('Sakit') ? 'block' : 'none'; }
    function submitIjin() {
      const t = v('ijTanggal'), n = v('ijNama'), j = v('ijJenis'), k = v('ijKeterangan'); if (!t || !n || !k) return toast('Lengkapi data', 'error');
      const btn = document.getElementById('btnSaveIjin'); btn.disabled = true; btn.textContent = '⏳...';
      const proceed = url => google.script.run.withSuccessHandler(res => { btn.disabled = false; btn.textContent = '💾 Ajukan Ijin'; if (res.success) { toast('Berhasil'); closeModal('modalIjin'); loadIjin(); resetForm(['ijKeterangan']); removeFile('ij'); } else toast(res.message, 'error'); }).addIjin(t, n, j, k, url, currentUser.username);
      if (j.includes('Sakit')) { const f = document.getElementById('ijFile').files[0] || window['_droppedFile_ij']; if (f) { window['_droppedFile_ij'] = null; uploadFileAndProceed('ij', f, 'Ijin', proceed, btn); } else proceed(v('ijBukti')); } else proceed('');
    }
    function delIjin(id) { if (confirm('Batalkan pengajuan?')) google.script.run.withSuccessHandler(res => { if (res.success) { toast('Dibatalkan'); loadIjin(); } else toast(res.message, 'error'); }).deleteIjin(id); }

    function loadLembur() {
      const isPimpinan = (currentUser.role === 'admin' || currentUser.role === 'HR' || currentUser.role === 'Supervisor' || currentUser.role === 'Vice Supervisor' || (currentUser.role || '').includes('Supervisor') || (currentUser.role || '').includes('HR'));
      if (!karyawanData.length && isPimpinan) loadKaryawan();

      switchLemburTab('list');

      // Loading Indicator (Sama kaya Karyawan/Approval)
      const tbLembur = document.getElementById('tableLembur');
      const tbBelum = document.getElementById('tableLemburBelum');
      if (tbLembur) {
        tbLembur.innerHTML = `
          <tr>
            <td colspan="8" style="text-align: center; padding: 20px; color: var(--text-muted);">
              <div style="font-size: 14px; font-weight: 600;">⚡ Memuat data lembur...</div>
            </td>
          </tr>`;
      }
      if (tbBelum) {
        tbBelum.innerHTML = `
          <tr>
            <td colspan="6" style="text-align: center; padding: 20px; color: var(--text-muted);">
              <div style="font-size: 14px; font-weight: 600;">⚡ Memuat data...</div>
            </td>
          </tr>`;
      }

      // Visibilitas tombol Pengaturan Tanggal Merah
      const role = (currentUser?.role || '').toLowerCase();
      let allowedMenus = []; try { allowedMenus = JSON.parse(currentUser.permissions || '[]'); } catch (e) { }
      const canManageTM = (role === 'admin' || role.includes('team leader') || role === 'tl' || role.includes('hr') || allowedMenus.includes('pengaturanTglMerah'));
      const btnTM = document.getElementById('btnTglMerahSetting');
      if (btnTM) btnTM.style.display = canManageTM ? 'inline-block' : 'none';

      // Tampilkan Semua Tanggal - Abaikan filter bulan
      const filterBulanEl = document.getElementById('lemburFilterBulan');
      if (filterBulanEl) filterBulanEl.style.display = 'none';

      google.script.run.withSuccessHandler(res => {
        if (!res.success) {
          toast('Gagal memuat data lembur', 'error');
          return;
        }

        // 1. Data Lembur
        if (res.lembur && res.lembur.success) lemburDataAll = res.lembur.data || [];
        else lemburDataAll = [];

        // 2. Data Laporan Kerja
        if (res.laporan && res.laporan.success) laporanKerjaData = res.laporan.data || [];
        else laporanKerjaData = [];

        // 3. Data Tanggal Merah
        if (res.tglMerah && res.tglMerah.success) tglMerahDataAll = res.tglMerah.data || [];
        else tglMerahDataAll = [];

        renderLemburTable();
      }).withFailureHandler(err => {
        console.error('Critical Error getLemburFullData:', err);
        toast('Gagal memuat data lembur dari server', 'error');
        renderLemburTable();
      }).getLemburFullData();
    }


    let currentLemburTab = 'list';
    function switchLemburTab(tab) {
      currentLemburTab = tab;
      const tList = document.getElementById('tabLemburList');
      const tBelum = document.getElementById('tabLemburBelum');
      const wList = document.getElementById('tableLemburWrap');
      const wBelum = document.getElementById('tableLemburBelumWrap');
      if (tList) tList.classList.toggle('active', tab === 'list');
      if (tBelum) tBelum.classList.toggle('active', tab === 'belum');
      if (wList) wList.style.display = tab === 'list' ? 'block' : 'none';
      if (wBelum) wBelum.style.display = tab === 'belum' ? 'block' : 'none';
    }

    function quickFillLembur(tanggal, jam, divisi) {
      openLemburNormal();
      autoFillLembur(tanggal, jam, divisi);
    }

    function renderLemburTable() {
      const tb = document.getElementById('tableLembur'); tb.innerHTML = '';
      const role = currentUser?.role || 'user';
      const isAdmin = role === 'admin';
      const isPimpinan = (role === 'admin' || role === 'HR' || role === 'Supervisor' || role === 'Vice Supervisor' || role.includes('Supervisor') || role.includes('HR'));
      const myName = (currentUser?.nama || '').toLowerCase();
      const filterDiv = v('filterLemburDivisi');

      // 1. Calculate missing overtime submissions (Belum Pengajuan)
      const missingMap = {};
      const activeSubmissions = (lemburDataAll || []).filter(d => d.status !== 'Dibatalkan');

      (laporanKerjaData || []).forEach(l => {
        if (!l.staffLemburNames || !l.tanggal) return;

        // Filter bulan dihapus agar tampil semua tanggal

        const names = l.staffLemburNames.split(',').map(n => n.trim());
        names.forEach(name => {
          if (!name) return;
          const nameLower = name.toLowerCase();

          // Non-pimpinan can only see their own missing submissions
          if (!isPimpinan && nameLower !== myName) return;

          const key = l.tanggal + '|' + nameLower;
          const hasSubmission = activeSubmissions.some(d => {
            return d.nama && d.nama.toLowerCase() === nameLower && d.tanggal === l.tanggal;
          });

          if (!hasSubmission) {
            if (!missingMap[key]) {
              missingMap[key] = {
                tanggal: l.tanggal,
                nama: name,
                divisi: l.divisi || '',
                jumlahJam: l.jamLembur || 0
              };
            } else {
              // Akumulasi jam lembur jika ada di beberapa laporan kerja pada hari yang sama
              missingMap[key].jumlahJam += (l.jamLembur || 0);
              if (l.divisi && !missingMap[key].divisi.includes(l.divisi)) {
                missingMap[key].divisi += ', ' + l.divisi;
              }
            }
          }
        });
      });

      let displayMissing = Object.values(missingMap);
      if (filterDiv) {
        displayMissing = displayMissing.filter(m => m.divisi === filterDiv);
      }

      // 2. Filter Submitted Overtime Data
      let displayData = lemburDataAll;

      // Role pimpinan (TL, Vice SPV, SPV, HR) bisa melihat semua data lembur untuk diproses
      const isLeaderRole = (
        role.includes('Team Leader') ||
        role === 'TL' ||
        role === 'Vice Supervisor' ||
        role === 'Vice SPV' ||
        role === 'Supervisor' ||
        role === 'SPV' ||
        role === 'HR'
      );

      if (!isAdmin && !isLeaderRole) {
        displayData = lemburDataAll.filter(d => (d.nama || '').toLowerCase() === myName);
      }
      if (filterDiv) {
        displayData = displayData.filter(d => d.divisi === filterDiv);
      }

      // 3. Compute Stats
      let totalApprovedHours = 0;
      let pendingCount = 0;
      let pendingForHR = 0;

      displayData.forEach(d => {
        if (d.status === 'Disetujui') {
          totalApprovedHours += parseFloat(d.jumlahJam) || 0;
        }
        if (d.status.includes('Pending')) {
          pendingCount++;
        }
        if (d.status === 'Pending HR') {
          pendingForHR++;
        }
      });

      // Check user permissions for user management (Masukan Ke Hak Akses manajemen User)
      let userPerms = []; try { userPerms = JSON.parse(currentUser.permissions || '[]'); } catch (e) { }
      const isHR = (role === 'HR');
      const isViceSPV = (role === 'Vice SPV' || role === 'Vice Supervisor');
      const canManageUsers = (isAdmin || isHR || isViceSPV || userPerms.includes('kelolaUser'));

      // Update Dashboard Counters
      const elAppr = document.getElementById('lemburStatApproved');
      const elPend = document.getElementById('lemburStatPending');
      const elBelum = document.getElementById('lemburStatBelum');
      const elCountBelum = document.getElementById('countLemburBelum');
      const cardBelum = document.getElementById('cardLemburBelum');
      const tabBelum = document.getElementById('tabLemburBelum');

      if (elAppr) elAppr.textContent = `${totalApprovedHours} Jam`;
      if (elPend) elPend.textContent = `${pendingCount} Pengajuan`;

      if (cardBelum) cardBelum.style.display = canManageUsers ? 'block' : 'none';
      if (tabBelum) tabBelum.style.display = canManageUsers ? 'flex' : 'none';

      if (canManageUsers) {
        if (elBelum) elBelum.textContent = `${displayMissing.length} Orang-Hari`;
        if (elCountBelum) elCountBelum.textContent = displayMissing.length;
      } else {
        switchLemburTab('list');
      }

      // 4. Render Main Table (Daftar Pengajuan Lembur)
      if (!displayData.length) {
        tb.innerHTML = '<tr><td colspan="8" class="empty-state">Belum ada riwayat pengajuan lembur yang sesuai</td></tr>';
      } else {
        displayData.sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal)).forEach(d => {
          let stsCls = 'badge-pending';
          if (d.status === 'Disetujui') stsCls = 'badge-approved';
          else if (d.status === 'Ditolak') stsCls = 'badge-rejected';
          else if (d.status.includes('Pending')) stsCls = 'badge-pending-stage';

          let actionHtml = `<button class="btn btn-ghost btn-sm" onclick="showHistoryModal('${encodeURIComponent(d.history)}')">Riwayat</button>`;
          let userPerms = []; try { userPerms = JSON.parse(currentUser.permissions || '[]'); } catch (e) { }
          const escapedNamaLembur = (d.nama || '').replace(/'/g, "\\'");
          const tanggalStrLembur = d.tanggal || '';
          const canAppr = canApprove(d.status, role, 'Lembur');

          if (canAppr) {
            actionHtml += `
              <button class="btn btn-teal btn-sm" onclick="processApproval('Lembur','${d.id}','Approve','${escapedNamaLembur}','${tanggalStrLembur}')">Setuju</button>
              <button class="btn btn-danger btn-sm" onclick="processApproval('Lembur','${d.id}','Reject','${escapedNamaLembur}','${tanggalStrLembur}')">Tolak</button>
            `;
          }

          if (isAdmin) {
            actionHtml += `<button class="btn btn-ghost btn-sm" onclick="editLemburAdmin('${d.id}')" title="Admin Edit">⚙️</button>`;
          }

          if (currentUser.username === d.createdBy && d.status === 'Pending Team Leader') {
            actionHtml += `<button class="btn btn-danger btn-sm" onclick="delLembur('${d.id}')">Batal</button>`;
          }

          const absenInfo = `<div style="font-size:10px; color:var(--gray);">IN: ${d.inTime || '-'}</div><div style="font-size:10px; color:var(--gray);">OUT: ${d.outTime || '-'}</div>`;

          tb.innerHTML += `<tr>
            <td>${formatDate(d.tanggal)}</td>
            <td><strong>${d.nama}</strong></td>
            <td>${d.divisi}</td>
            <td><span style="color:var(--accent);font-weight:700">⌛ ${d.jumlahJam} Jam</span></td>
            <td>${absenInfo}</td>
            <td>${d.keterangan}</td>
            <td><span class="${stsCls}">${d.status}</span></td>
            <td style="display:flex; gap:5px;">${actionHtml}</td>
          </tr>`;
        });
      }

      // 5. Render Missing Overtime Table (Belum Pengajuan)
      const tbBelum = document.getElementById('tableLemburBelum');
      if (tbBelum) {
        tbBelum.innerHTML = '';
        if (!displayMissing.length) {
          tbBelum.innerHTML = '<tr><td colspan="6" class="empty-state">Semua personil lembur sudah melakukan pengajuan</td></tr>';
        } else {
          displayMissing.sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal)).forEach(m => {
            let actionHtml = '';
            if (currentUser?.nama && m.nama.toLowerCase() === currentUser.nama.toLowerCase()) {
              actionHtml = `<button class="btn btn-primary btn-sm" onclick="quickFillLembur('${m.tanggal}', ${m.jumlahJam}, '${m.divisi}')">Ajukan Sekarang</button>`;
            } else {
              actionHtml = `<span class="text-muted" style="font-size:12px;">Menunggu tindakan karyawan</span>`;
            }

            tbBelum.innerHTML += `<tr>
              <td>${formatDate(m.tanggal)}</td>
              <td><strong>${m.nama}</strong></td>
              <td>${m.divisi}</td>
              <td><span style="color:var(--red);font-weight:700">⌛ ${m.jumlahJam} Jam</span></td>
              <td><span class="badge-rejected">Belum Diajukan</span></td>
              <td>${actionHtml}</td>
            </tr>`;
          });
        }
      }

      // Update HR Badge if exists
      const hBadge = document.getElementById('notifBadgeHR');
      if (hBadge) {
        if (pendingForHR > 0) { hBadge.textContent = pendingForHR; hBadge.style.display = 'flex'; }
        else { hBadge.style.display = 'none'; }
      }
    }


    function editLemburAdmin(id) {
      const d = lemburDataAll.find(x => x.id === id);
      if (!d) return;
      setVal('editAdminLbId', d.id);
      setVal('editAdminLbInfo', `${d.nama} (${d.divisi})`);
      setVal('editAdminLbJam', d.jumlahJam);
      setVal('editAdminLbStatus', d.status);
      setVal('editAdminLbNote', '');
      openModal('modalEditLemburAdmin');
    }

    function submitEditLemburAdmin() {
      const id = v('editAdminLbId'), jam = v('editAdminLbJam'), sts = v('editAdminLbStatus'), note = v('editAdminLbNote');
      const btn = event.target;
      btn.disabled = true; btn.textContent = '⏳';
      google.script.run.withSuccessHandler(res => {
        btn.disabled = false; btn.textContent = '💾 Update Data';
        if (res.success) {
          toast('Data terupdate oleh Admin');
          closeModal('modalEditLemburAdmin');
          loadLembur();
        } else toast(res.message, 'error');
      }).updateLemburAdmin(id, jam, sts, note, currentUser.nama);
    }

    function updateLemburStaffSelect() {
      const t = v('lbTanggal');
      const inp = document.getElementById('lbNama');
      const info = document.getElementById('lbNamesInfo');
      const ok = document.getElementById('lbNamesOk');

      if (!inp) return;
      const myName = currentUser?.nama || '';
      inp.value = myName;

      info.style.display = 'none'; ok.style.display = 'none';
      if (!t) return;

      const lapForDate = laporanKerjaData.filter(d => d.tanggal === t);
      let names = [];
      lapForDate.forEach(l => { if (l.staffLemburNames) names = [...names, ...l.staffLemburNames.split(',')]; });
      names = [...new Set(names)].map(n => n.trim().toLowerCase());

      const userPerms = JSON.parse(currentUser.permissions || '[]');
      const canDirect = userPerms.includes('lemburTanpaLaporan') || currentUser.role === 'admin';

      // Cek apakah terdaftar di Tanggal Merah
      const registeredTM = (tglMerahDataAll || []).some(d => d.tanggal === t && d.nama.toLowerCase() === myName.toLowerCase());

      if (!names.includes(myName.toLowerCase()) && !registeredTM) {
        if (canDirect) {
          ok.style.display = 'block';
          document.getElementById('lbJumlahJam').readOnly = false;
        } else {
          info.style.display = 'block';
          setVal('lbJumlahJam', '0');
        }
      } else {
        ok.style.display = 'block';
        syncJamLemburDariLaporan();
      }
    }

    function openLemburNormal() {
      const now = new Date();
      const today = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + String(now.getDate()).padStart(2, '0');
      setVal('lbTanggal', today);
      setVal('lbKeterangan', '');

      openModal('modalLembur');
      updateLemburStaffSelect();
      renderMissingLemburDates();
    }

    function renderMissingLemburDates() {
      const suggestBox = document.getElementById('missingLemburSuggest');
      const suggestList = document.getElementById('missingLemburSuggestList');
      const emptyBox = document.getElementById('missingLemburEmpty');
      if (!suggestBox || !suggestList || !emptyBox) return;

      suggestBox.style.display = 'none';
      emptyBox.style.display = 'none';
      suggestList.innerHTML = '';

      const myName = (currentUser?.nama || '').toLowerCase();
      if (!myName || !laporanKerjaData.length) return;

      // 1. Kumpulkan semua tanggal di Laporan Kerja di mana nama staff ini ada
      const datesInLaporan = [];
      const lapHoursMap = {}; // tanggal -> { jam, divisi }
      laporanKerjaData.forEach(l => {
        if (!l.staffLemburNames) return;
        const names = l.staffLemburNames.split(',').map(n => n.trim().toLowerCase());
        if (names.includes(myName)) {
          if (!datesInLaporan.includes(l.tanggal)) datesInLaporan.push(l.tanggal);
          lapHoursMap[l.tanggal] = { jam: l.jamLembur || 0, divisi: l.divisi || '' };
        }
      });

      if (!datesInLaporan.length) { emptyBox.style.display = 'block'; return; }

      // 2. Tanggal yang sudah punya pengajuan lembur (bukan Dibatalkan)
      const submittedDates = (lemburDataAll || []).filter(d => {
        return d.nama && d.nama.toLowerCase() === myName && d.status !== 'Dibatalkan';
      }).map(d => d.tanggal);

      // 3. Selisih = belum diajukan
      const missing = datesInLaporan.filter(t => !submittedDates.includes(t))
        .sort((a, b) => new Date(b) - new Date(a));

      if (!missing.length) { emptyBox.style.display = 'block'; return; }

      suggestBox.style.display = 'block';
      missing.forEach(tDate => {
        const info = lapHoursMap[tDate];
        const btn = document.createElement('button');
        btn.className = 'btn btn-ghost btn-sm';
        btn.style.cssText = 'background:rgba(245,158,11,0.12); border:1px solid var(--accent); color:var(--accent); border-radius:8px; padding:5px 12px; font-size:11px; font-weight:700; cursor:pointer; white-space:nowrap; transition:all .2s;';
        btn.innerHTML = `📅 ${formatDate(tDate)} &nbsp;•&nbsp; ${info.jam} Jam`;
        btn.onclick = () => autoFillLembur(tDate, info.jam, info.divisi);
        suggestList.appendChild(btn);
      });
    }

    function autoFillLembur(tanggal, jam, divisi) {
      // Auto-fill form fields
      setVal('lbTanggal', tanggal);
      setVal('lbDivisi', divisi);

      const jamEl = document.getElementById('lbJumlahJam');
      if (jamEl) jamEl.value = jam;

      // Update UI verification status
      updateLemburStaffSelect();

      // Visual feedback (Highlight button)
      const btns = document.querySelectorAll('#missingLemburSuggestList button');
      btns.forEach(b => {
        b.style.background = 'rgba(245,158,11,0.12)';
        b.style.transform = 'scale(1)';
        b.style.borderColor = 'var(--accent)';
      });

      const target = event.target.closest('button');
      if (target) {
        target.style.background = 'rgba(245,158,11,0.35)';
        target.style.transform = 'scale(1.05)';
        target.style.borderColor = '#ffffff';
      }

      toast('✅ Form diisi otomatis untuk tanggal: ' + formatDate(tanggal), 'success');
    }



    function syncJamLemburDariLaporan() {
      const t = v('lbTanggal');
      const nama = v('lbNama');
      const out = document.getElementById('lbJumlahJam');
      const divOut = document.getElementById('lbDivisi'); // automatic division sync
      if (!out) return;
      const userPerms = JSON.parse(currentUser.permissions || '[]');
      const canDirect = userPerms.includes('lemburTanpaLaporan') || currentUser.role === 'admin';

      out.value = '0';
      out.readOnly = !canDirect;

      if (!t || !nama) return;

      const lapForDate = laporanKerjaData.filter(d => d.tanggal === t);
      let jamFound = 0;
      let divFound = '';
      lapForDate.forEach(l => {
        if (l.staffLemburNames) {
          const names = l.staffLemburNames.split(',').map(x => x.trim().toLowerCase());
          if (names.includes(nama.toLowerCase())) {
            // Akumulasi jam jika nama muncul di beberapa divisi/shift laporan kerja
            jamFound += parseFloat(l.jamLembur) || 0;
            if (!divFound) divFound = l.divisi;
            else if (l.divisi && !divFound.includes(l.divisi)) divFound += ', ' + l.divisi;
          }
        }
      });

      // Akumulasikan dengan data Tanggal Merah jika terdaftar pada tanggal tersebut
      let jamTglMerah = 0;
      if (tglMerahDataAll && tglMerahDataAll.length) {
        const foundTM = tglMerahDataAll.find(d => d.tanggal === t && d.nama.toLowerCase() === nama.toLowerCase());
        if (foundTM) {
          jamTglMerah = parseFloat(foundTM.jamEstimasi) || 0;
        }
      }

      const totalJam = jamFound + jamTglMerah;
      out.value = totalJam;

      if (jamTglMerah > 0) {
        toast(`Akumulasi: ${jamFound} Jam Reguler + ${jamTglMerah} Jam Libur = ${totalJam} Jam`, 'info');
      }

      if (divOut) {
        let finalJab = '';
        if (karyawanData && karyawanData.length) {
          const emp = karyawanData.find(k => k.nama.toLowerCase() === nama.toLowerCase());
          if (emp) finalJab = emp.jabatan;
        }
        divOut.value = finalJab || currentUser?.jabatan || divFound || '';
      }
      updateLemburVerificationDisplay();
    }




    function submitLembur() {
      const t = v('lbTanggal'), n = v('lbNama'), d = v('lbDivisi'), jj = v('lbJumlahJam'), k = v('lbKeterangan');
      if (!t || !n || !jj || !k) return toast('Lengkapi data', 'error');

      const btn = document.querySelector('#modalLembur .btn-primary');
      btn.disabled = true; btn.textContent = '⏳...';

      google.script.run.withSuccessHandler(res => {
        btn.disabled = false; btn.textContent = '💾 Ajukan Lembur';
        if (res.success) { toast('Berhasil'); closeModal('modalLembur'); loadLembur(); resetForm(['lbKeterangan']); }
        else { toast(res.message, 'error'); }
      }).addLembur(t, n, d, jj, k, currentUser.username);
    }

    // ============================================================
    // LEMBUR TANGGAL MERAH LOGIC
    // ============================================================
    let tglMerahDataAll = [];

    function loadTglMerah() {
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          tglMerahDataAll = res.data || [];
          renderTglMerahList();
        }
      }).getTglMerahData();
    }

    let _tmDivisiAkunMode = false;

    function toggleTmDivisiMode() {
      _tmDivisiAkunMode = !_tmDivisiAkunMode;
      const track = document.getElementById('tmToggleTrack');
      const thumb = document.getElementById('tmToggleThumb');
      const label = document.getElementById('tmDivisiModeLabel');
      const chip = document.getElementById('tmAkunDivisiChip');
      const chipTxt = document.getElementById('tmAkunDivisiText');
      const selDiv = document.getElementById('tmFilterDivisi');

      if (_tmDivisiAkunMode) {
        track.style.background = 'var(--teal)';
        thumb.style.left = '20px';
        label.style.color = 'var(--teal)';
        label.textContent = 'Sesuai Akun';
        const akDivisi = currentUser?.jabatan || '';
        chipTxt.textContent = akDivisi || 'Jabatan tidak diatur';
        chip.style.display = 'block';
        // Auto-set the dropdown to match account divisi (jabatan)
        if (akDivisi) {
          const opts = selDiv.options;
          let matched = false;
          for (let i = 0; i < opts.length; i++) {
            if (opts[i].value.toLowerCase() === akDivisi.toLowerCase() ||
              akDivisi.toLowerCase().includes(opts[i].value.toLowerCase()) && opts[i].value !== '') {
              selDiv.value = opts[i].value; matched = true; break;
            }
          }
          if (!matched) selDiv.value = '';
        }
      } else {
        track.style.background = 'var(--border-color)';
        thumb.style.left = '2px';
        label.style.color = 'var(--gray)';
        label.textContent = 'Semua Jabatan';
        chip.style.display = 'none';
      }
      renderTglMerahStaffList();
    }

    function openTglMerahSetting() {
      _tmDivisiAkunMode = false;
      // Reset toggle UI
      const track = document.getElementById('tmToggleTrack');
      const thumb = document.getElementById('tmToggleThumb');
      const label = document.getElementById('tmDivisiModeLabel');
      const chip = document.getElementById('tmAkunDivisiChip');
      if (track) { track.style.background = 'var(--border-color)'; }
      if (thumb) { thumb.style.left = '2px'; }
      if (label) { label.textContent = 'Semua Jabatan'; label.style.color = 'var(--gray)'; }
      if (chip) { chip.style.display = 'none'; }

      setVal('tmTanggal', new Date().toISOString().split('T')[0]);
      setVal('tmJamEstimasi', '8');
      setVal('tmSearchKaryawan', '');
      setVal('tmFilterDivisi', '');

      renderTglMerahStaffList();
      loadTglMerah();
      openModal('modalTglMerahSetting');
    }


    function renderTglMerahStaffList() {
      const q = v('tmSearchKaryawan').toLowerCase();
      const divFilter = v('tmFilterDivisi');
      const div = document.getElementById('tmStaffList');
      div.innerHTML = '';
      document.getElementById('tmCheckAll').checked = false;

      // If Divisi Sesuai Akun mode, compute effective divisi filter
      let effectiveDivFilter = divFilter;
      if (_tmDivisiAkunMode) {
        const akDivisi = (currentUser?.jabatan || '').toLowerCase();
        if (akDivisi && !divFilter) effectiveDivFilter = '__akun__';
      }

      const filtered = karyawanData.filter(k => {
        const matchQ = k.nama.toLowerCase().includes(q);
        const divK = (k.jabatan || 'Staff').toLowerCase();
        let matchDiv = true;
        if (effectiveDivFilter === '__akun__') {
          const akDivisi = (currentUser?.jabatan || '').toLowerCase();
          matchDiv = akDivisi ? divK.includes(akDivisi) || akDivisi.includes(divK) : true;
        } else if (effectiveDivFilter) {
          matchDiv = divK === effectiveDivFilter.toLowerCase();
        }
        return matchQ && matchDiv;
      }).sort((a, b) => a.nama.localeCompare(b.nama));

      if (!filtered.length) {
        div.innerHTML = '<div style="text-align:center; color:var(--gray); font-size:12px; padding:20px;">Tidak ada karyawan ditemukan</div>';
        return;
      }

      filtered.forEach(k => {
        const kDivisi = k.jabatan || 'Staff';
        div.innerHTML += `
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:5px; padding:6px 4px; border-bottom:1px solid #ffffff05;">
            <input type="checkbox" class="tm-staff-check" value="${k.nama}" data-divisi="${kDivisi}" id="chk-tm-${k.id}">
            <label for="chk-tm-${k.id}" style="font-size:12px; cursor:pointer; flex:1;">
              ${k.nama}
              <span style="color:var(--gray); font-size:10px; margin-left:4px;">(${k.jabatan})</span>
              <span style="background:rgba(14,165,233,0.12); color:var(--teal); font-size:9px; font-weight:700; padding:1px 6px; border-radius:6px; margin-left:4px;">${kDivisi}</span>
            </label>
          </div>
        `;
      });
    }

    function sinkronkanDenganLembur() {
      const tanggal = v('tmTanggal');
      if (!tanggal) return toast('Pilih tanggal terlebih dahulu!', 'warning');

      // Gather names already in lembur for that date
      const namesInLembur = new Set();
      (laporanKerjaData || []).filter(l => l.tanggal === tanggal).forEach(l => {
        if (l.staffLemburNames) {
          l.staffLemburNames.split(',').map(n => n.trim().toLowerCase()).filter(Boolean).forEach(n => namesInLembur.add(n));
        }
      });
      // Also check lemburDataAll
      (lemburDataAll || []).filter(l => l.tanggal === tanggal && l.status !== 'Dibatalkan').forEach(l => {
        if (l.nama) namesInLembur.add(l.nama.toLowerCase());
      });

      if (!namesInLembur.size) return toast('Tidak ada data lembur ditemukan untuk tanggal ini.', 'info');

      let checked = 0;
      document.querySelectorAll('.tm-staff-check').forEach(cb => {
        if (namesInLembur.has(cb.value.toLowerCase())) { cb.checked = true; checked++; }
      });
      toast(`✅ ${checked} personel berhasil disinkronkan dari data lembur!`, 'success');
    }

    function toggleSelectAllTglMerah(el) {
      const checks = document.querySelectorAll('.tm-staff-check');
      checks.forEach(c => c.checked = el.checked);
    }


    function submitTglMerahSetting() {
      const t = v('tmTanggal'), jam = v('tmJamEstimasi');
      if (!t) return toast('Pilih tanggal!', 'warning');

      const checked = document.querySelectorAll('.tm-staff-check:checked');
      if (!checked.length) return toast('Pilih minimal 1 orang!', 'warning');

      const data = Array.from(checked).map(cb => ({
        nama: cb.value,
        divisi: cb.getAttribute('data-divisi'),
        jamEstimasi: jam
      }));

      const btn = document.querySelector('#modalTglMerahSetting .btn-warning');
      btn.disabled = true; btn.textContent = '⏳ Menyimpan...';

      google.script.run.withSuccessHandler(res => {
        btn.disabled = false; btn.textContent = '💾 Simpan Daftar Personel';
        if (res.success) {
          toast('Personel berhasil didaftarkan!', 'success');
          loadTglMerah();
          // Clear checks
          document.querySelectorAll('.tm-staff-check').forEach(c => c.checked = false);
        } else toast(res.message, 'error');
      }).addTglMerahPersonel(t, data, currentUser.username);
    }

    function renderTglMerahList() {
      const tb = document.getElementById('tblTglMerahList');
      tb.innerHTML = '';
      if (!tglMerahDataAll.length) { tb.innerHTML = '<tr><td colspan="4" class="empty-state">Belum ada data</td></tr>'; return; }

      [...tglMerahDataAll].sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal)).forEach(d => {
        tb.innerHTML += `
          <tr>
            <td>${formatDate(d.tanggal)}</td>
            <td><strong>${d.nama}</strong></td>
            <td>${d.divisi}</td>
            <td><button class="btn btn-ghost btn-sm" onclick="delTglMerah('${d.id}')" style="color:var(--red)">🗑️</button></td>
          </tr>
        `;
      });
    }

    function delTglMerah(id) {
      if (confirm('Hapus pendaftaran lembur ini?')) {
        google.script.run.withSuccessHandler(res => {
          if (res.success) { toast('Berhasil dihapus'); loadTglMerah(); }
          else toast(res.message, 'error');
        }).deleteTglMerah(id);
      }
    }

    function openLemburTglMerah() {
      setVal('tmUserTanggal', new Date().toISOString().split('T')[0]);
      setVal('tmUserNama', currentUser.nama);
      setVal('tmUserJam', '');
      setVal('tmUserKet', '');
      document.getElementById('tmWarning').style.display = 'none';
      document.getElementById('btnSubmitTglMerah').disabled = true;

      // Ensure data is loaded
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          tglMerahDataAll = res.data || [];
          validateTglMerahUser();
        }
      }).getTglMerahData();

      openModal('modalLemburTglMerah');
    }

    function validateTglMerahUser() {
      const t = v('tmUserTanggal'), n = currentUser.nama;
      const warn = document.getElementById('tmWarning');
      const btn = document.getElementById('btnSubmitTglMerah');

      const found = tglMerahDataAll.find(d => d.tanggal === t && d.nama === n);
      if (found) {
        warn.style.display = 'none';
        btn.disabled = false;

        // --- AKUMULASI OTOMATIS ---
        let jamLibur = parseFloat(found.jamEstimasi) || 0;
        let jamReguler = 0;

        // Cari di Laporan Jam Kerja
        const lapForDate = (laporanKerjaData || []).filter(l => l.tanggal === t);
        lapForDate.forEach(l => {
          if (l.staffLemburNames) {
            const names = l.staffLemburNames.split(',').map(x => x.trim().toLowerCase());
            if (names.includes(n.toLowerCase())) {
              jamReguler = parseFloat(l.jamLembur) || 0;
            }
          }
        });

        const totalJam = jamLibur + jamReguler;
        setVal('tmUserJam', totalJam);

        if (jamReguler > 0) {
          toast(`Akumulasi: ${jamLibur} Jam Libur + ${jamReguler} Jam Reguler = ${totalJam} Jam`, 'info');
        }

      } else {
        warn.style.display = 'block';
        btn.disabled = true;
        setVal('tmUserJam', '-');
      }
    }


    function submitLemburTglMerah() {
      const t = v('tmUserTanggal'), n = v('tmUserNama'), j = v('tmUserJam'), k = v('tmUserKet');
      if (!t || !n || !j || !k) return toast('Lengkapi keterangan tugas!', 'warning');

      const found = tglMerahDataAll.find(d => d.tanggal === t && d.nama === n);
      if (!found) return toast('Anda tidak terdaftar untuk tanggal ini!', 'error');

      const btn = document.getElementById('btnSubmitTglMerah');
      btn.disabled = true; btn.textContent = '⏳ Memproses...';

      google.script.run.withSuccessHandler(res => {
        btn.disabled = false; btn.textContent = '🚀 Ajukan Lembur Hari Libur';
        if (res.success) {
          toast('Berhasil diajukan!', 'success');
          closeModal('modalLemburTglMerah');
          loadLembur();
        } else toast(res.message, 'error');
      }).addLembur(t, n, found.divisi, j, k, currentUser.username);
    }


    function doExportLembur() {
      const start = v('exLbStart'), end = v('exLbEnd');
      if (!start || !end) return toast('Pilih rentang tanggal!', 'warning');
      const filtered = (lemburDataAll || []).filter(r => r.tanggal >= start && r.tanggal <= end);
      if (!filtered.length) return toast('Tidak ada data', 'info');

      // AGREGASI DATA (GABUNG NAMA & TANGGAL SAMA - HANYA YANG DISETUJUI)
      const aggregatedMap = {};
      filtered.forEach(r => {
        // Hanya yang berstatus "Disetujui" yang dihitung ke total payroll/rekap
        if (r.status !== 'Disetujui') return;

        const key = `${r.nama || 'Tanpa Nama'}|${r.tanggal}`;
        if (!aggregatedMap[key]) {
          aggregatedMap[key] = {
            tanggal: r.tanggal,
            nama: r.nama,
            divisi: r.divisi || '-',
            jumlahJam: 0,
            keterangan: [],
            status: 'Terverifikasi (Disetujui)',
            createdBy: r.createdBy
          };
        }
        aggregatedMap[key].jumlahJam += parseFloat(r.jumlahJam) || 0;
        if (r.keterangan && r.keterangan.trim()) {
          aggregatedMap[key].keterangan.push(r.keterangan.trim());
        }
      });

      // Konversi Map ke Array dan Sortir berdasarkan Tanggal
      // Filter out rows with 0 hours (e.g. if a day only had pending requests)
      const resultList = Object.values(aggregatedMap)
        .filter(r => r.jumlahJam > 0)
        .sort((a, b) => a.tanggal.localeCompare(b.tanggal));

      if (!resultList.length) return toast('Tidak ada data lembur yang disetujui pada periode ini', 'info');

      let html = `<table border="1"><thead><tr><th>Tanggal</th><th>Nama</th><th>Divisi</th><th>Total Jam Disetujui</th><th>Rincian Keterangan</th><th>Status</th><th>Oleh</th></tr></thead><tbody>`;
      resultList.forEach(r => {
        const ketStr = r.keterangan.length > 0 ? r.keterangan.join('; ') : '-';
        html += `<tr><td>${r.tanggal}</td><td>${r.nama}</td><td>${r.divisi}</td><td>${r.jumlahJam}</td><td>${ketStr}</td><td>${r.status}</td><td>${r.createdBy}</td></tr>`;
      });
      html += '</tbody></table>';
      exportToExcel(html, `Laporan_Lembur_Approval_Fix_${start}_sd_${end}.xls`);
      closeModal('modalExportLembur');
    }

    function exportToExcel(html, filename) {
      const uri = 'data:application/vnd.ms-excel;base64,';
      const template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body>{table}</body></html>';
      const base64 = (s) => window.btoa(unescape(encodeURIComponent(s)));
      const format = (s, c) => s.replace(/{(\w+)}/g, (m, p) => c[p]);
      const ctx = { worksheet: 'Sheet1', table: html };
      const link = document.createElement('a');
      link.href = uri + base64(format(template, ctx));
      link.download = filename;
      link.click();
    }

    function delLembur(id) { if (confirm('Batalkan pengajuan?')) google.script.run.withSuccessHandler(res => { if (res.success) { toast('Dibatalkan'); loadLembur(); } else toast(res.message, 'error'); }).deleteLembur(id); }

    // ORG CHART
    // ORG CHART
    let orgDataCache = []; // Cache untuk filtering

    function loadOrg() {
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          orgDataCache = res.data || [];
          renderOrgTree(orgDataCache);
          updateOrgStats(orgDataCache);
          populateOrgFilters(orgDataCache);
        } else {
          const tree = document.getElementById('orgTree');
          tree.innerHTML = '<div class="org-empty-state"><div style="font-size:48px;opacity:0.3;">❌</div><div style="color:var(--red);">Gagal memuat data organisasi</div></div>';
        }
      }).getOrganisasi();
    }

    function updateOrgStats(data) {
      const total = data.length;
      const depts = new Set(data.map(d => d.departemen || 'Lainnya').filter(Boolean));
      const leaders = data.filter(d => !d.atasan || d.atasan === '').length;

      document.getElementById('orgStatTotal').textContent = total;
      document.getElementById('orgStatDept').textContent = depts.size;
      document.getElementById('orgStatLeader').textContent = leaders;
    }

    function populateOrgFilters(data) {
      const depts = new Set(data.map(d => d.departemen || 'Lainnya').filter(Boolean));
      const deptFilter = document.getElementById('orgDeptFilter');
      const currentVal = deptFilter.value;

      deptFilter.innerHTML = '<option value="">Semua Departemen</option>';
      Array.from(depts).sort().forEach(dept => {
        deptFilter.innerHTML += `<option value="${dept}">${dept}</option>`;
      });
      deptFilter.value = currentVal;

      // Populate modal atasan dropdown
      const selAts = document.getElementById('orgAtasan');
      if (selAts) {
        selAts.innerHTML = '<option value="">— Tidak ada (Root) —</option>';
        data.sort((a, b) => (a.urutan || 0) - (b.urutan || 0))
          .forEach(d => selAts.innerHTML += `<option value="${d.jabatan}">${d.jabatan} - ${d.nama}</option>`);
      }
    }

    function renderOrgTree(data) {
      const tree = document.getElementById('orgTree');
      tree.innerHTML = '';

      if (!data || data.length === 0) {
        tree.innerHTML = `
          <div class="org-empty-state">
            <div style="font-size:64px;margin-bottom:16px;opacity:0.3;">🏗️</div>
            <div style="font-size:18px;font-weight:700;color:var(--text-main);margin-bottom:8px;">
              Belum Ada Struktur Organisasi
            </div>
            <div style="font-size:13px;color:var(--text-muted);">
              Klik tombol "Tambah Anggota" untuk membuat struktur organisasi
            </div>
          </div>`;
        return;
      }

      const buildNode = (jabatanAtasan) => {
        const nodes = data.filter(d => (jabatanAtasan ? d.atasan === jabatanAtasan : !d.atasan));
        if (!nodes.length) return '';

        let html = '<div class="oc-subtree">';
        nodes.forEach(n => {
          const childrenHTML = buildNode(n.jabatan);
          const hasChildren = childrenHTML !== '';
          const isLeader = !n.atasan || n.atasan === '';
          const initials = n.nama.split(' ').map(w => w[0]).join('').substring(0, 2).toUpperCase();

          html += `
            <div class="oc-container">
              <div class="oc-node ${isLeader ? 'leader' : ''}" data-id="${n.id}" data-dept="${n.departemen || ''}" data-name="${n.nama}">
                <div class="oc-avatar">
                  ${n.foto ? `<img src="${n.foto}" alt="${n.nama}">` : `<span>${initials}</span>`}
                </div>
                <div class="oc-name">${n.nama}</div>
                <div class="oc-title">${n.jabatan}</div>
                ${n.departemen ? `<div class="oc-dept">🏢 ${n.departemen}</div>` : ''}
                <div class="oc-actions">
                  <button class="oc-btn" onclick="editOrg('${n.id}')" title="Edit">✏️ Edit</button>
                  <button class="oc-btn delete" onclick="delOrg('${n.id}')" title="Hapus">🗑️</button>
                </div>
                ${hasChildren ? `<div class="oc-toggle" onclick="toggleOrgNode(this)" title="Expand/Collapse">−</div>` : ''}
              </div>
              ${hasChildren ? childrenHTML : ''}
            </div>
          `;
        });
        html += '</div>';
        return html;
      };

      tree.innerHTML = buildNode('');
    }

    function toggleOrgNode(btn) {
      const container = btn.closest('.oc-container');
      const subtree = container.querySelector('.oc-subtree');
      if (subtree) {
        subtree.classList.toggle('collapsed');
        btn.textContent = subtree.classList.contains('collapsed') ? '+' : '−';
      }
    }

    function expandAllOrg() {
      document.querySelectorAll('.oc-subtree').forEach(e => e.classList.remove('collapsed'));
      document.querySelectorAll('.oc-toggle').forEach(e => e.textContent = '−');
    }

    function collapseAllOrg() {
      document.querySelectorAll('.oc-subtree').forEach(e => e.classList.add('collapsed'));
      document.querySelectorAll('.oc-toggle').forEach(e => e.textContent = '+');
    }

    function filterOrgChart() {
      const searchTerm = (document.getElementById('orgSearchInput')?.value || '').toLowerCase();
      const deptFilter = document.getElementById('orgDeptFilter')?.value || '';

      const nodes = document.querySelectorAll('.oc-node');
      let visibleCount = 0;

      nodes.forEach(node => {
        const name = (node.getAttribute('data-name') || '').toLowerCase();
        const dept = node.getAttribute('data-dept') || '';
        const title = (node.querySelector('.oc-title')?.textContent || '').toLowerCase();

        const matchSearch = !searchTerm || name.includes(searchTerm) || title.includes(searchTerm);
        const matchDept = !deptFilter || dept === deptFilter;

        const container = node.closest('.oc-container');
        if (matchSearch && matchDept) {
          container.style.display = '';
          visibleCount++;
        } else {
          container.style.display = 'none';
        }
      });
    }

    function resetOrgFilter() {
      document.getElementById('orgSearchInput').value = '';
      document.getElementById('orgDeptFilter').value = '';
      filterOrgChart();
    }

    function openOrgModal(editId = null) {
      document.getElementById('orgId').value = editId || '';
      document.getElementById('orgModalTitle').textContent = editId ? '✏️ Edit Anggota' : '🏗️ Tambah Anggota Organisasi';

      if (editId) {
        const org = orgDataCache.find(o => o.id === editId);
        if (org) {
          document.getElementById('orgNama').value = org.nama || '';
          document.getElementById('orgJabatan').value = org.jabatan || '';
          document.getElementById('orgAtasan').value = org.atasan || '';
          document.getElementById('orgDept').value = org.departemen || '';
          document.getElementById('orgFoto').value = org.foto || '';
          document.getElementById('orgUrutan').value = org.urutan || '';
        }
      } else {
        document.getElementById('orgNama').value = '';
        document.getElementById('orgJabatan').value = '';
        document.getElementById('orgAtasan').value = '';
        document.getElementById('orgDept').value = '';
        document.getElementById('orgFoto').value = '';
        document.getElementById('orgUrutan').value = '';
      }

      openModal('modalOrg');
    }

    function editOrg(id) {
      openOrgModal(id);
    }

    function submitOrg() {
      const id = v('orgId');
      const n = v('orgNama');
      const j = v('orgJabatan');
      const a = v('orgAtasan');
      const d = v('orgDept');
      const f = v('orgFoto');
      const u = v('orgUrutan');

      if (!n || !j) return toast('Nama dan Jabatan wajib diisi!', 'error');

      const btn = document.querySelector('#modalOrg .btn-primary');
      btn.disabled = true;
      btn.textContent = '⏳ Menyimpan...';

      const fn = id ? 'updateOrganisasi' : 'addOrganisasi';

      google.script.run.withSuccessHandler(res => {
        btn.disabled = false;
        btn.textContent = '💾 Simpan';

        if (res.success) {
          toast(id ? 'Anggota berhasil diupdate!' : 'Anggota berhasil ditambahkan!', 'success');
          closeModal('modalOrg');
          loadOrg();
        } else {
          toast(res.message || 'Gagal menyimpan data', 'error');
        }
      }).withFailureHandler(err => {
        btn.disabled = false;
        btn.textContent = '💾 Simpan';
        toast('Error: ' + err.message, 'error');
      })[fn](id, n, j, a, d, f, u);
    }

    function delOrg(id) {
      if (!confirm('Yakin ingin menghapus anggota ini dari struktur organisasi?')) return;

      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          toast('Anggota berhasil dihapus!', 'success');
          loadOrg();
        } else {
          toast(res.message || 'Gagal menghapus anggota', 'error');
        }
      }).withFailureHandler(err => {
        toast('Error: ' + err.message, 'error');
      }).deleteOrganisasi(id);
    }

    // === OPERASIONAL & LAPORAN KERJA ===
    function updateLapFields() {
      const div = v('lapDivisi');
      const wOrd = document.getElementById('wrapLapOrder');
      const wDist = document.getElementById('wrapLapDistributor');
      const wInb = document.getElementById('wrapLapInbound');
      const wCon = document.getElementById('wrapLapConsumable');
      const labelOrd = document.getElementById('labelLapOrder');

      // Reset visibility
      if (wOrd) wOrd.style.display = 'none';
      if (wDist) wDist.style.display = 'none';
      if (wInb) wInb.style.display = 'none';
      if (wCon) wCon.style.display = 'none';

      if (div === 'Marketplace' || div === 'Return' || div === 'KOL' || div === 'Market Place SBY') {
        if (wOrd) wOrd.style.display = 'block';
        if (div === 'Marketplace' || div === 'Market Place SBY') labelOrd.textContent = 'Total Order Marketplace';
        else if (div === 'Return') labelOrd.textContent = 'Total Paket Return';
        else labelOrd.textContent = 'Total Order KOL';
      } else if (div === 'Distributor' || div === 'Distributor SBY') {
        if (wDist) wDist.style.display = 'block';
      } else if (div === 'Inbound') {
        if (wInb) wInb.style.display = 'block';
      } else if (div === 'Consumable') {
        if (wCon) wCon.style.display = 'block';
      }

      // Reset KPI label
      document.getElementById('wrapLapKPIOrder').querySelector('label').textContent = 'KPI (Total Output / Total Jam Kerja)';
      document.getElementById('wrapLapKPIQty').style.display = 'none';
      calcJamKerja();
    }

    function calcJamKerja() {
      const oU = parseInt(v('lapOrang')) || 0, a = parseInt(v('lapAdmin')) || 0, oPhl = parseInt(v('lapPhl')) || 0, jPhl = parseFloat(v('lapJamPhl')) || 0;
      const oB = parseInt(v('lapOrangBantu')) || 0, jB = parseFloat(v('lapJamBantu')) || 0;
      const oK = parseInt(v('lapOrangKurang')) || 0, jK = parseFloat(v('lapJamKurang')) || 0;
      const oL = parseInt(v('lapOrangLembur')) || 0, jL = parseFloat(v('lapLembur')) || 0;

      const div = v('lapDivisi');
      let output = 0;
      if (div === 'Distributor' || div === 'Distributor SBY' || div === 'Distributor Surabaya') output = parseInt(v('lapQty')) || 0;
      else if (div === 'Inbound') output = parseInt(v('lapInbound')) || 0;
      else if (div === 'Consumable') output = (parseInt(v('lapPotongBubble')) || 0) + (parseInt(v('lapBuatBubble')) || 0);
      else output = parseInt(v('lapOrder')) || 0;

      // Rumus: ((Utama + Admin) * 8) + (PHL * Jam) + (Bantu * Jam) - (Kurang * Jam) + (Lembur * Jam)
      const tot = ((oU + a) * 8) + (oPhl * jPhl) + (oB * jB) - (oK * jK) + (oL * jL);
      setVal('lapTotalJam', tot > 0 ? tot : 0);

      const kpiInput = document.getElementById('lapKPI');
      const kpiQtyInput = document.getElementById('lapKPIQty');

      if (div === 'Inbound' || div === 'Return' || div === 'KOL') {
        setVal('lapKPI', '-');
        if (kpiInput) kpiInput.style.opacity = '0.3';
      } else {
        setVal('lapKPI', tot > 0 ? (output / tot).toFixed(3) : 0);
        if (kpiInput) kpiInput.style.opacity = '1';
      }

      const lWrap = document.getElementById('lapLemburNamesWrap');
      if (oL > 0) {
        lWrap.style.display = 'block';
      } else {
        lWrap.style.display = 'none';
      }

      const lReasonWrap = document.getElementById('wrapLapAlasanLembur');
      if (oL > 0 || jL > 0) {
        lReasonWrap.style.display = 'block';
      } else {
        lReasonWrap.style.display = 'none';
      }

      const kWrap = document.getElementById('wrapLapAlasanKurang');
      if (oK > 0) {
        kWrap.style.display = 'block';
      } else {
        kWrap.style.display = 'none';
      }
    }
    function openLaporanKerjaModal() {
      // Reset fields
      resetForm(['lapOrang', 'lapAdmin', 'lapPhl', 'lapJamPhl', 'lapOrangBantu', 'lapJamBantu', 'lapOrangKurang', 'lapJamKurang', 'lapAlasanKurang', 'lapOrangLembur', 'lapLembur', 'lapAlasanLembur', 'lapOrder', 'lapPo', 'lapQty', 'lapInbound', 'lapQtyInb', 'lapKendala', 'lapPotongBubble', 'lapBuatBubble']);
      setVal('lapTotalJam', 0);
      setVal('lapKPI', 0);
      setVal('lapPic', currentUser ? currentUser.nama : '');
      setVal('lapTanggal', new Date().toISOString().split('T')[0]);

      // Ensure staff list is populated immediately
      populateLemburStaffList();

      // Hide lembur wrap initially
      document.getElementById('lapLemburNamesWrap').style.display = 'none';

      // Uncheck all checkboxes
      document.querySelectorAll('input[name="staffLembur"]').forEach(cb => cb.checked = false);

      setVal('lapEditId', '');
      document.getElementById('lapModalTitle').textContent = '📝 Input Laporan Pengerjaan Orderan';
      document.querySelector('#modalLaporanKerja .btn-primary').textContent = '💾 Simpan Laporan';

      updateLapFields();
      openModal('modalLaporanKerja');
    }

    function editLaporanKerja(id) {
      const d = laporanKerjaData.find(x => x.id === id);
      if (!d) return;

      setVal('lapEditId', d.id);
      document.getElementById('lapModalTitle').textContent = '✏️ Edit Laporan Kerja';
      document.querySelector('#modalLaporanKerja .btn-primary').textContent = '💾 Update Laporan';

      setVal('lapTanggal', d.tanggal);
      setVal('lapDivisi', d.divisi);
      updateLapFields(); // Panggil segera agar UI sinkron

      setVal('lapPic', d.pic);
      setVal('lapShift', d.shift || 'Pagi');
      setVal('lapOrang', d.totalOrang);
      setVal('lapAdmin', d.totalAdmin || 0);
      setVal('lapPhl', d.totalPHL || 0);
      setVal('lapJamPhl', d.jamKerjaPHL || 0);

      setVal('lapOrangBantu', (d.perbantuan > 0 ? 1 : 0));
      setVal('lapJamBantu', d.perbantuan);
      setVal('lapOrangKurang', (d.pengurangan > 0 ? 1 : 0));
      setVal('lapJamKurang', d.pengurangan);

      setVal('lapOrangLembur', d.totalStaff || 0);
      setVal('lapLembur', d.jamLembur);
      setVal('lapAlasanLembur', d.alasanLembur || '');

      setVal('lapQtyInb', d.totalQty || 0);
      setVal('lapKendala', d.kendala ? d.kendala.split('\n')[0] : '');
      setVal('lapAlasanKurang', d.alasanPengurangan || '');
      setVal('lapPotongBubble', d.pendapatanPotongBubble || 0);
      setVal('lapBuatBubble', d.pendapatanBuatBubble || 0);

      // Restore staff selection
      selectedLemburStaff = new Set();
      if (d.staffLemburNames) {
        d.staffLemburNames.split(',').forEach(n => selectedLemburStaff.add(n.trim()));
      }

      updateLapFields();
      populateLemburStaffList();
      openModal('modalLaporanKerja');
    }

    let selectedLemburStaff = new Set();
    function populateLemburStaffList() {
      const cont = document.getElementById('lapStaffLemburList');
      const search = document.getElementById('lapStaffSearch')?.value.toLowerCase() || '';
      const filterJab = document.getElementById('lapStaffJabatan')?.value || '';
      const selJab = document.getElementById('lapStaffJabatan');

      if (!cont) return;

      if (!karyawanData || karyawanData.length === 0) {
        cont.innerHTML = '<div style="color:var(--gray); grid-column: 1/-1; text-align:center; padding:10px;">⏳ Memuat data karyawan...</div>';
        google.script.run.withSuccessHandler(res => {
          if (res.success) {
            karyawanData = res.data;
            populateLemburStaffList();
          }
        }).getKaryawan();
        return;
      }

      // Populate jabatan filter options once
      if (selJab && selJab.options.length <= 1) {
        const jabatans = [...new Set(karyawanData.map(k => k.jabatan).filter(j => j))].sort();
        jabatans.forEach(j => {
          const opt = document.createElement('option'); opt.value = j; opt.textContent = j;
          selJab.appendChild(opt);
        });
      }

      // Filter data
      const filtered = karyawanData.filter(k => {
        const matchesSearch = k.nama.toLowerCase().includes(search);
        const matchesJab = filterJab === '' || k.jabatan === filterJab;
        return matchesSearch && matchesJab;
      });

      // Group by Level/Jabatan
      const grouped = {};
      filtered.forEach(k => {
        const level = k.jabatan || 'Lainnya';
        if (!grouped[level]) grouped[level] = [];
        grouped[level].push(k);
      });

      cont.innerHTML = '';
      if (filtered.length === 0) {
        cont.innerHTML = '<div style="color:var(--gray); grid-column: 1/-1; text-align:center; padding:10px;">Tidak ada karyawan yang cocok.</div>';
        return;
      }

      Object.keys(grouped).sort().forEach(level => {
        const levelHeader = document.createElement('div');
        levelHeader.style.cssText = 'grid-column: 1/-1; margin-top:10px; padding-bottom:4px; border-bottom:1px solid #ffffff10; color:var(--accent); font-size:11px; font-weight:800; text-transform:uppercase;';
        levelHeader.textContent = level;
        cont.appendChild(levelHeader);

        grouped[level].sort((a, b) => a.nama.localeCompare(b.nama)).forEach(k => {
          const isChecked = selectedLemburStaff.has(k.nama);
          const label = document.createElement('label');
          label.style.cssText = 'display:flex; align-items:center; gap:8px; cursor:pointer; font-size:12px; color:var(--white); padding:4px 0;';
          label.innerHTML = `<input type="checkbox" name="staffLembur" value="${k.nama}" ${isChecked ? 'checked' : ''} onchange="syncLemburCount()" style="accent-color:var(--accent);"> ${k.nama}`;
          cont.appendChild(label);
        });
      });
    }

    function syncLemburCount() {
      // Perbarui set staff terpilih
      document.querySelectorAll('input[name="staffLembur"]').forEach(cb => {
        if (cb.checked) selectedLemburStaff.add(cb.value);
        else selectedLemburStaff.delete(cb.value);
      });

      const count = selectedLemburStaff.size;
      setVal('lapOrangLembur', count);
      calcJamKerja();
    }
    function loadLaporanKerja() {
      // Ensure karyawanData is loaded for overtime selection
      if (!karyawanData || karyawanData.length === 0) {
        google.script.run.withSuccessHandler(res => { if (res.success) karyawanData = res.data; }).getKaryawan();
      }

      google.script.run.withSuccessHandler(res => {
        laporanKerjaData = res.data || []; renderLaporanTable(); renderLemburSummaryDashboard();
      }).getLaporanKerja();
    }

    function renderLaporanTable() {
      const tb = document.getElementById('tableLaporanKerja'); tb.innerHTML = '';
      if (!laporanKerjaData.length) {
        tb.innerHTML = '<tr><td colspan="13" class="empty-state">Belum ada laporan</td></tr>';
        document.getElementById('lapLemburSummaryList').innerHTML = '<div style="color:var(--gray)">Tidak ada data lembur harian</div>';
        return;
      }

      const grouped = {};
      const filterDiv = document.getElementById('filterLapDivisi')?.value;

      laporanKerjaData.forEach(d => {
        if (filterDiv && d.divisi !== filterDiv) return;
        const t = d.tanggal;
        if (!grouped[t]) {
          grouped[t] = {
            tanggal: t,
            divisi: new Set(),
            pic: new Set(),
            totalJamKerja: 0,
            jamLembur: 0,
            totalPHL: 0,
            totalStaffMasuk: 0,
            totalOrangMasuk: 0,
            totalOrder: 0,
            totalPO: 0,
            totalInbound: 0,
            totalQty: 0,
            staffLembur: new Set(),
            alasanLembur: new Set(),
            ids: []
          };
        }
        const g = grouped[t];
        g.divisi.add(d.divisi);
        g.pic.add(d.pic);
        if (d.staffLemburNames) {
          d.staffLemburNames.split(',').forEach(n => {
            const name = n.trim();
            if (name) g.staffLembur.add(name);
          });
        }
        g.totalJamKerja += (d.totalJamKerja || 0);
        g.jamLembur += (d.jamLembur || 0);
        g.totalPHL += (d.totalPHL || 0);
        const currentStaffMasuk = (d.totalOrang || 0) + (d.totalAdmin || 0);
        g.totalStaffMasuk += currentStaffMasuk;
        g.totalOrangMasuk += (currentStaffMasuk + (d.totalPHL || 0));
        if (d.alasanLembur) g.alasanLembur.add(d.alasanLembur);
        g.totalOrder += (d.totalOrder || 0);
        g.totalPO += (d.totalPo || 0);
        g.totalInbound += (d.totalInbound || 0);
        g.totalQty += (d.totalQty || (d.totalQtyInb || 0) || 0);

        // Track specific Distributor Qty for group KPI
        if (d.divisi === 'Distributor' || d.divisi === 'Distributor SBY' || d.divisi === 'Distributor Surabaya') {
          if (!g.distQty) g.distQty = 0;
          g.distQty += (d.totalQty || 0);
        }

        g.ids.push(d.id);
      });

      const sortedDates = Object.keys(grouped).sort((a, b) => new Date(b) - new Date(a));

      sortedDates.forEach(t => {
        const g = grouped[t];
        const divList = Array.from(g.divisi);
        const divStr = divList.join(', ');
        const picStr = Array.from(g.pic).join(', ');

        // Output sums: MP Orders + Dist Qty (user requirement) + Inb SJ
        const totalOutput = g.totalOrder + (g.distQty || 0) + g.totalInbound;

        // KPI calculation
        let kpi = '-';
        if (g.totalJamKerja > 0) {
          const hideKPI = divList.every(d => ['Inbound', 'Return', 'KOL'].includes(d));
          const rawKpi = !hideKPI ? (totalOutput / g.totalJamKerja) : 0;
          kpi = rawKpi > 0 ? Math.round(rawKpi) : '-';
        }

        const alasanHtml = Array.from(g.alasanLembur)
          .filter(Boolean)
          .map(a => `<span style="display:inline-block; background: rgba(255,255,255,0.08); border: 1px solid rgba(255,255,255,0.12); border-radius: 999px; padding: 4px 10px; margin: 2px 2px 2px 0; font-size: 11px; color: var(--text); max-width: 170px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${a}</span>`)
          .join('');

        tb.innerHTML += `<tr>
      <td>${formatDate(t)}</td>
      <td>
        <div style="font-weight:700; color:var(--accent);">${divStr}</div>
        <div style="font-size:11px; color:var(--text-muted);">${picStr}</div>
      </td>
      <td align="center"><strong>${g.totalJamKerja.toFixed(1)}</strong></td>
      <td align="center" style="color:var(--orange)">
        <strong>${g.jamLembur.toFixed(1)}</strong>
        ${g.staffLembur.size > 0 ? `<div style="font-size:10px; opacity:0.8;">(${g.staffLembur.size} Orang)</div>` : ''}
      </td>
      <td align="left" style="min-width:180px;">
        ${alasanHtml || '<span style="color:var(--gray); font-size:11px;">-</span>'}
      </td>
      <td align="center">${g.totalPHL} Org</td>
      <td align="center"><strong>${g.totalStaffMasuk}</strong> Org</td>
      <td align="center" style="color:var(--green); font-weight:700;">${g.totalOrangMasuk} Org</td>
      <td align="center">
        <div style="color:var(--teal); font-weight:700;">${totalOutput.toLocaleString('id-ID')}</div>
        <div style="font-size:10px; color:var(--gray)">${g.totalOrder > 0 ? 'Ord:' + g.totalOrder : ''} ${g.totalPO > 0 ? 'PO:' + g.totalPO : ''} ${g.totalInbound > 0 ? 'Inb:' + g.totalInbound : ''}</div>
      </td>
      <td align="center">${g.totalQty.toLocaleString('id-ID')}</td>
      <td align="center"><span class="badge" style="background:var(--green)22; color:var(--green);">${kpi}</span></td>
      <td>
        <button class="btn btn-ghost btn-sm" onclick="showDailyDetail('${t}')" title="Detail Harian">📋</button>
      </td>
    </tr>`;
      });
    }

    function showDailyDetail(tanggal) {
      const reports = laporanKerjaData.filter(d => d.tanggal === tanggal);
      if (!reports.length) return;

      document.getElementById('detailLapTitle').textContent = `Laporan Harian - ${formatDate(tanggal)}`;
      let html = `<div style="display:flex; flex-direction:column; gap:15px;">`;

      reports.forEach(r => {
        html += `<div style="background:var(--navy2); padding:15px; border-radius:10px; border:1px solid #ffffff15;">
      <div style="display:flex; justify-content:space-between; margin-bottom:10px;">
        <strong style="color:var(--accent)">${r.divisi} (${r.shift})</strong>
        <small style="color:var(--gray)">PIC: ${r.pic}</small>
      </div>
      <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px; font-size:12px;">
        <div>Staff: <b>${(r.totalOrang || 0) + (r.totalAdmin || 0)}</b> <span style="font-size:10px; color:var(--gray)">(${r.totalOrang}p + ${r.totalAdmin || 0}a)</span></div>
        <div>PHL: <b>${r.totalPHL || 0}</b></div>
        <div>Total Masuk: <b style="color:var(--green)">${((r.totalOrang || 0) + (r.totalAdmin || 0)) + (r.totalPHL || 0)}</b></div>
        <div>Lembur: ${r.jamLembur} Jam</div>
        <div style="grid-column: span 2; border-top:1px solid #ffffff0a; padding-top:5px; margin-top:5px;">
          Output: <b>${['Distributor', 'Distributor SBY', 'Distributor Surabaya'].includes(r.divisi) ? r.totalQty : (r.totalOrder || r.totalPO || r.totalInbound || 0)}</b> 
          ${!['Inbound', 'Return', 'KOL'].includes(r.divisi) ? `| KPI: <b>${r.totalJamKerja > 0 ? ((['Distributor', 'Distributor SBY', 'Distributor Surabaya'].includes(r.divisi) ? r.totalQty : (r.totalOrder || r.totalPO || r.totalInbound || 0)) / r.totalJamKerja).toFixed(3) : 0}</b>` : ''}
        </div>
      </div>
      ${r.staffLemburNames ? `
      <div style="margin-top:12px; padding-top:8px; border-top:1px dashed #ffffff15;">
        <div style="font-size:11px; color:var(--accent); font-weight:700; margin-bottom:5px;">👥 Personil Lembur:</div>
        <div style="display:flex; flex-wrap:wrap; gap:5px;">
          ${r.staffLemburNames.split(',').map(n => `<span style="background:var(--accent)22; color:var(--accent); padding:2px 8px; border-radius:4px; font-size:10px; border:1px solid var(--accent)33;">${n.trim()}</span>`).join('')}
        </div>
      </div>` : ''}
      ${r.alasanLembur ? `<div style="margin-top:12px; font-size:12px; color:var(--green); font-weight:700;">Alasan Lembur: ${r.alasanLembur}</div>` : ''}
      ${r.kendala ? `<div style="margin-top:10px; font-size:11px; color:var(--amber); font-style:italic;">Kendala: ${r.kendala}</div>` : ''}
      <div style="margin-top:10px; display:flex; justify-content:flex-end; gap:8px;">
         <button class="btn btn-ghost btn-sm" style="padding:2px 10px; font-size:11px; border:1px solid #ffffff22;" onclick="closeModal('modalDetailLaporan'); editLaporanKerja('${r.id}')">✏️ Edit</button>
         <button class="btn btn-danger btn-sm" style="padding:2px 8px; font-size:10px;" onclick="delLaporanKerja('${r.id}')">🗑️ Hapus Per Shift</button>
      </div>
    </div>`;
      });

      html += `</div>`;
      document.getElementById('detailLapContent').innerHTML = html;
      openModal('modalDetailLaporan');
    }
    function renderLemburSummaryDashboard() {
      const sumList = document.getElementById('lapLemburSummaryList');
      const dateInfo = document.getElementById('lapCurrentDate');

      // Selalu gunakan tanggal hari ini (Sync)
      const today = new Date().toISOString().split('T')[0];
      dateInfo.textContent = formatDate(today);

      if (!laporanKerjaData || !laporanKerjaData.length) {
        sumList.innerHTML = `<div style="font-size:13px; color:var(--gray);">Belum ada laporan untuk hari ini (${formatDate(today)}).</div>`;
        return;
      }

      const lapForDate = laporanKerjaData.filter(d => d.tanggal === today);
      let allNames = [];
      lapForDate.forEach(l => {
        if (l.staffLemburNames) {
          allNames = [...allNames, ...l.staffLemburNames.split(',')];
        }
      });

      const uniqueNames = [...new Set(allNames)].filter(n => n.trim() !== '');

      if (!uniqueNames.length) {
        sumList.innerHTML = `<div style="font-size:13px; color:var(--gray);">Tidak ada personil lembur yang tercatat hari ini (${formatDate(today)}).</div>`;
      } else {
        sumList.innerHTML = uniqueNames.sort().map(name => {
          return `<div style="background:rgba(245,158,11,0.15); border:1px solid var(--accent); color:var(--accent); font-size:12px; font-weight:700; padding:6px 14px; border-radius:30px;">⚡ ${name.trim()}</div>`;
        }).join('');
      }
    }
    function submitLaporanKerja() {
      const t = v('lapTanggal'), d = v('lapDivisi'), pic = v('lapPic'), shf = v('lapShift');
      const utama = v('lapOrang'), admin = v('lapAdmin'), phl = v('lapPhl'), jPhl = v('lapJamPhl');
      const oB = parseInt(v('lapOrangBantu')) || 0, jB = parseFloat(v('lapJamBantu')) || 0;
      const oK = parseInt(v('lapOrangKurang')) || 0, jK = parseFloat(v('lapJamKurang')) || 0;
      const oL = parseInt(v('lapOrangLembur')) || 0, jL = parseFloat(v('lapLembur')) || 0;

      const pBubble = v('lapPotongBubble'), bBubble = v('lapBuatBubble'), alsK = v('lapAlasanKurang'), alsL = v('lapAlasanLembur');
      const tot = v('lapTotalJam'); let ken = v('lapKendala');
      const ord = parseInt(v('lapOrder')) || 0, po = parseInt(v('lapPo')) || 0, qty = parseInt(v('lapQty')) || 0;
      const inb = parseInt(v('lapInbound')) || 0, qtyInb = parseInt(v('lapQtyInb')) || 0;

      const selectedStaffArr = Array.from(document.querySelectorAll('input[name="staffLembur"]:checked')).map(el => el.value);
      const selectedStaff = selectedStaffArr.join(',');

      const det = `[Bantu: ${oB}orgx${jB}j | Kurang: ${oK}orgx${jK}j | Lembur: ${oL}orgx${jL}j | Shift: ${shf} | PHL: ${phl}org]`;
      const finalKen = (ken ? ken + '\n' : '') + (alsK ? 'Alasan Kurang: ' + alsK + '\n' : '') + det;

      if (!t || !pic) return toast('Lengkapi data', 'error');
      if (oK > 0 && !alsK) return toast('Alasan Pengurangan Orang wajib diisi!', 'error');
      if (oL > 0 && !alsL) return toast('Alasan Lembur wajib diisi jika lembur > 0!', 'error');

      const editId = v('lapEditId');

      // Duplicate Check (Client Side)
      if (!editId && laporanKerjaData) {
        const duplicate = laporanKerjaData.find(x => x.tanggal === t && x.divisi === d && x.shift === shf);
        if (duplicate) return toast(`Laporan untuk ${d} (${shf}) tanggal ${t} sudah ada!`, 'error');
      }

      if (oL > 0 && selectedStaff === '') return toast('Pilih minimal 1 personil lembur', 'error');
      if (oL !== selectedStaffArr.length) return toast(`Jumlah Lembur (${oL}) tidak sesuai dengan jumlah nama yang diceklis (${selectedStaffArr.length})!`, 'error');

      const btn = document.querySelector('#modalLaporanKerja .btn-primary'); btn.disabled = true; btn.textContent = '⏳ Memproses...';

      const successCb = res => {
        btn.disabled = false; btn.textContent = editId ? '💾 Update Laporan' : '💾 Simpan Laporan';
        if (res.success) {
          toast(editId ? 'Laporan Diperbarui' : 'Laporan Berhasil Disimpan');
          closeModal('modalLaporanKerja'); loadLaporanKerja(); loadDashboard();
          resetForm(['lapPic', 'lapKendala', 'lapOrder', 'lapPo', 'lapQty', 'lapInbound', 'lapQtyInb', 'lapPotongBubble', 'lapBuatBubble', 'lapAlasanKurang', 'lapAlasanLembur']);
          selectedLemburStaff = new Set();
        } else toast(res.message, 'error');
      };

      if (editId) {
        const finalQty = (d === 'Inbound') ? qtyInb : qty;
        google.script.run.withSuccessHandler(successCb).updateLaporanKerja(editId, t, d, pic, utama, oB * jB, oK * jK, jL, tot, finalKen, oL, admin, ord, currentUser.username, 0, selectedStaff, shf, phl, jPhl, po, finalQty, inb, pBubble, bBubble, alsK, alsL);
      } else {
        const finalQty = (d === 'Inbound') ? qtyInb : qty;
        google.script.run.withSuccessHandler(successCb).addLaporanKerja(t, d, pic, utama, oB * jB, oK * jK, jL, tot, finalKen, oL, admin, ord, currentUser.username, 0, selectedStaff, shf, phl, jPhl, po, finalQty, inb, pBubble, bBubble, alsK, alsL);
      }
    }

    function printLaporanHarian() {
      const content = document.getElementById('detailLapContent').innerHTML;
      const title = document.getElementById('detailLapTitle').textContent;
      const printWin = window.open('', '_blank', 'width=900,height=700');

      printWin.document.write(`
        <html>
          <head>
            <title>${title}<\/title>
            <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&display=swap" rel="stylesheet">
            <style>
              body { font-family: 'Outfit', sans-serif; padding: 40px; color: #333; line-height: 1.6; }
              strong { color: #000; }
              .header { text-align: center; margin-bottom: 30px; border-bottom: 2px solid #333; padding-bottom: 10px; }
              h1 { font-size: 24px; margin: 0; }
              .shift-box { border: 1px solid #ccc; padding: 20px; border-radius: 8px; margin-bottom: 20px; page-break-inside: avoid; }
              .shift-title { font-size: 18px; font-weight: 700; border-bottom: 1px solid #eee; margin-bottom: 15px; padding-bottom: 5px; }
              .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; font-size: 14px; }
              .personil-list { margin-top: 15px; padding-top: 10px; border-top: 1px dashed #eee; font-size: 12px; }
              .badge { display: inline-block; background: #f0f0f0; border: 1px solid #ddd; padding: 2px 8px; border-radius: 4px; margin: 2px; }
              @media print {
                .no-print { display: none; }
                body { padding: 0; }
              }
            <\/style>
          <\/head>
          <body>
            <div class="header">
              <h1>WAREHOUSE FCL - LAPORAN KERJA HARIAN<\/h1>
              <div style="font-size: 14px; margin-top: 5px;">${title}<\/div>
            <\/div>
            ${content.replace(/btn-danger|btn-ghost/g, 'no-print')}
            <div style="margin-top: 50px; text-align: right; font-size: 12px;">
              Dicetak pada: ${new Date().toLocaleString('id-ID')} oleh ${currentUser.nama}
            <\/div>
            \x3Cscript>
              setTimeout(() => { window.print(); window.close(); }, 500);
            <\/script>
          <\/body>
        <\/html>
      `);
      printWin.document.close();
    }
    function checkLaporanExisting() {
      // Fungsi placeholder untuk menangani event onchange di input tanggal modal laporan kerja
      console.log("Checking existing reports for date:", v('lapTanggal'));
    }
    function delLaporanKerja(id) { if (confirm('Hapus?')) google.script.run.withSuccessHandler(res => { if (res.success) { toast('Dihapus'); loadLaporanKerja(); } else toast(res.message, 'error'); }).deleteLaporanKerja(id); }
    function showDetailLaporan(id) {
      const d = laporanKerjaData.find(x => x.id === id); if (!d) return;
      const staffMasuk = (parseInt(d.totalOrang) || 0) + (parseInt(d.totalAdmin) || 0);
      const totalMasuk = staffMasuk + (parseInt(d.totalPHL) || 0);
      document.getElementById('detailLapTitle').textContent = `Laporan ${d.divisi} - ${formatDate(d.tanggal)}`;
      document.getElementById('detailLapContent').innerHTML = `<b>PIC:</b> ${d.pic}<br><b>Staff Masuk:</b> ${staffMasuk} <span style="font-size:11px; color:var(--gray)">(${d.totalOrang}p + ${d.totalAdmin || 0}a)</span><br><b>PHL:</b> ${d.totalPHL || 0} | <b>Total Masuk:</b> <b style="color:var(--green)">${totalMasuk}</b><br><b>Total Jam Kerja:</b> ${d.totalJamKerja} Jam<br><b>Total Order:</b> ${d.totalOrder} ${!['Inbound', 'Return', 'KOL'].includes(d.divisi) ? `| <b>KPI:</b> ${d.totalJamKerja > 0 ? (d.totalOrder / d.totalJamKerja).toFixed(3) : 0}` : ''}<br><b>Kendala & Detail:</b><br><pre style="white-space:pre-wrap;background:#0f2040;padding:10px;border-radius:8px;border:1px solid #ffffff15;margin-top:5px;font-family:inherit;">${d.kendala || '-'}</pre>`;
      openModal('modalDetailLaporan');
    }

    // GRAFIK
    function loadGrafikLaporan() { google.script.run.withSuccessHandler(res => { if (res.success) { laporanKerjaData = res.data; filterGrafik(); } }).getLaporanKerja(); }
    function filterGrafik() {
      const ym = v('filterBulanGrafik'); const y = ym.split('-')[0], m = ym.split('-')[1];
      const filtered = laporanKerjaData.filter(d => d.tanggal.startsWith(ym));
      const maps = { Marketplace: {}, Distributor: {}, "Market Place SBY": {}, "Distributor SBY": {}, "Marketplace Surabaya": {}, "Distributor Surabaya": {}, Return: {}, KOL: {}, Inbound: {}, Consumable: {} };
      const dInM = new Date(y, m, 0).getDate();
      for (let i = 1; i <= dInM; i++) { const ds = i.toString().padStart(2, '0'); Object.keys(maps).forEach(k => maps[k][ds] = { ord: 0, jm: 0 }); }
      filtered.forEach(d => {
        const day = d.tanggal.split('-')[2].substring(0, 2);
        if (maps[d.divisi] && maps[d.divisi][day]) {
          const currentOutput = (d.divisi === 'Distributor' || d.divisi === 'Distributor SBY' || d.divisi === 'Distributor Surabaya') ? (d.totalQty || 0) : (d.divisi === 'Inbound' ? (d.totalInbound || 0) : (d.divisi === 'Consumable' ? ((d.pendapatanPotongBubble || 0) + (d.pendapatanBuatBubble || 0)) : (d.totalOrder || 0)));
          maps[d.divisi][day].ord += currentOutput;
          maps[d.divisi][day].jm += d.totalJamKerja;
        }
      });

      const labels = Object.keys(maps.Marketplace).sort();
      const cData = div => ({
        ord: labels.map(l => maps[div][l].ord),
        kpi: labels.map(l => {
          const o = maps[div][l].ord, j = maps[div][l].jm;
          return j > 0 ? Math.round(o / j) : 0;
        })
      });

      createChart('chartMarketplace', 'mp', labels, cData('Marketplace').ord, cData('Marketplace').kpi, '#f59e0b');
      createChart('chartDistributor', 'dist', labels, cData('Distributor').ord, cData('Distributor').kpi, '#0ea5e9');

      // Separate Surabaya Charts
      createChart('chartMarketplaceSurabaya', 'mpsby', labels, cData('Market Place SBY').ord, cData('Market Place SBY').kpi, '#38bdf8');
      createChart('chartDistributorSurabaya', 'distsby', labels, cData('Distributor SBY').ord, cData('Distributor SBY').kpi, '#0284c7');

      // Historical Surabaya (Legacy) - remains as fallback/audit
      // if (cData('Marketplace Surabaya').ord.some(v => v > 0)) { ... }

      createChart('chartReturn', 'ret', labels, cData('Return').ord, cData('Return').kpi, '#ef4444');
      if (cData('KOL').ord.length) createChart('chartKOL', 'kol', labels, cData('KOL').ord, cData('KOL').kpi, '#10b981');
      if (cData('Inbound').ord.length) createChart('chartInbound', 'inb', labels, cData('Inbound').ord, cData('Inbound').kpi, '#8b5cf6');
      if (cData('Consumable').ord.length) createChart('chartConsumable', 'con', labels, cData('Consumable').ord, cData('Consumable').kpi, '#f472b6');

      const summary = { Marketplace: { rep: 0, p: 0, l: 0, j: 0, o: 0 }, Distributor: { rep: 0, p: 0, l: 0, j: 0, o: 0 }, "Market Place SBY": { rep: 0, p: 0, l: 0, j: 0, o: 0 }, "Distributor SBY": { rep: 0, p: 0, l: 0, j: 0, o: 0 }, "Marketplace Surabaya": { rep: 0, p: 0, l: 0, j: 0, o: 0 }, "Distributor Surabaya": { rep: 0, p: 0, l: 0, j: 0, o: 0 }, Return: { rep: 0, p: 0, l: 0, j: 0, o: 0 }, KOL: { rep: 0, p: 0, l: 0, j: 0, o: 0 }, Inbound: { rep: 0, p: 0, l: 0, j: 0, o: 0 }, Consumable: { rep: 0, p: 0, l: 0, j: 0, o: 0 } };
      filtered.forEach(d => {
        if (summary[d.divisi]) {
          summary[d.divisi].rep++;
          const staffMasuk = (d.totalOrang || 0) + (d.totalAdmin || 0);
          const totalMasuk = staffMasuk + (d.totalPHL || 0);
          summary[d.divisi].p += totalMasuk;
          summary[d.divisi].l += d.jamLembur;
          summary[d.divisi].j += d.totalJamKerja;

          const currentOutput = (d.divisi === 'Distributor' || d.divisi === 'Distributor SBY' || d.divisi === 'Distributor Surabaya') ? (d.totalQty || 0) : (d.divisi === 'Inbound' ? (d.totalInbound || 0) : (d.divisi === 'Consumable' ? ((d.pendapatanPotongBubble || 0) + (d.pendapatanBuatBubble || 0)) : (d.totalOrder || 0)));
          summary[d.divisi].o += currentOutput;
        }
      });
      const tb = document.getElementById('tableRingkasanGrafik'); tb.innerHTML = '';
      Object.keys(summary).forEach(k => {
        const s = summary[k];
        let displayOutput = s.o.toLocaleString('id-ID');
        let displayKpi = '-';

        const rawKpi = (s.j > 0 && !['Inbound', 'Return', 'KOL', 'Consumable'].includes(k)) ? (s.o / s.j) : 0;
        displayKpi = rawKpi > 0 ? Math.round(rawKpi) : '-';

        let kpiStyle = 'color:var(--green)';
        let targetMsg = '';

        if (displayKpi !== '-') {
          if ((k === 'Marketplace' || k === 'Market Place SBY' || k === 'Marketplace Surabaya') && displayKpi < 95) {
            kpiStyle = 'color:var(--red)';
            targetMsg = '<br><small style="color:var(--red); font-size:9px; font-weight:800;">(Target Tidak Tercapai)</small>';
          } else if ((k === 'Distributor' || k === 'Distributor SBY' || k === 'Distributor Surabaya') && displayKpi < 1875) {
            kpiStyle = 'color:var(--red)';
            targetMsg = '<br><small style="color:var(--red); font-size:9px; font-weight:800;">(Target Tidak Tercapai)</small>';
          }
        }

        tb.innerHTML += `<tr><td><strong>${k}</strong></td><td>${s.rep} Laporan</td><td>${s.rep > 0 ? Math.round(s.p / s.rep) : 0} Org/hari</td><td>${s.l} Jam</td><td>${s.j} Jam</td><td style="color:var(--teal)"><strong>${displayOutput}</strong></td><td style="${kpiStyle}"><strong>${displayKpi}</strong>${targetMsg}</td></tr>`;
      });
    }
    function createChart(id, instKey, labels, dOrd, dKpi, color) {
      const typeLabel = (instKey === 'dist' || instKey === 'con') ? (instKey === 'dist' ? 'Total Qty' : 'Total Bubble') : 'Total Order';
      const kpiLabel = (instKey === 'dist' || instKey === 'con') ? (instKey === 'dist' ? 'KPI (Qty/Jam)' : 'KPI (Bbl/Jam)') : 'KPI (Ord/Jam)';
      const ctx = document.getElementById(id); if (chartInstances[instKey]) chartInstances[instKey].destroy();

      // Determine target coloring for points
      const target = (instKey === 'mp' || instKey === 'mpsby') ? 95 : (instKey === 'dist' || instKey === 'distsby' ? 1875 : 0);
      const pointColors = dKpi.map(v => (target > 0 && v < target) ? '#ef4444' : '#10b981');
      const pointBorderColors = dKpi.map(v => (target > 0 && v < target) ? '#ef4444' : '#10b981');

      chartInstances[instKey] = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: labels,
          datasets: [
            {
              label: typeLabel,
              data: dOrd,
              backgroundColor: color + '88',
              borderColor: color,
              borderWidth: 1,
              yAxisID: 'y'
            },
            {
              label: kpiLabel,
              data: dKpi,
              type: 'line',
              borderColor: '#10b981',
              backgroundColor: '#10b981',
              pointBackgroundColor: pointColors,
              pointBorderColor: pointBorderColors,
              pointRadius: 4,
              pointHoverRadius: 6,
              borderWidth: 2,
              tension: 0.3,
              yAxisID: 'y1'
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          scales: {
            y: { type: 'linear', display: true, position: 'left', grid: { color: '#ffffff10' } },
            y1: { type: 'linear', display: true, position: 'right', grid: { drawOnChartArea: false } },
            x: { grid: { color: '#ffffff10' } }
          },
          plugins: {
            legend: { labels: { color: '#e2e8f0' } }
          }
        }
      });
    }


    // STOCK CONTROL (REPLACING HANDOVER)
    function loadStockControl() {
      google.script.run.withSuccessHandler(res => {
        const tb = document.getElementById('tableHandover');
        tb.innerHTML = '';
        if (!res.data || !res.data.length) {
          tb.innerHTML = '<tr><td colspan="8" class="empty-state">Kosong</td></tr>';
        } else {
          res.data.forEach(d => {
            const safeArea = (d.area || '').toString().replace(/'/g, "\\'").replace(/\n/g, " ");
            const safeAlasan = (d.alasan || '').toString().replace(/'/g, "\\'").replace(/\n/g, " ");
            const safeKaryawan = (d.karyawan || '').toString().replace(/'/g, "\\'").replace(/\n/g, " ");
            const safeLog = (d.syncLog || '').toString().replace(/'/g, "\\'").replace(/\n/g, " ");
            const statusBadge = d.status === 'Disetujui' ? 'bg-success' : d.status === 'Ditolak' ? 'bg-danger' : 'bg-warning';

            tb.innerHTML += `<tr id="sc-master-${d.id}">
              <td>
                <button class="btn-toggle-row" onclick="toggleStockControlDetail('${d.id}','${safeArea}','${d.kategori}','${safeAlasan}','${safeKaryawan}','${safeLog}')">
                  <i class="bi bi-plus-circle"></i>
                </button>
                ${formatDate(d.tanggal)}
              </td>
              <td>${d.pic}</td>
              <td><strong>${d.area}</strong></td>
              <td>${d.kategori}</td>
              <td><span style="font-size:10px">${d.alasan || '-'}</span></td>
              <td><span class="badge ${statusBadge}">${d.status || 'Pending'}</span></td>
              <td style="text-align: center;">
                ${(d.status === 'Menunggu Approval' || d.status === 'Pending' || !d.status) ? `
                  <div class="d-flex gap-1 justify-content-center">
                    <button class="btn btn-success btn-sm" style="padding:2px 5px; font-size:10px;" onclick="updateStockControlStatus('${d.id}', 'Disetujui')">Setuju</button>
                    <button class="btn btn-danger btn-sm" style="padding:2px 5px; font-size:10px;" onclick="updateStockControlStatus('${d.id}', 'Ditolak')">Tolak</button>
                  </div>
                ` : '-'}
              </td>
              <td>
                <div class="d-flex gap-1">
                  <button class="btn btn-ghost btn-sm" onclick="syncSingleStock('${d.id}')" title="Singkron / Hitung Ulang" style="color:var(--teal)">🔄</button>
                  <button class="btn btn-ghost btn-sm" onclick="editStockControl('${d.id}','${d.tanggal}','${d.pic}','${d.area}','${d.kategori}','${safeAlasan}','${safeKaryawan}')" title="Input Admin" style="color:var(--accent)">📝</button>
                  <button class="btn btn-ghost btn-sm" onclick="openStockRecountForm('${d.id}','${d.tanggal}','${d.pic}','${d.area}','${d.kategori}','${safeAlasan}','${safeKaryawan}')" title="Hitung Ulang Staff" style="color:var(--teal)">🧑‍💻</button>
                  <button class="btn btn-danger btn-sm" onclick="delStockControl('${d.id}')" title="Hapus">🗑️</button>
                </div>
              </td>
            </tr>
            <tr id="sc-detail-${d.id}" class="row-detail" style="display:none;">
              <td colspan="8"><div class="detail-container"><div class="loading-inline">Memuat detail...</div></div></td>
            </tr>`;
          });
        }
      }).getStockControl();
    }

    function syncSingleStock(id) {
      toast('Sedang sinkronisasi...', 'info');
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          if (res.varianceChanged) {
            alert('⚠️ PERINGATAN!\n\nSelisih stock berubah dari data sebelumnya. Silakan cek detail item.');
            toast('✓ Selesai dengan peringatan perubahan data', 'warning');
          } else {
            toast('✓ ' + res.message, 'success');
          }
          loadStockControl();
          loadDashboardStock();
        } else {
          toast('Error: ' + res.message, 'error');
        }
      }).recalculateSingleStockControl(id);
    }

    function syncCurrentList() {
      const table = document.getElementById('tableHandover');
      if (!table) return;
      const ids = Array.from(table.querySelectorAll('tr[id^="sc-master-"]'))
        .map(tr => tr.id.replace('sc-master-', ''));

      if (ids.length === 0) {
        toast('Tidak ada data di tabel untuk disinkronkan', 'warning');
        return;
      }

      toast('Memulai sinkronisasi ' + ids.length + ' laporan...', 'info');
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          toast('✓ ' + res.message, 'success');
          loadStockControl();
          loadDashboardStock();
        } else {
          toast('Gagal: ' + res.message, 'error');
        }
      }).bulkRecalculateStockControl(ids);
    }

    function toggleStockControlDetail(id, area, kategori, alasan, karyawan, syncLog) {
      const detailRow = document.getElementById(`sc-detail-${id}`);
      const btn = document.querySelector(`#sc-master-${id} .btn-toggle-row`);
      if (detailRow.style.display === 'none') {
        const cont = detailRow.querySelector('.detail-container');

        if (kategori === 'Laporan Lapangan') {
          cont.innerHTML = `
            <div style="background:rgba(255,255,255,0.03); padding:16px; border-radius:12px; border:1px solid #ffffff10; line-height:1.6;">
              <div style="color:var(--accent); font-weight:800; font-size:12px; margin-bottom:12px; text-transform:uppercase; letter-spacing:1px;">🔎 Detail Temuan Lapangan</div>
              <div class="row">
                <div class="col-md-3 mb-2"><span style="color:var(--gray); font-size:11px;">PIC PELAPOR</span><br><strong>${btn.parentElement.nextElementSibling.textContent}</strong></div>
                <div class="col-md-3 mb-2"><span style="color:var(--gray); font-size:11px;">AREA / POSISI</span><br><strong>${area}</strong></div>
                <div class="col-md-3 mb-2"><span style="color:var(--gray); font-size:11px;">KARYAWAN TERLIBAT</span><br><strong>${karyawan || '-'}</strong></div>
                <div class="col-md-3 mb-2"><span style="color:var(--gray); font-size:11px;">KETERANGAN TEMUAN</span><br><div style="font-size:13px; color:var(--light);">${alasan || '-'}</div></div>
              </div>
            </div>`;
        } else if (cont.innerHTML.includes('Memuat detail...')) {
          google.script.run.withSuccessHandler(res => {
            if (res.success) {
              let hasAnyDiff = res.data.some(it => {
                const hasRecount = it.stockFisikStaff !== undefined && it.stockFisikStaff !== null && it.stockFisikStaff !== '' && it.stockFisikStaff !== 0;
                return hasRecount && (parseFloat(it.stockFisikStaff) !== parseFloat(it.stockFisik));
              });

              let html = `
                <div class="d-flex justify-content-between align-items-center mb-2">
                  <div style="font-weight:700; font-size:11px; color:var(--teal); text-transform:uppercase;">
                    📦 Detail Items - ${area} 
                    ${hasAnyDiff ? '<span style="margin-left:8px; padding:2px 8px; background:rgba(239,68,68,0.15); color:var(--red); border-radius:4px; font-size:9px;">⚠️ TERDETEKSI PERBEDAAN HITUNG STAFF</span>' : ''}
                  </div>
                  <div style="font-size:10px; font-style:italic; color:var(--gray);">${syncLog || ''}</div>
                </div>
                <table class="detail-table" style="font-size:10px; border-collapse: collapse;">
                  <thead>
                    <tr style="background: rgba(0,0,0,0.2);">
                      <th style="width:80px">Lokasi</th>
                      <th style="width:160px">SKU / Batch / Exp</th>
                      <th style="width:100px; text-align:center;">Kategori Data</th>
                      <th style="width:60px">Mabang</th>
                      <th style="width:60px">TTX</th>
                      <th style="width:70px">Fisik</th>
                      <th style="width:100px">Selisih M/T</th>
                      <th style="width:100px">Aksi</th>
                      <th style="width:120px">Alasan</th>
                    </tr>
                  </thead>
                  <tbody>`;
              res.data.forEach(it => {
                const selM = parseFloat(it.selisihMabang) || 0, selT = parseFloat(it.selisihTtx) || 0;
                const smS = parseFloat(it.selisihMabangStaff) || 0, stS = parseFloat(it.selisihTtxStaff) || 0;

                const fA = parseFloat(it.stockFisik) || 0;
                const fS = parseFloat(it.stockFisikStaff) || 0;
                const hasRecount = it.stockFisikStaff !== undefined && it.stockFisikStaff !== null && it.stockFisikStaff !== '' && it.stockFisikStaff !== 0;

                const diffAS = hasRecount ? (fS - fA) : 0;
                const badgeAS = hasRecount ? (diffAS !== 0 ? `<span style="color:var(--red); font-weight:800;">[⚠️ A-S: ${diffAS > 0 ? '+' : ''}${diffAS}]</span>` : '<span style="color:var(--green); font-weight:800;">[✅ A-S: Cocok]</span>') : '';

                // Row 1: ADMIN
                html += `
                <tr style="border-top:1px solid var(--border-color); background:rgba(245,158,11,0.03);">
                  <td rowspan="2" style="vertical-align:middle;">${it.lokasi || '-'}</td>
                  <td rowspan="2" style="vertical-align:middle;">
                    <strong>${it.sku}</strong><br>
                    <small style="color:var(--gray)">B: ${it.batch || '-'} | E: ${it.exp || '-'}</small>
                  </td>
                  <td style="color:var(--accent); font-weight:700; text-align:center;">STOCK AWAL (ADMIN)</td>
                  <td>${it.stockMabang}</td>
                  <td>${it.stockTtx}</td>
                  <td style="font-weight:800; color:var(--accent);">${fA}</td>
                  <td style="color:${(selM !== 0 || selT !== 0) ? 'var(--red)' : 'var(--green)'}; font-size:9px;">M:${selM > 0 ? '+' : ''}${selM} / T:${selT > 0 ? '+' : ''}${selT}</td>
                  <td rowspan="2" style="vertical-align:middle;">${it.aksi || '-'}</td>
                  <td rowspan="2" style="vertical-align:middle;"><small>${it.alasan || '-'}</small></td>
                </tr>
                <tr style="background:rgba(14,165,233,0.03);">
                  <td style="color:var(--teal); font-weight:700; text-align:center;">STOCK STAFF ${badgeAS}</td>
                  <td>${it.stockMabang}</td>
                  <td>${it.stockTtx}</td>
                  <td style="font-weight:800; color:var(--teal);">${it.stockFisikStaff || '-'}</td>
                  <td style="color:${(smS !== 0 || stS !== 0) ? 'var(--red)' : 'var(--green)'}; font-size:9px;">${it.stockFisikStaff ? `M:${smS > 0 ? '+' : ''}${smS} / T:${stS > 0 ? '+' : ''}${stS}` : '-'}</td>
                </tr>`;
              });
              html += `</tbody></table>`;
              cont.innerHTML = html;
            } else {
              cont.innerHTML = `<span style="color:var(--red)">Gagal memuat: ${res.message}</span>`;
            }
          }).getStockControlDetail(id);
        }
        detailRow.style.display = 'table-row';
        btn.innerHTML = '<i class="bi bi-dash-circle"></i>';
        btn.classList.add('active');
      } else {
        detailRow.style.display = 'none';
        btn.innerHTML = '<i class="bi bi-plus-circle"></i>';
        btn.classList.remove('active');
      }
    }

    function addStockControlInitialRow(data) {
      const tb = document.getElementById('hoInitialItemsList');
      const row = document.createElement('tr');
      const selM = (parseFloat(data.stockFisik || 0) - parseFloat(data.stockMabang || 0));
      const selT = (parseFloat(data.stockFisik || 0) - parseFloat(data.stockTtx || 0));
      row.innerHTML = `
        <td>${data.lokasi || '-'}</td>
        <td><strong>${data.sku}</strong></td>
        <td>${data.batch || '-'}</td>
        <td>${data.exp || '-'}</td>
        <td>${data.stockMabang}</td>
        <td>${data.stockTtx}</td>
        <td><strong style="color:var(--white)">${data.stockFisik}</strong></td>
        <td style="color:${(selM !== 0 || selT !== 0) ? '#ef4444' : '#10b981'}; font-weight:600;">M:${selM}/T:${selT}</td>
        <td><small>${data.aksi || '-'}</small></td>
      `;
      tb.appendChild(row);
    }

    function autoFillScRow(input) {
      const sku = input.value;
      const row = input.closest('.sc-item-row');
      if (!row || !sku) return;
      const s = stockData.find(x => x.sku === sku || x.barcode === sku);
      if (s) {
        input.value = s.sku; // Set to actual SKU
        const lok = row.querySelector('.sc-it-lok');
        const batch = row.querySelector('.sc-it-batch');
        const exp = row.querySelector('.sc-it-exp');
        if (lok && !lok.value) lok.value = s.lokasi || '';
        if (batch && !batch.value) batch.value = s.batch || '';
        if (exp && !exp.value) exp.value = s.expDate || '';
      }
    }

    function addStockControlItemRow(data = {}, recountMode = false) {
      const tb = document.getElementById('hoItemsList');
      const row = document.createElement('tr');
      row.className = 'sc-item-row';
      const selM = data.selisihMabang !== undefined ? data.selisihMabang : (parseFloat(data.stockFisik || 0) - parseFloat(data.stockMabang || 0));
      const selT = data.selisihTtx !== undefined ? data.selisihTtx : (parseFloat(data.stockFisik || 0) - parseFloat(data.stockTtx || 0));
      const selColor = (selM !== 0 || selT !== 0) ? 'color:#ef4444; font-weight:700;' : 'color:#10b981; font-weight:700;';
      const ro = recountMode ? 'readonly style="background:rgba(255,255,255,0.03); border:none;"' : '';

      row.innerHTML = `
        <td><input type="text" class="form-control form-control-sm sc-it-lok" style="font-size:11px; padding: 6px 4px;" value="${data.lokasi || ''}" ${ro}></td>
        <td><input type="text" class="form-control form-control-sm sc-it-sku" style="font-size:11px; padding: 6px 4px;" value="${data.sku || ''}" placeholder="SKU" list="listStockSKU" onchange="autoFillScRow(this)" ${ro}></td>
        <td><input type="text" class="form-control form-control-sm sc-it-batch" style="font-size:11px; padding: 6px 4px;" value="${data.batch || ''}" ${ro}></td>
        <td><input type="date" class="form-control form-control-sm sc-it-exp" style="font-size:11px; padding: 6px 4px;" value="${data.exp || ''}" ${ro}></td>
        <td><input type="number" class="form-control form-control-sm sc-it-ttx" style="font-size:11px; padding: 6px 4px;" value="${data.stockTtx !== undefined ? data.stockTtx : ''}" onkeyup="calcScSelisih(this)" onchange="calcScSelisih(this)"></td>
        <td><input type="number" class="form-control form-control-sm sc-it-mabang" style="font-size:11px; padding: 6px 4px;" value="${data.stockMabang !== undefined ? data.stockMabang : ''}" onkeyup="calcScSelisih(this)" onchange="calcScSelisih(this)"></td>
        <td><input type="number" class="form-control form-control-sm sc-it-fisik" style="font-size:11px; padding: 6px 4px;" value="${data.stockFisik !== undefined ? data.stockFisik : ''}" onkeyup="calcScSelisih(this)" onchange="calcScSelisih(this)" placeholder="Awal"></td>
        <td class="sc-col-staff" style="display: ${recountMode ? 'table-cell' : 'none'}"><input type="number" class="form-control form-control-sm sc-it-fisik-staff" style="font-size:11px; padding: 6px 4px; border:2px solid var(--teal); background:rgba(14,165,233,0.05)" value="${data.stockFisikStaff !== undefined ? data.stockFisikStaff : ''}" onkeyup="calcScSelisih(this)" onchange="calcScSelisih(this)" placeholder="Terbaru"></td>
        <td class="sc-it-selisih" style="${selColor}; font-size:10px; white-space:nowrap;"></td>
        <td>
          <select class="form-control form-control-sm sc-it-aksi" style="font-size:10px; padding: 6px 2px;" ${recountMode ? 'disabled' : ''}>
            <option value="Adjust Stock" ${data.aksi === 'Adjust Stock' ? 'selected' : ''}>Adjust</option>
            <option value="Cari Stock" ${data.aksi === 'Cari Stock' ? 'selected' : ''}>Cari</option>
            <option value="-" ${data.aksi === '-' ? 'selected' : ''}>-</option>
          </select>
        </td>
        <td><button class="btn btn-danger btn-sm" style="padding: 2px 8px;" onclick="this.parentElement.parentElement.remove()" ${recountMode ? 'disabled' : ''}>✕</button></td>
      `;
      tb.appendChild(row);
      calcScSelisih(row.querySelector('.sc-it-fisik'));
    }

    function calcScSelisih(el) {
      const row = el.closest('.sc-item-row');
      if (!row) return;
      const m = parseFloat(row.querySelector('.sc-it-mabang').value) || 0;
      const ttx = parseFloat(row.querySelector('.sc-it-ttx').value) || 0;
      const f = parseFloat(row.querySelector('.sc-it-fisik').value) || 0;
      const fs = parseFloat(row.querySelector('.sc-it-fisik-staff').value) || 0;

      const selM = f - m;
      const selT = f - ttx;
      const sms = fs ? (fs - m) : 0;
      const sts = fs ? (fs - ttx) : 0;

      const hasVar1 = (selM !== 0 || selT !== 0);
      const hasVar2 = fs ? (sms !== 0 || sts !== 0) : false;

      const cell = row.querySelector('.sc-it-selisih');
      if (cell) {
        cell.style.fontSize = '9px';
        cell.innerHTML = `
          <div style="color:${hasVar1 ? '#ef4444' : '#10b981'}; font-weight:700">A: TTX:${selT >= 0 ? '+' : ''}${selT} / MBG:${selM >= 0 ? '+' : ''}${selM}</div>
          ${fs ? `<div style="color:${hasVar2 ? '#ef4444' : '#10b981'}; font-weight:700">T: TTX:${sts >= 0 ? '+' : ''}${sts} / MBG:${sms >= 0 ? '+' : ''}${sms}</div>` : ''}
        `;
      }

      const aksiSel = row.querySelector('.sc-it-aksi');
      if (aksiSel && (hasVar1 || hasVar2) && aksiSel.value === '-') {
        aksiSel.value = 'Adjust Stock';
      }
    }

    function openStockControlForm() {
      resetForm(['hoId', 'hoTanggal', 'hoPic', 'hoArea', 'hoKategori', 'hoAlasan', 'hoKaryawan']);
      const modalTitle = document.querySelector('#modalHandover h3');
      if (modalTitle) modalTitle.innerHTML = '📦 Laporan Stock Control';

      const addBtn = document.querySelector('#modalHandover #btnHoAddItem');
      if (addBtn) addBtn.style.display = 'block';

      document.getElementById('hoSectionInitialItems').style.display = 'none';
      document.getElementById('hoLatestTitle').textContent = '📦 Daftar Item / SKU';
      document.getElementById('hoThFisikStaff').style.display = 'none';

      setToday('hoTanggal');
      document.getElementById('hoItemsList').innerHTML = '';
      document.getElementById('hoInitialItemsList').innerHTML = '';
      onStockControlKategoriChange();
      openModal('modalHandover');
      addStockControlItemRow();
    }

    function editStockControl(id, tanggal, pic, area, kategori, alasan, karyawan) {
      resetForm(['hoId', 'hoTanggal', 'hoPic', 'hoArea', 'hoKategori', 'hoAlasan', 'hoKaryawan']);
      const modalTitle = document.querySelector('#modalHandover h3');
      if (modalTitle) modalTitle.innerHTML = '📦 Laporan Stock Control - <span style="color:var(--accent)">Edit Admin</span>';

      const addBtn = document.querySelector('#modalHandover #btnHoAddItem');
      if (addBtn) addBtn.style.display = 'block';

      document.getElementById('hoSectionInitialItems').style.display = 'none';
      document.getElementById('hoLatestTitle').textContent = '📦 Daftar Item / SKU';
      document.getElementById('hoThFisikStaff').style.display = 'none';

      document.getElementById('hoId').value = id;
      document.getElementById('hoTanggal').value = tanggal;
      document.getElementById('hoPic').value = pic;
      document.getElementById('hoArea').value = area;
      document.getElementById('hoKategori').value = kategori;
      document.getElementById('hoAlasan').value = alasan;
      document.getElementById('hoKaryawan').value = karyawan;

      const list = document.getElementById('hoItemsList');
      list.innerHTML = '<tr><td colspan="11" class="text-center p-3">Memuat detail items...</td></tr>';
      document.getElementById('hoInitialItemsList').innerHTML = '';

      openModal('modalHandover');
      onStockControlKategoriChange();

      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          list.innerHTML = '';
          if (!res.data || res.data.length === 0) {
            addStockControlItemRow();
          } else {
            res.data.forEach(it => addStockControlItemRow(it));
          }
        } else {
          toast('Gagal memuat detail: ' + res.message, 'error');
        }
      }).getStockControlDetail(id);
    }

    function openStockRecountForm(id, tanggal, pic, area, kategori, alasan, karyawan) {
      resetForm(['hoId', 'hoTanggal', 'hoPic', 'hoArea', 'hoKategori', 'hoAlasan', 'hoKaryawan']);
      const modalTitle = document.querySelector('#modalHandover h3');
      if (modalTitle) modalTitle.innerHTML = '🧑‍💻 Hitung Ulang Staff - <span style="color:var(--teal)">' + area + '</span>';

      const addBtn = document.querySelector('#modalHandover #btnHoAddItem');
      if (addBtn) addBtn.style.display = 'none';

      document.getElementById('hoSectionInitialItems').style.display = 'block';
      document.getElementById('hoLatestTitle').textContent = '📦 2. Input Data Terbaru (Staff)';
      document.getElementById('hoThFisikStaff').style.display = 'table-cell';

      document.getElementById('hoId').value = id;
      document.getElementById('hoTanggal').value = tanggal;
      document.getElementById('hoPic').value = pic;
      document.getElementById('hoArea').value = area;
      document.getElementById('hoKategori').value = kategori;
      document.getElementById('hoAlasan').value = alasan;
      document.getElementById('hoKaryawan').value = karyawan;

      const list = document.getElementById('hoItemsList');
      const initialList = document.getElementById('hoInitialItemsList');
      list.innerHTML = '<tr><td colspan="11" class="text-center p-3">Memuat detail items untuk hitung ulang...</td></tr>';
      initialList.innerHTML = '';

      openModal('modalHandover');
      onStockControlKategoriChange();

      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          list.innerHTML = '';
          initialList.innerHTML = '';
          if (!res.data || res.data.length === 0) {
            addStockControlItemRow({}, true);
          } else {
            res.data.forEach(it => {
              addStockControlInitialRow(it);
              addStockControlItemRow(it, true);
            });
          }
        } else {
          toast('Gagal memuat detail: ' + res.message, 'error');
        }
      }).getStockControlDetail(id);
    }

    function submitStockControl() {
      const id = v('hoId'), t = v('hoTanggal'), p = v('hoPic'), a = v('hoArea'), cat = v('hoKategori'), alasan = v('hoAlasan'), karyawan = v('hoKaryawan');
      if (!t || !p || !a) return toast('Lengkapi data utama', 'error');

      const items = [];
      document.querySelectorAll('.sc-item-row').forEach(row => {
        const sku = row.querySelector('.sc-it-sku').value;
        if (sku || cat === 'Stock Opname') {
          items.push({
            lokasi: row.querySelector('.sc-it-lok').value,
            sku: sku,
            batch: row.querySelector('.sc-it-batch').value,
            exp: row.querySelector('.sc-it-exp').value,
            m: parseFloat(row.querySelector('.sc-it-mabang').value) || 0,
            ttx: parseFloat(row.querySelector('.sc-it-ttx').value) || 0,
            f: parseFloat(row.querySelector('.sc-it-fisik').value) || 0,
            fStaff: parseFloat(row.querySelector('.sc-it-fisik-staff').value) || 0,
            aksi: row.querySelector('.sc-it-aksi').value,
            alasan: ''
          });
        }
      });

      if (cat === 'Stock Opname' && items.length === 0) return toast('Tambahkan minimal 1 item', 'error');

      const btn = document.querySelector('#modalHandover .btn-primary[onclick*="submitStockControl"]');
      if (btn) { btn.disabled = true; btn.textContent = '⏳ Menyimpan...'; }
      google.script.run
        .withSuccessHandler(res => {
          if (btn) { btn.disabled = false; btn.textContent = '💾 Simpan Laporan'; }
          if (res.success) {
            toast(res.message || '✅ Laporan berhasil disimpan!', 'success');
            closeModal('modalHandover');
            loadStockControl();
            loadDashboardStock();
          } else {
            toast('❌ ' + (res.message || 'Gagal menyimpan'), 'error');
          }
        })
        .withFailureHandler(err => {
          if (btn) { btn.disabled = false; btn.textContent = '💾 Simpan Laporan'; }
          toast('❌ Error: ' + (err.message || err), 'error');
        })
        .saveStockControl(id, t, p, a, cat, alasan, karyawan, items, currentUser.username);
    }

    function onStockControlKategoriChange() {
      const cat = v('hoKategori');
      const sectionItems = document.getElementById('hoSectionItems');
      const sectionKaryawan = document.getElementById('hoSectionKaryawan');
      const labelAlasan = document.getElementById('hoAlasanLabel');

      if (cat === 'Laporan Lapangan') {
        sectionItems.style.display = 'none';
        sectionKaryawan.style.display = 'block';
        labelAlasan.textContent = 'Keterangan Temuan';
      } else {
        sectionItems.style.display = 'block';
        sectionKaryawan.style.display = 'flex'; // row
        labelAlasan.textContent = 'Keterangan Alasan';
      }
    }

    function downloadStockControlTemplate() {
      const headers = [["Tanggal (YYYY-MM-DD)", "PIC", "Area", "Kategori", "Alasan Utama", "Karyawan", "Lokasi Item", "SKU", "Batch", "Exp (YYYY-MM-DD)", "Stok Mabang", "Stok TTX", "Stok Fisik", "Aksi"]];
      const ws = XLSX.utils.aoa_to_sheet(headers);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Template");
      XLSX.writeFile(wb, "Template_Stock_Control.xlsx");
    }

    function importStockControlExcel(el) {
      const file = el.files[0]; if (!file) return;
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet);

        if (!rows.length) return toast('File kosong', 'error');

        // Group by Master fields (Tanggal, PIC, Area, Kategori)
        const groups = {};
        rows.forEach(r => {
          const key = `${r["Tanggal (YYYY-MM-DD)"]}_${r["PIC"]}_${r["Area"]}`;
          if (!groups[key]) {
            groups[key] = {
              t: r["Tanggal (YYYY-MM-DD)"], p: r["PIC"], a: r["Area"], cat: r["Kategori"] || 'Stock Opname',
              alasan: r["Alasan Utama"] || '', kar: r["Karyawan"] || '', items: []
            };
          }
          if (r["SKU"]) {
            const m = parseFloat(r["Stok Mabang"]) || 0, ttx = parseFloat(r["Stok TTX"]) || 0, f = parseFloat(r["Stok Fisik"]) || 0;
            groups[key].items.push({
              lokasi: r["Lokasi Item"] || '', sku: r["SKU"], batch: r["Batch"] || '', exp: r["Exp (YYYY-MM-DD)"] || '',
              stockMabang: m, stockTtx: ttx, stockFisik: f, selisihMabang: f - m, selisihTtx: f - ttx, aksi: r["Aksi"] || '-'
            });
          }
        });

        const promises = Object.values(groups).map(g => {
          return new Promise(resolve => {
            google.script.run.withSuccessHandler(resolve).importStockControl(g.t, g.p, g.a, g.cat, g.alasan, g.kar, g.items, currentUser.username);
          });
        });

        Promise.all(promises).then(() => {
          toast('Impor Berhasil'); loadStockControl(); el.value = '';
        });
      };
      reader.readAsArrayBuffer(file);
    }

    // DATA VALIDATION FUNCTION
    function showDataValidationForm() {
      const html = `
        <div style="padding: 20px;">
          <h4 style="margin-bottom: 15px;">✓ Validasi & Hitung Ulang Stock</h4>
          <div style="margin-bottom: 15px;">
            <label style="display: block; font-size: 12px; color: var(--text-muted); margin-bottom: 5px;">Pilih Kategori Stock</label>
            <select class="form-select" id="dvKategori" style="margin-bottom: 10px;">
              <option value="">-- Pilih Kategori --</option>
              <option value="Stock Opname">Stock Opname</option>
              <option value="Laporan Lapangan">Laporan Lapangan</option>
              <option value="Pengecekan Rutin">Pengecekan Rutin</option>
            </select>
          </div>
          <div style="margin-bottom: 15px;">
            <label style="display: block; font-size: 12px; color: var(--text-muted); margin-bottom: 5px;">Periode Waktu</label>
            <div style="display: flex; gap: 10px;">
              <input type="date" class="form-control" id="dvTglMulai" placeholder="Dari">
              <input type="date" class="form-control" id="dvTglAkhir" placeholder="Sampai">
            </div>
          </div>
          <div style="margin-bottom: 15px; padding: 10px; background: rgba(14, 165, 233, 0.1); border-left: 3px solid var(--teal); border-radius: 4px;">
            <small style="color: var(--text-muted);">💡 Sistem akan memvalidasi selisih stock dan menghitung ulang akurasi berdasarkan data fisik.</small>
          </div>
          <div style="display: flex; gap: 10px; justify-content: flex-end;">
            <button class="btn btn-ghost" onclick="closeModal('modalDataValidation')" style="padding: 8px 16px;">Batal</button>
            <button class="btn btn-primary" onclick="executeStockValidation()" style="padding: 8px 16px;">▶ Jalankan Validasi</button>
          </div>
        </div>
      `;

      let modal = document.getElementById('modalDataValidation');
      if (!modal) {
        modal = document.createElement('div');
        modal.id = 'modalDataValidation';
        modal.className = 'modal';
        modal.style.display = 'none';
        document.body.appendChild(modal);
      }
      modal.innerHTML = `<div class="modal-dialog"><div class="modal-content">${html}</div></div>`;
      openModal('modalDataValidation');
    }

    function executeStockValidation() {
      const kategori = v('dvKategori');
      const tglMulai = v('dvTglMulai');
      const tglAkhir = v('dvTglAkhir');

      if (!kategori || !tglMulai || !tglAkhir) return toast('Lengkapi semua field', 'error');

      toast('Sedang memproses validasi stock...', 'info');
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          toast('Validasi berhasil! ' + res.message, 'success');
          closeModal('modalDataValidation');
          loadStockControl();
          loadDashboardStock();
        } else {
          toast('Error: ' + res.message, 'error');
        }
      }).validateAndRecalculateStock(kategori, tglMulai, tglAkhir, currentUser.username);
    }

    // APPROVAL CENTER FUNCTIONS
    function renderApprovalCenter() {
      const status = v('acFilterStatus');
      const search = v('acSearch').toLowerCase();

      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          const tbody = document.getElementById('tableApprovalCenter');
          const data = res.approvals.filter(a =>
            (!status || a.status === status) &&
            (a.pic.toLowerCase().includes(search) || a.area.toLowerCase().includes(search))
          );

          // Update counts
          document.getElementById('acCountPending').textContent = res.stats.pending;
          document.getElementById('acCountApproved').textContent = res.stats.approved;
          document.getElementById('acCountRejected').textContent = res.stats.rejected;
          document.getElementById('acCountTotal').textContent = res.stats.total;

          tbody.innerHTML = data.map((a, i) => `
            <tr>
              <td>${a.tanggal}</td>
              <td>${a.pic}</td>
              <td>${a.area}</td>
              <td><small>${a.kategori}</small></td>
              <td><small>${a.temuan.substring(0, 30)}...</small></td>
              <td>${a.qty}</td>
              <td>
                <span class="badge ${a.status === 'pending' ? 'bg-warning' : a.status === 'approved' ? 'bg-success' : 'bg-danger'}">
                  ${a.status === 'pending' ? 'Pending' : a.status === 'approved' ? 'Approved' : 'Rejected'}
                </span>
              </td>
              <td>
                ${a.status === 'pending' ? `
                  <button class="btn btn-sm btn-success" onclick="approveStockAdjustment('${a.id}')" style="padding: 4px 8px; font-size: 10px;">✓ Approve</button>
                  <button class="btn btn-sm btn-danger" onclick="rejectStockAdjustment('${a.id}')" style="padding: 4px 8px; font-size: 10px;">✕ Reject</button>
                ` : '-'}
              </td>
            </tr>
          `).join('');
        }
      }).getApprovalCenterData(currentUser.username);
    }

    function approveStockAdjustment(id) {
      toast('Sedang memproses persetujuan...', 'info');
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          toast('✅ Stock adjustment approved!', 'success');
          renderApprovalCenter();
        } else toast('Error: ' + res.message, 'error');
      }).approveStockAdjustment(id, currentUser.username);
    }

    function rejectStockAdjustment(id) {
      const reason = prompt('Alasan penolakan:');
      if (!reason) return;
      toast('Sedang memproses penolakan...', 'info');
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          toast('❌ Stock adjustment ditolak!', 'success');
          renderApprovalCenter();
        } else toast('Error: ' + res.message, 'error');
      }).rejectStockAdjustment(id, currentUser.username, reason);
    }


    function togglePendingApproval(id, isChecked) {
      const status = isChecked ? 'Menunggu Approval' : 'Pending';
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          toast(isChecked ? '✓ Item marked for approval' : '✗ Approval removed', 'success');
          loadStockControl();
          if (isChecked) renderApprovalCenter();
        } else {
          toast('Error: ' + res.message, 'error');
          // Revert checkbox
          document.getElementById('chk-pending-' + id).checked = !isChecked;
        }
      }).updateStockControlStatus(id, status, currentUser.username);
    }

    function loadDashboardStock() {
      // 1. Fetch Stats & Trend
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          document.getElementById('dsTotalItem').textContent = res.stats.totalItems;
          document.getElementById('dsAccuracy').textContent = res.stats.accuracy.toFixed(1) + '%';

          const trendText = document.getElementById('dsAccuracyTrend');
          if (res.trend && res.trend.length >= 2) {
            const last = res.trend[res.trend.length - 1].accuracy;
            const prev = res.trend[res.trend.length - 2].accuracy;
            const diff = last - prev;
            trendText.innerHTML = `<span style="color:${diff >= 0 ? 'var(--green)' : 'var(--red)'}">${diff >= 0 ? '↑' : '↓'} ${Math.abs(diff).toFixed(1)}%</span> vs sesi sebelumnya`;
          } else {
            trendText.textContent = 'Data trend belum cukup';
          }

          // Render Trend Chart
          renderStockTrendChart(res.trend);
        }
      }).getStockControlStats();

      // 2. Fetch Master Records for Approval Summary & Laporan Lapangan
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          const pending = res.data.filter(d => d.status === 'Menunggu Approval').length;
          const approved = res.data.filter(d => d.status === 'Disetujui').length;
          const rejected = res.data.filter(d => d.status === 'Ditolak' || d.status === 'Dibatalkan').length;

          document.getElementById('dsCountPending').textContent = pending;
          document.getElementById('dsCountApproved').textContent = approved;
          document.getElementById('dsCountRejected').textContent = rejected;
          document.getElementById('dsVarianceItem').textContent = pending;

          // Render Laporan Lapangan Table
          const llData = res.data.filter(d => d.kategori === 'Laporan Lapangan').sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal)).slice(0, 5);
          const tb = document.getElementById('dsLaporanLapanganTable');
          if (tb) {
            tb.innerHTML = '';
            if (llData.length === 0) {
              tb.innerHTML = '<tr><td colspan="4" class="text-center text-muted">Belum ada laporan lapangan</td></tr>';
            } else {
              llData.forEach(d => {
                tb.innerHTML += `<tr>
                  <td><strong>${d.pic}</strong></td>
                  <td><span class="badge-tb">${d.area}</span></td>
                  <td style="font-size:12px;">${d.alasan}</td>
                  <td>${d.karyawan || '-'}</td>
                </tr>`;
              });
            }
          }
        }
      }).getStockControl();
    }

    let chartStockAcc = null;
    function renderStockTrendChart(trend) {
      const canvas = document.getElementById('chartStockAccuracy');
      if (!canvas) return;
      const ctx = canvas.getContext('2d');
      if (chartStockAcc) chartStockAcc.destroy();

      const labels = trend.map(t => formatDate(t.tanggal));
      const data = trend.map(t => t.accuracy.toFixed(1));

      chartStockAcc = new Chart(ctx, {
        type: 'line',
        data: {
          labels: labels,
          datasets: [{
            label: 'Akurasi (%)',
            data: data,
            borderColor: '#10b981',
            backgroundColor: 'rgba(16, 185, 129, 0.1)',
            fill: true,
            tension: 0.4,
            pointRadius: 4,
            pointBackgroundColor: '#10b981'
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          scales: {
            y: { beginAtZero: true, max: 100, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8' } },
            x: { grid: { display: false }, ticks: { color: '#94a3b8' } }
          },
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                label: function (context) {
                  return `Akurasi: ${context.parsed.y}%`;
                }
              }
            }
          }
        }
      });
    }

    function updateStockControlStatus(id, st) {
      const btn = event ? event.target : null;
      if (btn && btn.tagName === 'BUTTON') {
        btn.classList.add('btn-appr-animate');
        btn.disabled = true;
      }

      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          if (btn && btn.tagName === 'BUTTON') {
            btn.classList.remove('btn-appr-animate');
            if (st === 'Disetujui') {
              btn.classList.add('btn-appr-success');
            }
          }
          toast('Status Berhasil Diperbarui');
          loadStockControl();
          loadDashboardStock();
        } else {
          if (btn && btn.tagName === 'BUTTON') {
            btn.classList.remove('btn-appr-animate');
            btn.disabled = false;
          }
          toast(res.message, 'error');
        }
      }).updateStockControlStatus(id, st);
    }

    function delStockControl(id) {
      if (confirm('Apakah Anda yakin ingin menghapus laporan ini?')) {
        google.script.run.withSuccessHandler(res => {
          if (res.success) {
            toast('Laporan Dihapus');
            loadStockControl();
          }
        }).deleteStockControl(id);
      }
    }

    function loadKlaim() {
      google.script.run.withSuccessHandler(res => {
        const tb = document.getElementById('tableKlaim');
        tb.innerHTML = '';
        if (!res.data.length) {
          tb.innerHTML = '<tr><td colspan="8" class="empty-state">Kosong</td></tr>';
        } else {
          let p = 0, s = 0;
          res.data.forEach(d => {
            if (d.status === 'Selesai') s += d.harga; else p += d.harga;
            tb.innerHTML += `
              <tr id="klaim-main-${d.id}">
                <td>
                  <button class="btn-toggle-row" onclick="toggleKlaimDetail('${d.id}','${d.resi}')">
                    <i class="bi bi-plus-circle"></i>
                  </button>
                  ${formatDate(d.tanggal)}
                </td>
                <td>${d.pic}</td>
                <td><strong>${d.resi}</strong></td>
                <td class="rupiah">${formatRp(d.harga)}</td>
                <td>${d.keterangan || '-'}</td>
                <td>
                  <select onchange="updateKLStatus('${d.id}',this.value)" style="background:transparent;color:var(--white);border:1px solid #ffffff22;border-radius:4px;padding:2px">
                    <option value="Pending" ${d.status === 'Pending' ? 'selected' : ''}>Pending</option>
                    <option value="Selesai" ${d.status === 'Selesai' ? 'selected' : ''}>Selesai</option>
                  </select>
                </td>
                <td>
                  <div style="display:flex; gap:4px;">
                    <button class="btn btn-ghost btn-sm" onclick="editKlaim('${d.id}')">✏️</button>
                    <button class="btn btn-danger btn-sm" onclick="delKlaim('${d.id}')">🗑️</button>
                  </div>
                </td>
              </tr>
              <tr id="klaim-detail-${d.id}" class="row-detail" style="display:none;">
                <td colspan="8"><div class="detail-container"><div class="loading-inline">Memuat detail...</div></div></td>
              </tr>`;
          });
          klaimDataList = res.data; // Store globally for edit access
          document.getElementById('statKlaimPending').textContent = formatRp(p);
          document.getElementById('statKlaimSelesai').textContent = formatRp(s);
        }
      }).getKlaim();
    }

    function toggleKlaimDetail(id, resi) {
      const detailRow = document.getElementById(`klaim-detail-${id}`);
      const btn = document.querySelector(`#klaim-main-${id} .btn-toggle-row`);
      if (detailRow.style.display === 'none') {
        const cont = detailRow.querySelector('.detail-container');
        if (cont.innerHTML.includes('Memuat detail...')) {
          google.script.run.withSuccessHandler(res => {
            if (res.success) {
              let html = `
                <div style="margin-bottom:10px; font-weight:700; font-size:11px; color:var(--teal); text-transform:uppercase;">📦 Detail Item Klaim - ${resi}</div>
                <table class="detail-table">
                  <thead><tr><th>SKU</th><th>Harga</th></tr></thead>
                  <tbody>`;
              res.data.forEach(itm => {
                html += `<tr><td><strong>${itm.sku}</strong></td><td class="rupiah">${formatRp(itm.harga)}</td></tr>`;
              });
              html += `</tbody></table>`;
              cont.innerHTML = html;
            } else {
              cont.innerHTML = `<span style="color:var(--red)">Gagal memuat: ${res.message}</span>`;
            }
          }).getKlaimDetail(id);
        }
        detailRow.style.display = 'table-row';
        btn.innerHTML = '<i class="bi bi-dash-circle"></i>';
        btn.classList.add('active');
      } else {
        detailRow.style.display = 'none';
        btn.innerHTML = '<i class="bi bi-plus-circle"></i>';
        btn.classList.remove('active');
      }
    }

    function addKlaimItemRow(sku = '', harga = '') {
      const c = document.getElementById('klItemsList');
      const div = document.createElement('div');
      div.className = 'item-row mb-2';
      div.style.display = 'grid';
      div.style.gridTemplateColumns = '1fr 1fr 40px';
      div.style.gap = '8px';
      div.style.alignItems = 'center';
      div.innerHTML = `
        <input type="text" class="form-control sku-search" placeholder="Kode SKU" value="${sku}">
        <div class="rp-input-wrap">
          <span class="rp-prefix" style="padding: 0 8px; font-size: 12px;">Rp</span>
          <input type="text" class="form-control kl-item-harga" placeholder="0" oninput="formatRpInput(this); calculateKlaimTotal()" inputmode="numeric" value="${harga ? harga.toLocaleString('id-ID') : ''}">
        </div>
        <button class="btn btn-danger btn-sm" onclick="this.parentElement.remove(); calculateKlaimTotal()" style="height: 38px;">✕</button>
      `;
      c.appendChild(div);
    }

    function calculateKlaimTotal() {
      let total = 0;
      document.querySelectorAll('.kl-item-harga').forEach(inp => {
        total += getRpValue(inp.id || '', inp); // Pass element since it might not have ID
      });
      // Custom helper because getRpValue uses id
      const prices = document.querySelectorAll('.kl-item-harga');
      let sum = 0;
      prices.forEach(p => {
        const val = p.value.replace(/[^0-9]/g, '');
        sum += parseInt(val) || 0;
      });
      document.getElementById('klTotal').value = sum.toLocaleString('id-ID');
    }

    function openKlaimForm(editId = null) {
      if (!stockData || stockData.length === 0) loadStock();
      setVal('klId', '');
      document.getElementById('klaimModalTitle').textContent = '⚠️ Input Laporan Klaim Paket';
      document.getElementById('klItemsList').innerHTML = '';
      document.getElementById('klTotal').value = '0';
      resetForm(['klTanggal', 'klPic', 'klResi', 'klKeterangan']);
      setToday(); // Ensure date is set

      if (editId) {
        const d = klaimDataList.find(x => x.id === editId);
        if (!d) return;
        setVal('klId', d.id);
        document.getElementById('klaimModalTitle').textContent = '✏️ Edit Laporan Klaim Paket';
        setVal('klTanggal', d.tanggal);
        setVal('klPic', d.pic);
        setVal('klResi', d.resi);
        setVal('klKeterangan', d.keterangan || '');

        showLoading('Memuat rincian...');
        google.script.run.withSuccessHandler(res => {
          hideLoading();
          if (res.success) {
            res.data.forEach(itm => {
              addKlaimItemRow(itm.sku, itm.harga);
            });
            calculateKlaimTotal();
          } else toast(res.message, 'error');
        }).getKlaimDetail(editId);
      } else {
        addKlaimItemRow();
      }
      openModal('modalKlaim');
    }

    function editKlaim(id) {
      openKlaimForm(id);
    }

    function submitKlaim() {
      const t = v('klTanggal'), p = v('klPic'), r = v('klResi'), k = v('klKeterangan');
      const items = [];
      const rows = document.querySelectorAll('#klItemsList .item-row');

      let totalHarga = 0;
      rows.forEach(row => {
        const sku = row.querySelector('.sku-search').value;
        const harga = parseInt(row.querySelector('.kl-item-harga').value.replace(/[^0-9]/g, '')) || 0;
        if (sku && harga > 0) {
          items.push({ sku, harga });
          totalHarga += harga;
        }
      });

      if (!t || !p || !r) return toast('Lengkapi form (Tanggal, PIC, Resi)', 'error');
      if (items.length === 0) return toast('Tambahkan minimal 1 item (SKU & Harga)', 'error');

      const btn = document.querySelector('#modalKlaim .btn-primary');
      const oldTxt = btn.textContent;
      btn.disabled = true; btn.textContent = '⏳ Menyimpan...';

      const id = v('klId');
      const cb = res => {
        btn.disabled = false; btn.textContent = oldTxt;
        if (res.success) {
          toast('Berhasil');
          closeModal('modalKlaim');
          loadKlaim();
          resetForm(['klResi', 'klKeterangan', 'klId']);
          document.getElementById('klItemsList').innerHTML = '';
          document.getElementById('klTotal').value = '0';
        } else toast(res.message, 'error');
      };

      if (id) {
        google.script.run.withSuccessHandler(cb).updateKlaim(id, t, p, r, totalHarga, k, items, currentUser.username);
      } else {
        google.script.run.withSuccessHandler(cb).addKlaim(t, p, r, totalHarga, k, items, currentUser.username);
      }
    }

    function updateKLStatus(id, st) { google.script.run.withSuccessHandler(res => { if (res.success) { toast('Status update'); loadKlaim(); } else toast(res.message, 'error') }).updateKlaimStatus(id, st); }
    function delKlaim(id) { if (confirm('Hapus?')) google.script.run.withSuccessHandler(res => { if (res.success) { toast('Dihapus'); loadKlaim(); } }).deleteKlaim(id); }

    // ============================================================
    // TUGAS PROJECT WAREHOUSE
    // ============================================================
    function loadUsersDropdown() {
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          const sel = document.getElementById('tpAssignee'), fAsg = document.getElementById('tpFilterAssignee');
          if (sel && fAsg) {
            sel.innerHTML = '<option value="">-- Pilih User --</option>'; fAsg.innerHTML = '<option value="">Semua Assignee</option>';
            res.data.forEach(u => { sel.innerHTML += `<option value="${u.username}">${u.nama}</option>`; fAsg.innerHTML += `<option value="${u.username}">${u.nama}</option>`; });
          }
        }
      }).getUsers();
    }
    function loadTugasProject() {
      if (document.getElementById('tpAssignee').options.length <= 1) loadUsersDropdown();
      google.script.run.withSuccessHandler(res => { if (res.success) { tugasData = res.data; renderTugasProject(); } }).getTugasProject();
    }
    function openTugasModal(editId = null) {
      resetForm(['tpJudul', 'tpKategori', 'tpDeskripsi']); setVal('tpEditId', ''); setVal('tpTanggalMulai', new Date().toISOString().split('T')[0]);
      setVal('tpDeadline', ''); setVal('tpTargetHari', ''); setVal('tpAssignee', ''); setVal('tpPriority', 'Sedang'); setVal('tpStatus', 'Todo');
      document.getElementById('tpDurasiInfo').style.display = 'none'; document.getElementById('tugasModalTitle').textContent = '📋 Buat Tugas Baru';
      if (editId) {
        const t = tugasData.find(x => x.id === editId); if (!t) return;
        setVal('tpEditId', t.id); document.getElementById('tugasModalTitle').textContent = '✏️ Edit Tugas';
        setVal('tpJudul', t.judul); setVal('tpAssignee', t.assignee); setVal('tpPriority', t.prioritas); setVal('tpTanggalMulai', t.tanggalMulai);
        setVal('tpDeadline', t.deadline); setVal('tpTargetHari', t.targetHari); setVal('tpStatus', t.status); setVal('tpKategori', t.kategori); setVal('tpDeskripsi', t.deskripsi); hitungDurasi('tanggal');
      } openModal('modalTugasProject');
    }
    function hitungDurasi(source = 'tanggal') {
      const m = v('tpTanggalMulai'); let d = v('tpDeadline'); let t = v('tpTargetHari');
      const inf = document.getElementById('tpDurasiInfo'), txt = document.getElementById('tpDurasiText');
      if (m) {
        const dm = new Date(m);
        if (source === 'hari' && t) {
          const dd = new Date(dm); dd.setDate(dd.getDate() + parseInt(t));
          setVal('tpDeadline', dd.toISOString().split('T')[0]);
          inf.style.display = 'block'; txt.textContent = t + ' Hari';
        } else if (d) {
          const dd = new Date(d); const diff = Math.ceil((dd - dm) / 86400000);
          if (diff >= 0) { inf.style.display = 'block'; txt.textContent = diff + ' Hari'; setVal('tpTargetHari', diff); }
          else { inf.style.display = 'none'; setVal('tpTargetHari', ''); }
        } else { inf.style.display = 'none'; }
      } else inf.style.display = 'none';
    }
    function saveTugasProject() {
      const payload = { id: v('tpEditId'), judul: v('tpJudul'), assignee: v('tpAssignee'), assigneeName: document.getElementById('tpAssignee').options[document.getElementById('tpAssignee').selectedIndex]?.text, prioritas: v('tpPriority'), tanggalMulai: v('tpTanggalMulai'), deadline: v('tpDeadline'), targetHari: v('tpTargetHari'), status: v('tpStatus'), kategori: v('tpKategori'), deskripsi: v('tpDeskripsi'), createdBy: currentUser.username };

      if (!payload.judul) return toast('Judul Tugas wajib diisi!', 'error');
      if (!payload.assignee) return toast('Assignee (Penerima Tugas) wajib dipilih!', 'error');
      if (!payload.tanggalMulai || !payload.deadline) return toast('Tanggal Mulai dan Target Selesai wajib diisi!', 'error');

      const btn = document.querySelector('#modalTugasProject .btn-primary'); btn.disabled = true; btn.textContent = '⏳ Menyimpan...';
      const fn = payload.id ? 'updateTugasProject' : 'addTugasProject';
      google.script.run.withSuccessHandler(res => { btn.disabled = false; btn.textContent = '💾 Simpan Tugas'; if (res.success) { toast('Disimpan'); closeModal('modalTugasProject'); loadTugasProject(); } else toast(res.message, 'error'); })[fn](JSON.stringify(payload));
    }
    function renderTugasProject() {
      let filtered = tugasData;
      const fSts = v('tpFilterStatus'), fAsg = v('tpFilterAssignee'), fPri = v('tpFilterPriority'), q = v('tpSearch').toLowerCase(), view = v('tpView');
      if (fSts) filtered = filtered.filter(d => d.status === fSts); if (fAsg) filtered = filtered.filter(d => d.assignee === fAsg); if (fPri) filtered = filtered.filter(d => d.prioritas === fPri); if (q) filtered = filtered.filter(d => d.judul.toLowerCase().includes(q));

      let tot = filtered.length, prog = filtered.filter(d => d.status === 'In Progress').length, done = filtered.filter(d => d.status === 'Done').length, od = 0;
      document.getElementById('tpStatTotal').textContent = tot; document.getElementById('tpStatProgress').textContent = prog; document.getElementById('tpStatDone').textContent = done;

      const now = new Date();
      const makeCard = (d) => {
        let cCls = 'ontime', pCls = 'ontime', priCls = 'priority-medium'; if (d.prioritas === 'Tinggi') priCls = 'priority-high'; if (d.prioritas === 'Rendah') priCls = 'priority-low';
        if (d.status !== 'Done' && d.status !== 'Dibatalkan' && d.deadline) { const diff = Math.ceil((new Date(d.deadline) - now) / 86400000); if (diff < 0) { cCls = 'overdue'; pCls = 'overdue'; od++; } else if (diff <= 2) { cCls = 'warning'; pCls = 'warning'; } }
        if (d.status === 'Done') { cCls = 'ontime'; pCls = 'done'; }
        return `<div class="task-card ${cCls}" onclick="viewTugasDetail('${d.id}')"><div class="task-card-title">${d.judul}</div><div class="task-card-assignee"><span class="user-avatar" style="width:18px;height:18px;font-size:9px">${d.assigneeName.charAt(0)}</span> ${d.assigneeName}</div><div class="task-card-meta"><div class="task-card-date">Mulai: ${formatDate(d.tanggalMulai)}</div><div class="task-deadline-badge ${pCls}">🎯 ${formatDate(d.deadline)}</div></div><div style="margin-top:8px;display:flex;align-items:center;justify-content:space-between;"><span class="status-pill status-${d.status.toLowerCase().replace(' ', '')}">${d.status}</span><span style="font-size:11px;color:var(--gray);display:flex;align-items:center;gap:4px;"><span class="priority-dot ${priCls}"></span> ${d.prioritas}</span></div></div>`;
      };

      if (view === 'kanban') {
        document.getElementById('tpTableView').style.display = 'none'; document.getElementById('tpKanbanView').style.display = 'block';
        const cols = { 'Todo': [], 'In Progress': [], 'Review': [], 'Done': [] };
        filtered.forEach(d => { if (cols[d.status]) cols[d.status].push(d); });
        let kbHtml = ''; Object.keys(cols).forEach(k => { kbHtml += `<div class="kanban-col"><div class="kanban-col-header"><div class="kanban-col-title">${k === 'Todo' ? '📌' : k === 'In Progress' ? '🔄' : k === 'Review' ? '👀' : '✅'} ${k}</div><div class="kanban-count" style="background:var(--navy3);color:var(--white)">${cols[k].length}</div></div>${cols[k].map(makeCard).join('')}</div>`; });
        document.getElementById('tpKanban').innerHTML = kbHtml;
      } else {
        document.getElementById('tpKanbanView').style.display = 'none'; document.getElementById('tpTableView').style.display = 'block';
        const tb = document.getElementById('tpTableBody'); tb.innerHTML = '';
        filtered.forEach(d => { tb.innerHTML += `<tr><td><strong>${d.judul}</strong></td><td>${d.assigneeName}</td><td>${d.prioritas}</td><td>${formatDate(d.tanggalMulai)}</td><td>${formatDate(d.deadline)}</td><td>${d.targetHari} Hari</td><td><span class="status-pill status-${d.status.toLowerCase().replace(' ', '')}">${d.status}</span></td><td><button class="btn btn-ghost btn-sm" onclick="viewTugasDetail('${d.id}')">Lihat</button></td></tr>`; });
      }
      document.getElementById('tpStatOverdue').textContent = od;
    }
    function viewTugasDetail(id) {
      const d = tugasData.find(x => x.id === id); if (!d) return;
      const now = new Date(); let diff = 0; let progHtml = ''; let alertHtml = '';
      if (d.deadline && d.status !== 'Done') { diff = Math.ceil((new Date(d.deadline) - now) / 86400000); if (diff < 0) alertHtml = `<div class="task-overdue-alert"><div class="alert-icon">⚠️</div><div class="alert-text"><strong>Tugas Melewati Deadline!</strong><br>Tugas ini terlambat ${Math.abs(diff)} hari dari target penyelesaian.</div></div>`; }

      let logHtml = ''; try { const l = JSON.parse(d.log || '[]'); if (l.length) logHtml = l.map(x => `<div class="task-log-item"><div class="task-log-dot"></div><div><div class="task-log-text">${x.action}</div><div class="task-log-time">${formatDate(x.time)} | oleh ${x.by}</div></div></div>`).join(''); } catch (e) { }

      document.getElementById('tugasDetailContent').innerHTML = `
    ${alertHtml}
    <div class="task-detail-header">
      <div><h2 style="font-family:'Outfit';font-size:20px;margin-bottom:6px;line-height:1.3">${d.judul}</h2><div style="display:flex;gap:8px;"><span class="status-pill status-${d.status.toLowerCase().replace(' ', '')}">${d.status}</span><span style="font-size:12px;color:var(--gray)">Kategori: ${d.kategori || '-'}</span></div></div>
      <div style="text-align:right"><div style="font-size:11px;color:var(--gray);text-transform:uppercase;font-weight:700">Assignee</div><div style="display:flex;align-items:center;gap:6px;justify-content:flex-end;margin-top:4px;"><span class="user-avatar">${d.assigneeName.charAt(0)}</span><span style="font-weight:600;color:var(--teal)">${d.assigneeName}</span></div></div>
    </div>
    <div class="task-detail-body" style="margin-bottom:20px">
      <div class="detail-field"><span class="detail-label">Tgl Mulai</span><span class="detail-value">${formatDate(d.tanggalMulai)}</span></div>
      <div class="detail-field"><span class="detail-label">Target Selesai (Deadline)</span><span class="detail-value">${formatDate(d.deadline)}</span></div>
      <div class="detail-field"><span class="detail-label">Target Waktu</span><span class="detail-value">${d.targetHari} Hari</span></div>
      <div class="detail-field"><span class="detail-label">Prioritas</span><span class="detail-value">${d.prioritas}</span></div>
    </div>
    <div class="detail-field" style="margin-bottom:20px"><span class="detail-label">Deskripsi & Catatan</span><div class="task-notes-box">${escHtml(d.deskripsi)}</div></div>
    <div class="detail-field"><span class="detail-label">Riwayat Aktivitas</span><div class="task-log">${logHtml || '<div style="color:var(--gray);font-size:12px">Belum ada aktivitas.</div>'}</div></div>
  `;

      let footer = `<button class="btn btn-ghost" onclick="closeModal('modalTugasDetail')">Tutup</button>`;
      if (currentUser.role === 'admin' || currentUser.username === d.assignee || currentUser.username === d.createdBy) {
        footer = `
       <select class="form-select" id="dtStatusUpdate" style="padding:8px;background:#0f2040;color:#fff;border:1px solid #ffffff22;border-radius:6px;margin-right:auto;">
         <option value="Todo" ${d.status === 'Todo' ? 'selected' : ''}>📌 Todo</option><option value="In Progress" ${d.status === 'In Progress' ? 'selected' : ''}>🔄 In Progress</option>
         <option value="Review" ${d.status === 'Review' ? 'selected' : ''}>👀 Review</option><option value="Done" ${d.status === 'Done' ? 'selected' : ''}>✅ Done</option><option value="Dibatalkan" ${d.status === 'Dibatalkan' ? 'selected' : ''}>❌ Dibatalkan</option>
       </select>
       <button class="btn btn-primary" onclick="updateTugasStatus('${d.id}', document.getElementById('dtStatusUpdate').value)">Update Status</button>
       <button class="btn btn-ghost" onclick="closeModal('modalTugasDetail'); openTugasModal('${d.id}')">✏️ Edit</button>
       <button class="btn btn-danger" onclick="deleteTugasProject('${d.id}'); closeModal('modalTugasDetail')">🗑️</button>
       ` + footer;
      }
      document.getElementById('tugasDetailFooter').innerHTML = footer; openModal('modalTugasDetail');
    }
    function updateTugasStatus(id, st) { google.script.run.withSuccessHandler(res => { if (res.success) { toast('Status diperbarui'); closeModal('modalTugasDetail'); loadTugasProject(); } else toast(res.message, 'error') }).updateTugasStatus(id, st, currentUser.username); }
    function deleteTugasProject(id) { if (confirm('Hapus tugas ini?')) google.script.run.withSuccessHandler(res => { if (res.success) { toast('Dihapus'); loadTugasProject(); } }).deleteTugasProject(id); }

    // === SOP ===
    let sopDataCache = []; // Cache untuk filtering

    function loadSOP() {
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          sopDataCache = res.data || [];
          renderSOPList(sopDataCache);
        } else {
          const list = document.getElementById('sopList');
          list.innerHTML = '<div class="col-12" style="text-align:center;padding:40px;"><div style="font-size:48px;opacity:0.3;">❌</div><div style="color:var(--red);">Gagal memuat SOP</div></div>';
        }
      }).getSOP();
    }

    function renderSOPList(data) {
      const list = document.getElementById('sopList');
      list.innerHTML = '';

      if (!data || data.length === 0) {
        list.innerHTML = `
          <div class="col-12" style="text-align:center;padding:60px 20px;">
            <div style="font-size:64px;margin-bottom:16px;opacity:0.3;">📋</div>
            <div style="font-size:18px;font-weight:700;color:var(--text-main);margin-bottom:8px;">
              Belum Ada SOP
            </div>
            <div style="font-size:13px;color:var(--text-muted);">
              Klik tombol "Tambah SOP" untuk membuat Standard Operating Procedure baru
            </div>
          </div>`;
        return;
      }

      data.forEach(d => {
        const catClass = d.kategori.toLowerCase().replace(/ /g, '');
        const catIcon = d.kategori === 'Penerimaan Barang' ? '📥' :
          d.kategori === 'Pengeluaran Barang' ? '📤' :
            d.kategori === 'Keamanan' ? '🔒' :
              d.kategori === 'Keselamatan' ? '⚠️' : '📄';

        list.innerHTML += `
          <div class="col-md-6 col-lg-4">
            <div class="sop-item">
              <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px;">
                <div style="flex:1;" onclick="toggleSOP('${d.id}')">
                  <div class="sop-title">${catIcon} ${d.judul}</div>
                  <div class="sop-cat ${catClass}">${d.kategori}</div>
                </div>
                <div style="display:flex;gap:6px;">
                  <button class="btn btn-ghost btn-sm" onclick="editSOP('${d.id}')" 
                    style="padding:4px 8px;font-size:12px;color:var(--teal);" title="Edit SOP">
                    ✏️
                  </button>
                  <button class="btn btn-ghost btn-sm" onclick="delSOP('${d.id}')" 
                    style="padding:4px 8px;font-size:12px;color:var(--red);" title="Hapus SOP">
                    🗑️
                  </button>
                </div>
              </div>
              <div class="sop-body" id="sop-b-${d.id}">${escHtml(d.konten)}</div>
              <div class="sop-footer">
                <span>Dibuat: ${formatDate(d.updatedAt || d.createdAt || '')}</span>
                <span style="color:var(--teal);font-weight:600;">${d.createdBy || 'Admin'}</span>
              </div>
            </div>
          </div>`;
      });
    }

    function filterSOPList() {
      const searchTerm = (document.getElementById('sopSearchInput')?.value || '').toLowerCase();
      const categoryFilter = document.getElementById('sopCategoryFilter')?.value || '';

      let filtered = sopDataCache;

      if (searchTerm) {
        filtered = filtered.filter(d =>
          (d.judul || '').toLowerCase().includes(searchTerm) ||
          (d.konten || '').toLowerCase().includes(searchTerm) ||
          (d.kategori || '').toLowerCase().includes(searchTerm)
        );
      }

      if (categoryFilter) {
        filtered = filtered.filter(d => d.kategori === categoryFilter);
      }

      renderSOPList(filtered);
    }

    function toggleSOP(id) {
      const b = document.getElementById('sop-b-' + id);
      if (b) {
        const isVisible = b.style.display === 'block';
        b.style.display = isVisible ? 'none' : 'block';
      }
    }

    function openModalSOP(editId = null) {
      document.getElementById('sopId').value = editId || '';
      document.getElementById('sopModalTitle').textContent = editId ? '✏️ Edit SOP' : '📋 Tambah SOP';

      if (editId) {
        const sop = sopDataCache.find(s => s.id === editId);
        if (sop) {
          document.getElementById('sopJudul').value = sop.judul || '';
          document.getElementById('sopKategori').value = sop.kategori || '';
          document.getElementById('sopKonten').value = sop.konten || '';
        }
      } else {
        document.getElementById('sopJudul').value = '';
        document.getElementById('sopKategori').value = 'Lainnya';
        document.getElementById('sopKonten').value = '';
      }

      openModal('modalSOP');
    }

    function editSOP(id) {
      openModalSOP(id);
    }

    function submitSOP() {
      const id = v('sopId');
      const j = v('sopJudul');
      const k = v('sopKategori');
      const kt = v('sopKonten');

      if (!j || !kt) return toast('Judul & Konten wajib diisi!', 'error');

      const btn = document.querySelector('#modalSOP .btn-primary');
      btn.disabled = true;
      btn.textContent = '⏳ Menyimpan...';

      if (id) {
        // Update existing SOP
        google.script.run.withSuccessHandler(res => {
          btn.disabled = false;
          btn.textContent = '💾 Simpan';

          if (res.success) {
            toast('SOP berhasil diupdate!', 'success');
            closeModal('modalSOP');
            loadSOP();
          } else {
            toast(res.message || 'Gagal menyimpan SOP', 'error');
          }
        }).withFailureHandler(err => {
          btn.disabled = false;
          btn.textContent = '💾 Simpan';
          toast('Error: ' + err.message, 'error');
        }).updateSOP(id, j, kt, k);
      } else {
        // Add new SOP
        google.script.run.withSuccessHandler(res => {
          btn.disabled = false;
          btn.textContent = '💾 Simpan';

          if (res.success) {
            toast('SOP berhasil ditambahkan!', 'success');
            closeModal('modalSOP');
            loadSOP();
          } else {
            toast(res.message || 'Gagal menyimpan SOP', 'error');
          }
        }).withFailureHandler(err => {
          btn.disabled = false;
          btn.textContent = '💾 Simpan';
          toast('Error: ' + err.message, 'error');
        }).addSOP(j, kt, k, currentUser.username);
      }
    }

    function delSOP(id) {
      if (!confirm('Yakin ingin menghapus SOP ini?')) return;

      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          toast('SOP berhasil dihapus!', 'success');
          loadSOP();
        } else {
          toast(res.message || 'Gagal menghapus SOP', 'error');
        }
      }).withFailureHandler(err => {
        toast('Error: ' + err.message, 'error');
      }).deleteSOP(id);
    }

    function doExportSOP() {
      const btn = document.getElementById('btnExportSOP');
      if (!btn) return;

      btn.disabled = true;
      btn.textContent = '⏳ Mengekspor...';

      google.script.run.withSuccessHandler(res => {
        btn.disabled = false;
        btn.textContent = '📤 Ekspor ke Google Docs';

        if (res.success) {
          toast('SOP berhasil diekspor!', 'success');
          window.open(res.url, '_blank');
        } else {
          toast(res.message || 'Gagal mengekspor SOP', 'error');
        }
      }).withFailureHandler(err => {
        btn.disabled = false;
        btn.textContent = '📤 Ekspor ke Google Docs';
        toast('Error: ' + err.message, 'error');
      }).exportSOP();
    }

    // INVENTORY CORE (Stock, In/Out/Retur, Order)
    // =================================== =========================

    // === USER MANAGEMENT ===
    function filterUsersTable() {
      const query = (document.getElementById('searchUsers')?.value || '').toLowerCase().trim();
      const tbody = document.getElementById('tableUsers');
      if (!tbody) return;
      const rows = tbody.getElementsByTagName('tr');
      for (let i = 0; i < rows.length; i++) {
        const uCell = rows[i].cells[0]?.textContent || '';
        const nCell = rows[i].cells[1]?.textContent || '';
        const rCell = rows[i].cells[2]?.textContent || '';
        if (uCell.toLowerCase().includes(query) || nCell.toLowerCase().includes(query) || rCell.toLowerCase().includes(query)) {
          rows[i].style.display = '';
        } else {
          rows[i].style.display = 'none';
        }
      }
    }

    function loadUsers() {
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          usersData = res.data;
          const tb = document.getElementById('tableUsers');
          if (!tb) return;
          tb.innerHTML = '';

          const getRoleBadgeStyle = (role) => {
            const r = (role || '').toLowerCase();
            if (r === 'admin') return 'background: rgba(239, 68, 68, 0.12); color: #f87171; border: 1px solid rgba(239, 68, 68, 0.25);';
            if (r === 'hr') return 'background: rgba(245, 158, 11, 0.12); color: #fbbf24; border: 1px solid rgba(245, 158, 11, 0.25);';
            if (r.includes('supervisor') || r.includes('spv')) return 'background: rgba(99, 102, 241, 0.12); color: #818cf8; border: 1px solid rgba(99, 102, 241, 0.25);';
            if (r.includes('team leader') || r === 'tl') return 'background: rgba(20, 184, 166, 0.12); color: #2dd4bf; border: 1px solid rgba(20, 184, 166, 0.25);';
            return 'background: rgba(107, 114, 128, 0.12); color: #9ca3af; border: 1px solid rgba(107, 114, 128, 0.25);';
          };

          res.data.forEach(u => {
            const badgeStyle = getRoleBadgeStyle(u.role);
            const divisiText = u.divisi ? `<small style="color:var(--gray); display:block; margin-top:2px;">💼 ${u.divisi}</small>` : '';
            const actionButtons = u.username !== 'admin' ? `
              <button class="btn btn-ghost btn-sm" onclick="editUser('${u.id}')" style="padding: 4px 8px; border-radius: 6px;" title="Edit Akun">✏️</button>
              <button class="btn btn-danger btn-sm" onclick="delUser('${u.id}')" style="padding: 4px 8px; border-radius: 6px;" title="Hapus Akun">🗑️</button>
            ` : `<span class="text-muted" style="font-size: 12px; padding-right: 8px;">Sistem Utama</span>`;

            tb.innerHTML += `<tr>
              <td style="padding-left: 24px;"><strong>@${u.username}</strong></td>
              <td>${u.nama}</td>
              <td>
                <span class="badge-tb" style="padding: 4px 10px; font-size: 11px; border-radius: 20px; font-weight: 600; display: inline-block; ${badgeStyle}">${u.role}</span>
                ${divisiText}
              </td>
              <td>${formatDate(u.createdAt)}</td>
              <td style="text-align:right; padding-right:24px;">
                <div style="display:flex; gap:6px; justify-content: flex-end; align-items:center;">
                  ${actionButtons}
                </div>
              </td>
            </tr>`;
          });

          filterUsersTable();
        }
      }).getUsers();
    }
    function togglePermissions() { document.getElementById('uPermissionsWrap').style.display = v('uRole') !== 'admin' ? 'block' : 'none'; }
    function applyRolePresets() {
      const r = v('uRole'); if (r === 'admin' || r === 'user') return;
      const grid = document.getElementById('uPermissionsGrid');
      const checks = grid.querySelectorAll('input[type="checkbox"]');
      const set = (val) => { const c = grid.querySelector(`input[value="${val}"]`); if (c) c.checked = true; };

      // Default: Bersihkan dulu jika role berubah secara manual
      // checks.forEach(c => c.checked = false); // Opsional: Hapus komentar jika ingin reset total tiap ganti role

      if (r === 'HR') {
        ['karyawan', 'ijin', 'lembur', 'kpiKaryawan', 'pengajuanAsset', 'organisasi', 'editKaryawan', 'aksesHr', 'kelolaUser', 'aksesLemburLangsung', 'lemburTanpaLaporan', 'pengaturanTglMerah', 'dashboardValidasiLembur', 'absensiKaryawan', 'jadwalShift', 'aksesRepairAbsensi', 'antrianDistributor', 'aksesApprovalAsset', 'approvalDashboard'].forEach(set);
      } else if (r === 'Supervisor') {
        ['dashboard', 'kasGudang', 'teamBuilding', 'expense', 'pettyCash', 'paymentGudang', 'laporanKerja', 'grafikLaporan', 'stock', 'inbound', 'outbound', 'retur', 'order', 'antrianDistributor', 'bookingMobil', 'updateStatusBookingMobil', 'aksesSpv', 'kelolaUser', 'aksesLemburLangsung', 'lemburTanpaLaporan', 'pengaturanTglMerah', 'dashboardValidasiLembur', 'absensiKaryawan', 'jadwalShift', 'aksesRepairAbsensi', 'pengajuanAsset', 'aksesApprovalAsset', 'approvalDashboard', 'kpiKaryawan'].forEach(set);
      } else if (r === 'Vice Supervisor') {
        ['dashboard', 'kasGudang', 'pettyCash', 'paymentGudang', 'laporanKerja', 'grafikLaporan', 'stock', 'inbound', 'outbound', 'retur', 'order', 'antrianDistributor', 'bookingMobil', 'updateStatusBookingMobil', 'aksesViceSpv', 'kelolaUser', 'aksesLemburLangsung', 'lemburTanpaLaporan', 'absensiKaryawan', 'jadwalShift', 'aksesRepairAbsensi', 'pengajuanAsset', 'aksesApprovalAsset', 'approvalDashboard', 'kpiKaryawan'].forEach(set);
      } else if (r === 'Team Leader' || r.includes('Team Leader') || r === 'TL') {
        ['dashboard', 'laporanKerja', 'stock', 'inbound', 'outbound', 'antrianDistributor', 'aksesApproval', 'aksesLemburLangsung', 'lemburTanpaLaporan', 'pengajuanAsset', 'aksesApprovalAsset', 'approvalDashboard', 'kpiKaryawan'].forEach(set);
      }
    }
    function toggleCustomRole() { const s = v('uRole'); document.getElementById('uRoleCustom').style.display = s === 'Lainnya' ? 'block' : 'none'; }
    function updateKaryawanDatalist() {
      const dl = document.getElementById('listKaryawanNama');
      if (!dl) return;
      dl.innerHTML = '';
      if (karyawanData && karyawanData.length) {
        karyawanData.forEach(k => {
          const opt = document.createElement('option');
          opt.value = k.nama;
          dl.appendChild(opt);
        });
      }
    }
    function openUserModal() { updateKaryawanDatalist(); setVal('uId', ''); document.getElementById('userModalTitle').textContent = '⚙️ Tambah User'; resetForm(['uUsername', 'uPassword', 'uNama', 'uRoleCustom', 'uDivisi']); setVal('uRole', 'user'); togglePermissions(); toggleCustomRole(); document.querySelectorAll('#uPermissionsGrid input').forEach(c => c.checked = false); openModal('modalUser'); }
    function editUser(id) { updateKaryawanDatalist(); const u = usersData.find(x => x.id === id); if (!u) return; setVal('uId', u.id); document.getElementById('userModalTitle').textContent = '✏️ Edit User'; setVal('uUsername', u.username); setVal('uNama', u.nama); setVal('uPassword', ''); setVal('uDivisi', u.divisi || ''); const rOpts = Array.from(document.getElementById('uRole').options).map(o => o.value); if (rOpts.includes(u.role)) { setVal('uRole', u.role); } else { setVal('uRole', 'Lainnya'); setVal('uRoleCustom', u.role); } togglePermissions(); toggleCustomRole(); document.querySelectorAll('#uPermissionsGrid input').forEach(c => c.checked = false); try { const p = JSON.parse(u.permissions || '[]'); p.forEach(val => { const cb = document.querySelector(`#uPermissionsGrid input[value="${val}"]`); if (cb) cb.checked = true; }); } catch (e) { } openModal('modalUser'); }
    function submitUser() {
      const id = v('uId'), un = v('uUsername'), pw = v('uPassword'), nm = v('uNama'), rSel = v('uRole'), r = rSel === 'Lainnya' ? v('uRoleCustom') : rSel;
      if (!un || !nm || (!id && !pw)) return toast('Lengkapi data', 'error');
      const perms = []; document.querySelectorAll('#uPermissionsGrid input:checked').forEach(c => perms.push(c.value));
      const div = v('uDivisi');
      const cb = res => { if (res.success) { toast('Berhasil'); closeModal('modalUser'); loadUsers(); } else toast(res.message, 'error'); };
      if (id) google.script.run.withSuccessHandler(cb).updateUser(id, un, pw, nm, r, JSON.stringify(perms), div); else google.script.run.withSuccessHandler(cb).addUser(un, pw, nm, r, JSON.stringify(perms), div);
    }
    function delUser(id) { if (confirm('Hapus?')) google.script.run.withSuccessHandler(res => { if (res.success) { toast('Dihapus'); loadUsers(); } }).deleteUser(id); }
    function doChangePass() { const o = v('oldPass'), n = v('newPass'); if (!o || !n) return toast('Isi password', 'error'); google.script.run.withSuccessHandler(res => { if (res.success) { toast('Password diganti'); resetForm(['oldPass', 'newPass']); } else toast(res.message, 'error') }).changePassword(currentUser.username, o, n); }

    // === INVENTORY CORE (Stock, In/Out/Retur, Order) ===
    function exportKaryawanToExcel() {
      if (!karyawanData || !karyawanData.length) return toast('Tidak ada data', 'info');
      let html = `<table border="1"><thead><tr><th>Nama</th><th>Jabatan</th><th>Cabang</th><th>Telepon</th><th>Email</th><th>Tgl Masuk</th><th>Status</th><th>Sisa Cuti</th></tr></thead><tbody>`;
      karyawanData.sort((a, b) => a.nama.localeCompare(b.nama)).forEach(k => {
        html += `<tr><td>${k.nama}</td><td>${k.jabatan}</td><td>${k.cabang}</td><td>${k.telepon || ''}</td><td>${k.email || ''}</td><td>${k.tanggalMasuk}</td><td>${k.status}</td><td>${k.sisaCuti}</td></tr>`;
      });
      html += '</tbody></table>';
      exportToExcel(html, 'Data_Karyawan_GudangFCL.xls');
    }

    function loadStock() {
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          stockData = res.data; renderStock(stockData); updateStockStats();
          updateStockDatalist();
          // Pemicu Notifikasi Stok Rendah
          stockData.forEach(s => {
            if (s.stok <= (s.stokMin || 0) && s.stokMin > 0) {
              addNotification(`Stok Rendah!`, `${s.sku} - ${s.nama} sisa ${s.stok} ${s.satuan}`, 'danger', 'stock');
            }
          });
        }
      }).getStock();
    }
    function updateStockStats() {
      document.getElementById('statTotalSKU').textContent = stockData.length; document.getElementById('statTotalStok').textContent = stockData.reduce((s, d) => s + d.stok, 0);
      document.getElementById('statStokRendah').textContent = stockData.filter(d => d.stok <= d.stokMin).length;
      const now = new Date(); document.getElementById('statExpSoon').textContent = stockData.filter(d => { if (!d.expDate) return false; const df = (new Date(d.expDate) - now) / 86400000; return df >= 0 && df <= 30; }).length;
    }
    function renderStock(data) {
      const tb = document.getElementById('tableStock'); tb.innerHTML = ''; if (!data.length) { tb.innerHTML = '<tr><td colspan="12" class="empty-state">Kosong</td></tr>'; return; }
      const now = new Date(); data.forEach(d => {
        let sCls = 'stok-aman', sTxt = 'Aman'; if (d.stok <= 0) { sCls = 'stok-kritis'; sTxt = 'Habis'; } else if (d.stok <= d.stokMin) { sCls = 'stok-rendah'; sTxt = 'Rendah'; }
        let expHtml = formatDate(d.expDate); if (d.expDate) { const diff = (new Date(d.expDate) - now) / 86400000; if (diff <= 30) expHtml = `<span class="exp-warning">⚠️ ${expHtml}</span>`; }
        tb.innerHTML += `<tr><td><strong>${d.sku}</strong></td><td>${d.nama}</td><td>${d.barcode || '-'}</td><td>${d.batch || '-'}</td><td>${expHtml}</td><td><strong style="font-size:16px">${d.stok}</strong></td><td>${d.stokMin}</td><td>${d.satuan}</td><td>${d.kategori}</td><td>${d.lokasi}</td><td><span class="stok-badge ${sCls}">${sTxt}</span></td><td><div style="display:flex;gap:4px;"><button class="btn btn-ghost btn-sm" onclick="editStock('${d.id}')" title="Edit">✏️</button><button class="btn btn-teal btn-sm" onclick="openMoveStockModal('${d.id}')" title="Move Stock ke Lokasi Lain">🔄</button></div></td></tr>`;
      });
    }
    function filterStock() { const q = v('stockSearch').toLowerCase(); renderStock(stockData.filter(d => d.sku.toLowerCase().includes(q) || d.nama.toLowerCase().includes(q) || d.kategori.toLowerCase().includes(q))); }

    // FIX: Perbaikan editStock - SKU undefined bug karena mapping key salah
    function editStock(id) {
      const d = stockData.find(x => x.id === id); if (!d) return;
      setVal('stockId', d.id);
      document.getElementById('stockModalTitle').textContent = '✏️ Edit Barang';
      // Mapping eksplisit untuk menghindari bug capitalisation (sKU -> undefined)
      setVal('stSKU', d.sku || '');
      setVal('stNama', d.nama || '');
      setVal('stBarcode', d.barcode || '');
      setVal('stBatch', d.batch || '');
      setVal('stExpDate', d.expDate || '');
      setVal('stSatuan', d.satuan || 'PCS');
      setVal('stStok', d.stok !== undefined ? d.stok : '');
      setVal('stStokMin', d.stokMin !== undefined ? d.stokMin : '');
      setVal('stKategori', d.kategori || '');
      setVal('stLokasi', d.lokasi || '');
      openModal('modalStock');
    }
    function submitStock() {
      const id = v('stockId'), s = v('stSKU'), n = v('stNama'), bc = v('stBarcode'), b = v('stBatch'), exp = v('stExpDate'), sat = v('stSatuan'), stk = v('stStok'), stkm = v('stStokMin'), k = v('stKategori'), l = v('stLokasi');
      if (!n) return toast('Nama wajib', 'error');
      const cb = res => { if (res.success) { toast('Berhasil'); closeModal('modalStock'); loadStock(); } else toast(res.message, 'error') };
      if (id) google.script.run.withSuccessHandler(cb).updateStock(id, s, n, bc, b, exp, sat, stk, stkm, k, l); else google.script.run.withSuccessHandler(cb).addStock(s, n, bc, b, exp, sat, stk, stkm, k, l);
    }
    function doExportStock() {
      if (!stockData.length) return toast('Data kosong', 'error');
      let csv = 'SKU,Nama Barang,Barcode,Batch,Exp Date,Satuan,Stok,Stok Min,Kategori,Lokasi\n';
      stockData.forEach(d => csv += `"${d.sku}","${d.nama}","${d.barcode}","${d.batch}","${d.expDate}","${d.satuan}","${d.stok}","${d.stokMin}","${d.kategori}","${d.lokasi}"\n`);
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' }); const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = 'Data_Stock_FCL.csv'; document.body.appendChild(a); a.click(); document.body.removeChild(a);
    }

    // ITEM ROWS UTILS
    function updateStockDatalist() {
      const dl = document.getElementById('listStockSKU'); if (!dl) return;
      dl.innerHTML = '';
      stockData.forEach(s => {
        const opt = document.createElement('option');
        opt.value = s.sku;
        opt.textContent = `${s.sku} - ${s.nama}${s.batch ? ' [Batch: ' + s.batch + ']' : ''}`;
        dl.appendChild(opt);
      });
    }
    function addItemRow(containerId, type) {
      const c = document.getElementById(containerId); const div = document.createElement('div'); div.className = 'item-row';
      const isOrder = type === 'order';
      div.innerHTML = `
    <input type="hidden" class="stock-id">
    <input type="text" class="form-control" ${isOrder ? 'list="listLokasiSelect" onfocus="prepareLokasiList(this)" onchange="syncBatchExp(this)"' : 'readonly'} placeholder="Lokasi" style="text-align:center; background:rgba(255,255,255,0.03) !important;">
    <input type="text" class="form-control sku-search" list="listStockSKU" onchange="handleSkuSearch(this, '${type}')" placeholder="SKU...">
    <input type="text" class="form-control" readonly placeholder="Nama Barang"> 
    <input type="text" class="form-control" ${isOrder ? 'readonly' : ''} placeholder="Batch"> 
    <input type="date" class="form-control" ${isOrder ? 'readonly' : ''} placeholder="Exp"> 
    <div style="position:relative;">
      <input type="number" class="form-control" min="1" placeholder="Qty" style="font-weight:800; font-size:15px; border-color:var(--teal);">
      <small class="satuan-label" style="position:absolute; right:6px; bottom:2px; font-size:9px; color:var(--teal); font-weight:800; pointer-events:none;"></small>
    </div>
    <button class="btn-remove" onclick="removeItemRow(this)">✕</button>`;
      c.appendChild(div);
      return div;
    }
    function handleSkuSearch(input, type) {
      const sku = input.value, row = input.parentElement, inputs = row.querySelectorAll('input');
      const s = stockData.find(x => x.sku === sku || x.barcode === sku);

      // Reset dependent fields when SKU changes
      inputs[1].value = ''; // Lokasi
      inputs[4].value = ''; // Batch
      inputs[5].value = ''; // Exp
      row.querySelector('.stock-id').value = '';

      if (s) {
        inputs[3].value = s.nama;           // Nama
        row.querySelector('.satuan-label').textContent = s.satuan || '';
        if (type !== 'order') {
          // Non-order types (Inbound/Outbound) still auto-fill first match
          row.querySelector('.stock-id').value = s.id;
          inputs[1].value = s.lokasi || '-';
          inputs[4].value = s.batch || '';
          inputs[5].value = s.expDate || '';
        } else {
          toast('Pilih Lokasi Pengambilan', 'info');
        }
      } else {
        inputs[3].value = '';
        row.querySelector('.satuan-label').textContent = '';
      }
    }

    function handleOrderScan(inputEl) {
      const code = inputEl.value.trim().toLowerCase();
      if (!code) return;

      // Cari di stockData (sudah ada di frontend)
      const matches = stockData.filter(s =>
        (s.barcode && String(s.barcode).toLowerCase() === code) ||
        (s.sku && String(s.sku).toLowerCase() === code)
      );

      if (matches.length === 0) {
        toast('Barang tidak ditemukan!', 'error');
        inputEl.value = '';
        inputEl.focus();
        return;
      }

      // Pilih item pertama yang ada stoknya, atau item pertama jika semua kosong
      let item = matches.find(s => s.stok > 0) || matches[0];

      console.log('📦 Item ditemukan:', item.nama, '| Lokasi:', item.lokasi, '| Stok:', item.stok);

      // Cek apakah item ini sudah ada di daftar order
      const existingRows = document.querySelectorAll('#ordItems .item-row');
      let foundRow = null;
      existingRows.forEach(row => {
        if (row.querySelector('.stock-id').value === item.id) {
          foundRow = row;
        }
      });

      if (foundRow) {
        // Jika sudah ada, tambah qty nya
        const inputs = foundRow.querySelectorAll('input');
        const qtyInp = inputs[6];
        qtyInp.value = (parseInt(qtyInp.value) || 0) + 1;
        // Trigger validasi visual
        if (qtyInp.oninput) qtyInp.oninput();
        toast(`✅ ${item.nama} ditambahkan (+1)`, 'success');
      } else {
        // Jika belum ada, tambah baris baru dan isi datanya
        const newRow = addItemRow('ordItems', 'order');
        const inputs = newRow.querySelectorAll('input');

        newRow.querySelector('.stock-id').value = item.id;
        // inputs[1] = Lokasi — TIDAK diisi otomatis saat scan barcode
        inputs[2].value = item.sku;    // SKU
        inputs[3].value = item.nama;   // Nama
        inputs[4].value = item.batch || ''; // Batch

        // Konversi format tanggal untuk input type=date
        let d = item.expDate || '';
        if (d && d.includes('/')) {
          const p = d.split('/');
          if (p.length === 3) d = `${p[2]}-${p[1].padStart(2, '0')}-${p[0].padStart(2, '0')}`;
        }
        inputs[5].value = d;           // Exp
        inputs[6].value = 1;           // Qty

        newRow.querySelector('.satuan-label').textContent = item.satuan || '';

        // Set validasi stok (simpan info lokasi sebagai referensi tanpa mengisi field)
        inputs[6].setAttribute('data-max-stok', item.stok || 0);
        inputs[6].setAttribute('data-lokasi-label', `${item.lokasi || '-'} (Stok: ${item.stok})`);
        setOrderQtyValidation(inputs[6], item.stok);

        toast(`✅ Berhasil scan: ${item.nama}`, 'success');
      }

      inputEl.value = '';
      inputEl.focus();
    }

    function prepareLokasiList(input) {
      const row = input.parentElement, sku = row.querySelector('.sku-search').value;
      if (!sku) { input.blur(); return toast('Isi SKU terlebih dahulu', 'warning'); }

      const dl = document.getElementById('listLokasiSelect');
      dl.innerHTML = '';
      const available = stockData.filter(x => x.sku === sku);
      if (!available.length) return toast('SKU tidak ditemukan di stok', 'error');

      available.forEach(s => {
        const opt = document.createElement('option');
        // Use a unique pattern for identification
        opt.value = `${s.lokasi} | ${s.batch || '-'}`;
        opt.textContent = `📍 SKU: ${s.sku} | Stok: ${s.stok} | Exp: ${s.expDate || '-'}`;
        dl.appendChild(opt);
      });
    }

    function syncBatchExp(input) {
      const row = input.parentElement, val = input.value, sku = row.querySelector('.sku-search').value;
      const inputs = row.querySelectorAll('input');

      if (!val) return;
      const parts = val.split(' | ');
      const lok = parts[0];
      const batch = parts[1] === '-' ? '' : parts[1];

      // Find the specific item matching SKU, Location, and Batch
      const match = stockData.find(x => x.sku === sku && x.lokasi === lok && (x.batch || '') === (batch || ''));

      if (match) {
        row.querySelector('.stock-id').value = match.id;
        input.value = match.lokasi; // Clean up the input to show only Location
        inputs[4].value = match.batch || '';

        // Handle Date Conversion (Enforce YYYY-MM-DD for <input type="date">)
        let d = match.expDate || '';
        if (d && d.includes('/')) {
          const p = d.split('/');
          if (p.length === 3) d = `${p[2]}-${p[1].padStart(2, '0')}-${p[0].padStart(2, '0')}`;
        }
        inputs[5].value = d;

        // Simpan stok tersedia di data-attribute input qty untuk validasi
        const qtyInput = inputs[6];
        qtyInput.setAttribute('data-max-stok', match.stok || 0);
        qtyInput.setAttribute('data-lokasi-label', `${match.lokasi} (Stok: ${match.stok})`);
        qtyInput.max = match.stok || 9999;
        setOrderQtyValidation(qtyInput, match.stok);

        toast(`✅ Sync: ${match.lokasi} | Stok tersedia: ${match.stok} ${match.satuan || ''}`, 'success');
      } else {
        // Fallback for direct typing or partial matches
        const directMatch = stockData.find(x => x.sku === sku && x.lokasi === val);
        if (directMatch) {
          row.querySelector('.stock-id').value = directMatch.id;
          inputs[4].value = directMatch.batch || '';
          let d = directMatch.expDate || '';
          if (d && d.includes('/')) {
            const p = d.split('/');
            if (p.length === 3) d = `${p[2]}-${p[0].padStart(2, '0')}-${p[1].padStart(2, '0')}`;
          }
          inputs[5].value = d;
        } else {
          inputs[4].value = '';
          inputs[5].value = '';
          row.querySelector('.stock-id').value = '';
          if (val && !val.includes('|')) toast('Data tidak singkron', 'error');
        }
      }
    }
    function handleItemSelect(sel) { /* Legacy support if needed */ }
    function removeItemRow(btn) { btn.parentElement.remove(); }

    // Validasi qty order tidak melebihi stok di lokasi
    function setOrderQtyValidation(qtyInput, maxStok) {
      // Hapus listener lama dulu
      qtyInput.oninput = null;
      qtyInput.oninput = function () {
        const val = parseFloat(this.value) || 0;
        const max = parseFloat(this.getAttribute('data-max-stok')) || 0;
        const label = this.getAttribute('data-lokasi-label') || '';
        let warn = this.parentElement.querySelector('.stok-warn');
        if (!warn) {
          warn = document.createElement('small');
          warn.className = 'stok-warn';
          warn.style.cssText = 'color:var(--red);font-weight:700;font-size:10px;display:block;margin-top:2px;';
          this.parentElement.appendChild(warn);
        }
        if (max > 0 && val > max) {
          warn.textContent = `⚠️ Melebihi stok! Max: ${max}`;
          this.style.borderColor = 'var(--red)';
        } else if (max > 0) {
          warn.textContent = `✓ Stok: ${max}`;
          warn.style.color = 'var(--green)';
          this.style.borderColor = 'var(--teal)';
        } else {
          warn.textContent = '';
        }
      };
    }

    function collectItems(containerId, type) {
      const items = []; let valid = true;
      let overStockMsg = '';
      document.querySelectorAll(`#${containerId} .item-row`).forEach(row => {
        const sId = row.querySelector('.stock-id').value;
        const inputs = row.querySelectorAll('input');
        // Indexing: 0:hidden_stockid, 1:lokasi, 2:sku, 3:nama, 4:batch, 5:exp, 6:qty
        if (sId && inputs[6].value) {
          const qty = parseFloat(inputs[6].value) || 0;
          const maxStok = parseFloat(inputs[6].getAttribute('data-max-stok') || 0);
          // Validasi: hanya untuk form order, stok tidak boleh melebihi
          if (type === 'order' && maxStok > 0 && qty > maxStok) {
            overStockMsg = `SKU ${inputs[2].value}: Qty ${qty} melebihi stok di lokasi (${maxStok})`;
            valid = false;
            return;
          }
          const itm = {
            stockId: sId,
            sku: inputs[2].value,
            nama: inputs[3].value,
            qty: qty,
            satuan: row.querySelector('.satuan-label').textContent,
            batch: inputs[4].value,
            expDate: inputs[5].value,
            lokasi: inputs[1].value
          };
          items.push(itm);
        } else valid = false;
      });
      if (overStockMsg) { toast('❌ ' + overStockMsg, 'error'); return null; }
      return valid && items.length > 0 ? items : null;
    }

    // ============================================================
    // MOVE STOCK - Pindah Stok ke Lokasi Lain
    // ============================================================
    function openMoveStockModal(stockId) {
      // Populate SKU dropdown
      const sel = document.getElementById('msSkuSelect');
      sel.innerHTML = '<option value="">-- Pilih SKU --</option>';
      const seen = new Set();
      stockData.forEach(s => {
        const key = s.id;
        if (!seen.has(key)) {
          seen.add(key);
          sel.innerHTML += `<option value="${s.id}">${s.sku} — ${s.nama} (${s.lokasi}, Stok: ${s.stok})</option>`;
        }
      });
      // Reset fields
      setVal('msStockId', '');
      setVal('msJumlah', '');
      setVal('msLokasiTujuan', '');
      setVal('msKeterangan', '');
      document.getElementById('msInfoBox').style.display = 'none';
      document.getElementById('msQtyWarning').style.display = 'none';
      // Pre-select if stockId given
      if (stockId) { sel.value = stockId; onMoveStockSkuChange(); }
      openModal('modalMoveStock');
    }

    function onMoveStockSkuChange() {
      const sel = document.getElementById('msSkuSelect');
      const sId = sel.value;
      const d = stockData.find(x => x.id === sId);
      const box = document.getElementById('msInfoBox');
      if (!d) { box.style.display = 'none'; return; }
      setVal('msStockId', d.id);
      document.getElementById('msNamaInfo').textContent = d.nama;
      document.getElementById('msLokasiInfo').textContent = d.lokasi;
      document.getElementById('msStokInfo').textContent = d.stok + ' ' + (d.satuan || '');
      document.getElementById('msBatchInfo').textContent = d.batch || '-';
      const qtyInput = document.getElementById('msJumlah');
      qtyInput.max = d.stok;
      qtyInput.setAttribute('data-max-stok', d.stok);
      setVal('msJumlah', '');
      document.getElementById('msQtyWarning').style.display = 'none';
      box.style.display = 'block';
    }

    function validateMoveStockQty() {
      const qty = parseFloat(v('msJumlah')) || 0;
      const max = parseFloat(document.getElementById('msJumlah').getAttribute('data-max-stok') || 0);
      const warn = document.getElementById('msQtyWarning');
      warn.style.display = (max > 0 && qty > max) ? 'block' : 'none';
    }

    function submitMoveStock() {
      const sId = v('msStockId');
      const jumlah = parseFloat(v('msJumlah')) || 0;
      const lokasiTujuan = v('msLokasiTujuan').trim();
      const ket = v('msKeterangan').trim();
      if (!sId) return toast('Pilih barang (SKU) terlebih dahulu!', 'error');
      if (!jumlah || jumlah <= 0) return toast('Jumlah harus lebih dari 0!', 'error');
      if (!lokasiTujuan) return toast('Lokasi tujuan wajib diisi!', 'error');

      const d = stockData.find(x => x.id === sId);
      if (!d) return toast('Data stok tidak ditemukan!', 'error');
      if (jumlah > d.stok) return toast(`❌ Jumlah (${jumlah}) melebihi stok tersedia (${d.stok})!`, 'error');
      if (lokasiTujuan.toLowerCase() === (d.lokasi || '').toLowerCase()) return toast('Lokasi tujuan sama dengan lokasi asal!', 'warning');

      const btn = document.querySelector('#modalMoveStock .btn-teal');
      const oldTxt = btn.textContent; btn.disabled = true; btn.textContent = '⏳ Memproses...';

      google.script.run
        .withSuccessHandler(res => {
          btn.disabled = false; btn.textContent = oldTxt;
          if (res.success) {
            toast(`✅ Stok dipindah! ${jumlah} ${d.satuan} dari ${d.lokasi} → ${lokasiTujuan}`, 'success');
            closeModal('modalMoveStock');
            loadStock();
          } else toast('❌ ' + res.message, 'error');
        })
        .withFailureHandler(err => {
          btn.disabled = false; btn.textContent = oldTxt;
          toast('Error: ' + err, 'error');
        })
        .moveStock(sId, jumlah, lokasiTujuan, ket, currentUser.username);
    }

    // INBOUND, OUTBOUND, RETUR
    function loadInbound() {
      google.script.run.withSuccessHandler(res => {
        const tb = document.getElementById('tableInbound');
        tb.innerHTML = '';
        if (!res.data.length) {
          tb.innerHTML = '<tr><td colspan="6" class="empty-state">Kosong</td></tr>';
          return;
        }
        res.data.forEach(d => {
          tb.innerHTML += `
            <tr id="inbound-main-${d.id}">
              <td>
                <button class="btn-toggle-row" onclick="toggleSJItems('${d.id}','masuk','${d.noSJ}', 'inbound')">
                  <i class="bi bi-plus-circle"></i>
                </button>
                <strong>${d.noSJ}</strong>
              </td>
              <td>${formatDate(d.tanggal)}</td>
              <td>${d.supplier}</td>
              <td>${d.keterangan}</td>
              <td>${d.createdBy}</td>
              <td></td>
            </tr>
            <tr id="inbound-detail-${d.id}" class="row-detail" style="display:none;">
              <td colspan="6"><div class="detail-container"><div class="loading-inline">Memuat detail...</div></div></td>
            </tr>`;
        });
      }).withFailureHandler(err => {
        toast('Gagal memuat Inbound: ' + err, 'error');
      }).getSuratJalanMasuk();
    }
    function submitInbound() { const t = v('inbTanggal'), s = v('inbSupplier'), k = v('inbKet'), items = collectItems('inbItems', 'inbound'); if (!t || !s || !items) return toast('Lengkapi data & barang', 'error'); const btn = document.querySelector('#modalInbound .btn-primary'); btn.disabled = true; google.script.run.withSuccessHandler(res => { btn.disabled = false; if (res.success) { toast('Berhasil'); closeModal('modalInbound'); loadInbound(); loadStock(); document.getElementById('inbItems').innerHTML = ''; resetForm(['inbSupplier', 'inbKet']); } else toast(res.message, 'error') }).addSuratJalanMasuk(t, s, k, JSON.stringify(items), currentUser.username); }
    function renderDetailTable(items) {
      if (!items || !items.length) return '<div style="padding:10px; color:var(--text-muted);">Tidak ada barang.</div>';
      let html = `
        <table class="detail-table">
          <thead>
            <tr>
              <th>Lokasi</th>
              <th>SKU</th>
              <th>Nama Barang</th>
              <th>Batch</th>
              <th>Exp</th>
              <th>Qty</th>
              <th>Satuan</th>
            </tr>
          </thead>
          <tbody>`;
      items.forEach(itm => {
        html += `
          <tr>
            <td><span class="badge-tb" style="background:var(--navy3); color:var(--teal);">${itm.lokasi || '-'}</span></td>
            <td><strong>${itm.sku || '-'}</strong></td>
            <td>${itm.nama || '-'}</td>
            <td>${itm.batch || '-'}</td>
            <td>${itm.expDate || '-'}</td>
            <td><strong style="color:var(--accent)">${itm.qty || 0}</strong></td>
            <td>${itm.satuan || ''}</td>
          </tr>`;
      });
      html += `</tbody></table>`;
      return html;
    }

    function toggleSJItems(id, tipe, noSJ, prefix) {
      const detailRow = document.getElementById(`${prefix}-detail-${id}`);
      const btn = document.querySelector(`#${prefix}-main-${id} .btn-toggle-row`);
      if (detailRow.style.display === 'none') {
        const cont = detailRow.querySelector('.detail-container');
        if (cont.innerHTML.includes('Memuat detail...')) {
          google.script.run.withSuccessHandler(res => {
            if (res.success) {
              cont.innerHTML = `
                <div style="margin-bottom:10px; font-weight:700; font-size:11px; color:var(--teal); text-transform:uppercase;">📦 Detail Barang - ${noSJ}</div>
                ${renderDetailTable(res.data)}`;
            } else {
              cont.innerHTML = `<span style="color:var(--red)">Gagal memuat: ${res.message}</span>`;
            }
          }).withFailureHandler(err => {
            cont.innerHTML = `<span style="color:var(--red)">Gagal memuat detail: ${err}</span>`;
            console.error('SJ Detail Error:', err);
          }).getSJDetailData(id, tipe, noSJ);
        }
        detailRow.style.display = 'table-row';
        btn.innerHTML = '<i class="bi bi-dash-circle"></i>';
        btn.classList.add('active');
      } else {
        detailRow.style.display = 'none';
        btn.innerHTML = '<i class="bi bi-plus-circle"></i>';
        btn.classList.remove('active');
      }
    }

    function toggleReturItems(id, noRetur) {
      const detailRow = document.getElementById(`retur-detail-${id}`);
      const btn = document.querySelector(`#retur-main-${id} .btn-toggle-row`);
      if (detailRow.style.display === 'none') {
        const cont = detailRow.querySelector('.detail-container');
        if (cont.innerHTML.includes('Memuat detail...')) {
          google.script.run.withSuccessHandler(res => {
            if (res.success) {
              cont.innerHTML = `
                <div style="margin-bottom:10px; font-weight:700; font-size:11px; color:var(--teal); text-transform:uppercase;">📦 Detail Retur - ${noRetur}</div>
                ${renderDetailTable(res.data)}`;
            } else {
              cont.innerHTML = `<span style="color:var(--red)">Gagal memuat: ${res.message}</span>`;
            }
          }).withFailureHandler(err => {
            cont.innerHTML = `<span style="color:var(--red)">Gagal memuat detail: ${err}</span>`;
            console.error('Retur Detail Error:', err);
          }).getReturDetail(id, noRetur);
        }
        detailRow.style.display = 'table-row';
        btn.innerHTML = '<i class="bi bi-dash-circle"></i>';
        btn.classList.add('active');
      } else {
        detailRow.style.display = 'none';
        btn.innerHTML = '<i class="bi bi-plus-circle"></i>';
        btn.classList.remove('active');
      }
    }

    function toggleOrderItems(id, noOrder) {
      const detailRow = document.getElementById(`order-detail-${id}`);
      const btn = document.querySelector(`#order-main-${id} .btn-toggle-row`);
      if (detailRow.style.display === 'none') {
        const cont = detailRow.querySelector('.detail-container');
        if (cont.innerHTML.includes('Memuat item...')) {
          google.script.run.withSuccessHandler(res => {
            if (res.success) {
              // Pertahankan Header detail jika ada
              const headers = cont.querySelectorAll('div:not(.loading-inline)');
              let headerHtml = '';
              headers.forEach(h => headerHtml += h.outerHTML);

              cont.innerHTML = `
                ${headerHtml}
                ${renderDetailTable(res.data)}`;
            } else {
              cont.innerHTML = `<span style="color:var(--red)">Gagal memuat: ${res.message}</span>`;
            }
          }).withFailureHandler(err => {
            cont.innerHTML = `<span style="color:var(--red)">Gagal memuat detail: ${err}</span>`;
            console.error('Order Detail Error:', err);
          }).getOrderDetail(id, noOrder);
        }
        detailRow.style.display = 'table-row';
        btn.innerHTML = '<i class="bi bi-dash-circle"></i>';
        btn.classList.add('active');
      } else {
        detailRow.style.display = 'none';
        btn.innerHTML = '<i class="bi bi-plus-circle"></i>';
        btn.classList.remove('active');
      }
    }

    function loadOutbound() {
      google.script.run.withSuccessHandler(res => {
        const tb = document.getElementById('tableOutbound');
        tb.innerHTML = '';
        if (!res.data.length) {
          tb.innerHTML = '<tr><td colspan="6" class="empty-state">Kosong</td></tr>';
          return;
        }
        res.data.forEach(d => {
          const detailHtml = `
            <div style="margin-bottom:10px; font-weight:700; font-size:11px; color:var(--teal); text-transform:uppercase;">📦 Detail Barang - ${d.noSJ}</div>
            ${renderDetailTable(d.items)}`;

          tb.innerHTML += `
            <tr id="outbound-main-${d.id}">
              <td>
                <button class="btn-toggle-row" onclick="toggleSJItems('${d.id}','keluar','${d.noSJ}', 'outbound')">
                  <i class="bi bi-plus-circle"></i>
                </button>
                <strong>${d.noSJ}</strong>
              </td>
              <td>${formatDate(d.tanggal)}</td>
              <td><span class="badge-tb">${d.tujuan}</span></td>
              <td>${d.keterangan}</td>
              <td>${d.createdBy}</td>
              <td></td>
            </tr>
            <tr id="outbound-detail-${d.id}" class="row-detail" style="display:none;">
              <td colspan="6"><div class="detail-container">${detailHtml}</div></td>
            </tr>`;
        });
      }).withFailureHandler(err => {
        toast('Gagal memuat Outbound: ' + err, 'error');
      }).getSJKeluarWithDetails();
    }
    function submitOutbound() { const t = v('outTanggal'), tj = v('outTujuan'), k = v('outKet'), items = collectItems('outItems', 'outbound'); if (!t || !items) return toast('Lengkapi data', 'error'); const btn = document.querySelector('#modalOutbound .btn-danger'); btn.disabled = true; google.script.run.withSuccessHandler(res => { btn.disabled = false; if (res.success) { toast('Berhasil dipotong'); closeModal('modalOutbound'); loadOutbound(); loadStock(); document.getElementById('outItems').innerHTML = ''; resetForm(['outKet']); } else toast(res.message, 'error') }).addSuratJalanKeluar(t, tj, k, JSON.stringify(items), currentUser.username); }

    function loadRetur() {
      google.script.run.withSuccessHandler(res => {
        const tb = document.getElementById('tableRetur');
        tb.innerHTML = '';
        if (!res.data.length) {
          tb.innerHTML = '<tr><td colspan="8" class="empty-state">Kosong</td></tr>';
          return;
        }
        res.data.forEach(d => {
          const detailHtml = `
            <div style="margin-bottom:10px; font-weight:700; font-size:11px; color:var(--teal); text-transform:uppercase;">📦 Detail Retur - ${d.noRetur}</div>
            ${renderDetailTable(d.items)}`;

          tb.innerHTML += `
            <tr id="retur-main-${d.id}">
              <td>
                <button class="btn-toggle-row" onclick="toggleReturItems('${d.id}','${d.noRetur}')">
                  <i class="bi bi-plus-circle"></i>
                </button>
                <strong>${d.noRetur}</strong>
              </td>
              <td>${formatDate(d.tanggal)}</td>
              <td>${d.sumber}</td>
              <td><span class="badge-tb">${d.alasan}</span></td>
              <td>${d.keterangan}</td>
              <td>${d.createdBy}</td>
              <td></td>
              <td></td>
            </tr>
            <tr id="retur-detail-${d.id}" class="row-detail" style="display:none;">
              <td colspan="8"><div class="detail-container">${detailHtml}</div></td>
            </tr>`;
        });
      }).withFailureHandler(err => {
        toast('Gagal memuat Retur: ' + err, 'error');
      }).getReturWithDetails();
    }
    function submitRetur() { const t = v('retTanggal'), a = v('retAlasan'), s = v('retSumber'), k = v('retKet'), items = collectItems('retItems', 'retur'); if (!t || !s || !items) return toast('Lengkapi data', 'error'); const btn = document.querySelector('#modalRetur .btn-teal'); btn.disabled = true; google.script.run.withSuccessHandler(res => { btn.disabled = false; if (res.success) { toast('Berhasil'); closeModal('modalRetur'); loadRetur(); loadStock(); document.getElementById('retItems').innerHTML = ''; resetForm(['retSumber', 'retKet']); } else toast(res.message, 'error') }).addRetur(t, s, a, k, JSON.stringify(items), currentUser.username); }

    // ===== RETURN DISTRIBUTOR =====
    let returnDistributorData = [];
    let rdSettingsPenarikan = [];
    let rdSettingsBuyback = [];
    let rdSettingsBPOM = [];
    let rdItemRows = []; // array of {sku, batch, qty, expDate, kategoriReturn, keterangan}
    let rdDetailCache = {}; // { returnId: [detail items] }

    // ---- Helpers ----
    function rdDetectKategori(sku, batch) {
      // Pastikan sku dan batch adalah string
      const s = String(sku || '').trim().toLowerCase(), b = String(batch || '').trim().toLowerCase();
      if (!s) return 'Return Normal';
      const matchItem = (x) => {
        // Pastikan x.sku dan x.batch adalah string
        const xs = String(x.sku || '').toLowerCase(), xb = String(x.batch || '').toLowerCase();
        return xs === s && (xb === 'all' || xb === b);
      };
      if (rdSettingsPenarikan.some(matchItem)) return 'Return Penarikan';
      if (rdSettingsBuyback.some(matchItem)) return 'Return Buy Back';
      if (rdSettingsBPOM.some(matchItem)) return 'Return BPOM';
      return 'Return Normal';
    }

    // Format Rupiah untuk Return Distributor
    function formatRupiahInput(el) {
      let raw = el.value.replace(/[^\d]/g, '');
      if (!raw) { el.value = ''; return; }
      el.value = parseInt(raw, 10).toLocaleString('id-ID');
    }

    function getRupiahValue(id) {
      const raw = v(id);
      return parseFloat(String(raw).replace(/[^\d]/g, '')) || 0;
    }

    function formatRupiah(n) {
      return 'Rp ' + (parseFloat(n) || 0).toLocaleString('id-ID');
    }

    function rdKatBadgeInline(kat) {
      const map = {
        'Return Penarikan': 'background:rgba(239,68,68,0.15);color:#ef4444;border:1px solid rgba(239,68,68,0.4);',
        'Return Buy Back': 'background:rgba(245,158,11,0.15);color:#f59e0b;border:1px solid rgba(245,158,11,0.4);',
        'Return Normal': 'background:rgba(16,185,129,0.12);color:#10b981;border:1px solid rgba(16,185,129,0.3);',
        'Return Exp Date': 'background:rgba(14,165,233,0.12);color:#0ea5e9;border:1px solid rgba(14,165,233,0.3);',
        'Return BPOM': 'background:rgba(139,92,246,0.12);color:#8b5cf6;border:1px solid rgba(139,92,246,0.3);'
      };
      const style = map[kat] || 'background:rgba(148,163,184,0.1);color:#94a3b8;border:1px solid rgba(148,163,184,0.2);';
      const icons = { 'Return Penarikan': '⚠️', 'Return Buy Back': '💰', 'Return Normal': '✅', 'Return Exp Date': '📅', 'Return BPOM': '🏛️' };
      return `<span style="${style}padding:2px 8px;border-radius:12px;font-size:11px;font-weight:700;">${icons[kat] || ''} ${kat || '-'}</span>`;
    }

    // ---- Render item rows in modal ----
    function renderRDItemRows() {
      const tbody = document.getElementById('rdItemsBody');
      if (!tbody) return; // modal belum terbuka
      if (!rdItemRows.length) {
        tbody.innerHTML = `<tr><td colspan="8" style="text-align:center;color:var(--text-muted);padding:28px;font-size:13px;">
          Belum ada item. Klik <strong>+ Tambah Baris</strong> atau <strong>Impor Excel</strong>.
        </td></tr>`;
        const countEl = document.getElementById('rdItemCount');
        if (countEl) countEl.textContent = '0';
        return;
      }
      const countEl = document.getElementById('rdItemCount');
      if (countEl) countEl.textContent = rdItemRows.length;

      tbody.innerHTML = rdItemRows.map((row, idx) => {
        const katOpts = ['Return Normal', 'Return Exp Date', 'Return Penarikan', 'Return Buy Back', 'Return BPOM']
          .map(k => `<option value="${k}" ${row.kategoriReturn === k ? 'selected' : ''}>${k}</option>`).join('');
        const isFlagged = (row.kategoriReturn === 'Return Penarikan' || row.kategoriReturn === 'Return Buy Back' || row.kategoriReturn === 'Return BPOM');
        const flagBorder = isFlagged ? 'border-color:rgba(245,158,11,0.6);' : '';
        const flagBg = isFlagged ? 'background:rgba(245,158,11,0.05);' : '';
        const rowBg = isFlagged ? 'background:rgba(245,158,11,0.02);' : '';
        // Lebar input otomatis mengikuti konten (min 8ch)
        const skuW = Math.max(8, (row.sku || 'Kode SKU').length + 2) + 'ch';
        const batW = Math.max(8, (row.batch || 'No. Batch').length + 2) + 'ch';
        const qtyW = Math.max(4, String(row.qty || '0').length + 2) + 'ch';
        const ketW = Math.max(10, (row.keterangan || 'Opsional').length + 2) + 'ch';
        return `<tr id="rdRow_${idx}" style="border-bottom:1px solid var(--border-color);${rowBg}">
          <td style="padding:8px 8px;color:var(--text-muted);font-size:12px;text-align:center;white-space:nowrap;">${idx + 1}</td>
          <td style="padding:5px 5px;white-space:nowrap;">
            <input type="text" class="form-control form-control-sm"
              id="rdSku_${idx}"
              value="${(row.sku || '').replace(/"/g, '&quot;')}"
              placeholder="Kode SKU"
              style="${flagBorder}${flagBg}width:${skuW};min-width:100px;transition:width 0.1s;"
              oninput="rdItemRows[${idx}].sku=this.value; this.style.width=Math.max(8,this.value.length+2)+'ch'"
              onblur="rdRowFieldChange(${idx},'sku',this.value)">
          </td>
          <td style="padding:5px 5px;white-space:nowrap;">
            <input type="text" class="form-control form-control-sm"
              id="rdBatch_${idx}"
              value="${(row.batch || '').replace(/"/g, '&quot;')}"
              placeholder="No. Batch"
              style="${flagBorder}${flagBg}width:${batW};min-width:90px;transition:width 0.1s;"
              oninput="rdItemRows[${idx}].batch=this.value; this.style.width=Math.max(8,this.value.length+2)+'ch'"
              onblur="rdRowFieldChange(${idx},'batch',this.value)">
          </td>
          <td style="padding:5px 5px;white-space:nowrap;">
            <input type="number" class="form-control form-control-sm"
              id="rdQtyRow_${idx}"
              value="${row.qty || ''}"
              placeholder="0" min="1"
              style="text-align:center;width:${qtyW};min-width:60px;transition:width 0.1s;"
              oninput="rdItemRows[${idx}].qty=this.value; this.style.width=Math.max(4,this.value.length+2)+'ch'">
          </td>
          <td style="padding:5px 5px;white-space:nowrap;">
            <input type="date" class="form-control form-control-sm"
              id="rdExp_${idx}"
              value="${row.expDate || ''}"
              style="min-width:138px;"
              oninput="rdItemRows[${idx}].expDate=this.value">
          </td>
          <td style="padding:5px 5px;white-space:nowrap;">
            <select class="form-select form-select-sm"
              id="rdKat_${idx}"
              style="min-width:158px;"
              onchange="rdRowFieldChange(${idx},'kategoriReturn',this.value)">
              ${katOpts}
            </select>
          </td>
          <td style="padding:5px 5px;white-space:nowrap;">
            <input type="text" class="form-control form-control-sm"
              id="rdKet_${idx}"
              value="${(row.keterangan || '').replace(/"/g, '&quot;')}"
              placeholder="Opsional"
              style="width:${ketW};min-width:100px;transition:width 0.1s;"
              oninput="rdItemRows[${idx}].keterangan=this.value; this.style.width=Math.max(10,this.value.length+2)+'ch'">
          </td>
          <td style="padding:5px 5px;text-align:center;white-space:nowrap;">
            <button class="btn btn-danger btn-sm" style="padding:3px 8px;" onclick="removeRDItemRow(${idx})">🗑️</button>
          </td>
        </tr>`;
      }).join('');
    }

    // Dipanggil saat onblur — untuk auto-detect kategori setelah selesai ketik
    function rdRowFieldChange(idx, field, val) {
      if (!rdItemRows[idx]) return;
      rdItemRows[idx][field] = val;
      if (field === 'sku' || field === 'batch') {
        const detected = rdDetectKategori(rdItemRows[idx].sku, rdItemRows[idx].batch);
        if (detected !== rdItemRows[idx].kategoriReturn) {
          rdItemRows[idx].kategoriReturn = detected;
          const sel = document.getElementById(`rdKat_${idx}`);
          if (sel) sel.value = detected;
          const isFlagged = (detected === 'Return Penarikan' || detected === 'Return Buy Back' || detected === 'Return BPOM');
          const tr = document.getElementById(`rdRow_${idx}`);
          if (tr) tr.style.background = isFlagged ? 'rgba(245,158,11,0.02)' : '';
          const skuEl = document.getElementById(`rdSku_${idx}`);
          const batEl = document.getElementById(`rdBatch_${idx}`);
          const border = isFlagged ? 'rgba(245,158,11,0.6)' : '';
          const bg = isFlagged ? 'rgba(245,158,11,0.05)' : '';
          if (skuEl) { skuEl.style.borderColor = border; skuEl.style.background = bg; }
          if (batEl) { batEl.style.borderColor = border; batEl.style.background = bg; }
        }
      }
    }

    function addRDItemRow(prefill) {
      rdItemRows.push({
        sku: prefill ? (prefill.sku || '') : '',
        batch: prefill ? (prefill.batch || '') : '',
        qty: prefill ? (prefill.qty || '') : '',
        expDate: prefill ? (prefill.expDate || '') : '',
        kategoriReturn: prefill ? (prefill.kategoriReturn || 'Return Normal') : 'Return Normal',
        keterangan: prefill ? (prefill.keterangan || '') : ''
      });
      renderRDItemRows();
      // Fokus ke input SKU baris terakhir (pakai id yang benar)
      setTimeout(() => {
        const lastIdx = rdItemRows.length - 1;
        const el = document.getElementById(`rdSku_${lastIdx}`);
        if (el) { el.focus(); el.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); }
      }, 60);
    }

    function removeRDItemRow(idx) {
      rdItemRows.splice(idx, 1);
      renderRDItemRows();
    }

    // ---- Download Template Excel ----
    function downloadRDTemplate() {
      const headers = ['Tanggal (YYYY-MM-DD)', 'Nama Distributor', 'SKU', 'Batch', 'Qty', 'Exp Date (YYYY-MM-DD)', 'Kategori Return', 'Keterangan'];
      const sample = [
        ['2025-01-15', 'PT Distributor Maju', 'SKU-001', 'BATCH-A1', 10, '2025-12-31', 'Return Normal', 'Contoh keterangan'],
        ['2025-01-15', 'PT Distributor Maju', 'SKU-002', 'BATCH-B2', 5, '2025-06-30', 'Return Exp Date', ''],
        ['2025-01-15', 'PT Distributor Maju', 'SKU-003', 'BATCH-C3', 3, '', 'Return Penarikan', 'Produk ditarik'],
        ['2025-01-15', 'PT Distributor Maju', 'SKU-004', 'BATCH-D4', 8, '', 'Return Buy Back', 'Buy back program'],
        ['2025-01-15', 'PT Distributor Maju', 'SKU-005', 'BATCH-E5', 2, '', 'Return BPOM', 'Penarikan BPOM'],
      ];
      const note = [['CATATAN: Kolom Kategori Return diisi salah satu: Return Normal / Return Exp Date / Return Penarikan / Return Buy Back / Return BPOM']];
      const ws = XLSX.utils.aoa_to_sheet([headers, ...sample, [], ...note]);
      ws['!cols'] = [22, 22, 14, 14, 8, 20, 18, 24].map(w => ({ wch: w }));
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Template Return Distributor');
      XLSX.writeFile(wb, 'Template_ReturnDistributor.xlsx');
    }

    // ---- Import Excel ----
    function handleImportRDExcel(input) {
      const file = input.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = function (e) {
        try {
          const wb = XLSX.read(e.target.result, { type: 'binary', cellDates: true });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          if (rows.length < 2) return toast('File kosong atau tidak ada data', 'error');

          const headerRow = rows[0].map(h => String(h || '').toLowerCase().trim());
          const colIdx = {
            tanggal: headerRow.findIndex(h => h.includes('tanggal')),
            distributor: headerRow.findIndex(h => h.includes('distributor') || h.includes('nama')),
            sku: headerRow.findIndex(h => h === 'sku' || h.includes('sku')),
            batch: headerRow.findIndex(h => h.includes('batch')),
            qty: headerRow.findIndex(h => h === 'qty' || h.includes('qty') || h.includes('jumlah')),
            expDate: headerRow.findIndex(h => h.includes('exp')),
            kategori: headerRow.findIndex(h => h.includes('kategori')),
            keterangan: headerRow.findIndex(h => h.includes('keterangan') || h.includes('catatan'))
          };

          let imported = 0, skipped = 0;
          let firstTanggal = '', firstDistributor = '';

          for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.every(c => c === '')) continue;

            const tanggal = colIdx.tanggal >= 0 ? formatExcelDate(row[colIdx.tanggal]) : '';
            const distributor = colIdx.distributor >= 0 ? String(row[colIdx.distributor] || '').trim() : '';
            const sku = colIdx.sku >= 0 ? String(row[colIdx.sku] || '').trim() : '';
            const batch = colIdx.batch >= 0 ? String(row[colIdx.batch] || '').trim() : '';
            const qty = colIdx.qty >= 0 ? String(row[colIdx.qty] || '').trim() : '';
            const expDate = colIdx.expDate >= 0 ? formatExcelDate(row[colIdx.expDate]) : '';
            const kategoriRaw = colIdx.kategori >= 0 ? String(row[colIdx.kategori] || '').trim() : '';
            const keterangan = colIdx.keterangan >= 0 ? String(row[colIdx.keterangan] || '').trim() : '';

            if (!sku || !batch) { skipped++; continue; }

            if (!firstTanggal && tanggal) firstTanggal = tanggal;
            if (!firstDistributor && distributor) firstDistributor = distributor;

            const validKat = ['Return Normal', 'Return Exp Date', 'Return Penarikan', 'Return Buy Back', 'Return BPOM'];
            // Pastikan kategoriRaw adalah string sebelum toLowerCase
            const kategori = validKat.find(k => k.toLowerCase() === String(kategoriRaw).toLowerCase()) || rdDetectKategori(sku, batch);

            rdItemRows.push({ sku, batch, qty, expDate, kategoriReturn: kategori, keterangan });
            imported++;
          }

          if (firstTanggal && !v('rdTanggal')) setVal('rdTanggal', firstTanggal);
          if (firstDistributor && !v('rdNamaDistributor')) setVal('rdNamaDistributor', firstDistributor);

          renderRDItemRows();
          toast(`✅ ${imported} baris berhasil diimpor${skipped ? `, ${skipped} baris dilewati (SKU/Batch kosong)` : ''}`, 'success');
        } catch (err) {
          toast('Gagal membaca file: ' + err.message, 'error');
        }
        input.value = '';
      };
      reader.readAsBinaryString(file);
    }

    function formatExcelDate(val) {
      if (!val) return '';
      if (val instanceof Date) {
        const y = val.getFullYear(), m = String(val.getMonth() + 1).padStart(2, '0'), d = String(val.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
      }
      const s = String(val).trim();
      // Format DD/MM/YYYY
      const dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
      if (dmy) return `${dmy[3]}-${dmy[2].padStart(2, '0')}-${dmy[1].padStart(2, '0')}`;
      // Format YYYY-MM-DD sudah benar
      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
      return '';
    }

    function loadReturnDistributor() {
      // Load settings + PIC Sales bersamaan
      google.script.run.withSuccessHandler(sRes => {
        if (sRes && sRes.success) {
          rdSettingsPenarikan = sRes.data.penarikan || [];
          rdSettingsBuyback = sRes.data.buyback || [];
          rdSettingsBPOM = sRes.data.bpom || [];
        }
      }).withFailureHandler(err => {
        console.error('Error loading Return Distributor Settings:', err);
        toast('Gagal memuat pengaturan Return Distributor: ' + err, 'error');
      }).getReturnDistributorSettings();

      google.script.run.withSuccessHandler(pRes => {
        if (pRes && pRes.success) {
          rdPICSalesList = pRes.data || [];
          rdRefreshPICSalesDropdown();
        }
      }).withFailureHandler(err => {
        console.error('Error loading PIC Sales:', err);
        toast('Gagal memuat PIC Sales: ' + err, 'error');
      }).getReturnDistributorPICSales();

      google.script.run.withSuccessHandler(res => {
        if (!res || !res.success) {
          console.error('Error in getReturnDistributor response:', res);
          toast(res ? res.message : 'Gagal memuat data Return Distributor', 'error');
          return;
        }
        console.log('✅ Return Distributor data loaded:', res.data.length, 'items');
        if (res.data.length > 0) {
          console.log('Sample data:', res.data[0]);
        }
        returnDistributorData = res.data || [];
        rdDetailCache = {};
        filterReturnDistributor();
        updateRDStats();
      }).withFailureHandler(err => {
        console.error('Error calling getReturnDistributor:', err);
        toast('Gagal memuat data Return Distributor: ' + err, 'error');
        returnDistributorData = [];
        filterReturnDistributor();
        updateRDStats();
      }).getReturnDistributor();
    }

    let rdPICSalesList = [];

    // ---- PIC Sales helpers ----
    function rdRefreshPICSalesDropdown() {
      const sel = document.getElementById('rdPICSalesSelect');
      if (!sel) return;
      const cur = sel.value;
      sel.innerHTML = '<option value="">-- Pilih PIC Sales --</option>';
      rdPICSalesList.forEach(name => {
        const opt = document.createElement('option');
        opt.value = name; opt.textContent = name;
        sel.appendChild(opt);
      });
      if (cur) sel.value = cur;
    }

    function rdOnPICSalesSelect() {
      const sel = document.getElementById('rdPICSalesSelect');
      const inp = document.getElementById('rdPICSales');
      if (sel && inp && sel.value) inp.value = sel.value;
    }

    // ---- Setting PIC Sales ----
    function openRDPICSalesSettings() {
      renderRDPICSalesTable();
      openModal('modalRDPICSales');
    }

    function renderRDPICSalesTable() {
      const tb = document.getElementById('rdPICSalesTable');
      if (!rdPICSalesList.length) {
        tb.innerHTML = '<tr><td colspan="3" style="text-align:center;color:var(--text-muted);padding:20px;">Belum ada PIC Sales</td></tr>';
        return;
      }
      tb.innerHTML = rdPICSalesList.map((name, idx) => `
        <tr style="border-bottom:1px solid var(--border-color);">
          <td style="padding:8px 12px;color:var(--text-muted);font-size:12px;">${idx + 1}</td>
          <td style="padding:8px 12px;font-weight:600;">${name}</td>
          <td style="padding:8px 12px;text-align:right;">
            <button class="btn btn-danger btn-sm" style="padding:2px 8px;" onclick="removeRDPICSales(${idx})">🗑️</button>
          </td>
        </tr>`).join('');
    }

    function addRDPICSales() {
      const inp = document.getElementById('rdNewPICSales');
      const name = (inp.value || '').trim();
      if (!name) return toast('Isi nama PIC Sales', 'error');
      if (rdPICSalesList.some(x => x.toLowerCase() === name.toLowerCase()))
        return toast('Nama sudah ada di daftar', 'error');
      rdPICSalesList.push(name);
      renderRDPICSalesTable();
      inp.value = '';
      inp.focus();
    }

    function removeRDPICSales(idx) {
      rdPICSalesList.splice(idx, 1);
      renderRDPICSalesTable();
    }

    function saveRDPICSales() {
      const btn = document.getElementById('btnSaveRDPICSales');
      btn.disabled = true; btn.textContent = '⏳ Menyimpan...';
      google.script.run.withSuccessHandler(res => {
        btn.disabled = false; btn.textContent = '💾 Simpan';
        if (res.success) {
          rdRefreshPICSalesDropdown();
          toast('Daftar PIC Sales berhasil disimpan');
          closeModal('modalRDPICSales');
        } else toast(res.message, 'error');
      }).saveReturnDistributorPICSales(rdPICSalesList);
    }

    let rdCurrentView = 'card'; // 'card' | 'table'

    function setRDView(view) {
      rdCurrentView = view;
      document.getElementById('rdCardView').style.display = view === 'card' ? 'block' : 'none';
      document.getElementById('rdTableView').style.display = view === 'table' ? 'block' : 'none';
      const btnCard = document.getElementById('rdViewCard');
      const btnTable = document.getElementById('rdViewTable');
      btnCard.style.cssText = view === 'card'
        ? 'padding:5px 10px;background:var(--accent);color:#000;font-weight:700;border-radius:6px;'
        : 'padding:5px 10px;background:transparent;color:var(--text-muted);border-radius:6px;';
      btnTable.style.cssText = view === 'table'
        ? 'padding:5px 10px;background:var(--accent);color:#000;font-weight:700;border-radius:6px;'
        : 'padding:5px 10px;background:transparent;color:var(--text-muted);border-radius:6px;';
      filterReturnDistributor();
    }

    function updateRDStats() {
      const d = returnDistributorData;
      document.getElementById('rdStatTotal').textContent = d.length;
      // Hitung dari jenisReturn header (lebih akurat & cepat)
      let cP = 0, cB = 0, cN = 0, cBPOM = 0;
      d.forEach(h => {
        // Pastikan jenisReturn adalah string sebelum toLowerCase
        const j = String(h.jenisReturn || '').toLowerCase();
        if (j === 'penarikan') cP++;
        else if (j === 'buy back') cB++;
        else if (j === 'bpom') cBPOM++;
        else cN++; // Normal, Exp Date, Sample, kosong
      });
      document.getElementById('rdStatPenarikan').textContent = cP;
      document.getElementById('rdStatBuyBack').textContent = cB;
      document.getElementById('rdStatNormal').textContent = cN;
      document.getElementById('rdStatBPOM').textContent = cBPOM;
    }

    function filterReturnDistributor() {
      const q = (v('rdSearch') || '').toLowerCase().trim();
      const kat = v('rdFilterKategori') || '';
      const jenis = v('rdFilterJenis') || '';
      const filtered = returnDistributorData.filter(d => {
        // Pastikan semua field adalah string
        const namaDistributor = String(d.namaDistributor || '');
        const noReturn = String(d.noReturn || '');

        const matchQ = !q || namaDistributor.toLowerCase().includes(q) ||
          noReturn.toLowerCase().includes(q);
        const matchJenis = !jenis || (d.jenisReturn || '') === jenis;
        // Filter kategori SKU: cek di detail cache
        let matchKat = !kat;
        if (kat && rdDetailCache[d.id]) {
          matchKat = rdDetailCache[d.id].some(x => x.kategoriReturn === kat);
        } else if (kat) matchKat = true;
        return matchQ && matchJenis && matchKat;
      });
      if (rdCurrentView === 'card') renderRDCardView(filtered);
      else renderRDTableView(filtered);
    }

    // ---- Inisial avatar distributor ----
    function rdDistributorAvatar(nama) {
      // Pastikan nama adalah string
      const namaStr = String(nama || 'D').trim();
      const words = namaStr.split(/\s+/);
      const initials = words.length >= 2
        ? (words[0][0] + words[1][0]).toUpperCase()
        : namaStr.substring(0, 2).toUpperCase();
      // Warna konsisten berdasarkan nama
      const colors = ['#0ea5e9', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#14b8a6', '#f97316'];
      let hash = 0;
      for (let i = 0; i < namaStr.length; i++) hash = namaStr.charCodeAt(i) + ((hash << 5) - hash);
      const color = colors[Math.abs(hash) % colors.length];
      return { initials, color };
    }

    // ---- Card View: 1 card per transaksi ----
    function renderRDCardView(filtered) {
      const container = document.getElementById('rdCardContainer');
      if (!filtered.length) {
        container.innerHTML = `<div style="text-align:center;padding:40px;color:var(--text-muted);">
          <div style="font-size:40px;margin-bottom:12px;">📭</div>
          <div style="font-size:14px;">Tidak ada data return</div>
        </div>`;
        return;
      }
      const sorted = [...filtered].sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
      container.innerHTML = sorted.map((h, gi) => {
        // Validasi dan konversi data ke string
        const namaDistributor = String(h.namaDistributor || 'Unknown');
        const noReturn = String(h.noReturn || '-');
        const tanggal = h.tanggal || '';
        const jenisReturn = String(h.jenisReturn || '');
        const createdBy = String(h.createdBy || '-');
        const picSales = h.picSales ? String(h.picSales) : '';
        const noMabang = h.noMabang ? String(h.noMabang) : '';
        const noResi = h.noResi ? String(h.noResi) : '';
        const hargaOngkir = Number(h.hargaOngkir) || 0;
        const totalSKU = h.totalSKU || 0;
        const totalQty = h.totalQty || 0;
        const keterangan = h.keterangan ? String(h.keterangan) : '';

        const av = rdDistributorAvatar(namaDistributor);
        const det = rdDetailCache[h.id] || [];
        const cP = det.filter(x => x.kategoriReturn === 'Return Penarikan').length;
        const cB = det.filter(x => x.kategoriReturn === 'Return Buy Back').length;
        const cN = det.filter(x => x.kategoriReturn === 'Return Normal').length;
        const cE = det.filter(x => x.kategoriReturn === 'Return Exp Date').length;
        const cBPOM = det.filter(x => x.kategoriReturn === 'Return BPOM').length;
        const hasFlagged = cP > 0 || cB > 0 || cBPOM > 0;
        const borderColor = cP > 0 ? 'rgba(239,68,68,0.4)' : cBPOM > 0 ? 'rgba(139,92,246,0.4)' : cB > 0 ? 'rgba(245,158,11,0.4)' : 'var(--border-color)';
        const summaryTags = [
          cP ? `<span style="background:rgba(239,68,68,0.12);color:var(--red);border:1px solid rgba(239,68,68,0.3);padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;">⚠️ ${cP} Penarikan</span>` : '',
          cB ? `<span style="background:rgba(245,158,11,0.12);color:var(--accent);border:1px solid rgba(245,158,11,0.3);padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;">💰 ${cB} Buy Back</span>` : '',
          cBPOM ? `<span style="background:rgba(139,92,246,0.12);color:#8b5cf6;border:1px solid rgba(139,92,246,0.3);padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;">🏛️ ${cBPOM} BPOM</span>` : '',
          cN ? `<span style="background:rgba(16,185,129,0.12);color:var(--green);border:1px solid rgba(16,185,129,0.3);padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;">✅ ${cN} Normal</span>` : '',
          cE ? `<span style="background:rgba(14,165,233,0.12);color:var(--teal);border:1px solid rgba(14,165,233,0.3);padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;">📅 ${cE} Exp</span>` : ''
        ].filter(Boolean).join('');

        // Detail rows (jika sudah di-cache)
        const detailContent = det.length > 0 ? `
          <table style="width:100%;border-collapse:collapse;font-size:13px;">
            <thead>
              <tr style="background:var(--bg-panel);">
                <th style="padding:8px 12px;font-size:11px;font-weight:800;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.8px;">SKU</th>
                <th style="padding:8px 12px;font-size:11px;font-weight:800;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.8px;">Batch</th>
                <th style="padding:8px 12px;font-size:11px;font-weight:800;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.8px;text-align:center;">Qty</th>
                <th style="padding:8px 12px;font-size:11px;font-weight:800;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.8px;">Exp Date</th>
                <th style="padding:8px 12px;font-size:11px;font-weight:800;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.8px;">Kategori</th>
                <th style="padding:8px 12px;font-size:11px;font-weight:800;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.8px;">Keterangan</th>
              </tr>
            </thead>
            <tbody>${det.map(item => {
          const expW = item.expDate && (new Date(item.expDate) - new Date()) / 86400000 <= 30 ? 'color:var(--red);font-weight:700;' : '';
          return `<tr style="border-bottom:1px solid var(--border-color);">
                <td style="padding:8px 12px;"><code style="background:var(--input-bg);padding:2px 8px;border-radius:6px;font-size:12px;font-weight:700;">${item.sku || '-'}</code></td>
                <td style="padding:8px 12px;">${item.batch || '-'}</td>
                <td style="padding:8px 12px;text-align:center;"><span style="background:var(--bg-panel-light);border:1px solid var(--border-color);padding:2px 10px;border-radius:8px;font-weight:800;">${item.qty || '-'}</span></td>
                <td style="padding:8px 12px;${expW}">${item.expDate ? formatDate(item.expDate) : '-'}</td>
                <td style="padding:8px 12px;">${getRDKategoriBadge(item.kategoriReturn)}</td>
                <td style="padding:8px 12px;color:var(--text-muted);font-size:12px;">${item.keterangan || '-'}</td>
              </tr>`;
        }).join('')}</tbody>
          </table>` : `<div style="text-align:center;padding:20px;color:var(--text-muted);font-size:13px;">⏳ Klik "Lihat Detail" untuk memuat SKU</div>`;

        return `<div style="border:1px solid ${borderColor};border-radius:12px;margin-bottom:12px;overflow:hidden;" id="rdCard_${gi}">
          <div style="display:flex;align-items:center;gap:14px;padding:14px 16px;background:var(--bg-panel-light);">
            <div style="width:44px;height:44px;border-radius:12px;background:${av.color};display:flex;align-items:center;justify-content:center;font-size:16px;font-weight:800;color:#fff;flex-shrink:0;">${av.initials}</div>
            <div style="flex:1;min-width:0;">
              <div style="font-size:15px;font-weight:700;color:var(--text-main);">${namaDistributor}</div>
              <div style="font-size:12px;color:var(--text-muted);margin-top:2px;display:flex;gap:10px;flex-wrap:wrap;">
                <span>🔖 <strong style="color:var(--teal);">${noReturn}</strong></span>
                <span>📅 ${formatDate(tanggal)}</span>
                <span>👤 ${createdBy}</span>
                ${picSales ? `<span>🧑‍💼 <strong>${picSales}</strong></span>` : ''}
                ${noMabang ? `<span>📋 Mabang: <strong>${noMabang}</strong></span>` : ''}
                ${noResi ? `<span>📦 Resi: <strong>${noResi}</strong></span>` : ''}
                ${hargaOngkir > 0 ? `<span>💰 Ongkir: <strong>${formatRupiah(hargaOngkir)}</strong></span>` : ''}
                <span>📦 <strong>${totalSKU}</strong> SKU</span>
                <span>🔢 Total Qty: <strong>${totalQty}</strong></span>
              </div>
            </div>
            <div style="display:flex;gap:6px;flex-wrap:wrap;justify-content:flex-end;flex-shrink:0;">
              ${jenisReturn ? getRDJenisBadge(jenisReturn) : ''}
              ${summaryTags}
            </div>
            <div style="display:flex;gap:6px;flex-shrink:0;">
              <button class="btn btn-sm" style="background:rgba(14,165,233,0.1);color:var(--teal);border:1px solid rgba(14,165,233,0.25);padding:5px 12px;font-size:12px;"
                onclick="rdToggleDetail('${h.id}',${gi})">📋 Detail</button>
              <button class="btn btn-ghost btn-sm" style="padding:5px 10px;" onclick="editReturnDistributor('${h.id}')">✏️</button>
              <button class="btn btn-danger btn-sm" style="padding:5px 10px;" onclick="delReturnDistributor('${h.id}')">🗑️</button>
            </div>
          </div>
          <div id="rdCardDetail_${gi}" style="display:none;">${detailContent}</div>
        </div>`;
      }).join('');
    }

    function rdToggleDetail(returnId, gi) {
      const panel = document.getElementById(`rdCardDetail_${gi}`);
      if (!panel) return;
      const isOpen = panel.style.display !== 'none';
      if (isOpen) { panel.style.display = 'none'; return; }
      // Jika belum ada cache, load dulu
      if (!rdDetailCache[returnId]) {
        panel.innerHTML = `<div style="text-align:center;padding:20px;color:var(--text-muted);">⏳ Memuat detail...</div>`;
        panel.style.display = 'block';
        google.script.run.withSuccessHandler(res => {
          if (res && res.success) {
            rdDetailCache[returnId] = res.data || [];
            renderRDDetailPanel(panel, res.data || []);
          } else {
            panel.innerHTML = `<div style="text-align:center;padding:20px;color:var(--red);">Gagal memuat detail</div>`;
          }
        }).getReturnDistributorDetail(returnId);
      } else {
        renderRDDetailPanel(panel, rdDetailCache[returnId]);
        panel.style.display = 'block';
      }
    }

    function renderRDDetailPanel(panel, det) {
      if (!det.length) {
        panel.innerHTML = `<div style="text-align:center;padding:20px;color:var(--text-muted);">Tidak ada detail SKU</div>`;
        return;
      }
      panel.innerHTML = `<table style="width:100%;border-collapse:collapse;font-size:13px;">
        <thead>
          <tr style="background:var(--bg-panel);">
            <th style="padding:8px 12px;font-size:11px;font-weight:800;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.8px;">#</th>
            <th style="padding:8px 12px;font-size:11px;font-weight:800;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.8px;">SKU</th>
            <th style="padding:8px 12px;font-size:11px;font-weight:800;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.8px;">Batch</th>
            <th style="padding:8px 12px;font-size:11px;font-weight:800;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.8px;text-align:center;">Qty</th>
            <th style="padding:8px 12px;font-size:11px;font-weight:800;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.8px;">Exp Date</th>
            <th style="padding:8px 12px;font-size:11px;font-weight:800;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.8px;">Kategori</th>
            <th style="padding:8px 12px;font-size:11px;font-weight:800;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.8px;">Keterangan</th>
          </tr>
        </thead>
        <tbody>${det.map((item, i) => {
        const expW = item.expDate && (new Date(item.expDate) - new Date()) / 86400000 <= 30 ? 'color:var(--red);font-weight:700;' : '';
        return `<tr style="border-bottom:1px solid var(--border-color);">
            <td style="padding:8px 12px;color:var(--text-muted);font-size:12px;">${i + 1}</td>
            <td style="padding:8px 12px;"><code style="background:var(--input-bg);padding:2px 8px;border-radius:6px;font-size:12px;font-weight:700;">${item.sku || '-'}</code></td>
            <td style="padding:8px 12px;">${item.batch || '-'}</td>
            <td style="padding:8px 12px;text-align:center;"><span style="background:var(--bg-panel-light);border:1px solid var(--border-color);padding:2px 10px;border-radius:8px;font-weight:800;">${item.qty || '-'}</span></td>
            <td style="padding:8px 12px;${expW}">${item.expDate ? formatDate(item.expDate) : '-'}</td>
            <td style="padding:8px 12px;">${getRDKategoriBadge(item.kategoriReturn)}</td>
            <td style="padding:8px 12px;color:var(--text-muted);font-size:12px;">${item.keterangan || '-'}</td>
          </tr>`;
      }).join('')}</tbody>
      </table>`;
    }

    // ---- Table View (flat header) ----
    function renderRDTableView(filtered) {
      const tb = document.getElementById('tableReturnDistributor');
      if (!filtered.length) {
        tb.innerHTML = '<tr><td colspan="8" style="text-align:center;color:var(--text-muted);padding:30px;">Tidak ada data</td></tr>';
        return;
      }
      const sorted = [...filtered].sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
      tb.innerHTML = sorted.map(h => {
        // Validasi dan konversi data ke string
        const namaDistributor = String(h.namaDistributor || 'Unknown');
        const noReturn = String(h.noReturn || '-');
        const tanggal = h.tanggal || '';
        const jenisReturn = String(h.jenisReturn || '');
        const createdBy = String(h.createdBy || '');
        const picSales = h.picSales ? String(h.picSales) : '';
        const noMabang = h.noMabang ? String(h.noMabang) : '';
        const noResi = h.noResi ? String(h.noResi) : '';
        const hargaOngkir = Number(h.hargaOngkir) || 0;
        const totalSKU = h.totalSKU || 0;
        const totalQty = h.totalQty || 0;
        const keterangan = h.keterangan ? String(h.keterangan) : '';

        const av = rdDistributorAvatar(namaDistributor);
        const noRet = noReturn.replace(/'/g, "\\'");
        const namaDist = namaDistributor.replace(/'/g, "\\'");
        return `<tr>
          <td>
            <div style="display:flex;align-items:center;gap:8px;">
              <div style="width:32px;height:32px;border-radius:8px;background:${av.color};display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:800;color:#fff;flex-shrink:0;">${av.initials}</div>
              <div>
                <div style="font-weight:700;font-size:13px;">${namaDistributor}</div>
                <div style="font-size:11px;color:var(--text-muted);">${createdBy}</div>
              </div>
            </div>
          </td>
          <td>
            <code style="background:var(--input-bg);padding:2px 8px;border-radius:6px;font-size:12px;color:var(--teal);font-weight:700;">${noReturn}</code>
            <div style="font-size:11px;color:var(--text-muted);margin-top:2px;">${formatDate(tanggal)}</div>
            ${picSales ? `<div style="font-size:11px;color:var(--text-muted);">🧑‍💼 ${picSales}</div>` : ''}
            ${noMabang ? `<div style="font-size:11px;color:var(--text-muted);">📋 ${noMabang}</div>` : ''}
            ${noResi ? `<div style="font-size:11px;color:var(--text-muted);">📦 ${noResi}</div>` : ''}
            ${hargaOngkir > 0 ? `<div style="font-size:11px;color:var(--text-muted);">💰 ${formatRupiah(hargaOngkir)}</div>` : ''}
          </td>
          <td style="text-align:center;">${getRDJenisBadge(jenisReturn)}</td>
          <td style="text-align:center;font-weight:700;">${totalSKU}</td>
          <td style="text-align:center;font-weight:700;">${totalQty}</td>
          <td style="max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${keterangan}">${keterangan || '-'}</td>
          <td style="text-align:right;white-space:nowrap;">
            <button class="btn btn-sm" style="background:rgba(14,165,233,0.1);color:var(--teal);border:1px solid rgba(14,165,233,0.25);padding:4px 10px;font-size:12px;"
              onclick="openRDDetailModal('${h.id}','${noRet}','${namaDist}','${tanggal}','${jenisReturn}')">📋 Detail</button>
            <button class="btn btn-ghost btn-sm" onclick="editReturnDistributor('${h.id}')">✏️</button>
            <button class="btn btn-danger btn-sm" onclick="delReturnDistributor('${h.id}')">🗑️</button>
          </td>
        </tr>`;
      }).join('');
    }

    function getRDKategoriBadge(kat) {
      const map = {
        'Return Penarikan': '<span style="background:rgba(239,68,68,0.12);color:var(--red);border:1px solid rgba(239,68,68,0.3);padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700;">⚠️ Penarikan</span>',
        'Return Buy Back': '<span style="background:rgba(245,158,11,0.12);color:var(--accent);border:1px solid rgba(245,158,11,0.3);padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700;">💰 Buy Back</span>',
        'Return Normal': '<span style="background:rgba(16,185,129,0.12);color:var(--green);border:1px solid rgba(16,185,129,0.3);padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700;">✅ Normal</span>',
        'Return Exp Date': '<span style="background:rgba(14,165,233,0.12);color:var(--teal);border:1px solid rgba(14,165,233,0.3);padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700;">📅 Exp Date</span>',
        'Return BPOM': '<span style="background:rgba(139,92,246,0.12);color:#8b5cf6;border:1px solid rgba(139,92,246,0.3);padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700;">🏛️ BPOM</span>'
      };
      return map[kat] || `<span style="background:var(--input-bg);color:var(--text-muted);padding:3px 10px;border-radius:20px;font-size:12px;">${kat || '-'}</span>`;
    }

    function openReturnDistributorModal() {
      setVal('rdId', '');
      setVal('rdEditMode', '0');
      document.getElementById('rdModalTitle').textContent = '🔄 Tambah Return Distributor';
      const today = new Date().toISOString().split('T')[0];
      setVal('rdTanggal', today);
      setVal('rdNamaDistributor', '');
      setVal('rdJenisReturn', '');
      setVal('rdPICSales', '');
      setVal('rdNoMabang', '');
      setVal('rdHargaOngkir', '');
      setVal('rdNoResi', '');
      setVal('rdKeteranganHeader', '');
      const sel = document.getElementById('rdPICSalesSelect');
      if (sel) sel.value = '';
      document.getElementById('rdJenisReturnError').style.display = 'none';
      document.getElementById('rdMultiItemWrap').style.display = 'block';
      document.getElementById('rdEditSingleWrap').style.display = 'none';
      document.getElementById('btnSaveRD').textContent = '💾 Simpan Semua';
      rdItemRows = [];
      // Buka modal dulu, baru render (agar rdItemsBody sudah ada di DOM)
      openModal('modalReturnDistributor');
      setTimeout(() => {
        renderRDItemRows();
        addRDItemRow();
      }, 30);
    }

    function editReturnDistributor(id) {
      const h = returnDistributorData.find(x => x.id === id);
      if (!h) return;

      // Set semua field header
      setVal('rdId', h.id);
      setVal('rdEditMode', '1');
      document.getElementById('rdModalTitle').textContent = `✏️ Edit Return — ${h.noReturn || ''}`;
      setVal('rdTanggal', h.tanggal || '');
      setVal('rdNamaDistributor', h.namaDistributor || '');
      setVal('rdJenisReturn', h.jenisReturn || '');
      setVal('rdPICSales', h.picSales || '');
      setVal('rdNoMabang', h.noMabang || '');
      const hargaOngkirEl = document.getElementById('rdHargaOngkir');
      if (hargaOngkirEl) {
        hargaOngkirEl.value = h.hargaOngkir > 0 ? Number(h.hargaOngkir).toLocaleString('id-ID') : '';
      }
      setVal('rdNoResi', h.noResi || '');
      setVal('rdKeteranganHeader', h.keterangan || '');
      const sel = document.getElementById('rdPICSalesSelect');
      if (sel) sel.value = rdPICSalesList.includes(h.picSales) ? h.picSales : '';
      document.getElementById('rdJenisReturnError').style.display = 'none';
      document.getElementById('rdMultiItemWrap').style.display = 'block';
      document.getElementById('rdEditSingleWrap').style.display = 'none';
      document.getElementById('btnSaveRD').textContent = '💾 Simpan';
      rdItemRows = [];

      // Buka modal
      openModal('modalReturnDistributor');

      const applyItems = (items) => {
        if (items && items.length > 0) {
          rdItemRows = items.map(x => ({
            sku: x.sku || '',
            batch: x.batch || '',
            qty: x.qty || '',
            expDate: x.expDate || '',
            kategoriReturn: x.kategoriReturn || 'Return Normal',
            keterangan: x.keterangan || ''
          }));
          renderRDItemRows();
        } else {
          // Tidak ada detail — beri 1 baris kosong
          rdItemRows = [];
          renderRDItemRows();
          addRDItemRow();
        }
      };

      // Jika sudah ada di cache, langsung tampilkan
      if (rdDetailCache[id]) {
        applyItems(rdDetailCache[id]);
        return;
      }

      // Fetch dari server
      google.script.run
        .withSuccessHandler(res => {
          if (res && res.success) {
            rdDetailCache[id] = res.data || [];
            applyItems(res.data || []);
          } else {
            toast('Gagal memuat detail: ' + (res ? res.message : 'Error'), 'error');
            rdItemRows = [];
            renderRDItemRows();
            addRDItemRow();
          }
        })
        .withFailureHandler(() => {
          toast('Koneksi gagal saat memuat detail SKU', 'error');
          rdItemRows = [];
          renderRDItemRows();
          addRDItemRow();
        })
        .getReturnDistributorDetail(id);
    }

    function openRDDetailModal(returnId, noReturn, namaDistributor, tanggal, jenisReturn) {
      document.getElementById('rdDetailModalTitle').textContent = `📋 Detail Return — ${noReturn}`;
      const h = returnDistributorData.find(x => x.id === returnId) || {};
      document.getElementById('rdDetailModalSub').innerHTML =
        `${namaDistributor} &nbsp;·&nbsp; ${formatDate(tanggal)}` +
        (jenisReturn ? ` &nbsp;·&nbsp; ${getRDJenisBadge(jenisReturn)}` : '') +
        (h.picSales ? ` &nbsp;·&nbsp; <span style="font-size:12px;">🧑‍💼 <strong>${h.picSales}</strong></span>` : '') +
        (h.noMabang ? ` &nbsp;·&nbsp; <span style="font-size:12px;">📋 Mabang: <strong>${h.noMabang}</strong></span>` : '');
      document.getElementById('rdDetailModalBody').innerHTML = `<div style="text-align:center;padding:30px;color:var(--text-muted);">⏳ Memuat detail...</div>`;
      openModal('modalRDDetail');
      if (rdDetailCache[returnId]) {
        renderRDDetailPanel(document.getElementById('rdDetailModalBody'), rdDetailCache[returnId]);
      } else {
        google.script.run.withSuccessHandler(res => {
          if (res && res.success) {
            rdDetailCache[returnId] = res.data || [];
            renderRDDetailPanel(document.getElementById('rdDetailModalBody'), res.data || []);
          } else {
            document.getElementById('rdDetailModalBody').innerHTML = `<div style="text-align:center;padding:30px;color:var(--red);">Gagal memuat detail</div>`;
          }
        }).getReturnDistributorDetail(returnId);
      }
    }

    function getRDJenisBadge(jenis) {
      const map = {
        'Normal': { bg: 'rgba(16,185,129,0.12)', color: 'var(--green)', border: 'rgba(16,185,129,0.3)', icon: '✅' },
        'Exp Date': { bg: 'rgba(14,165,233,0.12)', color: 'var(--teal)', border: 'rgba(14,165,233,0.3)', icon: '📅' },
        'Penarikan': { bg: 'rgba(239,68,68,0.12)', color: 'var(--red)', border: 'rgba(239,68,68,0.3)', icon: '⚠️' },
        'Buy Back': { bg: 'rgba(245,158,11,0.12)', color: 'var(--accent)', border: 'rgba(245,158,11,0.3)', icon: '💰' },
        'Sample': { bg: 'rgba(249,115,22,0.12)', color: '#f97316', border: 'rgba(249,115,22,0.3)', icon: '🧪' },
        'BPOM': { bg: 'rgba(139,92,246,0.12)', color: '#8b5cf6', border: 'rgba(139,92,246,0.3)', icon: '🏛️' }
      };
      const s = map[jenis];
      if (!s) return `<span style="background:var(--input-bg);color:var(--text-muted);padding:3px 10px;border-radius:20px;font-size:12px;">${jenis || '-'}</span>`;
      return `<span style="background:${s.bg};color:${s.color};border:1px solid ${s.border};padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700;">${s.icon} ${jenis}</span>`;
    }

    function rdOnJenisReturnChange() {
      const val = v('rdJenisReturn');
      const errEl = document.getElementById('rdJenisReturnError');
      if (val) {
        errEl.style.display = 'none';
        // Highlight select sesuai jenis
        const sel = document.getElementById('rdJenisReturn');
        const colors = {
          'Normal': 'rgba(16,185,129,0.08)',
          'Exp Date': 'rgba(14,165,233,0.08)',
          'Penarikan': 'rgba(239,68,68,0.08)',
          'Buy Back': 'rgba(245,158,11,0.08)',
          'Sample': 'rgba(249,115,22,0.08)',
          'BPOM': 'rgba(139,92,246,0.08)'
        };
        sel.style.background = colors[val] || '';
      }
    }

    function submitReturnDistributor() {
      const tanggal = v('rdTanggal');
      const namaDistributor = v('rdNamaDistributor');
      const jenisReturn = v('rdJenisReturn');
      const picSales = (v('rdPICSales') || '').trim();
      const noMabang = (v('rdNoMabang') || '').trim();
      const hargaOngkir = getRupiahValue('rdHargaOngkir');
      const noResi = (v('rdNoResi') || '').trim();
      const keterangan = (v('rdKeteranganHeader') || '').trim();
      const editMode = v('rdEditMode') === '1';

      if (!tanggal || !namaDistributor) return toast('Lengkapi Tanggal dan Nama Distributor', 'error');
      if (!jenisReturn) {
        document.getElementById('rdJenisReturnError').style.display = 'block';
        document.getElementById('rdJenisReturn').focus();
        return toast('Jenis Return wajib dipilih', 'error');
      }
      document.getElementById('rdJenisReturnError').style.display = 'none';
      const btn = document.getElementById('btnSaveRD');
      btn.disabled = true;

      // Sync DOM → array
      rdItemRows.forEach((row, idx) => {
        const skuEl = document.getElementById(`rdSku_${idx}`);
        const batEl = document.getElementById(`rdBatch_${idx}`);
        const qtyEl = document.getElementById(`rdQtyRow_${idx}`);
        const expEl = document.getElementById(`rdExp_${idx}`);
        const ketEl = document.getElementById(`rdKet_${idx}`);
        const katEl = document.getElementById(`rdKat_${idx}`);
        if (skuEl) row.sku = skuEl.value;
        if (batEl) row.batch = batEl.value;
        if (qtyEl) row.qty = qtyEl.value;
        if (expEl) row.expDate = expEl.value;
        if (ketEl) row.keterangan = ketEl.value;
        if (katEl) row.kategoriReturn = katEl.value;
      });

      const validRows = rdItemRows.filter(r => r.sku && r.batch);
      if (!validRows.length) { btn.disabled = false; return toast('Minimal 1 baris SKU dan Batch harus diisi', 'error'); }

      const rec = {
        id: editMode ? v('rdId') : null,
        tanggal, namaDistributor, jenisReturn,
        picSales, noMabang, hargaOngkir, noResi, keterangan,
        createdBy: currentUser.username,
        items: validRows
      };

      btn.textContent = `⏳ Menyimpan ${validRows.length} item...`;
      google.script.run.withSuccessHandler(res => {
        btn.disabled = false; btn.textContent = editMode ? '💾 Simpan' : '💾 Simpan Semua';
        if (res.success) {
          toast(`✅ ${res.saved} SKU berhasil disimpan (${res.noReturn || ''})`);
          if (res.id) delete rdDetailCache[res.id];
          closeModal('modalReturnDistributor');
          loadReturnDistributor();
        } else toast(res.message, 'error');
      }).saveReturnDistributor(rec);
    }

    function delReturnDistributor(id) {
      if (!confirm('Hapus transaksi return ini beserta semua detail SKU-nya?')) return;
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          delete rdDetailCache[id];
          toast('Dihapus');
          loadReturnDistributor();
        } else toast(res.message, 'error');
      }).deleteReturnDistributor(id);
    }

    function checkRDSKUFlag() {
      const sku = (v('rdSKU') || '').trim();
      const batch = (v('rdBatch') || '').trim();
      const badge = document.getElementById('rdKategoriBadge');
      if (!sku || !batch) { badge.style.display = 'none'; return; }
      const detected = rdDetectKategori(sku, batch);
      if (detected === 'Return Penarikan') {
        badge.innerHTML = '<span style="background:rgba(239,68,68,0.15);color:var(--red);border:1px solid rgba(239,68,68,0.4);padding:4px 12px;border-radius:20px;font-size:12px;font-weight:700;">⚠️ SKU ini terdaftar sebagai PENARIKAN — kategori otomatis diubah</span>';
        badge.style.display = 'block';
        setVal('rdKategoriReturn', 'Return Penarikan');
      } else if (detected === 'Return Buy Back') {
        badge.innerHTML = '<span style="background:rgba(245,158,11,0.15);color:var(--accent);border:1px solid rgba(245,158,11,0.4);padding:4px 12px;border-radius:20px;font-size:12px;font-weight:700;">💰 SKU ini terdaftar sebagai BUY BACK — kategori otomatis diubah</span>';
        badge.style.display = 'block';
        setVal('rdKategoriReturn', 'Return Buy Back');
      } else {
        badge.style.display = 'none';
      }
    }

    // ============================================================
    // EXPORT RETURN DISTRIBUTOR
    // ============================================================
    let rdExportMode = 'all'; // 'all' | 'selected'

    function toggleRDExportMenu() {
      const menu = document.getElementById('rdExportMenu');
      menu.style.display = menu.style.display === 'none' ? 'block' : 'none';
      // Tutup saat klik di luar
      setTimeout(() => {
        const close = (e) => {
          if (!document.getElementById('rdExportWrap').contains(e.target)) {
            menu.style.display = 'none';
            document.removeEventListener('click', close);
          }
        };
        document.addEventListener('click', close);
      }, 10);
    }

    function openRDExportModal(mode) {
      document.getElementById('rdExportMenu').style.display = 'none';
      rdExportMode = mode;
      const infoEl = document.getElementById('rdExportModeInfo');
      const selectWrap = document.getElementById('rdExportSelectWrap');

      if (mode === 'all') {
        infoEl.innerHTML = `<div style="background:rgba(14,165,233,0.08);border:1px solid rgba(14,165,233,0.2);border-radius:8px;padding:10px 14px;font-size:13px;color:var(--teal);">
          📦 Akan mengexport <strong>${returnDistributorData.length}</strong> transaksi return beserta semua detail SKU-nya.
        </div>`;
        selectWrap.style.display = 'none';
      } else {
        infoEl.innerHTML = `<div style="background:rgba(245,158,11,0.08);border:1px solid rgba(245,158,11,0.2);border-radius:8px;padding:10px 14px;font-size:13px;color:var(--accent);">
          ☑️ Pilih transaksi return yang ingin di-export. Detail SKU semua item terpilih akan ikut ter-export.
        </div>`;
        selectWrap.style.display = 'block';
        renderRDExportList();
      }
      openModal('modalRDExport');
    }

    function renderRDExportList() {
      const tb = document.getElementById('rdExportList');
      const sorted = [...returnDistributorData].sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));
      tb.innerHTML = sorted.map(h => {
        const av = rdDistributorAvatar(h.namaDistributor);
        return `<tr style="border-bottom:1px solid var(--border-color);">
          <td style="padding:8px 12px;text-align:center;">
            <input type="checkbox" class="rdExportCheck" value="${h.id}" onchange="updateRDExportCount()">
          </td>
          <td style="padding:8px 12px;"><code style="background:var(--input-bg);padding:2px 6px;border-radius:4px;font-size:12px;color:var(--teal);font-weight:700;">${h.noReturn || '-'}</code></td>
          <td style="padding:8px 12px;font-size:13px;">${formatDate(h.tanggal)}</td>
          <td style="padding:8px 12px;">
            <div style="display:flex;align-items:center;gap:8px;">
              <div style="width:26px;height:26px;border-radius:6px;background:${av.color};display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:800;color:#fff;flex-shrink:0;">${av.initials}</div>
              <span style="font-size:13px;font-weight:600;">${h.namaDistributor}</span>
            </div>
          </td>
          <td style="padding:8px 12px;text-align:center;font-weight:700;">${h.totalSKU || 0}</td>
          <td style="padding:8px 12px;text-align:center;font-weight:700;">${h.totalQty || 0}</td>
        </tr>`;
      }).join('');
      updateRDExportCount();
    }

    function rdExportSelectAll(checked) {
      document.querySelectorAll('.rdExportCheck').forEach(cb => cb.checked = checked);
      const allCb = document.getElementById('rdExportCheckAll');
      if (allCb) allCb.checked = checked;
      updateRDExportCount();
    }

    function updateRDExportCount() {
      const checked = document.querySelectorAll('.rdExportCheck:checked').length;
      const total = document.querySelectorAll('.rdExportCheck').length;
      document.getElementById('rdExportSelectedCount').textContent = `${checked} dari ${total} transaksi dipilih`;
      const allCb = document.getElementById('rdExportCheckAll');
      if (allCb) allCb.checked = checked === total && total > 0;
    }

    function doRDExport() {
      const btn = document.getElementById('btnDoRDExport');
      const format = document.querySelector('input[name="rdExportFormat"]:checked')?.value || 'xlsx';

      // Tentukan list header yang akan di-export
      let headers;
      if (rdExportMode === 'all') {
        headers = returnDistributorData;
      } else {
        const selectedIds = new Set([...document.querySelectorAll('.rdExportCheck:checked')].map(cb => cb.value));
        headers = returnDistributorData.filter(h => selectedIds.has(h.id));
      }
      if (!headers.length) return toast('Pilih minimal 1 transaksi return', 'error');

      btn.disabled = true; btn.textContent = '⏳ Memuat detail...';

      // Ambil semua detail yang belum di-cache
      const needFetch = headers.filter(h => !rdDetailCache[h.id]);
      let fetched = 0;

      const proceed = () => {
        btn.disabled = false; btn.textContent = '📤 Export Sekarang';
        // Susun data flat: 1 baris per SKU
        const rows = [];
        headers.forEach(h => {
          const det = rdDetailCache[h.id] || [];
          const base = {
            noReturn: h.noReturn, tanggal: h.tanggal,
            namaDistributor: h.namaDistributor, jenisReturn: h.jenisReturn || '',
            picSales: h.picSales || '', noMabang: h.noMabang || '',
            hargaOngkir: h.hargaOngkir || 0, noResi: h.noResi || ''
          };
          if (det.length === 0) {
            rows.push({
              ...base, sku: '-', batch: '-', qty: '-', expDate: '-',
              kategoriReturn: '-', keterangan: h.keterangan || '', createdBy: h.createdBy
            });
          } else {
            det.forEach(item => {
              rows.push({
                ...base, sku: item.sku, batch: item.batch,
                qty: item.qty || '', expDate: item.expDate || '',
                kategoriReturn: item.kategoriReturn, keterangan: item.keterangan || '',
                createdBy: h.createdBy
              });
            });
          }
        });

        if (format === 'csv') {
          doRDExportCSV(rows);
        } else {
          doRDExportXLSX(rows, headers);
        }
        closeModal('modalRDExport');
      };

      if (needFetch.length === 0) { proceed(); return; }

      // Fetch detail yang belum ada di cache
      needFetch.forEach(h => {
        google.script.run.withSuccessHandler(res => {
          if (res && res.success) rdDetailCache[h.id] = res.data || [];
          fetched++;
          if (fetched === needFetch.length) proceed();
        }).getReturnDistributorDetail(h.id);
      });
    }

    function doRDExportCSV(rows) {
      const cols = ['No Return', 'Tanggal', 'Nama Distributor', 'Jenis Return', 'PIC Sales', 'No Mabang', 'Harga Ongkir', 'No Resi', 'SKU', 'Batch', 'Qty', 'Exp Date', 'Kategori Return', 'Keterangan', 'Dibuat Oleh'];
      let csv = cols.map(c => `"${c}"`).join(',') + '\n';
      rows.forEach(r => {
        csv += [r.noReturn, r.tanggal, r.namaDistributor, r.jenisReturn,
        r.picSales, r.noMabang, r.hargaOngkir || 0, r.noResi || '', r.sku, r.batch, r.qty,
        r.expDate, r.kategoriReturn, r.keterangan, r.createdBy]
          .map(x => `"${(x || '').toString().replace(/"/g, '""')}"`)
          .join(',') + '\n';
      });
      const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'ReturnDistributor_' + new Date().toISOString().split('T')[0] + '.csv';
      a.click();
    }

    function doRDExportXLSX(rows, headers) {
      const wb = XLSX.utils.book_new();

      // ---- Sheet 1: Detail Lengkap ----
      const detailHeader = ['No Return', 'Tanggal', 'Nama Distributor', 'Jenis Return', 'PIC Sales', 'No Mabang', 'Harga Ongkir', 'No Resi', 'SKU', 'Batch', 'Qty', 'Exp Date', 'Kategori Return', 'Keterangan', 'Dibuat Oleh'];
      const detailData = rows.map(r => [
        r.noReturn, r.tanggal, r.namaDistributor, r.jenisReturn || '',
        r.picSales || '', r.noMabang || '', r.hargaOngkir || 0, r.noResi || '', r.sku, r.batch,
        r.qty !== '' && r.qty !== '-' ? (parseFloat(r.qty) || r.qty) : r.qty,
        r.expDate, r.kategoriReturn, r.keterangan, r.createdBy
      ]);
      const wsDetail = XLSX.utils.aoa_to_sheet([detailHeader, ...detailData]);
      wsDetail['!cols'] = [14, 12, 22, 12, 14, 12, 12, 14, 14, 14, 8, 12, 18, 22, 16].map(w => ({ wch: w }));
      wsDetail['!freeze'] = { xSplit: 0, ySplit: 1 };
      XLSX.utils.book_append_sheet(wb, wsDetail, 'Detail Return');

      // ---- Sheet 2: Ringkasan per Transaksi ----
      const summaryHeader = ['No Return', 'Tanggal', 'Nama Distributor', 'Jenis Return', 'PIC Sales', 'No Mabang', 'Harga Ongkir', 'No Resi', 'Total SKU', 'Total Qty', 'Keterangan', 'Dibuat Oleh'];
      const summaryData = headers.map(h => [
        h.noReturn, h.tanggal, h.namaDistributor, h.jenisReturn || '',
        h.picSales || '', h.noMabang || '', h.hargaOngkir || 0, h.noResi || '',
        parseInt(h.totalSKU) || 0, parseFloat(h.totalQty) || 0,
        h.keterangan || '', h.createdBy || ''
      ]);
      const wsSummary = XLSX.utils.aoa_to_sheet([summaryHeader, ...summaryData]);
      wsSummary['!cols'] = [14, 12, 22, 12, 14, 12, 12, 14, 10, 10, 22, 16].map(w => ({ wch: w }));
      XLSX.utils.book_append_sheet(wb, wsSummary, 'Ringkasan');

      const fname = `ReturnDistributor_${new Date().toISOString().split('T')[0]}.xlsx`;
      XLSX.writeFile(wb, fname);
      toast(`✅ Export berhasil: ${rows.length} baris detail, ${headers.length} transaksi`, 'success');
    }

    // Legacy (tetap ada untuk backward compat)
    function exportReturnDistributor() { openRDExportModal('all'); }

    // ---- Settings SKU Bermasalah ----
    function openReturnDistributorSettings() {
      // Buka modal dulu dengan data yang sudah ada di cache
      renderRDSettingTable('penarikan');
      renderRDSettingTable('buyback');
      renderRDSettingTable('bpom');
      switchRDSettingTab('penarikan');
      openModal('modalRDSettings');

      // Refresh data dari server di background
      google.script.run
        .withSuccessHandler(res => {
          if (res && res.success) {
            rdSettingsPenarikan = res.data.penarikan || [];
            rdSettingsBuyback = res.data.buyback || [];
            rdSettingsBPOM = res.data.bpom || [];
            renderRDSettingTable('penarikan');
            renderRDSettingTable('buyback');
            renderRDSettingTable('bpom');
          }
        })
        .withFailureHandler(err => {
          console.warn('Gagal load settings SKU:', err);
        })
        .getReturnDistributorSettings();
    }

    function switchRDSettingTab(tab) {
      const isPenarikan = tab === 'penarikan';
      const isBuyback = tab === 'buyback';
      const isBPOM = tab === 'bpom';
      const isSync = tab === 'sync';
      document.getElementById('rdSettingPanelPenarikan').style.display = isPenarikan ? 'block' : 'none';
      document.getElementById('rdSettingPanelBuyback').style.display = isBuyback ? 'block' : 'none';
      document.getElementById('rdSettingPanelBPOM').style.display = isBPOM ? 'block' : 'none';
      document.getElementById('rdSettingPanelSync').style.display = isSync ? 'block' : 'none';
      const styles = {
        penarikan: 'background:rgba(239,68,68,0.12);color:var(--red);border:1px solid rgba(239,68,68,0.3);font-weight:700;',
        buyback: 'background:rgba(245,158,11,0.12);color:var(--accent);border:1px solid rgba(245,158,11,0.3);font-weight:700;',
        bpom: 'background:rgba(139,92,246,0.12);color:#8b5cf6;border:1px solid rgba(139,92,246,0.3);font-weight:700;',
        sync: 'background:rgba(14,165,233,0.12);color:var(--teal);border:1px solid rgba(14,165,233,0.3);font-weight:700;',
        off: 'background:var(--input-bg);color:var(--text-muted);border:1px solid var(--border-color);'
      };
      document.getElementById('rdSettingTabPenarikan').style.cssText = isPenarikan ? styles.penarikan : styles.off;
      document.getElementById('rdSettingTabBuyback').style.cssText = isBuyback ? styles.buyback : styles.off;
      document.getElementById('rdSettingTabBPOM').style.cssText = isBPOM ? styles.bpom : styles.off;
      document.getElementById('rdSettingTabSync').style.cssText = isSync ? styles.sync : styles.off;
      if (isSync) renderRDSyncTable();
    }

    // Hitung berapa kali SKU+Batch muncul di data return
    function rdCountInReturn(sku, batch) {
      if (!returnDistributorData || !returnDistributorData.length) return 0;
      // Pastikan sku dan batch adalah string
      const s = String(sku || '').toLowerCase(), b = String(batch || '').toLowerCase();
      // Hitung dari detail cache semua transaksi
      let count = 0;
      returnDistributorData.forEach(h => {
        const det = rdDetailCache[h.id] || [];
        det.forEach(x => {
          // Pastikan x.sku dan x.batch adalah string
          if (String(x.sku || '').toLowerCase() === s && String(x.batch || '').toLowerCase() === b) count++;
        });
      });
      return count;
    }

    function renderRDSettingTable(type) {
      const list = type === 'penarikan' ? rdSettingsPenarikan : type === 'buyback' ? rdSettingsBuyback : rdSettingsBPOM;
      const tbId = type === 'penarikan' ? 'rdSettingTablePenarikan' : type === 'buyback' ? 'rdSettingTableBuyback' : 'rdSettingTableBPOM';
      const tb = document.getElementById(tbId);
      if (!tb) return;
      if (!list.length) {
        tb.innerHTML = `<tr><td colspan="7" style="text-align:center;color:var(--text-muted);padding:20px;">Belum ada data</td></tr>`;
        return;
      }
      tb.innerHTML = list.map((item, idx) => {
        const count = rdCountInReturn(item.sku, item.batch);
        const isAll = (item.batch || '').toUpperCase() === 'ALL';
        const batchDisplay = isAll
          ? `<span style="background:rgba(239,68,68,0.15);color:var(--red);border:1px solid rgba(239,68,68,0.4);padding:2px 8px;border-radius:8px;font-size:11px;font-weight:800;">🔴 ALL BATCH</span>`
          : `<code style="background:var(--input-bg);padding:2px 6px;border-radius:4px;font-size:12px;">${item.batch}</code>`;
        const statusBadge = count > 0
          ? `<span style="background:rgba(16,185,129,0.12);color:var(--green);border:1px solid rgba(16,185,129,0.3);padding:2px 8px;border-radius:12px;font-size:11px;font-weight:700;">✅ Ada Return</span>`
          : `<span style="background:rgba(148,163,184,0.1);color:var(--text-muted);border:1px solid rgba(148,163,184,0.2);padding:2px 8px;border-radius:12px;font-size:11px;">— Belum Ada</span>`;
        return `<tr style="${isAll ? 'background:rgba(239,68,68,0.03);' : ''}">
          <td style="color:var(--text-muted);font-size:12px;">${idx + 1}</td>
          <td><code style="background:var(--input-bg);padding:2px 6px;border-radius:4px;font-size:12px;">${item.sku}</code></td>
          <td>${batchDisplay}</td>
          <td style="font-size:12px;color:var(--text-muted);">${item.manufaktur || '-'}</td>
          <td style="text-align:center;">${statusBadge}</td>
          <td style="text-align:center;font-weight:700;color:${count > 0 ? 'var(--green)' : 'var(--text-muted)'};">${count > 0 ? count : '-'}</td>
          <td style="text-align:right;">
            <button class="btn btn-danger btn-sm" style="padding:2px 8px;" onclick="removeRDSettingItem('${type}', ${idx})">🗑️</button>
          </td>
        </tr>`;
      }).join('');
    }

    function addRDSettingItem(type) {
      const skuId = type === 'penarikan' ? 'rdNewPenarikanSKU' : type === 'buyback' ? 'rdNewBuybackSKU' : 'rdNewBPOMSKU';
      const batchId = type === 'penarikan' ? 'rdNewPenarikanBatch' : type === 'buyback' ? 'rdNewBuybackBatch' : 'rdNewBPOMBatch';
      const mfgId = type === 'penarikan' ? 'rdNewPenarikanMfg' : type === 'buyback' ? 'rdNewBuybackMfg' : 'rdNewBPOMMfg';
      const sku = (v(skuId) || '').trim();
      const batchRaw = (v(batchId) || '').trim();
      const manufaktur = (v(mfgId) || '').trim();
      const batch = batchRaw.toUpperCase() === 'ALL' ? 'ALL' : batchRaw;
      if (!sku || !batch) return toast('Isi SKU dan Batch (atau ketik ALL untuk semua batch)', 'error');
      const list = type === 'penarikan' ? rdSettingsPenarikan : type === 'buyback' ? rdSettingsBuyback : rdSettingsBPOM;
      // Pastikan x.sku dan x.batch adalah string sebelum toLowerCase
      if (list.some(x => String(x.sku || '').toLowerCase() === sku.toLowerCase() && String(x.batch || '').toLowerCase() === batch.toLowerCase()))
        return toast('SKU + Batch sudah ada di daftar', 'error');
      list.push({ sku, batch, manufaktur });
      renderRDSettingTable(type);
      setVal(skuId, ''); setVal(batchId, ''); setVal(mfgId, '');
      toast(`✅ Ditambahkan${batch === 'ALL' ? ' (ALL Batch)' : ''}`, 'success');
    }

    function removeRDSettingItem(type, idx) {
      if (type === 'penarikan') rdSettingsPenarikan.splice(idx, 1);
      else if (type === 'buyback') rdSettingsBuyback.splice(idx, 1);
      else rdSettingsBPOM.splice(idx, 1);
      renderRDSettingTable(type);
    }

    // ---- Import Excel untuk Setting ----
    function handleImportRDSetting(input, type) {
      const file = input.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = function (e) {
        try {
          const wb = XLSX.read(e.target.result, { type: 'binary' });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          if (rows.length < 2) return toast('File kosong', 'error');
          const headerRow = rows[0].map(h => String(h || '').toLowerCase().trim());
          const skuCol = headerRow.findIndex(h => h.includes('sku'));
          const batchCol = headerRow.findIndex(h => h.includes('batch'));
          const mfgCol = headerRow.findIndex(h => h.includes('manufaktur') || h.includes('manufacturer') || h.includes('mfg'));
          if (skuCol < 0 || batchCol < 0) return toast('Kolom SKU atau Batch tidak ditemukan di file', 'error');
          const list = type === 'penarikan' ? rdSettingsPenarikan
            : type === 'buyback' ? rdSettingsBuyback
              : rdSettingsBPOM;
          let added = 0, skipped = 0;
          for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.every(c => c === '')) continue;
            const sku = String(row[skuCol] || '').trim();
            const batchRaw = String(row[batchCol] || '').trim();
            const manufaktur = mfgCol >= 0 ? String(row[mfgCol] || '').trim() : '';
            const batch = batchRaw.toUpperCase() === 'ALL' ? 'ALL' : batchRaw;
            if (!sku || !batch) { skipped++; continue; }
            if (list.some(x => x.sku.toLowerCase() === sku.toLowerCase() && x.batch.toLowerCase() === batch.toLowerCase())) { skipped++; continue; }
            list.push({ sku, batch, manufaktur });
            added++;
          }
          renderRDSettingTable(type);
          toast(`✅ ${added} SKU ditambahkan${skipped ? `, ${skipped} dilewati` : ''}`, 'success');
        } catch (err) { toast('Gagal membaca file: ' + err.message, 'error'); }
        input.value = '';
      };
      reader.readAsBinaryString(file);
    }

    function downloadRDSettingTemplate(type) {
      const labels = { penarikan: 'Penarikan', buyback: 'BuyBack', bpom: 'BPOM' };
      const label = labels[type] || type;
      const headers = ['SKU', 'Batch', 'Manufaktur'];
      const sample = [
        ['SKU-001', 'BATCH-A1', 'PT Manufaktur ABC'],
        ['SKU-002', 'ALL', 'PT Manufaktur XYZ'],  // ALL = semua batch
        ['SKU-003', 'BATCH-C3', ''],
      ];
      const note = [[''], ['CATATAN: Isi Batch dengan "ALL" untuk menandai semua batch dari SKU tersebut.']];
      const ws = XLSX.utils.aoa_to_sheet([headers, ...sample, ...note]);
      ws['!cols'] = [{ wch: 16 }, { wch: 14 }, { wch: 24 }];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Template ' + label);
      XLSX.writeFile(wb, `Template_SKU_${label}.xlsx`);
    }

    // ---- Panel Sinkronisasi ----
    let rdSyncAllRows = [];
    function renderRDSyncTable() {
      const tb = document.getElementById('rdSyncTable');
      tb.innerHTML = `<tr><td colspan="6" style="text-align:center;color:var(--text-muted);padding:20px;">⏳ Memuat data...</td></tr>`;
      if (!returnDistributorData || !returnDistributorData.length) {
        tb.innerHTML = `<tr><td colspan="6" style="text-align:center;color:var(--text-muted);padding:20px;">Tidak ada data Return Distributor</td></tr>`;
        return;
      }
      // Buat map unik SKU+Batch dari data return
      const map = {};
      returnDistributorData.forEach(d => {
        // Pastikan sku dan batch adalah string
        const key = String(d.sku || '').toLowerCase() + '||' + String(d.batch || '').toLowerCase();
        if (!map[key]) map[key] = { sku: d.sku, batch: d.batch, count: 0, kategori: d.kategoriReturn };
        map[key].count++;
        // Prioritaskan kategori bermasalah
        if (d.kategoriReturn === 'Return Penarikan' || d.kategoriReturn === 'Return Buy Back')
          map[key].kategori = d.kategoriReturn;
      });
      rdSyncAllRows = Object.values(map);
      filterRDSyncTable();
    }

    function filterRDSyncTable() {
      const q = (v('rdSyncSearch') || '').toLowerCase().trim();
      const tb = document.getElementById('rdSyncTable');
      const filtered = rdSyncAllRows.filter(r =>
        !q || (r.sku || '').toLowerCase().includes(q) || (r.batch || '').toLowerCase().includes(q)
      );
      if (!filtered.length) {
        tb.innerHTML = `<tr><td colspan="6" style="text-align:center;color:var(--text-muted);padding:20px;">Tidak ada data</td></tr>`;
        return;
      }
      tb.innerHTML = filtered.map(r => {
        const s = (r.sku || '').toLowerCase(), b = (r.batch || '').toLowerCase();
        const inP = rdSettingsPenarikan.some(x => x.sku.toLowerCase() === s && x.batch.toLowerCase() === b);
        const inB = rdSettingsBuyback.some(x => x.sku.toLowerCase() === s && x.batch.toLowerCase() === b);
        const inBPOM = rdSettingsBPOM.some(x => x.sku.toLowerCase() === s && x.batch.toLowerCase() === b);
        let statusBadge, aksi;
        if (inP) {
          statusBadge = `<span style="background:rgba(239,68,68,0.12);color:var(--red);border:1px solid rgba(239,68,68,0.3);padding:2px 8px;border-radius:12px;font-size:11px;font-weight:700;">⚠️ Penarikan</span>`;
          aksi = `<span style="color:var(--text-muted);font-size:12px;">Sudah terdaftar</span>`;
        } else if (inB) {
          statusBadge = `<span style="background:rgba(245,158,11,0.12);color:var(--accent);border:1px solid rgba(245,158,11,0.3);padding:2px 8px;border-radius:12px;font-size:11px;font-weight:700;">💰 Buy Back</span>`;
          aksi = `<span style="color:var(--text-muted);font-size:12px;">Sudah terdaftar</span>`;
        } else if (inBPOM) {
          statusBadge = `<span style="background:rgba(139,92,246,0.12);color:#8b5cf6;border:1px solid rgba(139,92,246,0.3);padding:2px 8px;border-radius:12px;font-size:11px;font-weight:700;">🏛️ BPOM</span>`;
          aksi = `<span style="color:var(--text-muted);font-size:12px;">Sudah terdaftar</span>`;
        } else {
          statusBadge = `<span style="background:rgba(148,163,184,0.1);color:var(--text-muted);border:1px solid rgba(148,163,184,0.2);padding:2px 8px;border-radius:12px;font-size:11px;">— Belum Terdaftar</span>`;
          const skuE = r.sku.replace(/'/g, "\\'"), batE = r.batch.replace(/'/g, "\\'");
          aksi = `<div style="display:flex;gap:4px;justify-content:flex-end;flex-wrap:wrap;">
            <button class="btn btn-sm" style="padding:2px 7px;background:rgba(239,68,68,0.1);color:var(--red);border:1px solid rgba(239,68,68,0.3);font-size:11px;"
              onclick="rdSyncAddToList('${skuE}','${batE}','penarikan')">⚠️ Penarikan</button>
            <button class="btn btn-sm" style="padding:2px 7px;background:rgba(245,158,11,0.1);color:var(--accent);border:1px solid rgba(245,158,11,0.3);font-size:11px;"
              onclick="rdSyncAddToList('${skuE}','${batE}','buyback')">💰 Buy Back</button>
            <button class="btn btn-sm" style="padding:2px 7px;background:rgba(139,92,246,0.1);color:#8b5cf6;border:1px solid rgba(139,92,246,0.3);font-size:11px;"
              onclick="rdSyncAddToList('${skuE}','${batE}','bpom')">🏛️ BPOM</button>
          </div>`;
        }
        const katBadge = getRDKategoriBadge(r.kategori);
        return `<tr>
          <td><code style="background:var(--input-bg);padding:2px 6px;border-radius:4px;font-size:12px;">${r.sku}</code></td>
          <td style="font-size:13px;">${r.batch}</td>
          <td style="text-align:center;font-weight:700;color:var(--teal);">${r.count}</td>
          <td style="text-align:center;">${katBadge}</td>
          <td style="text-align:center;">${statusBadge}</td>
          <td style="text-align:right;">${aksi}</td>
        </tr>`;
      }).join('');
    }

    function rdSyncAddToList(sku, batch, type) {
      const list = type === 'penarikan' ? rdSettingsPenarikan
        : type === 'buyback' ? rdSettingsBuyback
          : rdSettingsBPOM;
      const s = sku.toLowerCase(), b = batch.toLowerCase();
      if (list.some(x => x.sku.toLowerCase() === s && x.batch.toLowerCase() === b))
        return toast('Sudah ada di daftar', 'info');
      list.push({ sku, batch, manufaktur: '' });
      renderRDSettingTable(type);
      filterRDSyncTable();
      const label = type === 'penarikan' ? 'Penarikan' : type === 'buyback' ? 'Buy Back' : 'BPOM';
      toast(`✅ ${sku} / ${batch} ditambahkan ke daftar ${label}`, 'success');
    }

    function saveRDSettings() {
      const btn = document.getElementById('btnSaveRDSettings');
      btn.disabled = true; btn.textContent = '⏳ Menyimpan...';
      google.script.run.withSuccessHandler(res => {
        btn.disabled = false; btn.textContent = '💾 Simpan Pengaturan';
        if (res.success) {
          toast(res.message || 'Pengaturan berhasil disimpan');
          closeModal('modalRDSettings');
        } else toast(res.message, 'error');
      }).saveReturnDistributorSettings(rdSettingsPenarikan, rdSettingsBuyback, rdSettingsBPOM);
    }
    // ===== END RETURN DISTRIBUTOR =====

    // ORDER
    function loadOrder() {
      const skuQuery = (v('searchOrderSKU') || '').toLowerCase().trim();
      const callGet = skuQuery ? 'getOrdersBySku' : 'getOrders';
      google.script.run.withSuccessHandler(res => {
        if (!res || !res.success) {
          toast(res ? res.message : 'Gagal memuat data order', 'error');
          return;
        }

        let hMP = '', hDist = '', hStore = '', hHis = '';
        let cMP = 0, cDist = 0, cStore = 0, cHis = 0;

        const q = (v('searchOrder') || '').toLowerCase().trim();
        const filtered = res.data.filter(d =>
          (d.noOrder || '').toLowerCase().includes(q) ||
          (d.pelanggan || '').toLowerCase().includes(q) ||
          (d.noResi || '').toLowerCase().includes(q)
        );

        filtered.forEach(d => {
          try {
            const isSent = d.status === 'Terkirim';
            const kat = (d.kategori || 'Distributor').toLowerCase().trim();
            const hasBukti = d.buktiPacking && d.buktiPacking !== '';
            const statusCls = isSent ? 'order-terkirim' : 'order-pending';
            const tgl = formatDate(d.tanggal);

            const uploadBtn = `<button class="btn btn-accent btn-sm" onclick="openUploadPackingModal('${d.id}','${d.noOrder}')">${hasBukti ? '📸' : '📤'}</button>`;
            const buyBtn = hasBukti ? `<a href="${d.buktiPacking}" target="_blank" class="btn btn-ghost btn-sm" title="Foto">🖼️</a>` : '';
            const sendBtn = !isSent ? `<button class="btn btn-teal btn-sm" onclick="kirimOrder('${d.id}', '${d.noOrder}')">Kirim</button>` : '';
            const validateBtn = !isSent ? `<button class="btn btn-primary btn-sm" onclick="openValidationModal('${d.id}', '${d.noOrder}')" title="Validasi Fisik Pesanan">✅</button>` : '';
            const printBtn = `<button class="btn btn-ghost btn-sm" onclick="printOrder('${d.id}','${d.noOrder}')" title="Print / Cetak Order" style="color:var(--accent);">🖨️</button>`;

            const rowId = `order-main-${d.id}`;
            const detailRowId = `order-detail-${d.id}`;

            const detailHeaderHtml = `
                <div style="display:flex; justify-content:space-between; margin-bottom:12px; border-bottom:1px solid var(--border-light); padding-bottom:8px;">
                  <div style="font-weight:700; font-size:11px; color:var(--teal); text-transform:uppercase;">📦 Item Order: ${d.noOrder}</div>
                  <div style="font-size:11px; color:var(--text-muted);">
                    📍 <strong>Alamat:</strong> ${d.alamat || '-'} | 🏷️ <strong>Pelanggan:</strong> ${d.pelanggan || '-'}
                  </div>
                </div>`;

            if (isSent) {
              cHis++;
              hHis += `
                <tr id="${rowId}">
                  <td>
                    <button class="btn-toggle-row" onclick="toggleOrderItems('${d.id}', '${d.noOrder}')">
                      <i class="bi bi-plus-circle"></i>
                    </button>
                    <strong>${d.noOrder}</strong>
                  </td>
                  <td>${formatDate(d.sentAt || d.tanggal)}</td>
                  <td>${d.pelanggan}</td>
                  <td><span class="badge-tb">${d.kategori || '-'}</span></td>
                  <td><strong>${d.totalItem}</strong></td>
                  <td>${buyBtn || '-'}</td>
                  <td><div style="display:flex; gap:4px; align-items:center;">${printBtn}</div></td>
                </tr>
                <tr id="${detailRowId}" class="row-detail" style="display:none;">
                  <td colspan="7"><div class="detail-container">${detailHeaderHtml}<div class="loading-inline">Memuat item...</div></div></td>
                </tr>`;
            } else {
              const row = `
                <tr id="${rowId}">
                  <td>
                    <button class="btn-toggle-row" onclick="toggleOrderItems('${d.id}', '${d.noOrder}')">
                      <i class="bi bi-plus-circle"></i>
                    </button>
                    <strong>${d.noOrder}</strong>
                  </td>
                  <td>${tgl}</td>
                  <td><strong>${d.pelanggan}</strong></td>
                  <td>${kat === 'marketplace' ? (d.noResi || '-') : (d.alamat || '-')}</td>
                  <td><strong>${d.totalItem || 0}</strong></td>
                  <td><span class="${statusCls}">${d.status}</span></td>
                  <td><div style="display:flex; gap:4px; flex-wrap:wrap; align-items:center;">${validateBtn} ${sendBtn} ${uploadBtn} ${buyBtn} ${printBtn}</div></td>
                </tr>
                <tr id="${detailRowId}" class="row-detail" style="display:none;">
                  <td colspan="7"><div class="detail-container">${detailHeaderHtml}<div class="loading-inline">Memuat item...</div></div></td>
                </tr>`;

              if (kat === 'marketplace') { cMP++; hMP += row; }
              else if (kat === 'store') { cStore++; hStore += row; }
              else { cDist++; hDist += row; }
            }
          } catch (err) {
            console.error('Error rendering row:', err, d);
          }
        });

        const empty = '<tr><td colspan="7" class="empty-state">Belum ada orderan di kategori ini</td></tr>';
        document.getElementById('tableOrderMP').innerHTML = hMP || empty;
        document.getElementById('tableOrderDist').innerHTML = hDist || empty;
        document.getElementById('tableOrderStore').innerHTML = hStore || empty;
        document.getElementById('tableOrderHistory').innerHTML = hHis || empty;

        document.getElementById('countOrderMP').textContent = cMP;
        document.getElementById('countOrderDist').textContent = cDist;
        document.getElementById('countOrderStore').textContent = cStore;
        document.getElementById('countOrderHistory').textContent = cHis;
      }).withFailureHandler(err => {
        toast('Gagal memuat Order: ' + err, 'error');
        console.error('Order Load Error:', err);
      })[callGet](skuQuery || undefined);
    }

    // ============================================================
    // ORDER VALIDATION (SCANNING FISIK)
    // ============================================================
    let currentValidationOrderId = null;
    let currentValidationItems = [];

    function openValidationModal(orderId, noOrder) {
      currentValidationOrderId = orderId;
      document.getElementById('valOrderNo').textContent = noOrder;
      document.getElementById('tblValidationItems').innerHTML = '<tr><td colspan="5" class="text-center">Memuat item...</td></tr>';
      document.getElementById('valProgressBadge').textContent = 'Progress: 0 / 0';
      openModal('modalValidationOrder');

      google.script.run.withSuccessHandler(res => {
        if (!res.success) return toast(res.message, 'error');
        currentValidationItems = res.data;
        renderValidationItems();

        // Auto focus scan input
        setTimeout(() => document.getElementById('valScanInput').focus(), 500);
      }).getOrderDetail(orderId, noOrder);
    }

    function renderValidationItems() {
      const tb = document.getElementById('tblValidationItems');
      tb.innerHTML = '';
      let totalOrdered = 0;
      let totalScanned = 0;

      currentValidationItems.forEach((item, idx) => {
        totalOrdered += item.qty;
        totalScanned += item.packedQty || 0;

        const isDone = (item.packedQty || 0) >= item.qty;
        const rowClass = isDone ? 'table-success-dim' : '';
        const statusIcon = isDone ? '✅' : '⏳';

        tb.innerHTML += `
          <tr class="${rowClass}" id="val-row-${idx}">
            <td>
              <strong>${item.sku}</strong><br>
              <small class="text-muted">${item.nama}</small>
            </td>
            <td style="text-align:center"><span class="badge-tb">${item.lokasi || '-'}</span></td>
            <td style="text-align:center; font-weight:700;">${item.qty}</td>
            <td style="text-align:center;">
              <span id="val-qty-scanned-${idx}" style="font-size:16px; font-weight:800; color:${isDone ? 'var(--teal)' : 'var(--accent)'}">
                ${item.packedQty || 0}
              </span>
            </td>
            <td style="text-align:center;">${statusIcon}</td>
          </tr>`;
      });

      document.getElementById('valProgressBadge').textContent = `Progress: ${totalScanned} / ${totalOrdered}`;
      if (totalScanned >= totalOrdered && totalOrdered > 0) {
        document.getElementById('valProgressBadge').className = 'badge bg-success';
        toast('🎉 Semua item sudah lengkap divalidasi!', 'success');
      } else {
        document.getElementById('valProgressBadge').className = 'badge bg-teal';
      }
    }

    function handleValidationScan(inputEl) {
      const code = inputEl.value.trim().toLowerCase();
      if (!code) return;

      // 1. Cari SKU dari code (jika code adalah barcode)
      let targetSku = code;
      const s = stockData.find(x => String(x.barcode).toLowerCase() === code || String(x.sku).toLowerCase() === code);
      if (s) targetSku = s.sku.toLowerCase();

      // 2. Cari item di list validation yang SKU nya cocok
      let foundIndex = -1;
      for (let i = 0; i < currentValidationItems.length; i++) {
        if (currentValidationItems[i].sku.toLowerCase() === targetSku) {
          // Cari yang belum penuh dulu
          if ((currentValidationItems[i].packedQty || 0) < currentValidationItems[i].qty) {
            foundIndex = i;
            break;
          }
        }
      }

      // Jika tidak ada yang "belum penuh", cari lagi (mungkin user scan lebih)
      if (foundIndex === -1) {
        foundIndex = currentValidationItems.findIndex(it => it.sku.toLowerCase() === targetSku);
      }

      const scanQty = parseFloat(document.getElementById('valQtyInput').value) || 1;

      if (foundIndex !== -1) {
        const item = currentValidationItems[foundIndex];
        if ((item.packedQty || 0) >= item.qty) {
          toast(`⚠️ ${item.sku} sudah lengkap!`, 'warning');
        } else {
          const needed = item.qty - (item.packedQty || 0);
          const add = Math.min(scanQty, needed);
          item.packedQty = (item.packedQty || 0) + add;

          if (scanQty > needed) {
            toast(`⚠️ ${item.sku} hanya butuh ${needed} lagi (discan ${scanQty})`, 'warning');
          }

          renderValidationItems();
          toast(`✅ Scan: ${item.nama} (${add})`, 'success');
          document.getElementById('valQtyInput').value = 1;

          // AUTO SUBMIT IF COMPLETE (Syarat Terpenuhi: Semua Item & Foto)
          checkAutoSubmitValidation();
        }
      } else {
        toast(`❌ Barang "${code}" tidak ada dalam order ini!`, 'error');
      }

      inputEl.value = '';
      inputEl.focus();
    }

    function submitValidation() {
      const btn = document.getElementById('btnSubmitValidation');
      btn.disabled = true;
      btn.textContent = '⏳ Menyimpan & Mengirim...';
      const noOrder = document.getElementById('valOrderNo').textContent;
      const photoUrl = document.getElementById('valPhotoUrl').value;

      const validationData = currentValidationItems.map(it => ({
        id: it.id,
        packedQty: it.packedQty || 0
      }));

      // SINGLE CALL OPTIMIZATION: Menggabungkan semua proses server (Update Qty, Status, Foto, Stok) menjadi 1 call.
      google.script.run.withSuccessHandler(res => {
        btn.disabled = false;
        btn.textContent = '💾 Simpan Hasil Validasi';

        if (res.success) {
          toast(`✅ Order ${noOrder} Berhasil Divalidasi & Terkirim!`, 'success');
          closeModal('modalValidationOrder');
          loadOrder();
          loadStock();
          resetValPhoto();
        } else {
          toast(`❌ Gagal: ${res.message}`, 'error');
        }
      }).withFailureHandler(err => {
        btn.disabled = false;
        btn.textContent = '💾 Simpan Hasil Validasi';
        toast(`⚠️ Terjadi kesalahan: ${err.message}`, 'error');
      }).finalizeOrderValidation(currentValidationOrderId, noOrder, JSON.stringify(validationData), photoUrl);
    }

    // ============================================================
    // CAMERA & PHOTO FOR VALIDATION
    // ============================================================
    let valCameraScanner = null;

    function toggleValCamera(show) {
      const section = document.getElementById('valCameraSection');
      if (show) {
        section.style.display = 'block';
        if (!valCameraScanner) {
          valCameraScanner = new Html5Qrcode("valCameraReader");
        }
        valCameraScanner.start({ facingMode: "environment" }, { fps: 15, qrbox: 250 }, () => { }, () => { });
        document.getElementById('btnCaptureValPhoto').disabled = false;
        document.getElementById('btnCaptureValPhoto').textContent = '📸 AMBIL & UPLOAD';
      } else {
        section.style.display = 'none';
        if (valCameraScanner) {
          valCameraScanner.stop().catch(e => console.warn(e));
        }
      }
    }

    function captureValPhoto() {
      if (!valCameraScanner) return;
      const btn = document.getElementById('btnCaptureValPhoto');
      btn.disabled = true;
      btn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> MENGAMBIL...';

      const video = document.querySelector('#valCameraReader video');
      if (!video) return;

      const canvas = document.createElement('canvas');

      // OPTIMASI: Batasi resolusi agar upload cepat (Max 1024px)
      const maxDim = 1024;
      let w = video.videoWidth;
      let h = video.videoHeight;
      if (w > maxDim || h > maxDim) {
        if (w > h) { h = (maxDim / w) * h; w = maxDim; }
        else { w = (maxDim / h) * w; h = maxDim; }
      }

      canvas.width = w;
      canvas.height = h;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(video, 0, 0, w, h);

      // OPTIMASI: Kurangi kualitas (0.5 cukup untuk bukti packing)
      const base64 = canvas.toDataURL('image/jpeg', 0.5);

      uploadValPhoto(base64);
    }

    function handleValFileSelect(input) {
      const file = input.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (e) => uploadValPhoto(e.target.result);
      reader.readAsDataURL(file);
    }

    function uploadValPhoto(base64) {
      const b64Data = base64.split(',')[1];
      const btn = document.getElementById('btnCaptureValPhoto');
      if (btn) {
        btn.disabled = true;
        btn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> MENGUNGGAH...';
      }
      toast('⏳ Mengunggah foto packing...', 'info');

      let uId = '';
      // OPTIMASI: Tingkatkan chunk size (200KB) agar request lebih sedikit & cepat
      const chunkSize = 200000;
      const chunks = [];
      for (let i = 0; i < b64Data.length; i += chunkSize) chunks.push(b64Data.substring(i, i + chunkSize));

      let cIdx = 0;
      const sendChunk = () => {
        if (btn) btn.innerHTML = `<span class="spinner-border spinner-border-sm"></span> ${Math.round((cIdx / chunks.length) * 100)}%`;
        if (cIdx < chunks.length) {
          google.script.run.withSuccessHandler(res => {
            if (res.success) { uId = res.uploadId; cIdx++; sendChunk(); }
            else {
              toast(res.message, 'error');
              if (btn) { btn.disabled = false; btn.textContent = '📸 AMBIL & UPLOAD'; }
            }
          }).uploadChunk(chunks[cIdx], cIdx, uId);
        } else {
          if (btn) btn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> FINISHING...';
          google.script.run.withSuccessHandler(res => {
            if (res.success) {
              document.getElementById('valPhotoUrl').value = res.url;
              document.getElementById('valPhotoPreviewImg').src = res.url;
              document.getElementById('valPhotoPreviewSection').style.display = 'block';
              toggleValCamera(false);
              toast('✅ Foto packing berhasil diunggah!', 'success');

              // AUTO SUBMIT IF COMPLETE & PHOTO READY
              checkAutoSubmitValidation();
            } else {
              toast(res.message, 'error');
              if (btn) { btn.disabled = false; btn.textContent = '📸 AMBIL & UPLOAD'; }
            }
          }).finalizeChunkedUpload(uId, 'Packing_' + document.getElementById('valOrderNo').textContent + '.jpg', 'image/jpeg', 'Bukti Packing');
        }
      };
      sendChunk();
    }

    function resetValPhoto() {
      document.getElementById('valPhotoUrl').value = '';
      document.getElementById('valPhotoPreviewSection').style.display = 'none';
      if (valCameraScanner) {
        valCameraScanner.stop().catch(e => { });
      }
    }

    function checkAutoSubmitValidation() {
      const allScanned = currentValidationItems.length > 0 && currentValidationItems.every(it => (it.packedQty || 0) >= it.qty);
      const photoReady = document.getElementById('valPhotoUrl').value !== '';

      if (allScanned && photoReady) {
        setTimeout(() => {
          toast('🚀 Semua syarat terpenuhi, memproses pengiriman otomatis...', 'success');
          submitValidation();
        }, 800);
      } else if (allScanned && !photoReady) {
        toast('📝 Barang lengkap! Silakan ambil foto bukti packing untuk menyelesaikan.', 'info');
        toggleValCamera(true);
      }
    }

    // ============================================================
    // PRINT ORDER - Cetak Packing Slip Detail
    // ============================================================
    function printOrder(orderId, noOrder) {
      toast('⏳ Memproses cetak order...', 'info');
      google.script.run
        .withSuccessHandler(function (res) {
          if (!res || !res.success) {
            return toast('Gagal mengambil data: ' + (res ? res.message : 'Server error'), 'error');
          }
          _doPrintOrder(res.header, res.items);
        })
        .withFailureHandler(function (err) {
          toast('Error server: ' + err, 'error');
        })
        .getOrderDetailFull(orderId, noOrder);
    }

    function _doPrintOrder(header, items) {
      const printDate = new Date().toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
      const printTime = new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });

      const statusClass = header.status === 'Terkirim' ? 'status-sent' : 'status-pending';
      const statusLabel = header.status === 'Terkirim' ? '✅ TERKIRIM' : '⏳ PENDING';
      const kategoriIco = (header.kategori || '').toLowerCase() === 'marketplace' ? '🛒' : (header.kategori || '').toLowerCase() === 'store' ? '🏬' : '🏢';

      let totalQty = 0;
      const itemRows = (items || []).map((item, idx) => {
        totalQty += parseFloat(item.qty) || 0;
        const expDate = item.expDate && item.expDate !== '-' ? item.expDate : '-';
        const batch = item.batch && item.batch !== '-' ? item.batch : '-';
        const lokasi = item.lokasi && item.lokasi !== '-' ? item.lokasi : '-';
        return `
          <tr>
            <td style="text-align:center;font-weight:700;color:#666;">${idx + 1}</td>
            <td><strong style="color:#1a1a2e;font-size:13px;">${item.sku || '-'}</strong></td>
            <td style="max-width:200px;">${item.nama || '-'}</td>
            <td style="text-align:center;"><span class="lokasi-badge">${lokasi}</span></td>
            <td style="text-align:center;font-family:monospace;font-size:11px;color:#555;">${batch}</td>
            <td style="text-align:center;font-size:11px;color:${expDate !== '-' ? '#c0392b' : '#555'};">${expDate}</td>
            <td style="text-align:center;"><strong style="font-size:16px;color:#e67e22;">${item.qty || 0}</strong></td>
            <td style="text-align:center;color:#666;font-size:11px;">${item.satuan || '-'}</td>
          </tr>`;
      }).join('');

      const printHtml = `<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Print Order - ${header.noOrder}</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Outfit:wght@400;600;700;800&display=swap');
    * { margin:0; padding:0; box-sizing:border-box; }
    body { font-family:'Plus Jakarta Sans',sans-serif; background:#f8f9fa; color:#1a1a2e; font-size:13px; }
    .print-page { background:#fff; max-width:820px; margin:20px auto; padding:0; border-radius:12px; overflow:hidden; box-shadow:0 4px 24px rgba(0,0,0,0.12); }
    /* HEADER */
    .print-header { background:linear-gradient(135deg,#1a3a5c 0%,#0a1628 100%); color:#fff; padding:28px 36px; display:flex; align-items:center; justify-content:space-between; }
    .print-header .company { }
    .print-header .company h1 { font-family:'Outfit',sans-serif; font-size:22px; font-weight:800; letter-spacing:-0.5px; color:#fff; }
    .print-header .company p { font-size:12px; color:rgba(255,255,255,0.65); margin-top:3px; }
    .print-header .doc-info { text-align:right; }
    .print-header .doc-title { font-family:'Outfit',sans-serif; font-size:13px; font-weight:700; text-transform:uppercase; letter-spacing:2px; color:#f59e0b; }
    .print-header .doc-no { font-family:'Outfit',sans-serif; font-size:26px; font-weight:800; color:#fff; margin-top:4px; letter-spacing:-0.5px; }
    .print-header .doc-date { font-size:11px; color:rgba(255,255,255,0.6); margin-top:4px; }
    /* STATUS BANNER */
    .status-banner { padding:10px 36px; display:flex; align-items:center; justify-content:space-between; }
    .status-banner.status-sent { background:linear-gradient(90deg,#10b98115,#10b98108); border-bottom:2px solid #10b981; }
    .status-banner.status-pending { background:linear-gradient(90deg,#f59e0b15,#f59e0b08); border-bottom:2px solid #f59e0b; }
    .status-label { font-weight:800; font-size:13px; }
    .status-banner.status-sent .status-label { color:#059669; }
    .status-banner.status-pending .status-label { color:#d97706; }
    .print-meta { font-size:11px; color:#888; }
    /* INFO SECTION */
    .info-section { display:grid; grid-template-columns:1fr 1fr; gap:24px; padding:24px 36px; border-bottom:1px solid #eee; }
    .info-card { }
    .info-card h3 { font-family:'Outfit',sans-serif; font-size:10px; font-weight:800; text-transform:uppercase; letter-spacing:1.5px; color:#94a3b8; margin-bottom:12px; padding-bottom:6px; border-bottom:1px solid #f1f5f9; }
    .info-row { display:flex; gap:8px; margin-bottom:8px; align-items:flex-start; }
    .info-label { font-size:11px; color:#888; font-weight:600; min-width:80px; flex-shrink:0; padding-top:1px; }
    .info-value { font-size:13px; color:#1a1a2e; font-weight:600; flex:1; line-height:1.4; }
    .info-value.highlight { color:#1a3a5c; font-weight:800; font-family:'Outfit',sans-serif; }
    .info-value.resi { font-family:monospace; background:#f8f9fa; padding:2px 8px; border-radius:4px; font-size:12px; }
    /* ITEMS TABLE */
    .items-section { padding:0 36px 24px; }
    .items-section h3 { font-family:'Outfit',sans-serif; font-size:10px; font-weight:800; text-transform:uppercase; letter-spacing:1.5px; color:#94a3b8; padding:20px 0 12px; border-bottom:2px solid #f1f5f9; margin-bottom:0; }
    table.print-table { width:100%; border-collapse:collapse; margin-top:0; }
    .print-table thead tr { background:#f8f9fa; }
    .print-table th { padding:11px 12px; text-align:left; font-size:10px; font-weight:800; color:#94a3b8; text-transform:uppercase; letter-spacing:0.8px; border-bottom:2px solid #e2e8f0; white-space:nowrap; }
    .print-table td { padding:11px 12px; border-bottom:1px solid #f1f5f9; font-size:13px; vertical-align:middle; }
    .print-table tr:last-child td { border-bottom:2px solid #e2e8f0; }
    .print-table tr:nth-child(even) td { background:#fafafa; }
    .lokasi-badge { background:#e0f2fe; color:#0369a1; padding:2px 8px; border-radius:4px; font-size:11px; font-weight:700; font-family:monospace; }
    /* FOOTER */
    .print-footer { padding:20px 36px 28px; }
    .totals-row { background:#f8f9fa; border-radius:8px; padding:14px 20px; display:flex; justify-content:flex-end; align-items:center; gap:20px; margin-bottom:24px; }
    .total-label { font-size:13px; color:#555; font-weight:600; }
    .total-value { font-family:'Outfit',sans-serif; font-size:22px; font-weight:800; color:#1a3a5c; }
    .total-satuan { font-size:12px; color:#888; }
    .sign-section { display:grid; grid-template-columns:1fr 1fr 1fr; gap:24px; margin-top:20px; }
    .sign-box { text-align:center; }
    .sign-box .sign-title { font-size:11px; font-weight:800; color:#888; text-transform:uppercase; letter-spacing:0.8px; margin-bottom:60px; }
    .sign-box .sign-line { border-top:1.5px solid #444; padding-top:6px; font-size:12px; font-weight:600; color:#444; }
    .print-watermark { text-align:center; margin-top:24px; font-size:10px; color:#ccc; padding-top:16px; border-top:1px solid #f0f0f0; }
    /* PRINT CONTROLS OVERLAY */
    .print-controls { position:fixed; bottom:24px; right:24px; display:flex; gap:10px; z-index:9999; }
    .btn-print { background:linear-gradient(135deg,#1a3a5c,#0a1628); color:#fff; border:none; padding:14px 28px; border-radius:10px; font-family:'Outfit',sans-serif; font-size:15px; font-weight:700; cursor:pointer; box-shadow:0 4px 20px rgba(26,58,92,0.4); transition:all 0.2s; }
    .btn-print:hover { transform:translateY(-2px); box-shadow:0 8px 28px rgba(26,58,92,0.5); }
    .btn-close-print { background:#fff; color:#555; border:2px solid #ddd; padding:14px 20px; border-radius:10px; font-family:'Outfit',sans-serif; font-size:15px; font-weight:700; cursor:pointer; transition:all 0.2s; }
    .btn-close-print:hover { border-color:#999; color:#333; }
    @media print {
      body { background:#fff; }
      .print-page { box-shadow:none; margin:0; border-radius:0; }
      .print-controls { display:none !important; }
      @page { margin:10mm 12mm; size:A4; }
    }
  </style>
</head>
<body>
  <div class="print-controls">
    <button class="btn-close-print" onclick="window.close()">✕ Tutup</button>
    <button class="btn-print" onclick="window.print()">🖨️ Cetak Sekarang</button>
  </div>
  <div class="print-page">
    <!-- HEADER -->
    <div class="print-header">
      <div class="company">
        <h1>GUDANG FCL GROUP</h1>
        <p>Sistem Pengelola Gudang &amp; Distribusi</p>
      </div>
      <div class="doc-info">
        <div class="doc-title">Packing Slip / Surat Jalan</div>
        <div class="doc-no">${header.noOrder || '-'}</div>
        <div class="doc-date">Dicetak: ${printDate} ${printTime}</div>
      </div>
    </div>

    <!-- STATUS BANNER -->
    <div class="status-banner ${statusClass}">
      <div class="status-label">${statusLabel} — ${kategoriIco} ${header.kategori || 'Distributor'}</div>
      <div class="print-meta">Tanggal Order: <strong>${header.tanggal || '-'}</strong>${header.noResi ? ' &nbsp;|&nbsp; No. Resi: <strong>' + header.noResi + '</strong>' : ''}</div>
    </div>

    <!-- INFO SECTION -->
    <div class="info-section">
      <div class="info-card">
        <h3>📦 Informasi Pengiriman</h3>
        <div class="info-row">
          <span class="info-label">Pelanggan</span>
          <span class="info-value highlight">${header.pelanggan || '-'}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Alamat</span>
          <span class="info-value">${header.alamat || '-'}</span>
        </div>
        ${header.noResi ? `<div class="info-row"><span class="info-label">No. Resi</span><span class="info-value resi">${header.noResi}</span></div>` : ''}
        <div class="info-row">
          <span class="info-label">Kategori</span>
          <span class="info-value">${kategoriIco} ${header.kategori || 'Distributor'}</span>
        </div>
      </div>
      <div class="info-card">
        <h3>📋 Detail Dokumen</h3>
        <div class="info-row">
          <span class="info-label">No. Order</span>
          <span class="info-value highlight">${header.noOrder || '-'}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Tanggal</span>
          <span class="info-value">${header.tanggal || '-'}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Status</span>
          <span class="info-value">${header.status || '-'}</span>
        </div>
        ${header.keterangan ? `<div class="info-row"><span class="info-label">Keterangan</span><span class="info-value">${header.keterangan}</span></div>` : ''}
        <div class="info-row">
          <span class="info-label">Total Item</span>
          <span class="info-value highlight">${header.totalItem || items.length} Jenis</span>
        </div>
      </div>
    </div>

    <!-- ITEMS TABLE -->
    <div class="items-section">
      <h3>📦 Daftar Barang</h3>
      <table class="print-table">
        <thead>
          <tr>
            <th style="text-align:center;width:36px;">#</th>
            <th>SKU</th>
            <th>Nama Barang</th>
            <th style="text-align:center;">Lokasi</th>
            <th style="text-align:center;">Batch</th>
            <th style="text-align:center;">Exp Date</th>
            <th style="text-align:center;">Qty</th>
            <th style="text-align:center;">Satuan</th>
          </tr>
        </thead>
        <tbody>
          ${itemRows || '<tr><td colspan="8" style="text-align:center;color:#aaa;padding:20px;">Tidak ada item</td></tr>'}
        </tbody>
      </table>
    </div>

    <!-- FOOTER -->
    <div class="print-footer">
      <div class="totals-row">
        <span class="total-label">Total Qty Keseluruhan:</span>
        <span class="total-value">${totalQty}</span>
        <span class="total-satuan">Unit</span>
      </div>

      <div class="sign-section">
        <div class="sign-box">
          <div class="sign-title">Disiapkan Oleh</div>
          <div class="sign-line">( ................................ )</div>
        </div>
        <div class="sign-box">
          <div class="sign-title">Diperiksa Oleh</div>
          <div class="sign-line">( ................................ )</div>
        </div>
        <div class="sign-box">
          <div class="sign-title">Diterima Oleh</div>
          <div class="sign-line">( ................................ )</div>
        </div>
      </div>

      <div class="print-watermark">
        Dokumen ini dicetak secara otomatis oleh Sistem Gudang FCL Group &bull; ${printDate} ${printTime}
      </div>
    </div>
  </div>
</body>
</html>`;

      const printWin = window.open('', '_blank', 'width=900,height=700,scrollbars=yes');
      if (!printWin) return toast('Popup diblokir browser! Izinkan popup untuk website ini.', 'error');
      printWin.document.write(printHtml);
      printWin.document.close();
    }


    function switchOrderTab(tab) {
      document.querySelectorAll('.order-tab-panel').forEach(p => p.style.display = 'none');
      document.querySelectorAll('#page-order .kar-tab').forEach(t => t.classList.remove('active'));

      if (tab === 'MP') {
        document.getElementById('panelOrderMP').style.display = 'block';
        document.getElementById('tabOrdMP').classList.add('active');
      } else if (tab === 'Dist') {
        document.getElementById('panelOrderDist').style.display = 'block';
        document.getElementById('tabOrdDist').classList.add('active');
      } else if (tab === 'Store') {
        document.getElementById('panelOrderStore').style.display = 'block';
        document.getElementById('tabOrdStore').classList.add('active');
      } else if (tab === 'History') {
        document.getElementById('panelOrderHistory').style.display = 'block';
        document.getElementById('tabOrdHistory').classList.add('active');
      }
    }

    let distributorQueueData = [];

    function queueLinkHtml(url, label) {
      if (!url) return '-';
      const safeUrl = escHtml(url);
      return `<a href="${safeUrl}" target="_blank">${label || 'Buka Link'}</a>`;
    }

    function queueCell(value) {
      return escHtml(value || '-');
    }

    function queueSlaBadge(sla) {
      if (!sla) return '<span class="badge bg-secondary">-</span>';
      if (sla.isLate) return `<span class="badge bg-danger">Late${sla.lateDays ? ' +' + sla.lateDays + ' hari' : ''}</span>`;
      if (sla.status === 'On Time') return '<span class="badge bg-success">On Time</span>';
      if (sla.status === 'Pending') return '<span class="badge bg-warning text-dark">Pending</span>';
      return `<span class="badge bg-secondary">${queueCell(sla.status)}</span>`;
    }

    let dqChartInstance = null;
    let dqChartType = 'doughnut';

    function toggleChartType() {
      dqChartType = dqChartType === 'doughnut' ? 'bar' : 'doughnut';
      if (window._lastDqDash) renderDistributorQueueChart(window._lastDqDash);
    }

    function toggleFormCard() {
      const card = document.getElementById('dqFormCard');
      if (!card) return;
      const isHidden = card.style.display === 'none' || card.style.display === '';
      card.style.display = isHidden ? 'block' : 'none';
      const btn = document.querySelector('[onclick="toggleFormCard()"]');
      if (btn) btn.textContent = isHidden ? '✖ Tutup Form' : '📝 Tampilkan Form';
      if (isHidden) card.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    function renderDistributorQueueChart(dash) {
      window._lastDqDash = dash;
      const ctx = document.getElementById('dqChart');
      if (!ctx) return;

      const labels = ['PO Selesai', 'Belum Selesai', 'Belum Dikerjakan', 'Late', 'Terkirim', 'Ready to Pickup'];
      const values = [
        dash.selesai || 0,
        dash.belumSelesai || 0,
        dash.belumDikerjakan || 0,
        dash.late || 0,
        dash.totalTerkirim || 0,
        dash.totalReadyToPickup || 0
      ];
      const colors = [
        'rgba(16, 185, 129, 0.85)',
        'rgba(245, 158, 11, 0.85)',
        'rgba(148, 163, 184, 0.85)',
        'rgba(239, 68, 68, 0.85)',
        'rgba(16, 185, 129, 0.6)',
        'rgba(14, 165, 233, 0.85)'
      ];
      const borderColors = [
        'rgba(16, 185, 129, 1)',
        'rgba(245, 158, 11, 1)',
        'rgba(148, 163, 184, 1)',
        'rgba(239, 68, 68, 1)',
        'rgba(16, 185, 129, 1)',
        'rgba(14, 165, 233, 1)'
      ];

      if (dqChartInstance) { dqChartInstance.destroy(); dqChartInstance = null; }

      const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
      const textColor = isDark ? '#94a3b8' : '#64748b';
      const gridColor = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';

      if (dqChartType === 'doughnut') {
        dqChartInstance = new Chart(ctx, {
          type: 'doughnut',
          data: {
            labels: labels,
            datasets: [{
              data: values,
              backgroundColor: colors,
              borderColor: borderColors,
              borderWidth: 2,
              hoverOffset: 12
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: true,
            cutout: '62%',
            plugins: {
              legend: {
                position: 'right',
                labels: {
                  color: textColor,
                  font: { family: 'Plus Jakarta Sans', size: 12, weight: '600' },
                  padding: 16,
                  usePointStyle: true,
                  pointStyleWidth: 10
                }
              },
              tooltip: {
                backgroundColor: isDark ? '#0f2040' : '#fff',
                titleColor: isDark ? '#f8fafc' : '#0f172a',
                bodyColor: textColor,
                borderColor: isDark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)',
                borderWidth: 1,
                padding: 12,
                callbacks: {
                  label: function (ctx) {
                    const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                    const pct = total > 0 ? ((ctx.parsed / total) * 100).toFixed(1) : 0;
                    return ` ${ctx.label}: ${ctx.parsed} (${pct}%)`;
                  }
                }
              }
            }
          }
        });
      } else {
        dqChartInstance = new Chart(ctx, {
          type: 'bar',
          data: {
            labels: labels,
            datasets: [{
              label: 'Jumlah PO',
              data: values,
              backgroundColor: colors,
              borderColor: borderColors,
              borderWidth: 2,
              borderRadius: 8,
              borderSkipped: false
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
              legend: { display: false },
              tooltip: {
                backgroundColor: isDark ? '#0f2040' : '#fff',
                titleColor: isDark ? '#f8fafc' : '#0f172a',
                bodyColor: textColor,
                borderColor: isDark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)',
                borderWidth: 1,
                padding: 12
              }
            },
            scales: {
              x: {
                ticks: { color: textColor, font: { family: 'Plus Jakarta Sans', size: 12 } },
                grid: { color: gridColor }
              },
              y: {
                beginAtZero: true,
                ticks: { color: textColor, font: { family: 'Plus Jakarta Sans', size: 12 }, stepSize: 1 },
                grid: { color: gridColor }
              }
            }
          }
        });
      }
    }

    function openExportDQModal() {
      const now = new Date();
      setVal('dqExportMonth', now.getMonth() + 1);
      setVal('dqExportYear', now.getFullYear());
      openModal('modalExportDQ');
    }

    function doExportDistributorQueue() {
      const month = parseInt(v('dqExportMonth'), 10);
      const year = parseInt(v('dqExportYear'), 10);
      if (!month || !year || year < 2020) return toast('Pilih bulan dan tahun yang valid', 'error');

      const monthNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
      const monthLabel = monthNames[month - 1];

      // Filter data by order queue time month/year
      const filtered = (distributorQueueData || []).filter(item => {
        const raw = item.orderQueueTime || item.timeWib || '';
        if (!raw || raw === '-') return false;
        const d = new Date(raw);
        if (isNaN(d)) return false;
        return d.getMonth() + 1 === month && d.getFullYear() === year;
      });

      if (!filtered.length) {
        return toast(`Tidak ada data untuk ${monthLabel} ${year}`, 'info');
      }

      const btn = document.getElementById('btnDoExportDQ');
      btn.disabled = true;
      btn.textContent = '⏳ Memproses...';

      try {
        // ── Headers ──
        const headers = [
          'Order Queue Time', 'PIC Sales', 'Nama Distributor', 'Alamat', 'No. HP',
          'PO Number', 'No Mabang', 'Metode Pengiriman', 'Ongkir Dibayar Oleh', 'Note',
          'Time', 'Status Gudang', 'Jumlah Dus', 'Total Pcs', 'Packer', 'Validation',
          'Tgl Selesai Packing', 'Ship Date', 'Status Mabang', 'GDrive', 'Delivery Bill',
          'Nomor Resi', 'Bukti Pengiriman', 'Status SLA', 'Terlambat (Hari)', 'Catatan Late'
        ];

        // ── Build rows ──
        const rows = filtered.map(item => {
          const sla = item.sla || {};
          return [
            item.orderQueueTime || '',
            item.picSales || '',
            item.namaDistributor || '',
            item.alamat || '',
            item.noHp || '',
            item.poNumber || '',
            item.noMabang || '',
            item.metodePengiriman || '',
            item.ongkirDibayarOleh || '',
            item.note || '',
            item.timeWib || '',
            item.statusGudang || '',
            item.jumlahDus || '',
            item.totalPcs || '',
            item.packer || '',
            item.validation || '',
            item.tanggalSelesaiPacking || '',
            item.shipDate || '',
            item.statusMabang || '',
            item.gdrive || '',
            item.deliveryBill || '',
            item.nomorResi || '',
            item.buktiPengiriman || '',
            sla.status || '',
            sla.lateDays || 0,
            item.catatanLate || ''
          ];
        });

        // ── Build workbook with XLSX ──
        const wb = XLSX.utils.book_new();
        const wsData = [headers, ...rows];
        const ws = XLSX.utils.aoa_to_sheet(wsData);

        // ── Column widths ──
        ws['!cols'] = [
          { wch: 16 }, { wch: 14 }, { wch: 22 }, { wch: 28 }, { wch: 14 },
          { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 18 }, { wch: 24 },
          { wch: 16 }, { wch: 16 }, { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 12 },
          { wch: 18 }, { wch: 14 }, { wch: 16 }, { wch: 28 }, { wch: 28 },
          { wch: 18 }, { wch: 28 }, { wch: 12 }, { wch: 16 }, { wch: 30 }
        ];

        // ── Style header row (bold, dark bg, white text) ──
        const headerStyle = {
          font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 11 },
          fill: { fgColor: { rgb: '1A3A5C' } },
          alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
          border: {
            top: { style: 'thin', color: { rgb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { rgb: 'FFFFFF' } },
            left: { style: 'thin', color: { rgb: 'FFFFFF' } },
            right: { style: 'thin', color: { rgb: 'FFFFFF' } }
          }
        };

        // ── Style for normal data rows ──
        const normalStyle = {
          font: { sz: 10 },
          alignment: { vertical: 'center', wrapText: false },
          border: {
            top: { style: 'thin', color: { rgb: 'DDDDDD' } },
            bottom: { style: 'thin', color: { rgb: 'DDDDDD' } },
            left: { style: 'thin', color: { rgb: 'DDDDDD' } },
            right: { style: 'thin', color: { rgb: 'DDDDDD' } }
          }
        };

        // ── Style for LATE rows (red bg, white text) ──
        const lateStyle = {
          font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 10 },
          fill: { fgColor: { rgb: 'C0392B' } },
          alignment: { vertical: 'center', wrapText: false },
          border: {
            top: { style: 'thin', color: { rgb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { rgb: 'FFFFFF' } },
            left: { style: 'thin', color: { rgb: 'FFFFFF' } },
            right: { style: 'thin', color: { rgb: 'FFFFFF' } }
          }
        };

        // ── Style for LATE status cell (brighter red) ──
        const lateStatusStyle = {
          font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 10 },
          fill: { fgColor: { rgb: 'E74C3C' } },
          alignment: { horizontal: 'center', vertical: 'center' },
          border: {
            top: { style: 'thin', color: { rgb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { rgb: 'FFFFFF' } },
            left: { style: 'thin', color: { rgb: 'FFFFFF' } },
            right: { style: 'thin', color: { rgb: 'FFFFFF' } }
          }
        };

        const colCount = headers.length;
        const encodeCell = (r, c) => XLSX.utils.encode_cell({ r, c });

        // Apply header styles
        for (let c = 0; c < colCount; c++) {
          const cellRef = encodeCell(0, c);
          if (ws[cellRef]) ws[cellRef].s = headerStyle;
        }

        // Apply row styles
        filtered.forEach((item, rowIdx) => {
          const isLate = item.sla && item.sla.isLate;
          const excelRow = rowIdx + 1; // +1 for header
          for (let c = 0; c < colCount; c++) {
            const cellRef = encodeCell(excelRow, c);
            if (!ws[cellRef]) ws[cellRef] = { v: '', t: 's' };
            // Status SLA column (index 23) gets special style
            if (isLate && c === 23) {
              ws[cellRef].s = lateStatusStyle;
            } else if (isLate) {
              ws[cellRef].s = lateStyle;
            } else {
              ws[cellRef].s = normalStyle;
            }
          }
        });

        // Freeze header row
        ws['!freeze'] = { xSplit: 0, ySplit: 1 };

        XLSX.utils.book_append_sheet(wb, ws, `${monthLabel} ${year}`);

        // ── Summary sheet ──
        const lateCount = filtered.filter(i => i.sla && i.sla.isLate).length;
        const onTimeCount = filtered.filter(i => i.sla && i.sla.status === 'On Time').length;
        const terkirim = filtered.filter(i => i.shipDate && String(i.shipDate).trim() !== '').length;
        const readyPickup = filtered.filter(i => i.nomorResi && String(i.nomorResi).trim() !== '' && !(i.shipDate && String(i.shipDate).trim() !== '')).length;

        const summaryData = [
          ['RINGKASAN ANTRIAN DISTRIBUTOR'],
          [`Periode: ${monthLabel} ${year}`],
          [''],
          ['Keterangan', 'Jumlah'],
          ['Total PO', filtered.length],
          ['PO Selesai (Ship Date terisi)', terkirim],
          ['Ready to Pickup (Resi ada, belum Ship)', readyPickup],
          ['On Time', onTimeCount],
          ['Late Shipment', lateCount],
          ['', ''],
          ['Dibuat pada', new Date().toLocaleString('id-ID')]
        ];

        const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
        wsSummary['!cols'] = [{ wch: 36 }, { wch: 16 }];

        // Style summary title
        if (wsSummary['A1']) wsSummary['A1'].s = { font: { bold: true, sz: 14, color: { rgb: '1A3A5C' } } };
        if (wsSummary['A4']) wsSummary['A4'].s = { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1A3A5C' } } };
        if (wsSummary['B4']) wsSummary['B4'].s = { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1A3A5C' } } };
        // Late row red
        if (wsSummary['A9']) wsSummary['A9'].s = { font: { bold: true, color: { rgb: 'C0392B' } } };
        if (wsSummary['B9']) wsSummary['B9'].s = { font: { bold: true, color: { rgb: 'C0392B' } } };

        XLSX.utils.book_append_sheet(wb, wsSummary, 'Ringkasan');

        const filename = `Antrian_Distributor_${monthLabel}_${year}.xlsx`;
        XLSX.writeFile(wb, filename, { bookType: 'xlsx', cellStyles: true });

        toast(`Export berhasil: ${filtered.length} data (${lateCount} late)`, 'success');
        closeModal('modalExportDQ');
      } catch (err) {
        toast('Gagal export: ' + err.message, 'error');
      } finally {
        btn.disabled = false;
        btn.textContent = '📥 Export Excel';
      }
    }

    function renderDistributorQueueStatusCards(data) {
      try {
        console.log('renderDistributorQueueStatusCards called with data:', data);

        if (!data || !data.length) {
          console.log('No data available, showing empty state');
          // Set empty state untuk semua tabel
          ['dqTableBelumDikerjakan', 'dqTableBelumSelesai'].forEach(function (id) {
            const tbody = document.getElementById(id);
            if (tbody) tbody.innerHTML = '<tr><td colspan="4" style="text-align:center; color:var(--text-muted); padding:20px; font-size:12px;">Tidak ada data</td></tr>';
          });
          const tbodyReadyPickup = document.getElementById('dqTableReadyPickup');
          if (tbodyReadyPickup) tbodyReadyPickup.innerHTML = '<tr><td colspan="5" style="text-align:center; color:var(--text-muted); padding:20px; font-size:12px;">Tidak ada data</td></tr>';

          // Reset badges
          setVal('dqCardBadgeBelumDikerjakan', 0);
          setVal('dqCardBadgeBelumSelesai', 0);
          setVal('dqCardBadgeReadyPickup', 0);
          setVal('dqStatBelumDikerjakan', 0);
          setVal('dqStatBelumSelesai', 0);
          setVal('dqStatReadyPickup', 0);

          return;
        }

        console.log('Processing', data.length, 'items');

        // ── Classify rows (OPTIMIZED) ──
        const belumDikerjakan = [];
        const belumSelesai = [];
        const readyPickup = [];

        // Gunakan for loop untuk performa lebih baik
        for (let i = 0; i < data.length; i++) {
          const item = data[i];
          const statusRaw = String(item.statusGudang || '').trim();
          const statusNorm = statusRaw.toLowerCase().replace(/\s+/g, '');

          // PO Belum Dikerjakan: kolom L kosong
          if (statusRaw === '') {
            belumDikerjakan.push(item);
            continue;
          }

          // PO Ready Pickup: kolom L mengandung ready/pickup/siap
          if (statusNorm.includes('ready') || statusNorm.includes('pickup') || statusNorm.includes('siap')) {
            readyPickup.push(item);
            continue;
          }

          // PO Belum Selesai: kolom L = "Picking" (atau mengandung "picking")
          if (statusNorm.includes('picking')) {
            belumSelesai.push(item);
            continue;
          }

          // Status lain (Terkirim, Selesai, dll) — tidak masuk ke 3 card ini
        }

        // ── Helper untuk badge source sheet ──
        function getSheetBadge(sourceSheet) {
          const sheet = sourceSheet || 'Antrian Distributor';
          if (sheet === 'ANTRIAN FOCALSKIN') {
            return '<span style="background:#10b98122;color:#10b981;padding:2px 6px;border-radius:12px;font-size:9px;font-weight:700;">FOCALSKIN</span>';
          } else if (sheet === 'ANTRIAN MISTINE') {
            return '<span style="background:#0ea5e922;color:#0ea5e9;padding:2px 6px;border-radius:12px;font-size:9px;font-weight:700;">MISTINE</span>';
          } else if (sheet === 'ANTRIAN SBY') {
            return '<span style="background:#f59e0b22;color:#f59e0b;padding:2px 6px;border-radius:12px;font-size:9px;font-weight:700;">SBY</span>';
          } else {
            return '<span style="background:#94a3b822;color:#94a3b8;padding:2px 6px;border-radius:12px;font-size:9px;font-weight:700;">MAIN</span>';
          }
        }

        // ── Render helper untuk PO Belum Dikerjakan & PO Belum Selesai (tanpa Jumlah Dus) ──
        function renderRows(tbodyId, badgeId, rows, accentColor) {
          const tbody = document.getElementById(tbodyId);
          const badge = document.getElementById(badgeId);
          if (!tbody) return;
          if (badge) badge.textContent = rows.length;
          if (!rows.length) {
            tbody.innerHTML = '<tr><td colspan="4" style="text-align:center; color:var(--text-muted); padding:20px; font-size:12px;">Tidak ada data</td></tr>';
            return;
          }

          // Tampilkan semua data tanpa batasan
          const fragment = document.createDocumentFragment();

          rows.forEach(function (item) {
            const tr = document.createElement('tr');
            tr.style.cursor = 'pointer';
            tr.title = 'Klik untuk edit';
            tr.onclick = function () { editDistributorQueue(item.rowNumber); };

            const dist = escHtml(item.namaDistributor || '-');
            const po = escHtml(item.poNumber || '-');
            // Perbaikan: Pastikan totalPcs adalah angka valid sebelum format
            let pcs = '-';
            if (item.totalPcs && item.totalPcs !== '' && !isNaN(item.totalPcs)) {
              pcs = Number(item.totalPcs).toLocaleString('id-ID');
            }
            const sheetBadge = getSheetBadge(item.sourceSheet);

            tr.innerHTML = `
            <td style="padding:8px 12px;">${sheetBadge}</td>
            <td style="padding:8px 12px; max-width:120px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">
              <strong style="font-size:12px;">${dist}</strong>
            </td>
            <td style="padding:8px 12px;">
              <span style="font-size:11px; background:rgba(255,255,255,0.06); border:1px solid var(--border-color); border-radius:6px; padding:2px 8px;">${po}</span>
            </td>
            <td style="padding:8px 12px; text-align:right; font-weight:700; color:${accentColor}; font-size:13px;">${pcs}</td>`;

            fragment.appendChild(tr);
          });

          tbody.innerHTML = '';
          tbody.appendChild(fragment);
        }

        renderRows('dqTableBelumDikerjakan', 'dqCardBadgeBelumDikerjakan', belumDikerjakan, 'var(--teal)');
        renderRows('dqTableBelumSelesai', 'dqCardBadgeBelumSelesai', belumSelesai, 'var(--accent)');

        // Ready Pickup — custom render with Jumlah Dus column (OPTIMIZED)
        (function () {
          const tbody = document.getElementById('dqTableReadyPickup');
          const badge = document.getElementById('dqCardBadgeReadyPickup');
          if (!tbody) return;
          if (badge) badge.textContent = readyPickup.length;
          if (!readyPickup.length) {
            tbody.innerHTML = '<tr><td colspan="5" style="text-align:center; color:var(--text-muted); padding:20px; font-size:12px;">Tidak ada data</td></tr>';
            return;
          }

          // Batasi hanya 20 item pertama untuk performa
          const limitedRows = readyPickup.slice(0, 20);
          const fragment = document.createDocumentFragment();

          limitedRows.forEach(function (item) {
            const tr = document.createElement('tr');
            tr.style.cursor = 'pointer';
            tr.title = 'Klik untuk edit';
            tr.onclick = function () { editDistributorQueue(item.rowNumber); };

            const dist = escHtml(item.namaDistributor || '-');
            const po = escHtml(item.poNumber || '-');

            // Perbaikan: Validasi totalPcs dan jumlahDus sebelum format
            let pcs = '-';
            if (item.totalPcs && item.totalPcs !== '' && !isNaN(item.totalPcs)) {
              pcs = Number(item.totalPcs).toLocaleString('id-ID');
            }

            let dus = '-';
            if (item.jumlahDus && item.jumlahDus !== '' && !isNaN(item.jumlahDus)) {
              dus = Number(item.jumlahDus).toLocaleString('id-ID');
            }

            const sheetBadge = getSheetBadge(item.sourceSheet);

            tr.innerHTML = `
            <td style="padding:8px 12px;">${sheetBadge}</td>
            <td style="padding:8px 12px; max-width:110px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">
              <strong style="font-size:12px;">${dist}</strong>
            </td>
            <td style="padding:8px 12px;">
              <span style="font-size:11px; background:rgba(255,255,255,0.06); border:1px solid var(--border-color); border-radius:6px; padding:2px 8px;">${po}</span>
            </td>
            <td style="padding:8px 12px; text-align:right; font-weight:700; color:var(--green); font-size:13px;">${pcs}</td>
            <td style="padding:8px 12px; text-align:right; font-weight:700; color:var(--teal); font-size:13px;">${dus}</td>`;

            fragment.appendChild(tr);
          });

          tbody.innerHTML = '';
          tbody.appendChild(fragment);

          // Tampilkan info jika ada lebih banyak data
          if (readyPickup.length > 20) {
            const infoTr = document.createElement('tr');
            infoTr.innerHTML = `<td colspan="5" style="text-align:center; color:var(--text-muted); padding:10px; font-size:11px; background:rgba(16,185,129,0.05);">Menampilkan 20 dari ${readyPickup.length} data</td>`;
            tbody.appendChild(infoTr);
          }
        })();

        // ── Sync stat cards dengan jumlah dari tabel card (sumber kebenaran tunggal) ──
        setVal('dqStatBelumDikerjakan', belumDikerjakan.length);
        setVal('dqStatBelumSelesai', belumSelesai.length);
        setVal('dqStatReadyPickup', readyPickup.length);

        console.log('Rendering complete:', {
          belumDikerjakan: belumDikerjakan.length,
          belumSelesai: belumSelesai.length,
          readyPickup: readyPickup.length
        });

      } catch (error) {
        console.error('Error in renderDistributorQueueStatusCards:', error);
        toast('Error menampilkan data: ' + error.message, 'error');
      }
    }

    function openEditSLAModal() {
      if (!canEditSLARule()) return toast('Anda tidak punya hak edit SLA Rules', 'error');
      google.script.run
        .withSuccessHandler(function (res) {
          if (res && res.success && res.data) {
            setVal('slaInputDueDays', res.data.dueDays || 1);
            setVal('slaInputRuleDesc', res.data.ruleDescription || 'SLA H+1 dari Order queue time');
          }
          openModal('modalEditSLARules');
        })
        .withFailureHandler(function (err) {
          setVal('slaInputDueDays', 1);
          setVal('slaInputRuleDesc', 'SLA H+1 dari Order queue time');
          openModal('modalEditSLARules');
        })
        .getDistributorQueueSLASettings();
    }

    function saveSLARules() {
      if (!canEditSLARule()) return toast('Anda tidak punya hak edit SLA Rules', 'error');
      var dueDays = parseInt(v('slaInputDueDays'), 10);
      var ruleDescription = v('slaInputRuleDesc').trim();
      if (!dueDays || dueDays < 1) return toast('Batas hari SLA harus minimal 1', 'error');
      if (!ruleDescription) return toast('Deskripsi aturan tidak boleh kosong', 'error');
      var btn = document.getElementById('btnSaveSLARules');
      btn.disabled = true;
      btn.textContent = '⏳ Menyimpan...';
      google.script.run
        .withSuccessHandler(function (res) {
          btn.disabled = false;
          btn.textContent = '💾 Simpan SLA';
          if (!res || !res.success) return toast(res ? res.message : 'Gagal menyimpan SLA', 'error');
          toast(res.message || 'SLA Rules berhasil disimpan', 'success');
          closeModal('modalEditSLARules');
          loadDistributorQueue();
        })
        .withFailureHandler(function (err) {
          btn.disabled = false;
          btn.textContent = '💾 Simpan SLA';
          toast('Gagal menyimpan SLA: ' + err, 'error');
        })
        .saveDistributorQueueSLASettings({ dueDays: dueDays, ruleDescription: ruleDescription });
    }

    function openEditLateCatatanModal(rowNumber) {
      if (!canUpdateLateNote()) return toast('Anda tidak punya hak update keterangan late', 'error');
      var item = distributorQueueData.find(function (d) { return Number(d.rowNumber) === Number(rowNumber); });
      if (!item) return toast('Data antrian tidak ditemukan', 'error');
      setVal('catatanLateRowNumber', rowNumber);
      setVal('catatanLatePoNumber', item.poNumber || '');
      setVal('catatanLateSourceSheet', item.sourceSheet || 'Antrian Distributor');
      setVal('catatanLatePoDisplay', item.poNumber || '-');
      setVal('catatanLateText', item.catatanLate || '');
      setVal('catatanLateApprovalStatus', item.catatanLateStatus || 'Belum ada keterangan');
      openModal('modalEditCatatanLate');
    }

    function saveCatatanLate() {
      if (!canUpdateLateNote()) return toast('Anda tidak punya hak update keterangan late', 'error');
      var rowNumber = v('catatanLateRowNumber');
      var poNumber = v('catatanLatePoNumber');
      var sourceSheet = v('catatanLateSourceSheet');
      var note = v('catatanLateText').trim();
      if (!rowNumber) return toast('Data tidak valid', 'error');
      var btn = document.getElementById('btnSaveCatatanLate');
      btn.disabled = true;
      btn.textContent = '⏳ Menyimpan...';
      google.script.run
        .withSuccessHandler(function (res) {
          btn.disabled = false;
          btn.textContent = '💾 Simpan Keterangan';
          if (!res || !res.success) return toast(res ? res.message : 'Gagal menyimpan keterangan', 'error');
          toast(res.message || 'Keterangan berhasil disimpan', 'success');
          closeModal('modalEditCatatanLate');
          loadDistributorQueue();
        })
        .withFailureHandler(function (err) {
          btn.disabled = false;
          btn.textContent = '💾 Simpan Keterangan';
          toast('Gagal menyimpan keterangan: ' + err, 'error');
        })
        .saveDistributorQueueLateNote(poNumber, rowNumber, note, sourceSheet, currentUser ? currentUser.username : '');
    }

    // ============================================================
    // PRINT SURAT JALAN
    // ============================================================
    let sjSelectedPOs = [];

    // Debug function to check all statuses
    function debugDistributorQueueStatuses() {
      console.log('=== ALL DISTRIBUTOR QUEUE STATUSES ===');
      console.log('Total data:', (distributorQueueData || []).length);

      const statusMap = {};
      (distributorQueueData || []).forEach(item => {
        const status = String(item.statusGudang || '').trim() || '(kosong)';
        statusMap[status] = (statusMap[status] || 0) + 1;
      });

      console.log('Status breakdown:');
      Object.keys(statusMap).sort().forEach(status => {
        console.log(`  "${status}": ${statusMap[status]} PO`);
      });

      return statusMap;
    }

    function openPrintSuratJalanModal() {
      // Set tanggal hari ini
      const today = new Date().toISOString().split('T')[0];
      setVal('sjTanggal', today);
      setVal('sjKeterangan', '');

      // Debug: Log semua status yang ada
      console.log('=== DEBUG: Checking Distributor Queue Data ===');
      console.log('Total data:', (distributorQueueData || []).length);

      // Kumpulkan semua status unik
      const statusMap = {};
      (distributorQueueData || []).forEach(item => {
        const status = String(item.statusGudang || '').trim() || '(kosong)';
        statusMap[status] = (statusMap[status] || 0) + 1;
      });

      console.log('Status breakdown:', statusMap);

      // Filter PO yang Ready Pickup - gunakan field statusGudang
      const readyPickupPOs = (distributorQueueData || []).filter(item => {
        const statusRaw = String(item.statusGudang || '').trim();
        const statusNorm = statusRaw.toLowerCase().replace(/\s+/g, '');

        // Matches: "Ready Pickup", "Ready To Pickup", "Ready", "Pickup", "Siap", "Siap Pickup", etc.
        const isMatch = statusNorm === 'readypickup' || statusNorm === 'ready' || statusNorm === 'pickup' ||
          statusNorm.includes('readypickup') || statusNorm.includes('readytopickup') ||
          statusNorm.includes('ready') || statusNorm.includes('pickup') || statusNorm.includes('siap');

        if (isMatch) {
          console.log(`✓ Matched PO ${item.poNumber}: "${statusRaw}"`);
        }

        return isMatch;
      });

      console.log(`Found ${readyPickupPOs.length} Ready Pickup POs`);

      if (!readyPickupPOs.length) {
        // Tampilkan status yang tersedia untuk membantu user
        const uniqueStatuses = Object.keys(statusMap).filter(s => s !== '(kosong)');

        let errorMsg = '❌ Tidak ada PO dengan status Ready Pickup.\n\n';

        if (uniqueStatuses.length > 0) {
          errorMsg += '📋 Status yang tersedia:\n' + uniqueStatuses.slice(0, 10).map(s => `• ${s} (${statusMap[s]} PO)`).join('\n');
          errorMsg += '\n\n💡 Ubah kolom "Status Gudang" menjadi "Ready Pickup" atau "Siap Pickup" untuk PO yang ingin dicetak.';
        } else {
          errorMsg += '⚠️ Semua PO belum memiliki status.\n\n💡 Isi kolom "Status Gudang" dengan "Ready Pickup" atau "Siap Pickup".';
        }

        // Tampilkan di console untuk debugging
        console.error(errorMsg);

        return toast('Tidak ada PO dengan status Ready Pickup. Lihat Console (F12) untuk detail status yang tersedia.', 'error');
      }

      // Reset selection
      sjSelectedPOs = [];

      // Render list PO
      renderSJPOList(readyPickupPOs);
      updateSJSummary();

      openModal('modalPrintSuratJalan');
    }

    function renderSJPOList(pos) {
      const container = document.getElementById('sjPOList');
      if (!pos.length) {
        container.innerHTML = '<div style="text-align:center; padding:30px; color:var(--text-muted);">Tidak ada PO Ready Pickup</div>';
        return;
      }

      container.innerHTML = pos.map((item, idx) => {
        const dist = String(item.namaDistributor || '-');
        const po = String(item.poNumber || '-');
        const koli = Number(item.jumlahDus) || 0;
        const pcs = Number(item.totalPcs) || 0;
        const keterangan = String(item.note || item.keterangan || '-');

        return `
          <div style="border:1px solid var(--border-color); border-radius:8px; padding:12px; margin-bottom:8px; display:flex; align-items:center; gap:12px; background:var(--bg-panel-light);">
            <input type="checkbox" class="sjPOCheckbox" data-idx="${idx}" 
              data-po="${po}" 
              data-dist="${dist}"
              data-koli="${koli}"
              data-pcs="${pcs}"
              data-ket="${keterangan}"
              onchange="updateSJSelection()"
              style="width:18px; height:18px; cursor:pointer;">
            <div style="flex:1;">
              <div style="font-weight:700; font-size:13px; color:var(--text-main); margin-bottom:4px;">
                📦 PO: <span style="color:var(--teal);">${po}</span>
              </div>
              <div style="font-size:12px; color:var(--text-muted); display:flex; gap:12px; flex-wrap:wrap;">
                <span>🏢 ${dist}</span>
                <span>📦 <strong>${koli}</strong> Koli</span>
                <span>🔢 <strong>${pcs.toLocaleString('id-ID')}</strong> Pcs</span>
              </div>
              ${keterangan !== '-' ? `<div style="font-size:11px; color:var(--text-muted); margin-top:4px;">💬 ${keterangan}</div>` : ''}
            </div>
          </div>
        `;
      }).join('');
    }

    function sjSelectAll() {
      document.querySelectorAll('.sjPOCheckbox').forEach(cb => cb.checked = true);
      updateSJSelection();
    }

    function sjDeselectAll() {
      document.querySelectorAll('.sjPOCheckbox').forEach(cb => cb.checked = false);
      updateSJSelection();
    }

    function updateSJSelection() {
      sjSelectedPOs = [];
      document.querySelectorAll('.sjPOCheckbox:checked').forEach(cb => {
        sjSelectedPOs.push({
          po: cb.dataset.po,
          distributor: cb.dataset.dist,
          koli: Number(cb.dataset.koli) || 0,
          pcs: Number(cb.dataset.pcs) || 0,
          keterangan: cb.dataset.ket
        });
      });
      updateSJSummary();
    }

    function updateSJSummary() {
      const totalPO = sjSelectedPOs.length;
      const totalKoli = sjSelectedPOs.reduce((sum, po) => sum + po.koli, 0);
      document.getElementById('sjTotalPO').textContent = totalPO;
      document.getElementById('sjTotalKoli').textContent = totalKoli.toLocaleString('id-ID');
    }

    function printSuratJalan() {
      const tanggal = v('sjTanggal');
      const keterangan = v('sjKeterangan');

      if (!tanggal) {
        return toast('Tanggal Surat Jalan wajib diisi', 'error');
      }

      if (!sjSelectedPOs.length) {
        return toast('Pilih minimal 1 PO untuk dicetak', 'error');
      }

      // Generate HTML untuk print
      const totalPO = sjSelectedPOs.length;
      const totalKoli = sjSelectedPOs.reduce((sum, po) => sum + po.koli, 0);
      const totalPcs = sjSelectedPOs.reduce((sum, po) => sum + po.pcs, 0);

      const tanggalFormatted = formatDate(tanggal);
      const currentDateTime = new Date().toLocaleString('id-ID', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      });

      const poRows = sjSelectedPOs.map((po, idx) => `
        <tr>
          <td class="text-center">${idx + 1}</td>
          <td><strong>${po.po}</strong></td>
          <td class="text-center"><strong>${po.koli}</strong></td>
          <td class="text-right">${po.pcs.toLocaleString('id-ID')}</td>
          <td class="text-small">${po.keterangan !== '-' ? po.keterangan : ''}</td>
        </tr>
      `).join('');

      // Generate surat jalan content - FULL PAGE A4
      const suratJalanContent = `
        <div class="surat-jalan-page">
          <!-- Header dengan Logo dan Judul -->
          <div class="header-box">
            <div class="company-info">
              <h1 class="company-name">GUDANG FCL</h1>
              <p class="company-subtitle">Antrian Distributor</p>
            </div>
            <div class="doc-title">
              <h2>SURAT JALAN</h2>
              <p class="doc-date">${tanggalFormatted}</p>
            </div>
          </div>

          <!-- Info Section -->
          <div class="info-grid">
            <div class="info-item">
              <span class="info-label">Tanggal Kirim:</span>
              <span class="info-value">${tanggalFormatted}</span>
            </div>
            <div class="info-item">
              <span class="info-label">Waktu Cetak:</span>
              <span class="info-value">${currentDateTime}</span>
            </div>
            <div class="info-item">
              <span class="info-label">Total PO:</span>
              <span class="info-value"><strong>${totalPO}</strong></span>
            </div>
            <div class="info-item">
              <span class="info-label">Total Koli:</span>
              <span class="info-value"><strong>${totalKoli.toLocaleString('id-ID')}</strong></span>
            </div>
          </div>

          ${keterangan ? `<div class="keterangan-box"><strong>📝 Keterangan:</strong> ${keterangan}</div>` : ''}

          <!-- Tabel PO -->
          <table class="po-table">
            <thead>
              <tr>
                <th style="width:50px;">No</th>
                <th>PO Number</th>
                <th style="width:100px;">Koli</th>
                <th style="width:120px;">Pcs</th>
                <th style="width:200px;">Keterangan</th>
              </tr>
            </thead>
            <tbody>
              ${poRows}
            </tbody>
          </table>

          <!-- Summary Box -->
          <div class="summary-box">
            <div class="summary-item">
              <span>Total PO</span>
              <strong>${totalPO}</strong>
            </div>
            <div class="summary-item">
              <span>Total Koli</span>
              <strong>${totalKoli.toLocaleString('id-ID')}</strong>
            </div>
            <div class="summary-item">
              <span>Total Pcs</span>
              <strong>${totalPcs.toLocaleString('id-ID')}</strong>
            </div>
          </div>

          <!-- Signature Section -->
          <div class="signature-section">
            <div class="signature-col">
              <div class="sig-label">Pengirim</div>
              <div class="sig-space"></div>
              <div class="sig-line">(&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;)</div>
            </div>
            <div class="signature-col">
              <div class="sig-label">Security</div>
              <div class="sig-space"></div>
              <div class="sig-line">(&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;)</div>
            </div>
            <div class="signature-col">
              <div class="sig-label">Driver</div>
              <div class="sig-space"></div>
              <div class="sig-line">(&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;)</div>
            </div>
            <div class="signature-col">
              <div class="sig-label">Penerima</div>
              <div class="sig-space"></div>
              <div class="sig-line">(&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;)</div>
            </div>
          </div>

          <!-- Footer Note -->
          <div class="footer-note">
            Dokumen ini dicetak secara otomatis dari sistem Gudang FCL
          </div>
        </div>
      `;

      const printContent = `
        <!DOCTYPE html>
        <html>
        <head>
          <meta charset="UTF-8">
          <title>Surat Jalan - ${tanggalFormatted}</title>
          <style>
            /* ===== PRINT SETTINGS - PRESISI A4 ===== */
            @page { 
              size: A4 portrait;
              margin: 0;
            }
            
            @media print {
              body { margin: 0; padding: 0; }
              .no-print { display: none !important; }
              .surat-jalan-page { page-break-inside: avoid; }
            }
            
            * {
              margin: 0;
              padding: 0;
              box-sizing: border-box;
            }
            
            body {
              font-family: 'Arial', 'Helvetica', sans-serif;
              color: #000;
              background: #fff;
              line-height: 1.4;
            }
            
            /* ===== SURAT JALAN PAGE - FULL A4 ===== */
            .surat-jalan-page {
              width: 210mm;
              height: 297mm;
              padding: 15mm 20mm;
              position: relative;
              page-break-after: always;
            }
            
            .surat-jalan-page:last-child {
              page-break-after: auto;
            }
            
            /* ===== HEADER BOX ===== */
            .header-box {
              display: flex;
              justify-content: space-between;
              align-items: center;
              padding-bottom: 12px;
              margin-bottom: 20px;
              border-bottom: 4px double #000;
            }
            
            .company-info {
              flex: 1;
            }
            
            .company-name {
              font-size: 28px;
              font-weight: 900;
              letter-spacing: 2px;
              color: #000;
              margin-bottom: 5px;
            }
            
            .company-subtitle {
              font-size: 14px;
              color: #666;
              font-weight: 600;
            }
            
            .doc-title {
              text-align: right;
              border-left: 4px solid #000;
              padding-left: 20px;
            }
            
            .doc-title h2 {
              font-size: 32px;
              font-weight: 900;
              letter-spacing: 4px;
              margin-bottom: 5px;
            }
            
            .doc-date {
              font-size: 14px;
              font-weight: 700;
              color: #333;
            }
            
            /* ===== INFO GRID ===== */
            .info-grid {
              display: grid;
              grid-template-columns: 1fr 1fr 1fr 1fr;
              gap: 15px;
              margin-bottom: 20px;
              padding: 15px 20px;
              background: #f8f8f8;
              border: 2px solid #ddd;
              border-radius: 8px;
            }
            
            .info-item {
              font-size: 13px;
            }
            
            .info-label {
              display: block;
              color: #666;
              font-weight: 600;
              margin-bottom: 5px;
            }
            
            .info-value {
              display: block;
              font-size: 15px;
              font-weight: 700;
              color: #000;
            }
            
            /* ===== KETERANGAN BOX ===== */
            .keterangan-box {
              margin-bottom: 15px;
              padding: 12px 15px;
              background: #fffbea;
              border-left: 5px solid #f59e0b;
              font-size: 13px;
              line-height: 1.5;
            }
            
            /* ===== TABLE ===== */
            .po-table {
              width: 100%;
              border-collapse: collapse;
              margin-bottom: 20px;
              font-size: 13px;
            }
            
            .po-table thead th {
              background: #333;
              color: #fff;
              padding: 10px 12px;
              text-align: left;
              border: 1px solid #000;
              font-size: 12px;
              font-weight: 700;
              text-transform: uppercase;
            }
            
            .po-table tbody td {
              padding: 8px 12px;
              border: 1px solid #333;
              font-size: 12px;
              vertical-align: middle;
            }
            
            .text-center { text-align: center; }
            .text-right { text-align: right; }
            .text-small { font-size: 11px; color: #555; }
            
            /* ===== SUMMARY BOX ===== */
            .summary-box {
              display: flex;
              justify-content: space-around;
              padding: 15px 20px;
              background: #f0f0f0;
              border: 3px solid #000;
              border-radius: 8px;
              margin-bottom: 25px;
            }
            
            .summary-item {
              text-align: center;
              font-size: 14px;
            }
            
            .summary-item span {
              display: block;
              color: #666;
              font-weight: 600;
              margin-bottom: 5px;
            }
            
            .summary-item strong {
              display: block;
              font-size: 20px;
              font-weight: 900;
              color: #000;
            }
            
            /* ===== SIGNATURE SECTION ===== */
            .signature-section {
              display: flex;
              justify-content: space-between;
              gap: 20px;
              margin-top: 30px;
              margin-bottom: 20px;
            }
            
            .signature-col {
              flex: 1;
              text-align: center;
            }
            
            .sig-label {
              font-size: 13px;
              font-weight: 700;
              margin-bottom: 5px;
              text-transform: uppercase;
            }
            
            .sig-space {
              height: 60mm;
            }
            
            .sig-line {
              border-top: 2px solid #000;
              padding-top: 5px;
              font-size: 11px;
              color: #666;
            }
            
            /* ===== FOOTER NOTE ===== */
            .footer-note {
              position: absolute;
              bottom: 15mm;
              left: 20mm;
              right: 20mm;
              text-align: center;
              font-size: 10px;
              color: #999;
              border-top: 1px solid #ddd;
              padding-top: 10px;
            }
            
            /* ===== PRINT INFO (NO PRINT) ===== */
            .print-info {
              text-align: center;
              padding: 20px;
              background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
              color: white;
              margin: 20px;
              border-radius: 12px;
              box-shadow: 0 10px 30px rgba(0,0,0,0.2);
            }
            
            .print-info h2 {
              margin: 0 0 15px 0;
              font-size: 28px;
              font-weight: 900;
            }
            
            .print-info p {
              margin: 8px 0;
              font-size: 16px;
              opacity: 0.95;
            }
            
            .print-btn {
              margin-top: 20px;
              padding: 15px 40px;
              font-size: 18px;
              font-weight: 700;
              background: white;
              color: #667eea;
              border: none;
              border-radius: 8px;
              cursor: pointer;
              box-shadow: 0 4px 15px rgba(0,0,0,0.2);
              transition: all 0.3s ease;
            }
            
            .print-btn:hover {
              transform: translateY(-2px);
              box-shadow: 0 6px 20px rgba(0,0,0,0.3);
            }
          </style>
        </head>
        <body>
          <div class="no-print print-info">
            <h2>🖨️ Siap Print Surat Jalan</h2>
            <p><strong>Format:</strong> 1 Halaman A4 = 1 Surat Jalan</p>
            <p><strong>Total PO:</strong> ${totalPO} PO | <strong>Total Koli:</strong> ${totalKoli.toLocaleString('id-ID')} | <strong>Total Pcs:</strong> ${totalPcs.toLocaleString('id-ID')}</p>
            <p style="font-size:14px; opacity:0.9;">📄 Surat jalan akan dicetak dalam 1 halaman penuh A4</p>
            <button class="print-btn" onclick="window.print()">🖨️ Print Sekarang</button>
          </div>
          
          <!-- Surat Jalan -->
          ${suratJalanContent}
        </body>
        </html>
      `;

      // Open print window
      const printWindow = window.open('', '_blank', 'width=900,height=700');
      printWindow.document.write(printContent);
      printWindow.document.close();
    }

    // ============================================================

    function toLocalDateValue(date) {
      if (!date) return '';
      const dt = new Date(date);
      if (isNaN(dt)) return '';
      return dt.toISOString().slice(0, 10);
    }

    function toLocalDateTimeValue(date) {
      if (!date) return '';
      const dt = new Date(date);
      if (isNaN(dt)) return '';
      const offsetMs = dt.getTimezoneOffset() * 60000;
      return new Date(dt.getTime() - offsetMs).toISOString().slice(0, 16);
    }

    function resetDistributorQueueForm() {
      if (!canEditDistributorQueue()) return;
      [
        'dqEditRow', 'dqNo', 'dqPicSales', 'dqNamaDistributor', 'dqAlamat', 'dqNoHp', 'dqPoNumber', 'dqNoMabang',
        'dqMetodePengiriman', 'dqOngkirDibayarOleh', 'dqNote', 'dqStatusGudang', 'dqJumlahDus',
        'dqTotalPcs', 'dqPacker', 'dqValidation', 'dqStatusMabang', 'dqGdrive', 'dqDeliveryBill',
        'dqNomorResi', 'dqBuktiPengiriman', 'dqTanggalSelesaiPacking', 'dqShipDate'
      ].forEach(id => setVal(id, ''));
      setVal('dqOrderQueueTime', toLocalDateValue(new Date()));
      setVal('dqTimeWib', toLocalDateTimeValue(new Date()));
      const mode = document.getElementById('dqFormMode');
      if (mode) mode.textContent = 'Mode Baru';
    }

    function fillDistributorQueueForm(item) {
      if (!canEditDistributorQueue()) return;
      setVal('dqEditRow', item.rowNumber || '');
      setVal('dqNo', item.no || '');
      setVal('dqOrderQueueTime', item.orderQueueTime || '');
      setVal('dqPicSales', item.picSales || '');
      setVal('dqNamaDistributor', item.namaDistributor || '');
      setVal('dqAlamat', item.alamat || '');
      setVal('dqNoHp', item.noHp || '');
      setVal('dqPoNumber', item.poNumber || '');
      setVal('dqNoMabang', item.noMabang || '');
      setVal('dqMetodePengiriman', item.metodePengiriman || '');
      setVal('dqOngkirDibayarOleh', item.ongkirDibayarOleh || '');
      setVal('dqNote', item.note || '');
      setVal('dqTimeWib', item.timeWib ? item.timeWib.replace(' ', 'T').slice(0, 16) : '');
      setVal('dqStatusGudang', item.statusGudang || '');
      setVal('dqJumlahDus', item.jumlahDus || '');
      setVal('dqTotalPcs', item.totalPcs || '');
      setVal('dqPacker', item.packer || '');
      setVal('dqValidation', item.validation || '');
      setVal('dqTanggalSelesaiPacking', item.tanggalSelesaiPacking || '');
      setVal('dqShipDate', item.shipDate || '');
      setVal('dqStatusMabang', item.statusMabang || '');
      setVal('dqGdrive', item.gdrive || '');
      setVal('dqDeliveryBill', item.deliveryBill || '');
      setVal('dqNomorResi', item.nomorResi || '');
      setVal('dqBuktiPengiriman', item.buktiPengiriman || '');
      const mode = document.getElementById('dqFormMode');
      if (mode) mode.textContent = 'Mode Edit';
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    function editDistributorQueue(rowNumber) {
      if (!canEditDistributorQueue()) return toast('Anda tidak punya hak tambah/edit Antrian Distributor', 'error');
      const item = distributorQueueData.find(d => Number(d.rowNumber) === Number(rowNumber));
      if (!item) return toast('Data antrian tidak ditemukan', 'error');
      fillDistributorQueueForm(item);
    }

    function renderDistributorQueueDashboard(dashboard) {
      applyDistributorQueuePermissions();
      const dash = dashboard || {};
      setVal('dqStatTotal', dash.total || 0);
      setVal('dqStatSelesai', dash.selesai || 0);
      // JANGAN update stat cards ini - sudah diupdate oleh renderDistributorQueueStatusCards()
      // setVal('dqStatBelumSelesai', dash.belumSelesai || 0);
      // setVal('dqStatBelumDikerjakan', dash.belumDikerjakan || 0);
      // setVal('dqStatReadyPickup', dash.totalReadyToPickup || 0);
      setVal('dqStatLate', dash.late || 0);
      setVal('dqStatTerkirim', dash.totalTerkirim || 0);
      setVal('dqStatPoHariIni', dash.poKeluarHariIni || 0);
      setVal('dqStatPoMingguIni', dash.poKeluarMingguIni || 0);
      setVal('dqStatPoBulanIni', dash.poKeluarBulanIni || 0);
      setVal('dqLateBadge', `${dash.late || 0} Late`);

      // Update tanggal untuk card detail PO
      const today = new Date();
      const bulanIndo = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
      const hariIndo = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];

      // PO Hari Ini
      const hariIni = `${hariIndo[today.getDay()]}, ${today.getDate()} ${bulanIndo[today.getMonth()]} ${today.getFullYear()}`;
      setVal('dqPoHariIniDate', hariIni);

      // PO Minggu Ini (Senin - Minggu)
      const firstDayOfWeek = new Date(today);
      firstDayOfWeek.setDate(today.getDate() - today.getDay() + 1); // Senin
      const lastDayOfWeek = new Date(firstDayOfWeek);
      lastDayOfWeek.setDate(firstDayOfWeek.getDate() + 6); // Minggu
      const mingguIni = `${firstDayOfWeek.getDate()} - ${lastDayOfWeek.getDate()} ${bulanIndo[lastDayOfWeek.getMonth()]} ${lastDayOfWeek.getFullYear()}`;
      setVal('dqPoMingguIniDate', mingguIni);

      // PO Bulan Ini
      const bulanIni = `${bulanIndo[today.getMonth()]} ${today.getFullYear()}`;
      setVal('dqPoBulanIniDate', bulanIni);

      renderDistributorQueueChart(dash);

      const lateBody = document.getElementById('dqLateTable');
      if (!lateBody) return;
      if (!canViewLateShipmentDashboard()) {
        lateBody.innerHTML = '<tr><td colspan="8" class="empty-state">Anda tidak punya hak melihat Dashboard Late Shipment</td></tr>';
        return;
      }

      const items = (dash.lateItems || []).slice(0, 30);
      if (!items.length) {
        lateBody.innerHTML = '<tr><td colspan="9" class="empty-state">Belum ada shipment yang melewati SLA</td></tr>';
        return;
      }

      lateBody.innerHTML = items.map(item => {
        let statusBadge = '';
        const status = item.catatanLateStatus || '';
        if (status.includes('Pending')) {
          statusBadge = `<span class="badge bg-warning">${status}</span>`;
        } else if (status === 'Disetujui') {
          statusBadge = `<span class="badge bg-success">Disetujui</span>`;
        } else if (status === 'Ditolak') {
          statusBadge = `<span class="badge bg-danger">Ditolak</span>`;
        } else {
          statusBadge = `<span class="badge bg-secondary">Belum ada</span>`;
        }

        return `
        <tr>
          <td><strong>${queueCell(item.poNumber)}</strong></td>
          <td>${queueCell(item.namaDistributor)}</td>
          <td>${queueCell(item.orderQueueTime)}</td>
          <td>${queueCell(item.tanggalSelesaiPacking || '-')}</td>
          <td>${queueSlaBadge(item.sla)}</td>
          <td><strong style="color:var(--red);">${item.sla?.lateDays || 0} hari</strong></td>
          <td><small>${queueCell(item.catatanLate || '-')}</small></td>
          <td>${statusBadge}</td>
          <td>${canUpdateLateNote() ? `<button class="btn btn-ghost btn-sm" onclick="openEditLateCatatanModal(${item.rowNumber})">📝 Edit Keterangan</button>` : '-'}</td>
        </tr>
      `;
      }).join('');
    }

    function renderDistributorQueueTable() {
      const body = document.getElementById('dqTableBody');
      if (!body) return;

      if (!distributorQueueData.length) {
        body.innerHTML = '<tr><td colspan="25" class="empty-state">Belum ada data antrian distributor</td></tr>';
        return;
      }

      body.innerHTML = distributorQueueData.map(item => {
        // Badge warna untuk Source Sheet
        let sheetBadge = '';
        const sourceSheet = item.sourceSheet || 'Antrian Distributor';
        if (sourceSheet === 'ANTRIAN FOCALSKIN') {
          sheetBadge = '<span style="background:#10b98122;color:#10b981;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:700;">FOCALSKIN</span>';
        } else if (sourceSheet === 'ANTRIAN MISTINE') {
          sheetBadge = '<span style="background:#0ea5e922;color:#0ea5e9;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:700;">MISTINE</span>';
        } else if (sourceSheet === 'ANTRIAN SBY') {
          sheetBadge = '<span style="background:#f59e0b22;color:#f59e0b;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:700;">SBY</span>';
        } else {
          sheetBadge = '<span style="background:#94a3b822;color:#94a3b8;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:700;">MAIN</span>';
        }

        return `
        <tr>
          <td>${sheetBadge}</td>
          <td>${queueCell(item.no)}</td>
          <td>${queueCell(item.orderQueueTime)}</td>
          <td>${queueCell(item.picSales)}</td>
          <td><strong>${queueCell(item.namaDistributor)}</strong></td>
          <td>${queueCell(item.alamat)}</td>
          <td>${queueCell(item.noHp)}</td>
          <td><strong>${queueCell(item.poNumber)}</strong></td>
          <td>${queueCell(item.noMabang)}</td>
          <td>${queueCell(item.metodePengiriman)}</td>
          <td>${queueCell(item.ongkirDibayarOleh)}</td>
          <td>${queueCell(item.note)}</td>
          <td>${queueCell(item.timeWib)}</td>
          <td>${queueCell(item.statusGudang)}</td>
          <td>${queueCell(item.jumlahDus)}</td>
          <td>${queueCell(item.totalPcs)}</td>
          <td>${queueCell(item.packer)}</td>
          <td>${queueCell(item.validation)}</td>
          <td>${queueCell(item.tanggalSelesaiPacking)}</td>
          <td>${queueCell(item.shipDate)}</td>
          <td>${queueCell(item.statusMabang)}</td>
          <td>${queueLinkHtml(item.gdrive, 'GDrive')}</td>
          <td>${queueLinkHtml(item.deliveryBill, 'Delivery Bill')}</td>
          <td>${queueCell(item.nomorResi)}</td>
          <td>${queueLinkHtml(item.buktiPengiriman, 'Bukti')}</td>
        </tr>
      `;
      }).join('');
    }

    function loadDistributorQueue() {
      if (!canViewDistributorQueue()) return toast('Anda tidak punya hak melihat Antrian Distributor', 'error');

      // Set timeout 60 detik untuk data besar (30K+ rows)
      let loadingTimeout = setTimeout(function () {
        toast('⚠️ Loading memakan waktu lama - data sangat besar. Menampilkan data yang sudah ter-load...', 'warning');
        // Jangan tampilkan empty state, biarkan data yang sudah ada
      }, 60000); // 60 detik timeout

      // ── Phase 1: Fast dashboard (instant stats + chart) ──
      google.script.run
        .withSuccessHandler(function (fastRes) {
          if (fastRes && fastRes.success && fastRes.dashboard) {
            renderDistributorQueueDashboard(fastRes.dashboard);
            toast('✅ Dashboard loaded! Loading detail data...', 'info');
          }
        })
        .withFailureHandler(function (err) {
          console.error('Fast dashboard failed:', err);
        })
        .getDistributorQueueDashboardFast();

      // ── Phase 2: Full data (table + SLA + late items) ──
      google.script.run
        .withSuccessHandler(function (res) {
          clearTimeout(loadingTimeout); // Clear timeout jika berhasil

          if (!res || !res.success) {
            toast(res ? res.message : 'Gagal memuat antrian distributor', 'error');
            // Tampilkan empty state
            renderDistributorQueueStatusCards([]);
            return;
          }

          distributorQueueData = res.data || [];
          console.log('✅ Data loaded successfully:', distributorQueueData.length, 'items');

          // PENTING: Render status cards DULU sebelum dashboard
          // Agar stat cards (Belum Dikerjakan, Belum Selesai, Ready Pickup) akurat
          renderDistributorQueueStatusCards(distributorQueueData);
          renderDistributorQueueDashboard(res.dashboard || {});
          renderDistributorQueueTable();

          toast('✅ Data berhasil dimuat: ' + distributorQueueData.length + ' items (dari 30K+ total)', 'success');

          const orderField = document.getElementById('dqOrderQueueTime');
          if (orderField && !orderField.value) resetDistributorQueueForm();
        })
        .withFailureHandler(function (err) {
          clearTimeout(loadingTimeout); // Clear timeout jika error
          console.error('❌ Load distributor queue failed:', err);
          toast('❌ Gagal memuat antrian distributor: ' + err, 'error');
          // Tampilkan empty state
          renderDistributorQueueStatusCards([]);
        })
        .getDistributorQueueData();
    }

    function submitDistributorQueue() {
      if (!canEditDistributorQueue()) return toast('Anda tidak punya hak tambah/edit Antrian Distributor', 'error');
      const payload = {
        rowNumber: v('dqEditRow'),
        no: v('dqNo'),
        orderQueueTime: v('dqOrderQueueTime'),
        picSales: v('dqPicSales'),
        namaDistributor: v('dqNamaDistributor'),
        alamat: v('dqAlamat'),
        noHp: v('dqNoHp'),
        poNumber: v('dqPoNumber'),
        noMabang: v('dqNoMabang'),
        metodePengiriman: v('dqMetodePengiriman'),
        ongkirDibayarOleh: v('dqOngkirDibayarOleh'),
        note: v('dqNote'),
        timeWib: v('dqTimeWib'),
        statusGudang: v('dqStatusGudang'),
        jumlahDus: v('dqJumlahDus'),
        totalPcs: v('dqTotalPcs'),
        packer: v('dqPacker'),
        validation: v('dqValidation'),
        tanggalSelesaiPacking: v('dqTanggalSelesaiPacking'),
        shipDate: v('dqShipDate'),
        statusMabang: v('dqStatusMabang'),
        gdrive: v('dqGdrive'),
        deliveryBill: v('dqDeliveryBill'),
        nomorResi: v('dqNomorResi'),
        buktiPengiriman: v('dqBuktiPengiriman')
      };

      if (!payload.orderQueueTime || !payload.namaDistributor || !payload.poNumber) {
        return toast('Order queue time, Nama Distributor, dan PO number wajib diisi', 'error');
      }

      const btn = document.getElementById('btnSaveDistributorQueue');
      btn.disabled = true;
      btn.textContent = '⏳ Menyimpan...';

      google.script.run.withSuccessHandler(res => {
        btn.disabled = false;
        btn.textContent = '💾 Simpan ke Spreadsheet';
        if (!res || !res.success) return toast(res ? res.message : 'Gagal menyimpan antrian distributor', 'error');
        toast(res.message || 'Antrian distributor berhasil disimpan', 'success');
        resetDistributorQueueForm();
        loadDistributorQueue();
      }).withFailureHandler(err => {
        btn.disabled = false;
        btn.textContent = '💾 Simpan ke Spreadsheet';
        toast('Gagal menyimpan antrian distributor: ' + err, 'error');
      }).saveDistributorQueue(JSON.stringify(payload), currentUser ? currentUser.username : '');
    }

    function downloadOrderTemplate() {
      const headers = [["Tanggal", "Kategori", "Pelanggan", "Alamat", "SKU", "Qty", "Lokasi", "No Order (Marketplace)", "No Resi (Marketplace)"]];
      const data = [["2024-03-31", "Marketplace", "Nama Pembeli", "Shopee / Tokopedia", "SKU001", 10, "G01-A1", "ORD-CUSTOM-001", "JP123456789"]];
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(headers.concat(data));
      XLSX.utils.book_append_sheet(wb, ws, "Template Order");
      XLSX.writeFile(wb, "Template_Order_FCL.xlsx");
    }

    // --- APPROVAL DASHBOARD FUNCTIONS ---
    let _approvalData = { lembur: [], ijin: [], asset: [], stockOpname: [], assetOpname: [] };
    let _activeApprovalTab = 'lembur';

    function switchApprovalTab(tab) {
      _activeApprovalTab = tab;
      document.querySelectorAll('.kar-tab').forEach(b => b.classList.remove('active'));
      const activeBtn = document.getElementById('tabAppr' + tab.charAt(0).toUpperCase() + tab.slice(1));
      if (activeBtn) activeBtn.classList.add('active');

      document.querySelectorAll('.approval-panel').forEach(p => p.style.display = 'none');
      const activePanel = document.getElementById('panelAppr' + (tab === 'stockOpname' ? 'StockOpname' : tab.charAt(0).toUpperCase() + tab.slice(1)));
      if (activePanel) activePanel.style.display = 'block';

      renderApprovalDashboardData();
    }

    function loadApprovalDashboard() {
      toast('⏳ Memuat data approval...', 'info');
      google.script.run.withSuccessHandler(res => {
        if (res && res.success) {
          _approvalData = res.data;
          updateApprovalStats();
          renderApprovalDashboardData();
          loadOpnameReportMeta();
        } else {
          toast('Gagal memuat data approval: ' + (res?.message || 'Unknown error'), 'error');
        }
      }).getApprovalDashboardData(currentUser.role);
    }

    function updateApprovalStats() {
      const getCount = (arr) => (arr || []).length;
      setVal('apprStatLembur', getCount(_approvalData.lembur));
      setVal('apprStatIjin', getCount(_approvalData.ijin));
      setVal('apprStatAsset', getCount(_approvalData.asset));
      setVal('apprStatSO', getCount(_approvalData.stockOpname) + getCount(_approvalData.assetOpname));
    }

    function renderApprovalDashboardData() {
      const tab = _activeApprovalTab;
      const data = _approvalData[tab] || [];
      const bodyId = 'bodyAppr' + (tab === 'stockOpname' ? 'SO' : tab.charAt(0).toUpperCase() + tab.slice(1));
      const body = document.getElementById(bodyId);
      if (!body) return;

      if (!data.length) {
        body.innerHTML = `<tr><td colspan="10" class="empty-state">Tidak ada data pending untuk ${tab}</td></tr>`;
        return;
      }

      body.innerHTML = data.map(item => {
        let cells = '';
        const id = item.id || item.rowNumber || '';
        const status = item.status || 'Pending';

        if (tab === 'lembur') {
          const absenInfo = `<div style="font-size:10px; color:var(--gray);">IN: ${item.inTime || '-'}</div><div style="font-size:10px; color:var(--gray);">OUT: ${item.outTime || '-'}</div>`;
          cells = `<td>${item.tanggal}</td><td>${item.nama}</td><td>${item.divisi}</td><td>${item.jumlahJam}</td><td>${absenInfo}</td><td>${status}</td>`;
        } else if (tab === 'ijin') {
          cells = `<td>${item.tanggal}</td><td>${item.nama}</td><td>${item.jenis}</td><td>${item.keterangan || '-'}</td><td>${status}</td>`;
        } else if (tab === 'asset') {
          cells = `<td>${item.tanggal}</td><td>${item.nama}</td><td>${item.jenisAsset}</td><td>${item.estimasiHarga}</td><td>${status}</td>`;
        } else if (tab === 'stockOpname') {
          const pic = item.pic || item.createdBy || '';
          const areaItem = item.area || (item.sku ? `${item.sku} - ${item.nama}` : '');
          cells = `<td>${item.tanggal}</td><td>${pic}</td><td>${areaItem}</td><td>${item.selisih || '-'}</td><td>${status}</td>`;
        } else if (tab === 'assetOpname') {
          cells = `<td>${item.tanggal}</td><td>${item.createdBy}</td><td>${item.divisi}</td><td>${item.terscan} / ${item.totalAsset}</td><td>${status}</td>`;
        }

        // Penentuan Label Tombol Berdasarkan Status
        let apprLabel = '✔️ Setuju';
        if (status === 'Pending Team Leader') apprLabel = '✔️ Setuju TL';
        else if (status === 'Pending Vice Supervisor') apprLabel = '✔️ Setuju Vice SPV';
        else if (status === 'Pending Supervisor') apprLabel = '✔️ Setuju SPV';
        else if (status === 'Pending HR') apprLabel = '✔️ Setuju HR';

        const btnAppr = `<button class="btn btn-success btn-sm me-1" onclick="processDashboardAppr('${tab}', '${id}', 'Approve', '${item.nama || item.pemohon || item.pic || item.auditor || item.createdBy}', '${item.tanggal}')">${apprLabel}</button>`;
        const btnRej = `<button class="btn btn-danger btn-sm" onclick="processDashboardAppr('${tab}', '${id}', 'Reject', '${item.nama || item.pemohon || item.pic || item.auditor || item.createdBy}', '${item.tanggal}')">✖️ Tolak</button>`;

        return `<tr>${cells}<td>${btnAppr}${btnRej}</td></tr>`;
      }).join('');
    }

    function processDashboardAppr(tipe, id, action, pemohon, tanggal) {
      // Tanpa Pop-Up Konfirmasi Sesuai Permintaan
      const btn = event ? event.target : null;

      // Add click animation immediately
      if (btn && btn.tagName === 'BUTTON') {
        btn.classList.add('btn-appr-animate');
        btn.disabled = true;
      }

      toast(`⏳ Memproses ${action}...`, 'info');
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          // Add success/reject animation
          if (btn && btn.tagName === 'BUTTON') {
            btn.classList.remove('btn-appr-animate');
            if (action === 'Approve') {
              btn.classList.add('btn-appr-success');
            } else {
              btn.classList.add('btn-appr-reject');
            }
          }
          toast(`✅ Berhasil ${action === 'Approve' ? 'disetujui' : 'ditolak'}`, 'success');
          setTimeout(() => {
            loadApprovalDashboard(); // Reload data
          }, 500);
        } else {
          if (btn && btn.tagName === 'BUTTON') {
            btn.classList.remove('btn-appr-animate');
            btn.disabled = false;
          }
          toast('Gagal: ' + res.message, 'error');
        }
      }).processApprovalStatus(tipe, id, action, currentUser.nama, currentUser.role, '', pemohon, tanggal);
    }

    function doExportOrder() {
      google.script.run.withSuccessHandler(res => {
        if (!res.success || !res.data.length) return toast('Tidak ada data untuk diekspor', 'info');

        let html = `
          <table border="1">
            <thead>
              <tr style="background:#f4f4f4">
                <th>No. Order</th>
                <th>Tanggal</th>
                <th>Pelanggan</th>
                <th>Alamat</th>
                <th>SKU</th>
                <th>Barang</th>
                <th>Qty</th>
                <th>Satuan</th>
                <th>Status</th>
                <th>Keterangan</th>
              </tr>
            </thead>
            <tbody>`;

        res.data.forEach(d => {
          // Flattening order items for Excel readability
          const rowSpan = d.items && d.items.length > 0 ? d.items.length : 1;
          if (d.items && d.items.length > 0) {
            d.items.forEach((itm, idx) => {
              html += `<tr>`;
              if (idx === 0) {
                html += `
                  <td rowspan="${rowSpan}">${d.noOrder}</td>
                  <td rowspan="${rowSpan}">${formatDate(d.tanggal)}</td>
                  <td rowspan="${rowSpan}">${d.pelanggan}</td>
                  <td rowspan="${rowSpan}">${d.alamat || '-'}</td>`;
              }
              html += `
                <td>${itm.sku}</td>
                <td>${itm.nama}</td>
                <td>${itm.qty}</td>
                <td>${itm.satuan}</td>`;
              if (idx === 0) {
                html += `
                  <td rowspan="${rowSpan}">${d.status}</td>
                  <td rowspan="${rowSpan}">${d.keterangan || '-'}</td>`;
              }
              html += `</tr>`;
            });
          } else {
            html += `<tr><td>${d.noOrder}</td><td>${formatDate(d.tanggal)}</td><td>${d.pelanggan}</td><td>${d.alamat || '-'}</td><td colspan="4">-</td><td>${d.status}</td><td>${d.keterangan || '-'}</td></tr>`;
          }
        });

        html += '</tbody></table>';
        exportToExcel(html, 'Data_Orderan_FCL.xls');
      }).getOrdersWithDetails(); // Calling a slightly different method to get items flatten if possible, or handle in JS
    }

    function exportToExcel(html, filename) {
      const blob = new Blob(['\ufeff', html], { type: 'application/vnd.ms-excel' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }

    let importDraftData = [];
    function handleImportOrder(input) {
      const f = input.files[0]; if (!f) return;
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const firstSheet = workbook.SheetNames[0];
        const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { header: 1 });

        if (sheetData.length < 2) return toast('File kosong!', 'error');

        importDraftData = [];
        for (let i = 1; i < sheetData.length; i++) {
          const row = sheetData[i];
          if (!row[2] || !row[4]) continue; // Skip if pelanggan or SKU empty

          importDraftData.push({
            tanggal: row[0] || new Date().toISOString().split('T')[0],
            kategori: row[1] || 'Distributor',
            pelanggan: row[2],
            alamat: row[3] || '',
            sku: row[4],
            qty: parseFloat(row[5]) || 0,
            lokasiTujuan: row[6] || '',
            customNoOrder: row[7] || '',
            noResi: row[8] || '',
            stockId: '',
            batch: '',
            expDate: ''
          });
        }
        renderImportPreview();
      };
      reader.readAsArrayBuffer(f);
      input.value = '';
    }

    function renderImportPreview() {
      const tb = document.getElementById('tblImportPreview');
      tb.innerHTML = '';
      if (!importDraftData.length) return toast('Tidak ada data valid', 'error');

      importDraftData.forEach((d, idx) => {
        const matches = stockData.filter(s => s.sku === d.sku);
        let options = `<option value="">-- Pilih Lokasi --</option>`;
        let autoSelectedId = '';

        matches.forEach(m => {
          const selected = (m.lokasi === d.lokasiTujuan) ? 'selected' : '';
          if (selected) autoSelectedId = m.id;
          options += `<option value="${m.id}" ${selected}>${m.lokasi} (Stok: ${m.stok})</option>`;
        });

        const statusStyle = autoSelectedId ? '' : 'border:2px solid var(--red);';
        const resiInfo = (d.customNoOrder || d.noResi) ?
          `<div style="font-size:11px; line-height:1.2;">
            <div>O: <span style="color:var(--accent)">${d.customNoOrder || '-'}</span></div>
            <div>R: <span style="color:var(--teal)">${d.noResi || '-'}</span></div>
          </div>` : '<small class="text-muted">-</small>';

        tb.innerHTML += `
          <tr id="import-row-${idx}">
            <td>${d.tanggal}</td>
            <td><span class="badge-tb">${d.kategori}</span></td>
            <td><strong>${d.pelanggan}</strong></td>
            <td><code style="color:var(--teal)">${d.sku}</code></td>
            <td><strong>${d.qty}</strong></td>
            <td>
              <select class="form-select form-select-sm" style="${statusStyle}" onchange="syncImportRowData(${idx}, this)">
                ${options}
              </select>
            </td>
            <td>${resiInfo}</td>
            <td id="import-row-info-${idx}">
              <small class="text-muted">-</small>
            </td>
          </tr>`;

        if (autoSelectedId) {
          setTimeout(() => {
            const select = tb.querySelector(`#import-row-${idx} select`);
            if (select) syncImportRowData(idx, select);
          }, 100);
        }
      });

      document.getElementById('importCountDisplay').textContent = `${importDraftData.length} Baris Terdeteksi`;
      openModal('modalOrderImportPreview');
    }


    function syncImportRowData(idx, select) {
      const sId = select.value;
      const rowInfo = document.getElementById(`import-row-info-${idx}`);
      const row = document.getElementById(`import-row-${idx}`);

      if (!sId) {
        importDraftData[idx].stockId = '';
        importDraftData[idx].batch = '';
        importDraftData[idx].expDate = '';
        rowInfo.innerHTML = '<small class="text-danger">Lokasi Belum Dipilih</small>';
        select.style.border = '2px solid var(--red)';
        return;
      }

      const match = stockData.find(s => s.id === sId);
      if (match) {
        importDraftData[idx].stockId = match.id;
        importDraftData[idx].batch = match.batch || '-';
        importDraftData[idx].expDate = match.expDate || '-';
        rowInfo.innerHTML = `
          <div style="font-size:11px;">
            <div>B: <strong>${match.batch || '-'}</strong></div>
            <div style="color:var(--accent)">E: ${match.expDate || '-'}</div>
          </div>`;
        select.style.border = '1px solid var(--border-color)';
      }
    }

    function submitImportOrders() {
      const incomplete = importDraftData.some(d => !d.stockId);
      if (incomplete) return toast('Lengkapi semua lokasi pengambilan!', 'error');

      // Group by Order (Pelanggan + Tanggal + Kategori + Resi)
      const groups = {};
      importDraftData.forEach(d => {
        const key = `${d.pelanggan}_${d.tanggal}_${d.kategori}_${d.customNoOrder}_${d.noResi}`;
        if (!groups[key]) {
          groups[key] = {
            tanggal: d.tanggal,
            pelanggan: d.pelanggan,
            alamat: d.alamat,
            kategori: d.kategori,
            customNoOrder: d.customNoOrder,
            noResi: d.noResi,
            items: []
          };
        }
        groups[key].items.push({
          stockId: d.stockId,
          sku: d.sku,
          qty: d.qty,
          batch: d.batch,
          expDate: d.expDate
        });
      });

      const finalOrders = Object.values(groups);
      const btn = document.getElementById('btnConfirmImport');
      btn.disabled = true; btn.textContent = '⏳ Menyimpan...';

      google.script.run.withSuccessHandler(res => {
        btn.disabled = false; btn.textContent = '🚀 Konfirmasi & Simpan Semua';
        if (res.success) {
          toast(`✅ ${res.count} Order berhasil diimport!`, 'success');
          closeModal('modalOrderImportPreview');
          loadOrder();
          loadStock();
        } else toast(res.message, 'error');
      }).importOrdersBulk(JSON.stringify(finalOrders), currentUser.username);
    }



    function handleImportInbound(input) {
      const f = input.files[0]; if (!f) return;
      toast('Sedang mengimpor...', 'info');
      const r = new FileReader(); r.onload = e => {
        const t = e.target.result;
        const lines = t.split(/\r?\n/);
        const inbs = []; let cur = {};
        for (let i = 1; i < lines.length; i++) {
          let line = lines[i].trim(); if (!line) continue;
          const l = line.split(/[;,]/).map(x => x.replace(/^["']|["']$/g, '').trim());
          if (l.length < 4) continue;
          const [tgl, sup, ket, sku, qty, batch, exp] = l;
          if (sup && sup !== cur.supplier) {
            if (cur.supplier) inbs.push(cur);
            cur = { tanggal: tgl || new Date().toISOString().split('T')[0], supplier: sup, keterangan: ket || 'Import Excel', items: [] };
          }
          if (sku) cur.items.push({ sku: sku, qty: parseFloat(qty) || 0, batch: batch || '', expDate: exp || '' });
        }
        if (cur.supplier) inbs.push(cur);
        if (!inbs.length) return toast('Format salah! Gunakan: Tgl, Supplier, Ket, SKU, Qty, Batch, Exp', 'error');
        google.script.run.withSuccessHandler(res => {
          if (res.success) { toast(res.count + ' Inbound berhasil diimport ✅'); loadInbound(); loadStock(); }
          else toast(res.message, 'error');
        }).importInboundBulk(JSON.stringify(inbs), currentUser.username);
      }; r.readAsText(f); input.value = '';
    }

    function handleImportRetur(input) {
      const f = input.files[0]; if (!f) return;
      toast('Sedang mengimpor...', 'info');
      const r = new FileReader(); r.onload = e => {
        const t = e.target.result;
        const lines = t.split(/\r?\n/);
        const rets = []; let cur = {};
        for (let i = 1; i < lines.length; i++) {
          let line = lines[i].trim(); if (!line) continue;
          const l = line.split(/[;,]/).map(x => x.replace(/^["']|["']$/g, '').trim());
          if (l.length < 4) continue;
          const [tgl, sumber, ala, ket, sku, qty, batch, exp] = l;
          if (sumber && sumber !== cur.sumber) {
            if (cur.sumber) rets.push(cur);
            cur = { tanggal: tgl || new Date().toISOString().split('T')[0], sumber: sumber, alasan: ala || 'Retur', keterangan: ket || 'Import Excel', items: [] };
          }
          if (sku) cur.items.push({ sku: sku, qty: parseFloat(qty) || 0, batch: batch || '', expDate: exp || '' });
        }
        if (cur.sumber) rets.push(cur);
        if (!rets.length) return toast('Format salah! Gunakan: Tgl, Sumber, Alasan, Ket, SKU, Qty, Batch, Exp', 'error');
        google.script.run.withSuccessHandler(res => {
          if (res.success) { toast(res.count + ' Retur berhasil diimport ✅'); loadRetur(); loadStock(); }
          else toast(res.message, 'error');
        }).importReturBulk(JSON.stringify(rets), currentUser.username);
      }; r.readAsText(f); input.value = '';
    }

    function downloadOrderXlsxTemplate() {
      const headers = [
        ["Tanggal", "Kategori", "Pelanggan", "Alamat", "SKU", "Qty", "Lokasi Pengambilan", "No Order Marketplace", "No Resi Marketplace"]
      ];
      const data = [
        ["2024-04-01", "Marketplace", "Shopee Customer", "Jl. Contoh No. 123", "SKU-001", 5, "RAK-A1", "240401SKU01", "RESI123456789"]
      ];
      const worksheet = XLSX.utils.aoa_to_sheet(headers.concat(data));
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Template Order");
      XLSX.writeFile(workbook, "Template_Impor_Order_FCL.xlsx");
    }

    function downloadInboundTemplate() {
      const headers = "Tanggal;Supplier;Keterangan;SKU;Qty;Batch;Exp";
      const sample = "\n2024-03-31;Vendor Utama;Masok Barang Baru;SKU-001;100;BATCH-XY;2026-12-31";
      const csvContent = "data:text/csv;charset=utf-8,sep=;\n" + headers + sample;
      const encodedUri = encodeURI(csvContent);
      const link = document.createElement("a");
      link.setAttribute("href", encodedUri);
      link.setAttribute("download", "Template_Impor_Inbound_FCL.csv");
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }

    function downloadReturTemplate() {
      const headers = "Tanggal;Sumber;Alasan;Keterangan;SKU;Qty;Batch;Exp";
      const sample = "\n2024-03-31;Customer A;Salah Ukuran;Retur Penjualan;SKU-001;2;BATCH-XY;2026-12-31";
      const csvContent = "data:text/csv;charset=utf-8,sep=;\n" + headers + sample;
      const encodedUri = encodeURI(csvContent);
      const link = document.createElement("a");
      link.setAttribute("href", encodedUri);
      link.setAttribute("download", "Template_Impor_Retur_FCL.csv");
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }

    function loadAnalisis() {
      google.script.run.withSuccessHandler(res => {
        if (!res.success) return;
        const tW = document.getElementById('tblAnalisisMinggu'), tM = document.getElementById('tblAnalisisBulan'), tF = document.getElementById('tblAnalisisFull');
        tW.innerHTML = ''; tM.innerHTML = ''; tF.innerHTML = '';
        const wD = [...res.data].sort((a, b) => b.minggu - a.minggu).slice(0, 5); const mD = [...res.data].sort((a, b) => b.bulan - a.bulan).slice(0, 5);
        wD.forEach(d => tW.innerHTML += `<tr><td>${d.sku}</td><td>${d.nama}</td><td style="color:var(--accent);font-weight:700">${d.minggu}</td><td>${d.satuan}</td></tr>`);
        mD.forEach(d => tM.innerHTML += `<tr><td>${d.sku}</td><td>${d.nama}</td><td style="color:var(--teal);font-weight:700">${d.bulan}</td><td>${d.satuan}</td></tr>`);
        res.data.forEach(d => { const cls = d.statusStok === 'Aman' ? 'stok-aman' : d.statusStok === 'Rendah' ? 'stok-rendah' : 'stok-kritis'; tF.innerHTML += `<tr><td><strong>${d.sku}</strong></td><td>${d.nama}</td><td><strong style="font-size:16px">${d.stokSaat}</strong></td><td style="color:var(--accent)">${d.minggu}</td><td style="color:var(--teal)">${d.bulan}</td><td>${d.rataHarian} / hr</td><td>${d.satuan}</td><td><span class="stok-badge ${cls}">${d.statusStok}</span></td></tr>`; });
      }).getAnalisisStock();
    }

    // ============================================================
    // PENGAJUAN ASSET
    // ============================================================
    let assetDataAll = [];

    function loadAsset() {
      console.log('loadAsset called. currentUser:', currentUser);
      google.script.run.withSuccessHandler(res => {
        console.log('Asset data received:', res);
        if (!res.success) return;
        assetDataAll = res.data;
        const tb = document.getElementById('tableAsset'); tb.innerHTML = '';
        let tot = res.data.length, pend = 0, appr = 0, rej = 0;
        if (!res.data.length) { tb.innerHTML = '<tr><td colspan="9" class="empty-state"><div class="emoji">📦</div>Belum ada pengajuan asset</td></tr>'; }
        else {
          res.data.sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal)).forEach(d => {
            const stsCls = d.status === 'Disetujui' ? 'badge-approved' : d.status === 'Ditolak' ? 'badge-rejected' : 'badge-pending';
            const priColor = d.prioritas === 'Kritis' ? '#ef4444' : d.prioritas === 'Mendesak' ? '#f59e0b' : '#10b981';
            const priLabel = d.prioritas === 'Kritis' ? '🔴 Kritis' : d.prioritas === 'Mendesak' ? '🟡 Mendesak' : '🟢 Normal';
            let userPerms = []; try { userPerms = JSON.parse(currentUser.permissions || '[]'); } catch (e) { }
            const isCreator = currentUser.username === d.createdBy;
            const canEdit = isCreator && !['Disetujui', 'Ditolak', 'Rejected', 'Approved', 'Tolak', 'Disetujui Admin', 'Approved Admin'].includes(d.status);
            const canDel = isCreator && ['Pending HR', 'Pending Team Leader', 'Pending Vice Supervisor', 'Pending Supervisor', 'Pending'].includes(d.status);
            const canApr = canApprove(d.status, currentUser.role, 'Asset');
            if (d.status === 'Disetujui') appr++;
            else if (d.status === 'Ditolak') rej++;
            else pend++;
            const escapedNama = (d.nama || '').replace(/'/g, "\\'");
            const tanggalStr = d.tanggal || '';
            tb.innerHTML += `<tr>
          <td>${formatDate(d.tanggal)}</td>
          <td><strong>${d.nama}</strong></td>
          <td><span class="badge-tb">${d.jenisAsset}</span></td>
          <td class="rupiah">${formatRp(d.estimasiHarga)}</td>
          <td><span style="color:${priColor};font-weight:700">${priLabel}</span></td>
          <td><span class="hover-marquee" style="max-width:180px;display:inline-block">${d.deskripsi || '-'}</span></td>
          <td>${d.bukti ? `<a href="${d.bukti}" target="_blank" class="link-bukti">📎 Lihat</a>` : '-'}</td>
          <td><span class="${stsCls}">${d.status}</span></td>
          <td style="display:flex;gap:4px;flex-wrap:wrap;">
            <button class="btn btn-ghost btn-sm" onclick="showHistoryModal('${encodeURIComponent(d.history)}')">Riwayat</button>
            <button class="btn btn-ghost btn-sm" onclick="openAssetCommentPrompt('${d.id}','${escapedNama}')">🗨️ Komentar</button>
            ${canEdit ? `<button class="btn btn-warning btn-sm" onclick="openAssetEditModal('${d.id}')">✏️ Edit</button>` : ''}
            ${canApr ? `<button class="btn btn-success btn-sm" onclick="processApproval('asset','${d.id}','Approve','${escapedNama}','${tanggalStr}')">✅</button><button class="btn btn-danger btn-sm" onclick="processApproval('asset','${d.id}','Reject','${escapedNama}','${tanggalStr}')">❌</button>` : ''}
            ${canDel ? `<button class="btn btn-danger btn-sm" onclick="delAsset('${d.id}')">Batal</button>` : ''}
          </td>
        </tr>`;
          });
        }
        document.getElementById('statAssetTotal').textContent = tot;
        document.getElementById('statAssetPending').textContent = pend;
        document.getElementById('statAssetApproved').textContent = appr;
        document.getElementById('statAssetRejected').textContent = rej;
      }).getAsset();
    }

    function submitAsset() {
      const t = v('assetTanggal'), n = v('assetNama'), j = v('assetJenis'), h = getRpValue('assetHarga'), p = v('assetPrioritas'), ds = v('assetDeskripsi');
      if (!t || !n || !ds) return toast('Tanggal, Nama, dan Deskripsi wajib diisi', 'error');
      const btn = document.getElementById('btnSaveAsset'); btn.disabled = true; btn.textContent = '⏳...';
      const editId = document.getElementById('assetEditId').value;
      const proceed = url => {
        const callback = res => {
          btn.disabled = false;
          btn.textContent = editId ? '💾 Update Asset' : '📤 Ajukan Asset';
          if (res.success) {
            toast(editId ? 'Pengajuan asset berhasil diperbarui! ✅' : 'Pengajuan asset berhasil dikirim! ✅');
            closeModal('modalAsset'); loadAsset(); checkPendingApprovals();
            resetForm(['assetNama', 'assetDeskripsi']); setVal('assetHarga', ''); removeFile('ast');
            document.getElementById('assetEditId').value = '';
          } else toast(res.message, 'error');
        };
        if (editId) {
          google.script.run.withSuccessHandler(callback).updateAsset(editId, t, n, j, ds, h, p, url, currentUser.username);
        } else {
          google.script.run.withSuccessHandler(callback).addAsset(t, n, j, ds, h, p, url, currentUser.username);
        }
      };
      const f = document.getElementById('astFile').files[0] || window['_droppedFile_ast'];
      if (f) { window['_droppedFile_ast'] = null; uploadFileAndProceed('ast', f, 'Pengajuan Asset', proceed, btn); } else proceed(v('astBukti'));
    }

    function delAsset(id) {
      if (confirm('Batalkan pengajuan asset ini?'))
        google.script.run.withSuccessHandler(res => {
          if (res.success) { toast('Pengajuan dibatalkan'); loadAsset(); checkPendingApprovals(); }
          else toast(res.message, 'error');
        }).deleteAsset(id);
    }

    function openNewAssetModal() {
      document.getElementById('assetEditId').value = '';
      if (typeof setToday === 'function') setToday('assetTanggal');
      else document.getElementById('assetTanggal').value = new Date().toISOString().split('T')[0];

      const assetNamaEl = document.getElementById('assetNama');
      assetNamaEl.value = currentUser.nama || currentUser.username || '';
      assetNamaEl.readOnly = true;
      assetNamaEl.style.background = 'var(--input-bg)';
      assetNamaEl.style.opacity = '0.8';

      document.getElementById('assetJenis').value = 'Laptop / PC';
      document.getElementById('assetPrioritas').value = 'Normal';
      document.getElementById('assetHarga').value = '';
      document.getElementById('assetDeskripsi').value = '';
      document.getElementById('astBukti').value = '';
      document.getElementById('astFile').value = '';
      document.getElementById('ast-fname').textContent = '-';
      document.getElementById('ast-fsize').textContent = '';
      document.getElementById('btnSaveAsset').textContent = '📤 Ajukan Asset';
      openModal('modalAsset');
    }

    function openAssetEditModal(id) {
      const asset = assetDataAll.find(a => a.id === id);
      if (!asset) return toast('Data asset tidak ditemukan', 'error');
      if (currentUser.username !== asset.createdBy) return toast('Hanya pembuat yang dapat mengedit.', 'error');
      if (['Disetujui', 'Ditolak', 'Rejected', 'Approved', 'Tolak', 'Disetujui Admin', 'Approved Admin'].includes(asset.status)) {
        return toast('Pengajuan tidak bisa diedit karena sudah selesai.', 'error');
      }
      document.getElementById('assetEditId').value = asset.id;
      document.getElementById('assetTanggal').value = asset.tanggal || '';
      document.getElementById('assetNama').value = asset.nama || '';
      document.getElementById('assetJenis').value = asset.jenisAsset || 'Laptop / PC';
      document.getElementById('assetPrioritas').value = asset.prioritas || 'Normal';
      document.getElementById('assetHarga').value = asset.estimasiHarga ? asset.estimasiHarga.toString() : '';
      document.getElementById('assetDeskripsi').value = asset.deskripsi || '';
      document.getElementById('astBukti').value = asset.bukti || '';
      document.getElementById('astFile').value = '';
      document.getElementById('ast-fname').textContent = asset.bukti ? asset.bukti.split('/').pop() : '-';
      document.getElementById('ast-fsize').textContent = '';
      document.getElementById('btnSaveAsset').textContent = '💾 Update Asset';
      openModal('modalAsset');
    }

    function openAssetCommentPrompt(id, nama) {
      const comment = prompt('Masukkan komentar untuk pengajuan ini (kosong untuk batal):');
      if (comment === null) return; // canceled
      const choice = prompt('Pilih tindakan setelah komentar:\n0 = Hanya komentar\n1 = Set status -> Pending Yang Membuat\n2 = Set status -> Pending Team Leader\n3 = Set status -> Kembalikan ke Pending Team Leader\nMasukkan angka (0-3):', '0');
      if (choice === null) return;
      const c = String(choice).trim();
      const by = currentUser?.nama || currentUser?.username || 'System';
      const role = currentUser?.role || '';
      // First, add comment
      google.script.run.withSuccessHandler(res => {
        if (!res.success) return toast(res.message || 'Gagal menambahkan komentar', 'error');
        // Then handle status change based on choice
        if (c === '0') {
          toast('Komentar tersimpan');
          loadAsset();
        } else {
          let newStatus = '';
          if (c === '1') newStatus = 'Pending Yang Membuat';
          else if (c === '2') newStatus = 'Pending Team Leader';
          else if (c === '3') newStatus = 'Pending Team Leader';
          if (newStatus) {
            google.script.run.withSuccessHandler(r2 => {
              if (!r2.success) return toast(r2.message || 'Gagal mengubah status', 'error');
              toast('Komentar dan status tersimpan');
              loadAsset(); checkPendingApprovals();
            }).updateAssetStatus(id, newStatus, by, role, comment);
          } else {
            toast('Komentar tersimpan');
            loadAsset();
          }
        }
      }).addAssetComment(id, comment, by, role);
    }


    // STOCK OPNAME LOGIC
    function loadStockOpname() {
      google.script.run.withSuccessHandler(res => {
        stockOpnameDataAll = res.data || [];
        filterStockOpname();
      }).getStockOpname();
    }
    function filterStockOpname() {
      const m = v('filterMonthOpname');
      const filtered = m ? stockOpnameDataAll.filter(d => d.tanggal.startsWith(m)) : stockOpnameDataAll;
      renderStockOpname(filtered);
    }
    function renderStockOpname(data) {
      const tb = document.getElementById('tableStockOpname'); tb.innerHTML = '';
      if (!data.length) { tb.innerHTML = '<tr><td colspan="11" class="empty-state">Belum ada data</td></tr>'; return; }
      data.sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal)).forEach(d => {
        const sel = parseFloat(d.selisih) || 0;
        const selClass = sel > 0 ? 'positive' : (sel < 0 ? 'negative' : '');
        const stClass = d.status === 'Approved' ? 'order-terkirim' : 'order-pending';
        const role = currentUser.role || '';
        const canApproveOp = d.status === 'Pending' && (role === 'admin' || role === 'Supervisor' || role === 'SPV' || role.includes('Supervisor') || role.includes('SPV'));
        tb.innerHTML += `<tr>
      <td>${formatDate(d.tanggal)}</td>
      <td><strong>${d.sku}</strong><br><small>${d.nama}</small></td>
      <td><span class="badge-tb">${d.lokasi || '-'}</span></td>
      <td>${d.batch || '-'}</td>
      <td>${d.expDate || '-'}</td>
      <td>${d.stokSistem}</td>
      <td><strong>${d.stokFisik}</strong></td>
      <td><span class="${selClass}">${sel > 0 ? '+' : ''}${sel}</span></td>
      <td><span class="${stClass}">${d.status}</span></td>
      <td>${d.createdBy}</td>
      <td class="no-print">
        ${canApproveOp ? `<button class="btn btn-teal btn-sm" onclick="procOpname('${d.id}','Approved')">✅</button>
        <button class="btn btn-danger btn-sm" onclick="procOpname('${d.id}','Rejected')">✕</button>` : '-'}
      </td>
    </tr>`;
      });
    }
    function printStockOpname() {
      const m = v('filterMonthOpname');
      const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
      let periodText = "-";
      if (m) {
        const parts = m.split('-');
        periodText = monthNames[parseInt(parts[1]) - 1] + " " + parts[0];
      }
      document.getElementById('opnamePrintSubtitle').textContent = 'Periode: ' + periodText;
      window.print();
    }
    function handleSoScan(el) {
      const q = el.value.trim().toLowerCase(); if (!q) return;
      el.disabled = true;
      const s = stockData.find(x => (x.sku || '').toLowerCase() === q || (x.barcode || '').toLowerCase() === q);
      if (s) {
        setVal('soStockId', s.id);
        handleSoSelect(document.getElementById('soStockId'));
        el.value = ''; // Clear scan after found
        toast('✅ Barang ditemukan: ' + s.nama, 'success');
      } else {
        toast('❌ Barang "' + q + '" tidak ditemukan', 'error');
        el.select();
      }
      el.disabled = false; el.focus();
    }
    function handleSoSelect(sel) {
      const opt = sel.options[sel.selectedIndex];
      const info = document.getElementById('soInfoBarang');
      if (!opt.value) { setVal('soSistem', ''); info.style.display = 'none'; return; }
      const s = stockData.find(x => x.id === opt.value);
      if (s) {
        info.style.display = 'block';
        setVal('soSistem', s.stok);
        document.getElementById('soSKU').textContent = s.sku;
        document.getElementById('soNama').textContent = s.nama;
        document.getElementById('soLokasi').textContent = s.lokasi || '-';
        document.getElementById('soBatch').textContent = s.batch || '-';
        document.getElementById('soExp').textContent = s.expDate || '-';
        calcSoSelisih();
        // Auto-focus physical info after scanning
        setTimeout(() => { document.getElementById('soFisik').focus(); }, 300);
      }
    }
    function calcSoSelisih() {
      const sis = parseFloat(v('soSistem')) || 0, fis = parseFloat(v('soFisik')) || 0, sel = fis - sis;
      const el = document.getElementById('soSelisihWrap'), txt = document.getElementById('soSelisihText');
      el.style.display = 'block'; txt.textContent = (sel > 0 ? '+' : '') + sel; txt.className = sel > 0 ? 'positive' : (sel < 0 ? 'negative' : '');
    }
    function submitStockOpname(isQueue = false) {
      const t = v('soTanggal'), id = v('soStockId'), sis = v('soSistem'), fis = v('soFisik'), cat = v('soCatatan');
      if (!t || (isQueue ? false : !id) || fis === '') return toast('Masukkan jumlah fisik', 'error');

      let s = isQueue ? opnameQueue[currentOpnameIndex] : stockData.find(x => x.id === id);
      let stockId = isQueue ? s.id : id;

      const btn = isQueue ? document.getElementById('btnNextSo') : document.querySelector('#soFooterSingle .btn-primary');
      btn.disabled = true; btn.textContent = '⏳...';

      google.script.run.withSuccessHandler(res => {
        btn.disabled = false; btn.textContent = isQueue ? '♻️ Simpan & Lanjut' : '💾 Ajukan Opname';
        if (res.success) {
          toast('Audit ' + s.sku + ' disimpan ✅');
          if (isQueue) {
            currentOpnameIndex++;
            if (currentOpnameIndex < opnameQueue.length) { renderQueueItem(); }
            else { toast('🎉 Seluruh Tugas Selesai!', 'success'); closeModal('modalStockOpname'); loadStockOpname(); }
          } else {
            closeModal('modalStockOpname'); loadStockOpname();
          }
        } else toast(res.message, 'error');
      }).submitStockOpname(t, stockId, s.sku, s.nama, s.lokasi, s.batch, s.expDate, sis, fis, cat, currentUser.username);
    }

    // BULK SO LOGIC
    function openSoSelection() {
      document.getElementById('searchSoSelect').value = '';
      document.getElementById('checkAllSo').checked = false;
      renderSoSelectionList();
      openModal('modalSoSelection');
    }
    function renderSoSelectionList() {
      const q = v('searchSoSelect').toLowerCase();
      const tb = document.getElementById('tblSoSelection'); tb.innerHTML = '';
      stockData.filter(s => s.sku.toLowerCase().includes(q) || s.nama.toLowerCase().includes(q)).forEach(s => {
        tb.innerHTML += `<tr>
      <td><input type="checkbox" class="so-check" value="${s.id}" onchange="updateSoSelectionCount()"></td>
      <td><strong>${s.sku}</strong><br><small>${s.nama}</small></td>
      <td>${s.batch || '-'}</td>
      <td><span class="badge-tb">${s.lokasi || '-'}</span></td>
      <td>${s.stok}</td>
    </tr>`;
      });
      updateSoSelectionCount();
    }
    function toggleAllSo(el) {
      document.querySelectorAll('.so-check').forEach(c => c.checked = el.checked);
      updateSoSelectionCount();
    }
    function updateSoSelectionCount() {
      const count = document.querySelectorAll('.so-check:checked').length;
      document.getElementById('btnStartSoQueue').textContent = `🏁 Mulai Sesi Opname (${count} Terpilih)`;
      document.getElementById('btnStartSoQueue').disabled = count === 0;
    }
    function startSoQueueFromSelection() {
      const ids = Array.from(document.querySelectorAll('.so-check:checked')).map(c => c.value);
      const items = stockData.filter(s => ids.includes(s.id));
      closeModal('modalSoSelection');
      startQueuedOpname(items);
    }
    function startAllSoQueue() {
      if (confirm('Mulai audit untuk seluruh gudang? (' + stockData.length + ' item)')) {
        startQueuedOpname(stockData);
      }
    }
    function startQueuedOpname(items) {
      opnameQueue = items; currentOpnameIndex = 0;
      document.getElementById('soQueueInfo').style.display = 'block';
      document.getElementById('soFooterQueue').style.display = 'flex';
      document.getElementById('soSelectWrap').style.display = 'none';
      document.getElementById('soFooterSingle').style.display = 'none';
      renderQueueItem();
      openModal('modalStockOpname');
    }
    function renderQueueItem() {
      const s = opnameQueue[currentOpnameIndex];
      document.getElementById('soQueueCurrent').textContent = currentOpnameIndex + 1;
      document.getElementById('soQueueTotal').textContent = opnameQueue.length;

      // Fill Info
      document.getElementById('soInfoBarang').style.display = 'block';
      document.getElementById('soSKU').textContent = s.sku;
      document.getElementById('soNama').textContent = s.nama;
      document.getElementById('soLokasi').textContent = s.lokasi || '-';
      document.getElementById('soBatch').textContent = s.batch || '-';
      document.getElementById('soExp').textContent = s.expDate || '-';

      setVal('soSistem', s.stok);
      setVal('soFisik', '');
      setVal('soCatatan', '');
      document.getElementById('soSelisihWrap').style.display = 'none';

      // Update button text if last item
      document.getElementById('btnNextSo').textContent = (currentOpnameIndex === opnameQueue.length - 1) ? '💾 Simpan & Selesai' : '♻️ Simpan & Lanjut';

      // PDA Auto-focus
      setTimeout(() => { document.getElementById('soFisik').focus(); }, 300);
    }
    function skipSoQueue() {
      currentOpnameIndex++;
      if (currentOpnameIndex < opnameQueue.length) renderQueueItem();
      else { toast('Sesi opname selesai'); closeModal('modalStockOpname'); loadStockOpname(); }
    }
    function finishSoQueue() {
      if (confirm('Akhiri sesi opname sekarang? Progres yang sudah disimpan tetap masuk.')) {
        closeModal('modalStockOpname'); loadStockOpname();
      }
    }
    function procOpname(id, st) {
      const btn = event ? event.target : null;
      if (btn && btn.tagName === 'BUTTON') {
        btn.classList.add('btn-appr-animate');
        btn.disabled = true;
      }

      if (confirm('Konfirmasi ' + st + ' opname ini?' + (st === 'Approved' ? ' Stok master akan diperbarui.' : '')))
        google.script.run.withSuccessHandler(res => {
          if (res.success) {
            if (btn && btn.tagName === 'BUTTON') {
              btn.classList.remove('btn-appr-animate');
              if (st === 'Approved') {
                btn.classList.add('btn-appr-success');
              } else {
                btn.classList.add('btn-appr-reject');
              }
            }
            toast('Berhasil diproses');
            loadStockOpname();
            loadStock();
          } else {
            if (btn && btn.tagName === 'BUTTON') {
              btn.classList.remove('btn-appr-animate');
              btn.disabled = false;
            }
            toast(res.message, 'error');
          }
        })
          .approveStockOpname(id, st, currentUser.username);
    }

    // PACKING LIST LOGIC
    function loadPackingList() {
      google.script.run.withSuccessHandler(res => {
        const tb = document.getElementById('tablePackingList'); tb.innerHTML = '';
        if (!res.data.length) { tb.innerHTML = '<tr><td colspan="6" class="empty-state">Belum ada dokumen</td></tr>'; return; }
        res.data.sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal)).forEach(d => {
          tb.innerHTML += `<tr>
        <td>${formatDate(d.tanggal)}</td>
        <td><strong>${d.noPL}</strong></td>
        <td>${d.noOrder || '-'}</td>
        <td>${d.supplier || '-'}</td>
        <td>${d.keterangan}</td>
        <td>${d.createdBy}</td>
        <td>${d.fileUrl ? `<a href="${d.fileUrl}" target="_blank" class="btn btn-ghost btn-sm">🔗 Buka</a>` : '-'}</td>
        <td><button class="btn btn-danger btn-sm" onclick="delPackingList('${d.id}')">✕</button></td>
      </tr>`;
        });
      }).getPackingList();
    }
    function submitPackingList() {
      const t = v('plTanggal'), no = v('plNo'), noOrd = v('plOrder'), supp = v('plSupplier'), ket = v('plKet');
      if (!t || !no) return toast('Lengkapi data', 'error');

      const btn = document.getElementById('btnSavePl');
      btn.setAttribute('data-orig', '💾 Simpan Dokumen');
      btn.disabled = true; btn.textContent = '⏳...';

      const proceed = url => {
        const user = (currentUser && currentUser.username) ? currentUser.username : 'User';
        google.script.run.withSuccessHandler(res => {
          btn.disabled = false; btn.textContent = '💾 Simpan Dokumen';
          if (res.success) {
            toast('Dokumen Tersimpan'); closeModal('modalPackingList'); loadPackingList(); resetForm(['plNo', 'plOrder', 'plSupplier', 'plKet', 'plUrl']); removeFile('pl');
          } else {
            console.error('Save Error:', res.message);
            toast(res.message, 'error');
          }
        }).withFailureHandler(err => {
          btn.disabled = false; btn.textContent = '💾 Simpan Dokumen';
          console.error('Script Run Error:', err);
          toast('Terjadi kesalahan sistem: ' + err.message, 'error');
        }).addPackingList(t, no, noOrd, supp, ket, url, user);
      };

      const f = document.getElementById('plFile').files[0] || window['_droppedFile_pl'];
      const panel = document.getElementById('pl-panel-upload').classList.contains('active');
      const manualUrl = v('plUrl');

      if (panel && f) {
        window['_droppedFile_pl'] = null;
        uploadFileAndProceed('pl', f, 'Packing List', proceed, btn);
      } else if (!panel && manualUrl) {
        proceed(manualUrl);
      } else if (panel && !f) {
        btn.disabled = false; btn.textContent = '💾 Simpan Dokumen';
        toast('Silakan pilih file terlebih dahulu', 'error');
      } else {
        proceed('');
      }
    }
    function delPackingList(id) { if (confirm('Hapus dokumen ini?')) google.script.run.withSuccessHandler(res => { if (res.success) { toast('Dihapus'); loadPackingList(); } else toast(res.message, 'error'); }).deleteRow('PackingList', id); }

    // ============================================================
    // BOOKING MOBIL LOGIC
    // ============================================================
    function loadBookingMobil() {
      const tb = document.getElementById('tableBookingMobil');
      if (tb && tb.innerHTML === '') {
        tb.innerHTML = '<tr><td colspan="8" style="text-align:center; padding:30px;"><div class="spinner-border text-primary" role="status"><span class="visually-hidden">Loading...</span></div><br><small>Memuat data booking...</small></td></tr>';
      }

      google.script.run.withSuccessHandler(res => {
        console.info('Booking Mobil Response:', res);
        if (!res.success) return toast(res.message, 'error');
        if (!tb) return;

        const data = Array.isArray(res.data) ? res.data : [];
        console.table(data); // Log data dalam bentuk tabel di konsol browser

        if (data.length === 0) {
          tb.innerHTML = '<tr><td colspan="8" class="empty-state">Belum ada booking mobil</td></tr>';
          return;
        }

        // Check permission untuk update status
        const canUpdateStatus = hasPermission('updateStatusBookingMobil');
        console.log('🔐 Permission Check:', {
          username: currentUser?.username,
          role: currentUser?.role,
          permissions: getCurrentUserPermissions(),
          canUpdateStatus: canUpdateStatus
        });

        let htmlRows = '';
        data.sort((a, b) => {
          const tA = a.tanggal || '';
          const tB = b.tanggal || '';
          const dA = tA ? new Date(tA).getTime() : 0;
          const dB = tB ? new Date(tB).getTime() : 0;
          return dB - dA;
        }).forEach(d => {
          const totalBiaya = Number(d.totalBiaya) || 0;
          const totalBiayaText = totalBiaya > 0 ? `Rp ${totalBiaya.toLocaleString('id-ID')}` : '-';
          const buktiPembayaranBtnDisabled = !d.buktiPembayaranUrl ? '' : ' disabled';
          const buktiPembayaranBtnText = d.buktiPembayaranUrl ? '✅ Sudah Upload' : '📤 Upload Bukti';

          // Status button dengan urutan
          const statusBtn = renderBookingStatusButton(d.status, d.id, canUpdateStatus, d);

          htmlRows += `<tr id="bm-master-${d.id}" data-booking-id="${d.id}" data-booking-parkir="${d.parkir}" data-booking-tol="${d.tol}" data-booking-bensin="${d.bensin}" data-booking-pkbm="${d.pkbm || 0}" data-booking-lainlain="${d.lainLain}" data-booking-notes="${(d.driverNotes || '').replace(/"/g, '&quot;')}">
            <td>
              <button class="btn-toggle-row" onclick="toggleBookingDetail('${d.id}')" style="margin-right:8px;">
                <i class="bi bi-plus-circle"></i>
              </button>
              ${formatDate(d.tanggal)}
            </td>
            <td><strong>${d.pic}</strong></td>
            <td>${d.jamBerangkat}</td>
            <td><strong>${d.tujuan}</strong></td>
            <td><small>${d.keterangan || '-'}</small></td>
            <td id="bm-total-master-${d.id}" style="font-weight:600; color:var(--accent);">${totalBiayaText}</td>
            <td id="bm-status-${d.id}">
              ${statusBtn}
            </td>
            <td style="white-space:nowrap;">
              <button class="btn btn-primary btn-sm" onclick="openUploadPaymentProofModal('${d.id}', '${d.tanggal}')" title="Upload Bukti Pembayaran Master" ${buktiPembayaranBtnDisabled}>${buktiPembayaranBtnText}</button>
              ${d.buktiPembayaranUrl ? `<a href="${d.buktiPembayaranUrl}" target="_blank" class="btn btn-teal btn-sm" title="Lihat Bukti Master">👁️ Lihat</a>` : ''}
              <button class="btn btn-ghost btn-sm" onclick="printSuratJalanBooking('${d.id}')" title="Print Surat Jalan">🖨️ Print</button>
              <button class="btn btn-danger btn-sm" onclick="doDeleteBooking('${d.id}')" title="Hapus">✕</button>
            </td>
          </tr>
          <tr id="bm-detail-${d.id}" class="row-detail" style="display:none;">
            <td colspan="8">
              <div class="detail-container">
                <div class="loading-inline">Memuat rincian PO...</div>
              </div>
            </td>
          </tr>`;
        });
        tb.innerHTML = htmlRows;
      }).withFailureHandler(err => {
        console.error('Booking Mobil Critical Error:', err);
        toast('Gagal memuat data: ' + err.message, 'error');
      }).getBookingMobil();
    }

    function openBookingModal() {
      setVal('bmTanggal', new Date().toISOString().split('T')[0]);
      setVal('bmPic', currentUser.nama);
      setVal('bmJamBerangkat', '');
      setVal('bmTujuan', '');
      setVal('bmKeterangan', '');
      setVal('bmRute', '');
      document.getElementById('bmPoTableBody').innerHTML = '';
      addBookingPoRow();
      calculateTotalCostBooking();
      openModal('modalBookingMobil');
    }

    // ============================================================
    // STATUS BOOKING MOBIL - URUTAN OTOMATIS
    // ============================================================
    const BOOKING_STATUS_FLOW = [
      { value: 'Belum Jalan', label: '⏳ Belum Jalan', color: '#6c757d', next: 'Sedang Dalam Perjalanan Ke Tujuan' },
      { value: 'Sedang Dalam Perjalanan Ke Tujuan', label: '🚗 Dalam Perjalanan', color: '#0ea5e9', next: 'Sudah Tiba Di Tempat Tujuan' },
      { value: 'Sudah Tiba Di Tempat Tujuan', label: '📍 Tiba Di Tujuan', color: '#10b981', next: 'Kembali Ke Warehouse JKT' },
      { value: 'Kembali Ke Warehouse JKT', label: '🔙 Kembali Ke WH', color: '#f59e0b', next: 'Sudah Sampai Di Warehouse' },
      { value: 'Sudah Sampai Di Warehouse', label: '✅ Sampai Di WH', color: '#22c55e', next: null }
    ];

    function renderBookingStatusButton(currentStatus, bookingId, canUpdate, bookingData) {
      console.log('🔍 renderBookingStatusButton called:', { currentStatus, bookingId, canUpdate });

      const statusInfo = BOOKING_STATUS_FLOW.find(s => s.value === currentStatus) || BOOKING_STATUS_FLOW[0];
      const isCompleted = statusInfo.next === null;

      console.log('📊 Status Info:', statusInfo);
      console.log('✅ Can Update:', canUpdate);
      console.log('🏁 Is Completed:', isCompleted);

      // Tampilkan timestamp jika ada
      let timestampHtml = '';
      if (currentStatus === 'Sedang Dalam Perjalanan Ke Tujuan' && bookingData.jamMulaiPerjalanan) {
        timestampHtml = `<br><small style="font-size:10px;opacity:0.8;">🕐 ${bookingData.jamMulaiPerjalanan}</small>`;
      } else if (currentStatus === 'Sudah Tiba Di Tempat Tujuan' && bookingData.jamTibaTujuan) {
        timestampHtml = `<br><small style="font-size:10px;opacity:0.8;">🕐 ${bookingData.jamTibaTujuan}</small>`;
      } else if (currentStatus === 'Kembali Ke Warehouse JKT' && bookingData.jamKembaliWarehouse) {
        timestampHtml = `<br><small style="font-size:10px;opacity:0.8;">🕐 ${bookingData.jamKembaliWarehouse}</small>`;
      } else if (currentStatus === 'Sudah Sampai Di Warehouse' && bookingData.jamSampaiWarehouse) {
        timestampHtml = `<br><small style="font-size:10px;opacity:0.8;">🕐 ${bookingData.jamSampaiWarehouse}</small>`;
      }

      if (isCompleted) {
        // Status terakhir - tampilkan badge saja
        console.log('✅ Rendering completed status badge');
        return `<span class="badge" style="background:${statusInfo.color};color:white;padding:8px 12px;border-radius:6px;font-size:12px;font-weight:600;display:inline-block;">${statusInfo.label}${timestampHtml}</span>`;
      }

      if (!canUpdate) {
        // User tidak punya akses - tampilkan badge saja
        console.log('🚫 User cannot update - showing badge only');
        return `<span class="badge" style="background:${statusInfo.color};color:white;padding:8px 12px;border-radius:6px;font-size:12px;font-weight:600;display:inline-block;">${statusInfo.label}${timestampHtml}</span>`;
      }

      // User punya akses dan belum selesai - tampilkan tombol untuk next status
      const nextStatus = BOOKING_STATUS_FLOW.find(s => s.value === statusInfo.next);

      if (!nextStatus) {
        // Fallback jika next status tidak ditemukan
        console.error('❌ Next status not found for:', statusInfo.next);
        return `<span class="badge" style="background:${statusInfo.color};color:white;padding:8px 12px;border-radius:6px;font-size:12px;font-weight:600;display:inline-block;">${statusInfo.label}${timestampHtml}</span>`;
      }

      console.log('🎯 Rendering button with next status:', nextStatus);

      return `
        <div style="display:flex;flex-direction:column;gap:6px;align-items:flex-start;">
          <span class="badge" style="background:${statusInfo.color};color:white;padding:6px 10px;border-radius:6px;font-size:11px;font-weight:600;">${statusInfo.label}${timestampHtml}</span>
          <button class="btn btn-sm" style="background:${nextStatus.color};color:white;font-size:11px;padding:4px 10px;border:none;border-radius:6px;font-weight:600;cursor:pointer;transition:all 0.2s;" onclick="advanceBookingStatus('${bookingId}', '${statusInfo.next}')" onmouseover="this.style.opacity='0.8'" onmouseout="this.style.opacity='1'" title="Klik untuk lanjut ke status berikutnya">
            ➡️ ${nextStatus.label}
          </button>
        </div>
      `;
    }

    function advanceBookingStatus(bookingId, nextStatus) {
      // Gunakan SweetAlert untuk konfirmasi (tidak ada popup Chrome)
      Swal.fire({
        title: 'Ubah Status?',
        html: `Ubah status menjadi:<br><strong>"${nextStatus}"</strong><br><br>⏰ Waktu akan dicatat secara otomatis.`,
        icon: 'question',
        showCancelButton: true,
        confirmButtonColor: '#0ea5e9',
        cancelButtonColor: '#6c757d',
        confirmButtonText: '✅ Ya, Ubah Status',
        cancelButtonText: '❌ Batal',
        reverseButtons: true
      }).then((result) => {
        if (result.isConfirmed) {
          // User klik Ya
          const statusCell = document.getElementById(`bm-status-${bookingId}`);
          if (statusCell) {
            statusCell.innerHTML = '<div class="spinner-border spinner-border-sm text-primary" role="status"></div>';
          }

          google.script.run
            .withSuccessHandler(function (res) {
              if (res.success) {
                // Tampilkan notifikasi sukses dengan SweetAlert
                Swal.fire({
                  title: 'Berhasil!',
                  html: `Status berhasil diubah menjadi:<br><strong>"${nextStatus}"</strong><br><br>🕐 Waktu: ${res.timestamp || 'Tercatat'}`,
                  icon: 'success',
                  timer: 2000,
                  showConfirmButton: false
                });
                console.log('📅 Timestamp:', res.timestamp);
                loadBookingMobil(); // Reload untuk update tampilan
              } else {
                Swal.fire('Gagal', res.message, 'error');
                if (statusCell) statusCell.innerHTML = '<span class="badge badge-danger">Error</span>';
              }
            })
            .withFailureHandler(function (err) {
              console.error('Error updating status:', err);
              Swal.fire('Error', 'Gagal update status: ' + err.message, 'error');
              if (statusCell) statusCell.innerHTML = '<span class="badge badge-danger">Error</span>';
            })
            .updateBookingStatus(bookingId, nextStatus);
        }
      });
    }

    function addBookingPoRow(data = {}) {
      const tbody = document.getElementById('bmPoTableBody');
      const row = document.createElement('tr');
      row.className = 'bm-po-row';
      const fmtVal = (v) => v ? Number(v).toLocaleString('id-ID') : '';
      row.innerHTML = `
        <td><input type="text" class="form-control form-control-sm bm-po-customer" value="${data.namaCustomer || ''}" placeholder="Nama Customer"></td>
        <td><input type="text" class="form-control form-control-sm bm-po-no" value="${data.noPo || ''}" placeholder="PO-123"></td>
        <td><input type="number" class="form-control form-control-sm bm-po-cartoon" value="${data.totalCartoon || 0}" min="0"></td>
        <td><input type="text" class="form-control form-control-sm bm-po-parkir rp-input" value="${fmtVal(data.parkir)}" placeholder="-" oninput="onRpInput(this); calculateTotalCostBooking()"></td>
        <td><input type="text" class="form-control form-control-sm bm-po-tol rp-input" value="${fmtVal(data.tol)}" placeholder="-" oninput="onRpInput(this); calculateTotalCostBooking()"></td>
        <td><input type="text" class="form-control form-control-sm bm-po-pkbm rp-input" value="${fmtVal(data.pkbm)}" placeholder="-" oninput="onRpInput(this); calculateTotalCostBooking()"></td>
        <td><input type="text" class="form-control form-control-sm bm-po-lain rp-input" value="${fmtVal(data.lainLain)}" placeholder="-" oninput="onRpInput(this); calculateTotalCostBooking()"></td>
        <td class="text-center"><button class="btn btn-danger btn-sm" onclick="this.closest('tr').remove(); calculateTotalCostBooking();">✕</button></td>
      `;
      tbody.appendChild(row);
    }

    // ============================================================
    // IMPORT PO READY PICKUP KE BOOKING MOBIL
    // ============================================================

    let poReadyPickupData = []; // Cache data PO Ready Pickup

    function openImportPOReadyPickupModal() {
      openModal('modalImportPOReadyPickup');
      loadPOReadyPickupData();
    }

    function loadPOReadyPickupData() {
      const tbody = document.getElementById('poReadyPickupTableBody');
      tbody.innerHTML = '<tr><td colspan="7" style="text-align:center; padding:20px;"><i class="bi bi-arrow-repeat" style="animation: spin 1s linear infinite;"></i> Memuat data...</td></tr>';

      // Ambil data dari distributorQueueData yang sudah di-load
      if (!distributorQueueData || distributorQueueData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align:center; padding:20px; color:var(--text-muted);">Data Antrian Distributor belum dimuat. Silakan buka menu Antrian Distributor terlebih dahulu.</td></tr>';
        return;
      }

      // Filter hanya PO Ready Pickup
      poReadyPickupData = distributorQueueData.filter(function (item) {
        const status = String(item.statusGudang || '').toLowerCase();
        return status.includes('ready') || status.includes('pickup') || status.includes('siap');
      });

      console.log('PO Ready Pickup found:', poReadyPickupData.length);

      if (poReadyPickupData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align:center; padding:20px; color:var(--text-muted);">Tidak ada PO dengan status Ready Pickup</td></tr>';
        return;
      }

      renderPOReadyPickupTable();
    }

    function renderPOReadyPickupTable() {
      const tbody = document.getElementById('poReadyPickupTableBody');
      tbody.innerHTML = '';

      poReadyPickupData.forEach(function (item, index) {
        const tr = document.createElement('tr');
        tr.className = 'po-ready-row';
        tr.dataset.index = index;

        const sheetBadge = getSheetBadgeSimple(item.sourceSheet);
        const distributor = escHtml(item.namaDistributor || '-');
        const poNumber = escHtml(item.poNumber || '-');
        const totalPcs = item.totalPcs ? Number(item.totalPcs).toLocaleString('id-ID') : '-';
        const jumlahDus = item.jumlahDus ? Number(item.jumlahDus).toLocaleString('id-ID') : '-';
        const alamat = escHtml((item.alamat || '-').substring(0, 50));

        tr.innerHTML = `
          <td class="text-center">
            <input type="checkbox" class="po-checkbox" data-index="${index}">
          </td>
          <td>${sheetBadge}</td>
          <td><strong>${distributor}</strong></td>
          <td><span style="background:rgba(14,165,233,0.1); padding:2px 6px; border-radius:4px; font-size:10px;">${poNumber}</span></td>
          <td class="text-end">${totalPcs}</td>
          <td class="text-end">${jumlahDus}</td>
          <td style="font-size:10px; color:var(--text-muted);">${alamat}</td>
        `;

        tbody.appendChild(tr);
      });

      updateSelectedPOCount();
    }

    function getSheetBadgeSimple(sourceSheet) {
      const sheet = sourceSheet || 'Antrian Distributor';
      if (sheet === 'ANTRIAN FOCALSKIN') {
        return '<span style="background:#10b98122;color:#10b981;padding:2px 6px;border-radius:4px;font-size:9px;font-weight:700;">FOCALSKIN</span>';
      } else if (sheet === 'ANTRIAN MISTINE') {
        return '<span style="background:#0ea5e922;color:#0ea5e9;padding:2px 6px;border-radius:4px;font-size:9px;font-weight:700;">MISTINE</span>';
      } else if (sheet === 'ANTRIAN SBY') {
        return '<span style="background:#f59e0b22;color:#f59e0b;padding:2px 6px;border-radius:4px;font-size:9px;font-weight:700;">SBY</span>';
      } else {
        return '<span style="background:#94a3b822;color:#94a3b8;padding:2px 6px;border-radius:4px;font-size:9px;font-weight:700;">MAIN</span>';
      }
    }

    function toggleSelectAllPO(checkbox) {
      const checkboxes = document.querySelectorAll('.po-checkbox');
      checkboxes.forEach(cb => cb.checked = checkbox.checked);
      updateSelectedPOCount();
    }

    function updateSelectedPOCount() {
      const checked = document.querySelectorAll('.po-checkbox:checked').length;
      document.getElementById('selectedPOCount').textContent = checked;
    }

    // Update count saat checkbox individual diklik
    document.addEventListener('change', function (e) {
      if (e.target.classList.contains('po-checkbox')) {
        updateSelectedPOCount();
      }
    });

    function filterPOReadyPickupTable() {
      const search = document.getElementById('searchPOReadyPickup').value.toLowerCase();
      const rows = document.querySelectorAll('.po-ready-row');

      rows.forEach(function (row) {
        const text = row.textContent.toLowerCase();
        row.style.display = text.includes(search) ? '' : 'none';
      });
    }

    function importSelectedPO() {
      const checkboxes = document.querySelectorAll('.po-checkbox:checked');

      if (checkboxes.length === 0) {
        toast('Pilih minimal 1 PO untuk diimport', 'warning');
        return;
      }

      let imported = 0;
      checkboxes.forEach(function (checkbox) {
        const index = parseInt(checkbox.dataset.index);
        const item = poReadyPickupData[index];

        if (item) {
          // Tambahkan ke tabel Booking Mobil
          addBookingPoRow({
            namaCustomer: item.namaDistributor || '',
            noPo: item.poNumber || '',
            totalCartoon: item.jumlahDus || 0,
            parkir: 0,
            tol: 0,
            pkbm: 0,
            lainLain: 0
          });
          imported++;
        }
      });

      calculateTotalCostBooking();
      closeModal('modalImportPOReadyPickup');
      toast(`✅ Berhasil import ${imported} PO ke Booking Mobil`, 'success');
    }

    // Format rupiah input: hapus non-digit, tambah pemisah ribuan
    function onRpInput(el) {
      var raw = el.value.replace(/[^0-9]/g, '');
      el.value = raw ? Number(raw).toLocaleString('id-ID') : '';
    }

    function rpVal(el) {
      return Number((el.value || '0').replace(/[^0-9]/g, '')) || 0;
    }

    function calculateTotalCostBooking() {
      let parkir = 0, tol = 0, bensin = 0, pkbm = 0, lainLain = 0;

      // Tambahkan dari baris PO (rupiah formatted)
      document.querySelectorAll('.bm-po-row').forEach(row => {
        parkir += rpVal(row.querySelector('.bm-po-parkir'));
        tol += rpVal(row.querySelector('.bm-po-tol'));
        pkbm += rpVal(row.querySelector('.bm-po-pkbm'));
        lainLain += rpVal(row.querySelector('.bm-po-lain'));
      });

      const total = parkir + tol + bensin + pkbm + lainLain;
      document.getElementById('bmTotalBiaya').textContent = 'Rp ' + total.toLocaleString('id-ID');
    }

    function updateRouteSuggestion() {
      const jamStr = v('bmJamBerangkat');
      if (!jamStr) return;
      const jam = parseInt(jamStr.split(':')[0]);
      let rute = "";

      if (jam >= 5 && jam < 10) {
        rute = "Jalur Arteri / Non-Tol";
      } else if (jam >= 10 && jam < 15) {
        rute = "Tol Kota / Utama";
      } else if (jam >= 15 && jam < 20) {
        rute = "Lingkar Luar / Jalur Alternatif";
      } else {
        rute = "Tol / Jalur Utama Cepet";
      }

      setVal('bmRute', rute); // Isi otomatis tapi pengguna bisa hapus/ubah
    }

    function submitBookingMobil() {
      const t = v('bmTanggal'), p = v('bmPic'), j = v('bmJamBerangkat'), tj = v('bmTujuan'), k = v('bmKeterangan'), r = v('bmRute');
      // Biaya master dinonaktifkan (diset 0) karena pindah ke rincian PO
      const parkir = 0, tol = 0, bensin = 0, pkbm = 0, lainLain = 0;

      if (!t || !j || !tj) return toast('Lengkapi Tanggal, Jam, dan Tujuan', 'error');

      const details = [];
      document.querySelectorAll('.bm-po-row').forEach(row => {
        const namaCustomer = row.querySelector('.bm-po-customer').value.trim();
        const noPo = row.querySelector('.bm-po-no').value.trim();
        if (namaCustomer || noPo) {
          details.push({
            namaCustomer: namaCustomer,
            noPo: noPo,
            totalCartoon: row.querySelector('.bm-po-cartoon').value,
            parkir: rpVal(row.querySelector('.bm-po-parkir')),
            tol: rpVal(row.querySelector('.bm-po-tol')),
            pkbm: rpVal(row.querySelector('.bm-po-pkbm')),
            lainLain: rpVal(row.querySelector('.bm-po-lain'))
          });
        }
      });

      const btn = document.getElementById('btnSaveBooking');
      const oldTxt = btn.textContent;
      btn.disabled = true; btn.textContent = '⏳ Menyimpan...';

      google.script.run.withSuccessHandler(res => {
        btn.disabled = false; btn.textContent = oldTxt;
        if (res.success) {
          toast('Booking Mobil berhasil disimpan! ✅');
          closeModal('modalBookingMobil');
          loadBookingMobil();
          if (res.id) {
            setTimeout(() => {
              const detailRow = document.getElementById(`bm-detail-${res.id}`);
              if (detailRow) {
                const masterRow = detailRow.previousElementSibling;
                const toggleBtn = masterRow?.querySelector('.btn-toggle-row');
                if (toggleBtn && detailRow.style.display === 'none') {
                  toggleBookingDetail(res.id);
                }
              }
            }, 300);
          }
        } else toast(res.message, 'error');
      }).addBookingMobil(t, p, j, tj, k, r, currentUser.username, parkir, tol, bensin, pkbm, lainLain, details);
    }

    function toggleBookingDetail(id) {
      const detailRow = document.getElementById(`bm-detail-${id}`);
      const btn = document.querySelector(`#bm-master-${id} .btn-toggle-row`);
      if (detailRow.style.display === 'none') {
        const cont = detailRow.querySelector('.detail-container');
        if (cont.innerHTML.includes('Memuat rincian PO...')) {
          renderBookingDetailContent(id);
        } else {
          // Tetap panggil render jika ingin data selalu fresh saat dibuka
          renderBookingDetailContent(id);
        }
        detailRow.style.display = 'table-row';
        btn.innerHTML = '<i class="bi bi-dash-circle"></i>';
        btn.classList.add('active');
      } else {
        detailRow.style.display = 'none';
        // Reset so next open re-fetches
        const cont = detailRow.querySelector('.detail-container');
        cont.innerHTML = '<div class="loading-inline">Memuat rincian PO...</div>';
        btn.innerHTML = '<i class="bi bi-plus-circle"></i>';
        btn.classList.remove('active');
      }
    }

    // ============================================================
    // EDIT BIAYA PER PO (DETAIL)
    // ============================================================
    let _editBiayaPoId = null;
    let _editBiayaPoBookingId = null;

    function openEditBiayaPoModal(detailId, namaCustomer, noPo, parkir, tol, pkbm, lainLain, bookingId) {
      _editBiayaPoId = detailId;
      _editBiayaPoBookingId = bookingId; // Simpan bookingId untuk refresh detail

      console.log('openEditBiayaPoModal called:', {
        detailId: detailId,
        namaCustomer: namaCustomer,
        noPo: noPo,
        parkir: parkir,
        tol: tol,
        pkbm: pkbm,
        lainLain: lainLain,
        bookingId: bookingId
      });

      document.getElementById('ebpNamaCustomer').textContent = namaCustomer || '-';
      document.getElementById('ebpNoPo').textContent = noPo || '-';

      // Tampilkan nilai actual, bahkan jika 0 (jangan kosongkan)
      const formatVal = (v) => {
        const num = Number(v) || 0;
        return num > 0 ? Number(num).toLocaleString('id-ID') : '0';
      };

      document.getElementById('ebpParkir').value = formatVal(parkir);
      document.getElementById('ebpTol').value = formatVal(tol);
      document.getElementById('ebpPkbm').value = formatVal(pkbm);
      document.getElementById('ebpLainLain').value = formatVal(lainLain);

      calculateEditBiayaPo();
      openModal('modalEditBiayaPo');
    }

    function calculateEditBiayaPo() {
      const get = (id) => {
        const val = document.getElementById(id).value || '0';
        // Hapus titik ribuan dan karakter non-digit
        return Number(String(val).replace(/\./g, '').replace(/[^0-9]/g, '')) || 0;
      };
      const total = get('ebpParkir') + get('ebpTol') + get('ebpPkbm') + get('ebpLainLain');
      document.getElementById('ebpTotal').textContent = total > 0 ? 'Rp ' + total.toLocaleString('id-ID') : 'Rp 0';
    }

    function submitEditBiayaPo() {
      if (!_editBiayaPoId) return;

      const getRaw = (id) => {
        const val = document.getElementById(id).value || '0';
        return Number(String(val).replace(/\./g, '').replace(/[^0-9]/g, '')) || 0;
      };

      const parkir = getRaw('ebpParkir'),
        tol = getRaw('ebpTol'),
        pkbm = getRaw('ebpPkbm'),
        lainLain = getRaw('ebpLainLain');

      // Validation & Debug Logging
      console.log('submitEditBiayaPo - Sending values:', {
        detailId: _editBiayaPoId,
        bookingId: _editBiayaPoBookingId,
        parkir: parkir,
        tol: tol,
        pkbm: pkbm,
        lainLain: lainLain
      });

      // Minimal validation
      if (parkir < 0 || tol < 0 || pkbm < 0 || lainLain < 0) {
        toast('❌ Nilai biaya tidak boleh negatif!', 'error');
        return;
      }

      const btn = document.getElementById('btnSaveEditBiayaPo');
      btn.disabled = true; btn.textContent = '⏳ Menyimpan...';

      google.script.run
        .withFailureHandler(err => {
          btn.disabled = false;
          btn.textContent = '💾 Simpan Perubahan';
          console.error('Backend Error:', err);
          toast('❌ Gagal menyimpan: ' + err.message, 'error');
        })
        .withSuccessHandler(res => {
          btn.disabled = false; btn.textContent = '💾 Simpan Perubahan';
          console.log('Backend Response:', res);

          if (res.success) {
            toast('✅ Biaya PO berhasil diupdate! Sedang menyinkronkan...', 'success');
            closeModal('modalEditBiayaPo');

            // 1. Update Master Total Cell secara langsung agar sinkron
            if (_editBiayaPoBookingId && res.finalTotal !== undefined) {
              const masterTotalCell = document.getElementById(`bm-total-master-${_editBiayaPoBookingId}`);
              if (masterTotalCell) {
                masterTotalCell.textContent = res.finalTotal > 0 ? `Rp ${res.finalTotal.toLocaleString('id-ID')}` : '-';
                masterTotalCell.style.animation = 'pulse-accent 1s'; // Beri visual feedback
                setTimeout(() => masterTotalCell.style.animation = '', 1000);
              }

              // 2. Refresh Detail Container dengan delay agar data ter-sync di backend
              setTimeout(() => {
                renderBookingDetailContent(_editBiayaPoBookingId);
              }, 500);
            } else {
              loadBookingMobil();
            }
          } else {
            toast('❌ ' + (res.message || 'Gagal menyimpan biaya PO'), 'error');
            console.error('Save failed:', res.message);
          }
        })
        .updateBookingMobilDetailBiaya(_editBiayaPoId, parkir, tol, pkbm, lainLain);
    }

    // Fungsi helper untuk merender ulang isi detail tanpa menutup row
    function renderBookingDetailContent(bookingId) {
      google.script.run.withSuccessHandler(res => {
        const detailRow = document.getElementById(`bm-detail-${bookingId}`);
        if (!detailRow) return;
        const cont = detailRow.querySelector('.detail-container');

        if (res.success && res.data.length > 0) {
          const rp = (v) => (v && v > 0) ? 'Rp ' + Number(v).toLocaleString('id-ID') : '<span style="color:#666">-</span>';
          let totalAll = 0;
          let rows = '';
          res.data.forEach(p => {
            const sub = (p.parkir || 0) + (p.tol || 0) + (p.pkbm || 0) + (p.lainLain || 0);
            totalAll += sub;
            let buktiBtn = '';
            if (p.buktiUrls) {
              try {
                const urls = JSON.parse(p.buktiUrls);
                if (urls.length > 0) {
                  buktiBtn = `<div style="display:flex; flex-direction:column; gap:2px;">
                    <button onclick="viewDetailBukti('${p.id}')" class="btn btn-teal btn-xs" style="font-size:10px; padding:2px 6px;">👁️ Lihat (${urls.length})</button>
                    <button onclick="openDetailBuktiModal('${p.id}','${p.bookingId}')" class="btn btn-ghost btn-xs" style="font-size:10px; padding:2px 4px; border:none; opacity:0.7;">+ Tambah</button>
                  </div>`;
                } else {
                  buktiBtn = `<button onclick="openDetailBuktiModal('${p.id}','${p.bookingId}')" class="btn btn-ghost btn-xs" style="font-size:10px; padding:2px 6px;">📎 Upload</button>`;
                }
              } catch (e) {
                buktiBtn = `<button onclick="openDetailBuktiModal('${p.id}','${p.bookingId}')" class="btn btn-ghost btn-xs" style="font-size:10px; padding:2px 6px;">📎 Upload</button>`;
              }
            } else {
              buktiBtn = `<button onclick="openDetailBuktiModal('${p.id}','${p.bookingId}')" class="btn btn-ghost btn-xs" style="font-size:10px; padding:2px 6px;">📎 Upload</button>`;
            }

            rows += `<tr>
              <td><strong>${p.namaCustomer || '-'}</strong></td>
              <td>${p.noPo || '-'}</td>
              <td style="text-align:center">${p.totalCartoon || 0}</td>
              <td>${rp(p.parkir)}</td>
              <td>${rp(p.tol)}</td>
              <td>${rp(p.pkbm)}</td>
              <td>${rp(p.lainLain)}</td>
              <td style="font-weight:700; color:var(--accent)">${sub > 0 ? 'Rp ' + sub.toLocaleString('id-ID') : '-'}</td>
              <td style="white-space:nowrap">
                <button onclick="openEditBiayaPoModal('${p.id}','${p.namaCustomer || ''}','${p.noPo || ''}',${p.parkir || 0},${p.tol || 0},${p.pkbm || 0},${p.lainLain || 0},'${p.bookingId}')" class="btn btn-warning btn-xs" style="font-size:10px; padding:2px 6px;">✏️ Edit</button>
                ${buktiBtn}
              </td>
            </tr>`;
          });

          let html = `<div style="padding:12px 16px;">
            <div style="color:var(--accent); font-weight:800; font-size:12px; margin-bottom:10px; text-transform:uppercase; letter-spacing:1px;">📦 Detail Rincian PO</div>
            <div style="overflow-x:auto">
            <table class="table table-sm table-bordered" style="font-size:11px; min-width:800px;">
              <thead class="table-dark"><tr>
                <th>Nama Customer</th><th>No. PO</th><th>Cartoon</th>
                <th>Parkir</th><th>Tol</th><th>PKBM</th><th>Lain-lain</th>
                <th>Sub Total</th><th>Aksi</th>
              </tr></thead>
              <tbody>${rows}</tbody>
              <tfoot><tr style="font-weight:700; background:rgba(255,255,255,0.05);">
                <td colspan="7" style="text-align:right">Total Biaya PO:</td>
                <td style="color:var(--accent)">Rp ${totalAll.toLocaleString('id-ID')}</td>
                <td></td>
              </tr></tfoot>
            </table></div>
            <div class="mt-2" style="text-align:right">
              <button class="btn btn-teal btn-sm" onclick="printSuratJalanBooking('${bookingId}')">🖨️ Print Surat Jalan</button>
            </div>
          </div>`;
          cont.innerHTML = html;
        } else {
          cont.innerHTML = `<div class="p-3" style="color:var(--text-muted)">Tidak ada rincian PO.</div>`;
        }
      }).getBookingMobilDetail(bookingId);
    }

    // Upload bukti per detail PO
    let _uploadDetailId = null;
    let _uploadDetailBookingId = null;
    let _uploadDetailFiles = [];

    function openDetailBuktiModal(detailId, bookingId) {
      _uploadDetailId = detailId;
      _uploadDetailBookingId = bookingId;
      _uploadDetailFiles = [];
      document.getElementById('ubkBookingId').value = bookingId;
      document.getElementById('ubkNoBooking').textContent = bookingId;
      document.getElementById('ubkFileList').innerHTML = '';
      document.getElementById('ubkFilePreview').style.display = 'none';
      document.getElementById('ubkUploadDone').style.display = 'none';
      document.getElementById('ubkUploadProgress').style.display = 'none';
      document.getElementById('btnDoUploadPayment').disabled = true;
      document.getElementById('btnDoUploadPayment').onclick = function () { handleUploadDetailBukti(); };
      openModal('modalUploadBuktiBayarBooking');
    }

    function viewDetailBukti(detailId) {
      const cont = document.getElementById('vbpContent');
      cont.innerHTML = '<div class="loading-inline">Memuat bukti...</div>';
      openModal('modalViewBuktiPo');

      // Ambil data detail untuk mendapatkan list URL
      // Kita bisa ambil dari data yang sudah ada di memori jika mau, 
      // tapi panggil server lagi memastikan data paling baru.
      google.script.run.withSuccessHandler(res => {
        // Cari baris yang cocok dengan detailId
        const detail = (res.data || []).find(d => String(d.id) === String(detailId));
        if (detail && detail.buktiUrls) {
          try {
            const urls = JSON.parse(detail.buktiUrls);
            if (urls.length > 0) {
              let html = '<div style="display:flex; flex-direction:column; gap:10px;">';
              urls.forEach((url, i) => {
                const fileName = url.split('/').pop().split('?')[0] || `Bukti ${i + 1}`;
                html += `
                  <div style="padding:10px; background:rgba(255,255,255,0.05); border-radius:8px; border:1px solid rgba(255,255,255,0.1); display:flex; justify-content:space-between; align-items:center;">
                    <div style="overflow:hidden; text-overflow:ellipsis; white-space:nowrap; margin-right:10px;">
                      <span style="font-size:16px; margin-right:8px;">📄</span>
                      <span style="font-weight:600;">Bukti ${i + 1}</span>
                    </div>
                    <a href="${url}" target="_blank" class="btn btn-teal btn-sm">Buka 🔗</a>
                  </div>`;
              });
              html += '</div>';
              cont.innerHTML = html;
            } else {
              cont.innerHTML = '<div style="text-align:center; padding:20px; color:var(--text-muted);">Belum ada bukti diupload.</div>';
            }
          } catch (e) {
            cont.innerHTML = '<div style="text-align:center; padding:20px; color:var(--red);">Gagal memproses data bukti.</div>';
          }
        } else {
          cont.innerHTML = '<div style="text-align:center; padding:20px; color:var(--text-muted);">Belum ada bukti diupload.</div>';
        }
      }).getBookingMobilDetailMaster(); // Saya butuh semua detail untuk cari by detailId
    }

    function printSuratJalanBooking(id) {
      google.script.run.withSuccessHandler(function (res) {
        if (!res.success) return toast(res.message, 'error');
        var d = res.data;

        google.script.run.withSuccessHandler(function (detRes) {
          var details = detRes.data || [];
          var poHtml = '';
          details.forEach(function (p) {
            poHtml += '<tr>'
              + '<td>' + (p.namaCustomer || '-') + '</td>'
              + '<td><strong>' + (p.noPo || '-') + '</strong></td>'
              + '<td>' + (p.totalCartoon || 0) + '</td>'
              + '<td>' + formatRp(p.parkir) + '</td>'
              + '<td>' + formatRp(p.tol) + '</td>'
              + '<td>' + formatRp(p.pkbm) + '</td>'
              + '<td>' + formatRp(p.lainLain) + '</td>'
              + '</tr>';
          });
          if (!poHtml) poHtml = '<tr><td colspan="7" style="text-align:center">Tidak ada rincian PO</td></tr>';

          var buktiHtml = d.buktiPembayaranUrl
            ? '<a href="' + d.buktiPembayaranUrl + '" target="_blank">' + d.buktiPembayaranUrl + '</a>'
            : 'Belum Upload';

          var html = '<!DOCTYPE html><html><head>'
            + '<title>Surat Jalan - ' + d.pic + ' - ' + (d.tanggal || '') + '</title>'
            + '<style>'
            + 'body{font-family:Arial,sans-serif;padding:24px;color:#222;font-size:13px;}'
            + '.header{text-align:center;border-bottom:2px solid #333;padding-bottom:10px;margin-bottom:18px;}'
            + '.header h2{margin:0;font-size:18px;letter-spacing:1px;}'
            + '.header p{margin:4px 0 0;color:#555;font-size:12px;}'
            + '.info-grid{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:14px;}'
            + '.info-box{background:#f9f9f9;padding:10px 14px;border-radius:6px;border:1px solid #ddd;}'
            + '.info-box p{margin:3px 0;}'
            + 'table{width:100%;border-collapse:collapse;margin-top:12px;font-size:12px;}'
            + 'th,td{border:1px solid #ccc;padding:6px 10px;text-align:left;}'
            + 'th{background:#f2f2f2;font-weight:700;}'
            + '.footer{margin-top:50px;display:flex;justify-content:space-around;}'
            + '.sign-box{text-align:center;width:180px;}'
            + '.sign-line{margin-top:60px;border-top:1px solid #333;padding-top:4px;font-size:11px;color:#555;}'
            + '@media print{.no-print{display:none;}}'
            + '</style>'
            + '</head><body>'
            + '<div class="header">'
            + '<h2>SURAT JALAN OPERASIONAL</h2>'
            + '<p>Gudang Focallure &mdash; Sistem Manajemen Logistik</p>'
            + '</div>'
            + '<div class="info-grid">'
            + '<div class="info-box">'
            + '<p><strong>ID Booking:</strong> ' + d.id + '</p>'
            + '<p><strong>Tanggal:</strong> ' + (d.tanggal || '-') + '</p>'
            + '<p><strong>PIC:</strong> ' + (d.pic || '-') + '</p>'
            + '<p><strong>Status:</strong> ' + (d.status || '-') + '</p>'
            + '</div>'
            + '<div class="info-box">'
            + '<p><strong>Tujuan:</strong> ' + (d.tujuan || '-') + '</p>'
            + '<p><strong>Jam Berangkat:</strong> ' + (d.jamBerangkat || '-') + '</p>'
            + '<p><strong>Rute:</strong> ' + (d.rute || '-') + '</p>'
            + '</div>'
            + '</div>'
            + '<div class="info-box" style="margin-bottom:14px;">'
            + '<p><strong>Keterangan:</strong> ' + (d.keterangan || '-') + '</p>'
            + '</div>'
            + '<h3 style="margin:0 0 4px;font-size:13px;">Rincian Daftar PO</h3>'
            + '<table>'
            + '<thead><tr><th>Nama Customer</th><th>No. PO</th><th>Total Cartoon</th><th>Parkir</th><th>Tol</th><th>PKBM</th><th>Lain-lain</th></tr></thead>'
            + '<tbody>' + poHtml + '</tbody>'
            + '</table>'
            + '<div class="info-box" style="margin-top:14px;font-size:12px;">'
            + '<p><strong>Catatan Driver:</strong> ' + (d.driverNotes || '-') + '</p>'
            + '<p><strong>Bukti Pembayaran:</strong> ' + buktiHtml + '</p>'
            + '</div>'
            + '<div class="footer">'
            + '<div class="sign-box"><div class="sign-line">PIC Pengirim</div></div>'
            + '<div class="sign-box"><div class="sign-line">Driver</div></div>'
            + '<div class="sign-box"><div class="sign-line">Penerima</div></div>'
            + '</div>'
            + '\x3Cscript\x3Ewindow.onload=function(){window.print();setTimeout(function(){window.close();},500);}\x3C\/script\x3E'
            + '</body></html>';

          var printWin = window.open('', '_blank');
          if (printWin) {
            printWin.document.write(html);
            printWin.document.close();
          } else {
            toast('Pop-up diblokir browser. Izinkan pop-up untuk fitur print.', 'error');
          }
        }).getBookingMobilDetail(id);
      }).getBookingMobilById(id);
    }

    function doDeleteBooking(id) {
      if (!confirm('Hapus booking ini?')) return;
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          toast('Booking dihapus');
          loadBookingMobil();
        } else toast(res.message, 'error');
      }).deleteBookingMobil(id);
    }

    function changeBookingStatus(id, newStatus) {
      toast('⏳ Memperbarui status...', 'success');
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          toast('Status diperbarui! ✅');
          loadBookingMobil();
        } else toast(res.message, 'error');
      }).updateBookingStatus(id, newStatus);
    }

    // ============================================================
    // UPLOAD BUKTI PEMBAYARAN BOOKING MOBIL (MULTI-FILE)
    // ============================================================
    let _paymentFiles = [];
    let _paymentBookingId = null;

    function openUploadPaymentProofModal(id, noBooking) {
      _paymentBookingId = id;
      _paymentFiles = [];
      document.getElementById('ubkBookingId').value = id;
      document.getElementById('ubkNoBooking').textContent = noBooking;

      // Reset UI
      document.getElementById('ubkUploadDone').style.display = 'none';
      document.getElementById('ubkUploadProgress').style.display = 'none';
      document.getElementById('ubkProgressBar').style.width = '0%';
      renderPaymentFileList();

      // Set default action to master upload
      document.getElementById('btnDoUploadPayment').onclick = function () { handleUploadPaymentProof(); };

      openModal('modalUploadBuktiBayarBooking');
    }

    function handleFileSelectBookingPayment(input) {
      const files = Array.from(input.files || []);
      if (!files.length) return;

      files.forEach(f => {
        // Cek duplikat nama & size agar tidak double input
        const exists = _paymentFiles.some(pf => pf.name === f.name && pf.size === f.size);
        if (exists) return;

        if (f.size > 20 * 1024 * 1024) {
          toast(`File ${f.name} terlalu besar (Maks 20MB)`, 'error');
        } else {
          _paymentFiles.push(f);
        }
      });

      renderPaymentFileList();
      // Reset input agar bisa pilih file yang sama jika dihapus
      if (input.tagName === 'INPUT') input.value = '';
    }

    function renderPaymentFileList() {
      const list = document.getElementById('ubkFileList');
      const preview = document.getElementById('ubkFilePreview');
      const btn = document.getElementById('btnDoUploadPayment');

      if (_paymentFiles.length === 0) {
        list.innerHTML = '<div style="text-align:center; color:var(--gray); padding:10px;">Belum ada file dipilih</div>';
        preview.style.display = 'none';
        btn.disabled = true;
        return;
      }

      preview.style.display = 'block';
      btn.disabled = false;

      let html = '';
      _paymentFiles.forEach((f, idx) => {
        const sizeStr = f.size > 1024 * 1024
          ? (f.size / (1024 * 1024)).toFixed(2) + ' MB'
          : (f.size / 1024).toFixed(0) + ' KB';

        html += `<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px; padding:8px 12px; background:rgba(255,255,255,0.05); border-radius:8px; border:1px solid rgba(255,255,255,0.1);">
          <div style="display:flex; align-items:center; gap:10px; overflow:hidden;">
            <span style="font-size:18px;">📄</span>
            <div style="overflow:hidden;">
              <div style="font-weight:600; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${f.name}</div>
              <div style="font-size:10px; color:var(--gray);">${sizeStr}</div>
            </div>
          </div>
          <button onclick="removePaymentFile(${idx})" class="btn btn-ghost btn-sm" style="color:var(--red); padding:4px 8px; font-size:14px;" title="Hapus">✕</button>
        </div>`;
      });
      list.innerHTML = html;
    }

    function removePaymentFile(idx) {
      _paymentFiles.splice(idx, 1);
      renderPaymentFileList();
    }

    function clearFileSelectionPaymentBooking() {
      _paymentFiles = [];
      renderPaymentFileList();
    }

    // Helper untuk upload satu file secara async (Chunking if needed)
    function uploadOneFilePromise(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function (e) {
          const b64 = e.target.result.split(',')[1];
          const fileName = file.name;
          const fileType = file.type;

          // Jika file kecil < 2MB, gunakan upload instan
          if (file.size < 2 * 1024 * 1024) {
            google.script.run
              .withFailureHandler(reject)
              .withSuccessHandler(res => {
                if (res.success) resolve(res.url);
                else reject(res.message);
              }).uploadFileInstant(b64, fileName, fileType, 'Bukti Pembayaran Booking');
          } else {
            // Chunked upload untuk file besar
            const chunkSize = 90000;
            const chunks = [];
            for (let i = 0; i < b64.length; i += chunkSize) {
              chunks.push(b64.substring(i, i + chunkSize));
            }

            let cIdx = 0;
            let uId = '';
            const sendChunk = () => {
              if (cIdx < chunks.length) {
                google.script.run
                  .withFailureHandler(reject)
                  .withSuccessHandler(res => {
                    if (res.success) {
                      uId = res.uploadId;
                      cIdx++;
                      sendChunk();
                    } else reject(res.message);
                  }).uploadChunk(chunks[cIdx], cIdx, uId);
              } else {
                google.script.run
                  .withFailureHandler(reject)
                  .withSuccessHandler(res => {
                    if (res.success) resolve(res.url);
                    else reject(res.message);
                  }).finalizeChunkedUpload(uId, fileName, fileType, 'Bukti Pembayaran Booking');
              }
            };
            sendChunk();
          }
        };
        reader.onerror = reject;
        reader.readAsDataURL(file);
      });
    }

    async function handleUploadPaymentProof() {
      if (!_paymentFiles.length) return;
      if (!_paymentBookingId) return toast('ID Booking tidak ditemukan', 'error');

      const btn = document.getElementById('btnDoUploadPayment');
      const progress = document.getElementById('ubkUploadProgress');
      const pBar = document.getElementById('ubkProgressBar');
      const pText = document.getElementById('ubkProgressText');

      const oldTxt = btn.textContent;
      btn.disabled = true;
      btn.innerHTML = '⏳ Menyiapkan...';
      progress.style.display = 'block';

      const urls = [];
      const total = _paymentFiles.length;

      for (let i = 0; i < total; i++) {
        pText.textContent = `${i + 1} / ${total}`;
        pBar.style.width = `${Math.round((i / total) * 100)}%`;
        btn.innerHTML = `⏳ Uploading (${i + 1}/${total})...`;

        try {
          const url = await uploadOneFilePromise(_paymentFiles[i]);
          urls.push(url);
        } catch (err) {
          console.error('Upload error:', err);
          toast(`Gagal upload file: ${_paymentFiles[i].name}`, 'error');
        }
      }

      pBar.style.width = '100%';
      pText.textContent = 'Menyimpan...';
      btn.innerHTML = '⏳ Memproses...';

      const finalUrl = urls.join('\n'); // Simpan sebagai list baris baru di sheet

      google.script.run.withSuccessHandler(res => {
        btn.disabled = false;
        btn.textContent = oldTxt;
        if (res.success) {
          toast('✅ Bukti Pembayaran Berhasil Disimpan!', 'success');
          document.getElementById('ubkUploadDone').style.display = 'block';
          document.getElementById('ubkUploadDoneCount').textContent = `${urls.length} file berhasil diupload.`;
          _paymentFiles = [];
          renderPaymentFileList();
          loadBookingMobil();
        } else {
          toast(res.message, 'error');
        }
      }).updateBuktiPembayaranBookingUrl(_paymentBookingId, finalUrl);
    }

    async function handleUploadDetailBukti() {
      if (!_paymentFiles.length) return;
      if (!_uploadDetailId) return toast('ID Detail tidak ditemukan', 'error');

      const btn = document.getElementById('btnDoUploadPayment');
      const progress = document.getElementById('ubkUploadProgress');
      const pBar = document.getElementById('ubkProgressBar');
      const pText = document.getElementById('ubkProgressText');

      const oldTxt = btn.textContent;
      btn.disabled = true;
      btn.innerHTML = '⏳ Menyiapkan...';
      progress.style.display = 'block';

      const urls = [];
      const total = _paymentFiles.length;

      for (let i = 0; i < total; i++) {
        pText.textContent = `${i + 1} / ${total}`;
        pBar.style.width = `${Math.round((i / total) * 100)}%`;
        btn.innerHTML = `⏳ Detail Upload (${i + 1}/${total})...`;

        try {
          const url = await uploadOneFilePromise(_paymentFiles[i]);
          urls.push(url);
        } catch (err) {
          console.error('Upload error:', err);
          toast(`Gagal upload file: ${_paymentFiles[i].name}`, 'error');
        }
      }

      pBar.style.width = '100%';
      pText.textContent = 'Menyimpan...';
      btn.innerHTML = '⏳ Memproses...';

      const urlsJson = JSON.stringify(urls);

      google.script.run.withSuccessHandler(res => {
        btn.disabled = false;
        btn.textContent = oldTxt;
        if (res.success) {
          toast('✅ Bukti PO Berhasil Disimpan!', 'success');
          document.getElementById('ubkUploadDone').style.display = 'block';
          document.getElementById('ubkUploadDoneCount').textContent = `${urls.length} file berhasil diupload ke PO.`;
          _paymentFiles = [];
          renderPaymentFileList();
          // Refresh detail container
          if (_uploadDetailBookingId) {
            const detRow = document.getElementById('bm-detail-' + _uploadDetailBookingId);
            if (detRow) detRow.querySelector('.detail-container').innerHTML = '<div class="loading-inline">Memuat rincian PO...</div>';
            toggleBookingDetail(_uploadDetailBookingId); // re-open to refresh
          }
        } else {
          toast(res.message, 'error');
        }
      }).updateBookingMobilDetailBukti(_uploadDetailId, urlsJson);
    }



    // ============================================================
    // INPUT BIAYA DRIVER
    // ============================================================
    let currentDriverCostBookingId = null;

    function openDriverCostFromRow(btn) {
      const row = btn.closest('tr');
      const id = row.dataset.bookingId;
      const tanggal = row.cells[0].textContent;
      const parkir = row.dataset.bookingParkir;
      const tol = row.dataset.bookingTol;
      const bensin = row.dataset.bookingBensin;
      const pkbm = row.dataset.bookingPkbm;
      const lainLain = row.dataset.bookingLainlain;
      const notes = row.dataset.bookingNotes.replace(/&quot;/g, '"');

      openDriverCostModal(id, tanggal, parkir, tol, bensin, pkbm, lainLain, notes);
    }

    function openDriverCostModal(id, noBooking, existingParkir, existingTol, existingBensin, existingPkbm, existingLainLain, existingNotes) {
      currentDriverCostBookingId = id;
      document.getElementById('ibdBookingId').value = id;
      document.getElementById('ibdNoBooking').textContent = noBooking;
      document.getElementById('ibdParkir').value = existingParkir || 0;
      document.getElementById('ibdTol').value = existingTol || 0;
      document.getElementById('ibdBensin').value = existingBensin || 0;
      document.getElementById('ibdPkbm').value = existingPkbm || 0;
      document.getElementById('ibdLainLain').value = existingLainLain || 0;
      document.getElementById('ibdCatatan').value = existingNotes || '';
      calculateTotalCostDriver();
      openModal('modalInputBiayaDriver');
    }

    function calculateTotalCostDriver() {
      const parkir = Number(document.getElementById('ibdParkir').value) || 0;
      const tol = Number(document.getElementById('ibdTol').value) || 0;
      const bensin = Number(document.getElementById('ibdBensin').value) || 0;
      const pkbm = Number(document.getElementById('ibdPkbm').value) || 0;
      const lainLain = Number(document.getElementById('ibdLainLain').value) || 0;
      const total = parkir + tol + bensin + pkbm + lainLain;
      document.getElementById('ibdTotalBiaya').textContent = 'Rp ' + total.toLocaleString('id-ID');
    }

    function submitDriverCostUpdate() {
      const parkir = document.getElementById('ibdParkir').value;
      const tol = document.getElementById('ibdTol').value;
      const bensin = document.getElementById('ibdBensin').value;
      const pkbm = document.getElementById('ibdPkbm').value;
      const lainLain = document.getElementById('ibdLainLain').value;
      const catatan = document.getElementById('ibdCatatan').value;

      const btn = document.getElementById('btnSaveDriverCost');
      const oldTxt = btn.textContent;
      btn.disabled = true;
      btn.textContent = '⏳ Menyimpan...';

      google.script.run.withSuccessHandler(res => {
        btn.disabled = false;
        btn.textContent = oldTxt;
        if (res.success) {
          toast('Biaya perjalanan berhasil diupdate! ✅');
          closeModal('modalInputBiayaDriver');
          loadBookingMobil();
        } else {
          toast(res.message, 'error');
        }
      }).updateBiayaBooking(currentDriverCostBookingId, parkir, tol, bensin, pkbm, lainLain);

      // Update driver notes
      google.script.run.withSuccessHandler(() => {
        // Notes sudah tersimpan, tidak perlu notifikasi tambahan
      }).updateDriverNotesBooking(currentDriverCostBookingId, catatan);
    }

    // ============================================================
    // PACKING LOGIC
    // ============================================================
    let currentPackingOrderId = null;
    let currentPackingItems = [];



    // ============================================================
    // UPLOAD BUKTI PACKING
    // ============================================================
    let currentUploadOrderId = null;
    let currentUploadOrderNo = null;
    let selectedPackingFile = null;

    function openUploadPackingModal(id, no) {
      currentUploadOrderId = id;
      currentUploadOrderNo = no;
      document.getElementById('upNoOrder').textContent = no;
      clearFileSelection();
      openModal('modalUploadPacking');
    }

    function handleFileSelect(input) {
      const file = input.files[0];
      if (!file) return;
      if (file.size > 20 * 1024 * 1024) {
        toast('Ukuran file maksimal 20MB', 'error');
        input.value = ''; return;
      }
      selectedPackingFile = file;
      document.getElementById('upFileName').textContent = file.name;
      document.getElementById('upFilePreview').style.display = 'block';
      document.getElementById('upDropZone').style.display = 'none';
      document.getElementById('btnDoUpload').disabled = false;
    }

    function clearFileSelection() {
      selectedPackingFile = null;
      document.getElementById('upFileInput').value = '';
      document.getElementById('upFileName').textContent = '';
      document.getElementById('upFilePreview').style.display = 'none';
      document.getElementById('upDropZone').style.display = 'block';
      document.getElementById('btnDoUpload').disabled = true;
    }

    function handleUploadPacking() {
      if (!selectedPackingFile) { toast('Silakan pilih file / ambil foto terlebih dahulu.', 'warning'); return; }
      if (!currentUploadOrderId && !currentUploadOrderNo) { toast('Data ID Order tidak ditemukan. Coba refresh halaman.', 'error'); return; }
      const btn = document.getElementById('btnDoUpload');
      const originalText = btn.textContent;
      btn.disabled = true; btn.innerHTML = '⏳ Menyiapkan...';

      const errorHandler = (err) => {
        toast('Terjadi kesalahan jaringan: ' + (err.message || err), 'error');
        btn.disabled = false; btn.textContent = originalText;
      };

      const reader = new FileReader();
      reader.onload = function (e) {
        const b64 = e.target.result.split(',')[1];
        const chunkSize = 90000;
        const chunks = [];
        for (let i = 0; i < b64.length; i += chunkSize) chunks.push(b64.substring(i, i + chunkSize));

        let cIdx = 0; let uId = '';
        const sendChunk = () => {
          btn.innerHTML = `⏳ Mengupload... ${Math.round((cIdx / chunks.length) * 100)}%`;
          if (cIdx < chunks.length) {
            google.script.run
              .withFailureHandler(errorHandler)
              .withSuccessHandler(res => {
                if (res.success) {
                  uId = res.uploadId; cIdx++; sendChunk();
                } else {
                  toast(res.message, 'error'); btn.disabled = false; btn.textContent = originalText;
                }
              }).uploadChunk(chunks[cIdx], cIdx, uId);
          } else {
            btn.innerHTML = '⏳ Memproses Berkas...';
            google.script.run
              .withFailureHandler(errorHandler)
              .withSuccessHandler(res => {
                if (res.success) {
                  google.script.run
                    .withFailureHandler(errorHandler)
                    .withSuccessHandler(res2 => {
                      btn.disabled = false; btn.textContent = originalText;
                      if (res2.success) {
                        toast('✅ Bukti Packing Berhasil Disimpan!', 'success');
                        clearFileSelection();
                        closeModal('modalUploadPacking');
                        loadOrder();
                      } else toast(res2.message, 'error');
                    }).updateBuktiPackingUrl(currentUploadOrderId, currentUploadOrderNo, res.url);
                } else {
                  toast(res.message, 'error'); btn.disabled = false; btn.textContent = originalText;
                }
              }).finalizeChunkedUpload(uId, selectedPackingFile.name, selectedPackingFile.type, 'Bukti Packing');
          }
        };
        sendChunk();
      };
      reader.readAsDataURL(selectedPackingFile);
    }

    // ============================================================
    // CAMERA LOGIC — Netlify Popup Integration
    // ============================================================
    const NETLIFY_CAM_URL = 'https://fcl-camera-app.vercel.app/?mode=cam';
    let netlifyPopup = null;

    // Listener: terima foto dari Netlify popup via postMessage
    window.addEventListener('message', function (event) {
      const data = event.data;
      if (!data) return;

      // Handle Camera Photo (Order/Packing)
      if (data.type === 'FCL_CAMERA_PHOTO') {
        const dataUrl = data.imageBase64;
        if (!dataUrl || !dataUrl.startsWith('data:image')) {
          toast('Format foto tidak valid dari kamera.', 'error');
          return;
        }
        toast('📷 Foto diterima, memulai upload...', 'success');
        uploadBase64Photo(dataUrl);
        return;
      }

    });



    function toggleCamera() {
      if (netlifyPopup && !netlifyPopup.closed) {
        netlifyPopup.focus();
        return;
      }
      const popupWidth = 500;
      const popupHeight = 700;
      const left = Math.round((screen.width - popupWidth) / 2);
      const top = Math.round((screen.height - popupHeight) / 2);
      netlifyPopup = window.open(
        NETLIFY_CAM_URL,
        'FCLCamera',
        `width=${popupWidth},height=${popupHeight},left=${left},top=${top},resizable=yes,scrollbars=no`
      );
      if (!netlifyPopup) {
        toast('Pop-up diblokir browser. Izinkan pop-up untuk situs ini lalu coba lagi.', 'error');
      }
    }

    function uploadBase64Photo(dataUrl) {
      if (!currentUploadOrderId && !currentUploadOrderNo) {
        toast('Data ID Order tidak ditemukan. Coba refresh halaman.', 'error');
        return;
      }
      const btn = document.getElementById('btnDoUpload');
      const originalText = btn ? btn.textContent : '';
      if (btn) { btn.disabled = true; btn.innerHTML = '⏳ Menyiapkan...'; }

      const errorHandler = (err) => {
        toast('Kesalahan jaringan: ' + (err.message || err), 'error');
        if (btn) { btn.disabled = false; btn.textContent = originalText; }
      };

      const b64 = dataUrl.split(',')[1];
      const chunkSize = 90000;
      const chunks = [];
      for (let i = 0; i < b64.length; i += chunkSize) chunks.push(b64.substring(i, i + chunkSize));

      let cIdx = 0; let uId = '';
      const sendChunk = () => {
        if (btn) btn.innerHTML = `⏳ Mengupload... ${Math.round((cIdx / chunks.length) * 100)}%`;
        if (cIdx < chunks.length) {
          google.script.run
            .withFailureHandler(errorHandler)
            .withSuccessHandler(res => {
              if (res.success) { uId = res.uploadId; cIdx++; sendChunk(); }
              else { toast(res.message, 'error'); if (btn) { btn.disabled = false; btn.textContent = originalText; } }
            }).uploadChunk(chunks[cIdx], cIdx, uId);
        } else {
          if (btn) btn.innerHTML = '⏳ Memproses Berkas...';
          google.script.run
            .withFailureHandler(errorHandler)
            .withSuccessHandler(res => {
              if (res.success) {
                google.script.run
                  .withFailureHandler(errorHandler)
                  .withSuccessHandler(res2 => {
                    if (btn) { btn.disabled = false; btn.textContent = originalText; }
                    if (res2.success) {
                      toast('✅ Bukti Packing Berhasil Disimpan!', 'success');
                      closeModal('modalUploadPacking');
                      loadOrder();
                    } else toast(res2.message, 'error');
                  }).updateBuktiPackingUrl(currentUploadOrderId, currentUploadOrderNo, res.url);
              } else {
                toast(res.message, 'error');
                if (btn) { btn.disabled = false; btn.textContent = originalText; }
              }
            }).finalizeChunkedUpload(uId, 'BuktiPacking_' + Date.now() + '.jpg', 'image/jpeg', 'Bukti Packing');
        }
      };
      sendChunk();
    }

    // Legacy stubs — tidak dipakai tapi agar tidak error jika masih direferensikan
    function stopCamera() { }
    function capturePhoto() { toggleCamera(); }

    function toggleOrderFields() {
      const kat = v('ordKategori');
      const mpFields = document.getElementById('ordFieldsMP');
      const stdFields = document.getElementById('ordFieldsStandard');
      const lblPel = document.getElementById('lblOrdPelanggan');
      const lblAlm = document.getElementById('lblOrdAlamat');

      if (kat === 'Marketplace') {
        mpFields.style.display = 'block';
        stdFields.style.display = 'block'; // Tetap tampilkan agar bisa isi Nama Pelanggan
        lblPel.textContent = 'Nama Pembeli (Marketplace)';
        lblAlm.textContent = 'Alamat / Platform';
        if (!v('ordAlamat')) s('ordAlamat', 'Platform Marketplace');
      } else {
        mpFields.style.display = 'none';
        stdFields.style.display = 'block';
        if (kat === 'Store') {
          lblPel.textContent = 'Nama Toko';
          lblAlm.textContent = 'Alamat Toko';
        } else {
          lblPel.textContent = 'Nama Pelanggan';
          lblAlm.textContent = 'Alamat Pengiriman';
        }
      }
    }


    function submitOrder() {
      const kat = v('ordKategori');
      const tgl = v('ordTanggal');
      const ket = v('ordKet');
      let pel = '', alm = '', resiObj = { customNoOrder: '', noResi: '' };

      if (kat === 'Marketplace') {
        pel = v('ordPelanggan') || 'Marketplace Customer'; // Mengambil nama pelanggan jika ada
        resiObj.customNoOrder = v('ordCustomNo');
        resiObj.noResi = v('ordNoResi');
        alm = 'Platform Marketplace';
      } else {
        pel = v('ordPelanggan');
        alm = v('ordAlamat');
        resiObj = ''; // Bukan objek untuk kategori non-MP
      }

      const items = collectItems('ordItems', 'order');
      if (!tgl || !pel || !items) return toast('Lengkapi data & barang!', 'error');

      const btn = document.querySelector('#modalOrder .btn-primary');
      btn.disabled = true; btn.textContent = '⏳ Menyimpan...';

      google.script.run.withSuccessHandler(res => {
        btn.disabled = false; btn.textContent = '💾 Simpan Order';
        if (res.success) {
          toast('✅ Order Berhasil Dibuat!', 'success');
          closeModal('modalOrder');

          // Switch to the correct tab automatically
          let targetTab = 'Dist';
          if (kat === 'Marketplace') targetTab = 'MP';
          else if (kat === 'Store') targetTab = 'Store';
          switchOrderTab(targetTab);

          loadOrder();
          document.getElementById('ordItems').innerHTML = '';
          resetForm(['ordPelanggan', 'ordAlamat', 'ordKet', 'ordCustomNo', 'ordNoResi']);
        } else toast(res.message, 'error');
      }).addOrder(tgl, pel, alm, ket, JSON.stringify(items), currentUser.username, kat, resiObj);
    }

    function kirimOrder(id, noOrder) {
      if (!confirm(`Konfirmasi Kirim Order: ${noOrder}?\nStok akan otomatis dikurangi setelah diklik.`)) return;

      const btn = event.target;
      const originalText = btn.textContent;
      btn.disabled = true; btn.textContent = '⏳...';

      google.script.run.withSuccessHandler(res => {
        btn.disabled = false; btn.textContent = originalText;
        if (res.success) {
          toast(`✅ Order ${noOrder} Berhasil Dikirim!`, 'success');
          loadOrder();
        } else {
          toast(res.message, 'error');
        }
      }).withFailureHandler(err => {
        btn.disabled = false; btn.textContent = originalText;
        toast('Gagal mengirim order: ' + err, 'error');
      }).kirimOrder(id, noOrder);
    }







    function printLemburDetail() {
      window.print();
    }

    /**
     * Memuat ringkasan absensi pribadi untuk user yang sedang login
     * Tampil di sidebar di bawah profil
     */
    function loadMyAttendanceSummary() {
      const elCard = document.getElementById('sbAttendanceCard');
      const elDate = document.getElementById('sbAttDate');
      const elShift = document.getElementById('sbAttShiftBadge');
      const elIn = document.getElementById('sbAttTimeIn');
      const elOut = document.getElementById('sbAttTimeOut');
      const elStatus = document.getElementById('sbAttStatus');

      if (!elCard || !currentUser || !currentUser.nama) return;

      if (elStatus) elStatus.innerHTML = 'Memuat (' + currentUser.nama.split(' ')[0] + ')... <button onclick="loadMyAttendanceSummary()" class="btn-icon-tiny" title="Refresh" style="border:none;background:none;padding:0;cursor:pointer;font-size:12px;">🔄</button>';

      const timeoutToken = setTimeout(() => {
        if (elStatus && elStatus.textContent.includes('Memuat')) {
          elStatus.innerHTML = 'Koneksi Lambat <button onclick="loadMyAttendanceSummary()" class="btn-icon-tiny" style="border:none;background:none;padding:0;cursor:pointer;font-size:12px;">🔄</button>';
        }
      }, 25000);

      google.script.run
        .withSuccessHandler(function (res) {
          clearTimeout(timeoutToken);
          if (res.success && res.data) {
            const d = res.data;
            if (elDate) elDate.textContent = d.tanggal || '-';
            if (elShift) {
              let shiftHtml = d.shift || '-';
              if (d.shift === 'PAGI') shiftHtml = '<span class="att-badge badge-pagi">PAGI</span>';
              else if (d.shift === 'MALAM') shiftHtml = '<span class="att-badge badge-malam">MALAM</span>';
              else if (d.shift === 'OFF') shiftHtml = '<span class="att-badge badge-off">OFF</span>';
              elShift.innerHTML = shiftHtml;
            }
            if (elIn) elIn.textContent = d.in || '--:--';
            if (elOut) elOut.textContent = d.out || '--:--';
            if (elStatus) {
              const refreshBtn = ' <button onclick="loadMyAttendanceSummary()" class="btn-icon-tiny" style="border:none;background:none;padding:0;cursor:pointer;font-size:12px;">🔄</button>';
              if (d.status === 'Sudah Absen') {
                elStatus.innerHTML = '<span style="color:#10b981;font-weight:700;">Sudah Absen</span>' + refreshBtn;
              } else {
                elStatus.innerHTML = '<span style="color:var(--gray);text-transform:none;">' + d.status + '</span>' + refreshBtn;
              }
            }
          } else {
            if (elStatus) elStatus.innerHTML = '- <button onclick="loadMyAttendanceSummary()" class="btn-icon-tiny" title="Refresh" style="border:none;background:none;padding:0;cursor:pointer;font-size:12px;">🔄</button>';
          }
        })
        .withFailureHandler(function (err) {
          clearTimeout(timeoutToken);
          if (elStatus) elStatus.innerHTML = 'Err <button onclick="loadMyAttendanceSummary()" class="btn-icon-tiny" title="Refresh" style="border:none;background:none;padding:0;cursor:pointer;font-size:12px;">🔄</button>';
        })
        .getMyAttendanceToday(currentUser.nama);
    }

    function toggleSidebar() {
      const sb = document.querySelector('.sidebar');
      const mc = document.querySelector('.main');
      if (sb.classList.contains('show')) {
        sb.classList.remove('show');
      } else {
        sb.classList.add('show');
      }
    }

    // Global modal overrides & registrations
    const originalOpenModal = typeof openModal === 'function' ? openModal : (id) => { document.getElementById(id).style.display = 'flex'; };
    const originalCloseModal = typeof closeModal === 'function' ? closeModal : (id) => { document.getElementById(id).style.display = 'none'; };

    window.openModal = function (id) {
      originalOpenModal(id);
      if (id === 'modalLaporanKerja') updateLapFields();
      if (id === 'modalExportDQ') {
        const now = new Date();
        const monthSel = document.getElementById('dqExportMonth');
        const yearInp = document.getElementById('dqExportYear');
        if (monthSel) monthSel.value = now.getMonth() + 1;
        if (yearInp) yearInp.value = now.getFullYear();
      }
    };

    window.closeModal = function (id) {
      originalCloseModal(id);
    };

    // === EXPORT LAPORAN KERJA ===
    function openExportLaporanModal() {
      const today = new Date().toISOString().split('T')[0];
      setVal('exportLapStart', today);
      setVal('exportLapEnd', today);
      originalOpenModal('modalExportLaporan');
    }

    function doExportLaporanExcel() {
      const start = v('exportLapStart'), end = v('exportLapEnd');
      const exportDivisi = v('exportLapDivisi');

      if (!start || !end) return toast('Pilih rentang tanggal', 'error');
      if (new Date(start) > new Date(end)) return toast('Tanggal mulai tidak boleh lebih besar dari tanggal selesai', 'error');

      const btn = document.getElementById('btnDoExportLap');
      btn.disabled = true; btn.textContent = '⏳ Memproses...';

      const filtered = (laporanKerjaData || []).filter(d => {
        const dt = (d.tanggal || '').substring(0, 10);
        const matchDivisi = !exportDivisi || d.divisi === exportDivisi;
        return dt >= start && dt <= end && matchDivisi;
      });

      if (!filtered.length) {
        btn.disabled = false; btn.textContent = '🚀 Ekspor Sekarang';
        return toast('Tidak ada data pada rentang tanggal tersebut', 'error');
      }

      let html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
      <head>
        <meta charset="utf-8">
        <title>Laporan Kerja ${start} sd ${end}</title>
      </head>
      <body>
        <table border="1">
          <thead>
            <tr>
              <th colspan="16" style="font-size: 20px; font-weight: bold; text-align: center; height: 40px; vertical-align: middle; background-color: #f8fafc; color: #0f172a;">
                LAPORAN KERJA TANGGAL ${start} s/d ${end}
              </th>
            </tr>
            <tr style="background-color:#0ea5e9;color:white;font-weight:bold;">
              <th>Tanggal</th><th>PIC</th><th>Divisi</th><th>Staff Masuk</th><th>PHL</th><th>Total Orang Masuk</th><th>Jumlah Orang Lembur</th><th>Jam Kerja</th><th>Total Jam Lembur</th><th>Total Order</th><th>Qty</th><th>KPI</th><th>Orang Pengurangan</th><th>Jam Pengurangan</th><th>Orang Perbantuan</th><th>Jam Perbantuan</th><th>Alasan Pengurangan</th>
            </tr>
          </thead>
          <tbody>`;

      filtered.sort((a, b) => (a.tanggal || '').localeCompare(b.tanggal || '')).forEach(d => {
        let output = d.divisi === 'Distributor' ? (d.totalQty || 0) : (d.totalOrder || d.totalPO || d.totalInbound || 0);
        let kpi = (d.totalJamKerja > 0 && !['Inbound', 'Return', 'KOL'].includes(d.divisi)) ? (output / d.totalJamKerja).toFixed(3) : '-';
        let orgKurang = d.pengurangan > 0 ? 1 : 0;
        let orgBantu = d.perbantuan > 0 ? 1 : 0;
        let totalJamLembur = (parseInt(d.totalStaff) || 0) * (parseFloat(d.jamLembur) || 0);

        let staffMasuk = (parseInt(d.totalOrang) || 0) + (parseInt(d.totalAdmin) || 0);
        let phl = parseInt(d.totalPHL) || 0;
        let totalMasuk = staffMasuk + phl;

        html += `<tr><td>${d.tanggal || '-'}</td><td>${d.pic || '-'}</td><td>${d.divisi || '-'}</td><td>${staffMasuk}</td><td>${phl}</td><td>${totalMasuk}</td><td>${d.totalStaff || 0}</td><td>${d.totalJamKerja || 0}</td><td>${totalJamLembur}</td><td>${d.divisi === 'Distributor' || d.divisi === 'Distributor SBY' ? 0 : output}</td><td>${d.totalQty || 0}</td><td>${kpi}</td><td>${orgKurang}</td><td>${d.pengurangan || 0}</td><td>${orgBantu}</td><td>${d.perbantuan || 0}</td><td>${d.alasanPengurangan || '-'}</td></tr>`;
      });

      html += `</tbody></table></body></html>`;

      try {
        const blob = new Blob(['\ufeff' + html], { type: 'application/vnd.ms-excel;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Laporan_Kerja_${start}_sd_${end}.xls`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        toast('Ekspor Berhasil! 📄');
        if (typeof originalCloseModal === 'function') originalCloseModal('modalExportLaporan');
        else closeModal('modalExportLaporan');
      } catch (err) {
        toast('Gagal ekspor: ' + err.message, 'error');
      } finally {
        if (btn) { btn.disabled = false; btn.textContent = '🚀 Ekspor Sekarang'; }
      }
    }


    // =================================== =========================
    // ASSET WAREHOUSE LOGIC
    // =================================== =========================
    let assetWarehouseData = [];

    function loadAssetWarehouse() {
      const tbody = document.getElementById('bodyAssetWarehouse');
      if (!tbody) return;
      tbody.innerHTML = '<tr><td colspan="7" style="text-align:center; padding:30px;"><div class="spinner"></div><br>Memuat Asset...</td></tr>';

      // Fetch Assets
      google.script.run.withSuccessHandler(res => {
        if (!res.success) {
          tbody.innerHTML = `<tr><td colspan="7" style="text-align:center; color:var(--red);">${res.message}</td></tr>`;
          return;
        }
        assetWarehouseData = res.data;

        // After fetching assets, fetch Map Zones
        google.script.run.withSuccessHandler(mapRes => {
          if (mapRes.success) {
            awMapZones = Array.isArray(mapRes.data) ? mapRes.data : [];
          }
          renderAssetWarehouseTable();
          awRenderMap();
          loadOpnameReportMeta();
        }).getWarehouseMapData();

      }).getAssetWarehouseData();
    }

    function loadOpnameReportMeta() {
      const btn = document.getElementById('btnOpnameReports');
      const approvalCount = document.getElementById('approvalOpnameReportCount');
      const awCount = document.getElementById('awOpnameReportCount');
      if (!btn && !approvalCount && !awCount) return;

      google.script.run.withSuccessHandler(res => {
        if (!res || !res.success) return;
        const count = Array.isArray(res.data) ? res.data.length : 0;
        if (btn) btn.style.display = 'inline-flex';
        if (approvalCount) approvalCount.textContent = count;
        if (awCount) awCount.textContent = count;
      }).getAuditReports();
    }

    function renderAssetWarehouseTable() {
      const q = v('searchAsset').toLowerCase();
      const divFilter = v('filterAssetDivisi');
      const tbody = document.getElementById('bodyAssetWarehouse');
      if (!tbody) return;
      tbody.innerHTML = '';

      // 1. Filter Initial Data
      const kategoriFilter = v('filterAssetKategori');
      const statusFilter = v('filterAssetStatus');
      let filtered = assetWarehouseData.filter(d => {
        const strNama = d.nama != null ? String(d.nama).toLowerCase() : '';
        const strCode = d.code != null ? String(d.code).toLowerCase() : '';
        const strKategori = d.kategori != null ? String(d.kategori).toLowerCase() : '';
        const matchesSearch = strNama.includes(q) || strCode.includes(q) || strKategori.includes(q);
        const matchesDivisi = !divFilter || String(d.divisi) === divFilter;
        const matchesKategori = !kategoriFilter || String(d.kategori) === kategoriFilter;
        const matchesStatus = !statusFilter || String(d.status) === statusFilter;

        return matchesSearch && matchesDivisi && matchesKategori && matchesStatus;
      });

      // Update Stats (Global)
      if (document.getElementById('awStatTotal')) {
        const totalQty = assetWarehouseData.reduce((s, h) => s + (parseInt(h.qty) || 0), 0);
        const mpQty = assetWarehouseData.filter(x => String(x.divisi || '').toLowerCase() === 'marketplace').reduce((s, h) => s + (parseInt(h.qty) || 0), 0);
        const distQty = assetWarehouseData.filter(x => String(x.divisi || '').toLowerCase() === 'distributor').reduce((s, h) => s + (parseInt(h.qty) || 0), 0);
        const kolQty = assetWarehouseData.filter(x => String(x.divisi || '').toLowerCase() === 'kol').reduce((s, h) => s + (parseInt(h.qty) || 0), 0);
        const offQty = assetWarehouseData.filter(x => ['office', 'offICE', 'office'].includes(String(x.divisi || '').toLowerCase())).reduce((s, h) => s + (parseInt(h.qty) || 0), 0);
        const returnQty = assetWarehouseData.filter(x => String(x.divisi || '').toLowerCase() === 'return').reduce((s, h) => s + (parseInt(h.qty) || 0), 0);
        const inboundQty = assetWarehouseData.filter(x => String(x.divisi || '').toLowerCase() === 'inbound').reduce((s, h) => s + (parseInt(h.qty) || 0), 0);

        const activeCount = assetWarehouseData.filter(x => String(x.status || '').toLowerCase() === 'aktif').length;
        const repairCount = assetWarehouseData.filter(x => String(x.status || '').toLowerCase().includes('perbaikan')).length;
        const brokenCount = assetWarehouseData.filter(x => String(x.status || '').toLowerCase().includes('rusak')).length;

        setVal('awStatTotal', totalQty);
        setVal('awStatMp', mpQty);
        setVal('awStatDist', distQty);
        setVal('awStatKOL', kolQty);
        setVal('awStatOffice', offQty);
        setVal('awStatReturn', returnQty);
        setVal('awStatInbound', inboundQty);
        setVal('awStatActiveCount', activeCount);
        setVal('awStatRepairCount', repairCount);
        setVal('awStatBrokenCount', brokenCount);

        // Render / Update Chart
        try {
          const ctx = document.getElementById('awChart').getContext('2d');
          const labels = ['Marketplace', 'Distributor', 'KOL', 'OFFICE', 'Return', 'Inbound', 'Other'];
          const otherQty = Math.max(0, totalQty - (mpQty + distQty + kolQty + offQty + returnQty + inboundQty));
          const data = [mpQty, distQty, kolQty, offQty, returnQty, inboundQty, otherQty];

          if (!window.awChartInstance) {
            window.awChartInstance = new Chart(ctx, {
              type: 'doughnut',
              data: {
                labels: labels,
                datasets: [{
                  data: data,
                  backgroundColor: ['#0ea5e9', '#f59e0b', '#7c3aed', '#10b981', '#ef4444', '#06b6d4', '#94a3b8'],
                  hoverOffset: 8,
                  borderWidth: 0
                }]
              },
              options: {
                plugins: { legend: { position: 'bottom', labels: { usePointStyle: true } } },
                maintainAspectRatio: false,
                responsive: true
              }
            });
          } else {
            window.awChartInstance.data.datasets[0].data = data;
            window.awChartInstance.update();
          }
        } catch (e) { console.error('Chart render error', e); }
      }

      if (filtered.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align:center; padding:30px; color:var(--gray);">Data tidak ditemukan</td></tr>';
        return;
      }

      // 2. Grouping Logic by Kategori if tersedia; otherwise fallback ke kode asset prefix
      const groups = {};
      filtered.forEach(item => {
        const categoryName = String(item.kategori || '').trim();
        const fullCode = String(item.code || '').trim();
        const codePrefix = fullCode !== '' ? (fullCode.includes('-') ? fullCode.split('-')[0] : fullCode) : (item.nama || 'Tanpa Nama');
        const groupKey = categoryName || codePrefix;

        if (!groups[groupKey]) {
          groups[groupKey] = {
            identity: groupKey,
            items: [],
            totalQty: 0,
            divisiList: new Set()
          };
        }
        groups[groupKey].items.push(item);
        groups[groupKey].totalQty += (parseInt(item.qty) || 0);
        groups[groupKey].divisiList.add(item.divisi);
      });

      // Convert groups to array and sort
      const groupList = Object.values(groups).sort((a, b) => a.identity.localeCompare(b.identity));

      groupList.forEach((group, index) => {
        const safeId = 'grp_' + index;
        const divisiSummary = Array.from(group.divisiList).filter(x => x && x.toString().trim() !== '').join(', ') || 'Tanpa Divisi';
        const categories = Array.from(new Set(group.items.map(it => (it.kategori || '').toString().trim()).filter(Boolean)));
        const groupCategory = categories.length === 1 ? categories[0] : (categories.length === 0 ? '-' : 'Beragam');

        // Parent Row Rendering (Simetris 5 Kolom)
        tbody.innerHTML += `
          <tr class="asset-row" onclick="toggleAssetDetail('${safeId}')" style="border-bottom:1px solid var(--border-color); cursor:pointer;">
            <td style="padding:15px; width:50px; text-align:center;">
              <div id="icon-${safeId}" style="width:24px; height:24px; border-radius:50%; background:rgba(14, 165, 233, 0.1); color:var(--teal); display:flex; align-items:center; justify-content:center; font-weight:800; font-size:14px;">+</div>
            </td>
              <td style="padding:15px;">
              <div style="font-weight:700; color:var(--accent); font-size:15px; letter-spacing:1px; text-transform:uppercase;">${groupCategory}</div>
              <div style="font-size:11px; color:var(--text-muted); margin-top:2px;">${group.identity} — ${group.items.length} unit</div>
            </td>
            <td style="padding:15px; text-align:center; width:120px;">
              <div style="font-size:16px; font-weight:800; color:var(--teal);">${group.totalQty}</div>
              <div style="font-size:9px; color:var(--text-muted); font-weight:700; text-transform:uppercase;">Unit</div>
            </td>
            <td style="padding:15px;">
              <div style="font-size:11px; color:var(--text-muted);">
                <i class="bi bi-geo-alt"></i> Tersebar di: <span style="color:var(--teal); font-weight:600;">${divisiSummary}</span>
              </div>
            </td>
            <td style="padding:15px; text-align:right; width:120px;">
               <span class="badge" style="background:rgba(14,165,233,0.1); color:var(--teal); font-size:10px; padding:4px 8px; border-radius:6px;">${group.items.length} Records <i class="bi bi-chevron-right"></i></span>
            </td>
          </tr>
          <!-- Group Detail Row -->
          <tr id="detail-${safeId}" class="asset-detail-row">
            <td colspan="5" class="asset-detail-container">
              <div style="font-size:10px; font-weight:800; color:var(--text-muted); margin-bottom:12px; text-transform:uppercase; letter-spacing:0.5px; padding-left:5px; border-left:3px solid var(--teal); line-height:1;">&nbsp; Rincian Distribusi & Lokasi Rack:</div>
              <div style="display:flex; flex-direction:column; gap:12px;">
                ${group.items.map(item => {
          const statusKey = item.status ? String(item.status).replace(/\s+/g, '').toLowerCase() : 'aktif';
          let statusColor = '#10b981';
          if (statusKey.includes('rusak')) statusColor = '#ef4444';
          if (statusKey.includes('perbaikan')) statusColor = '#f59e0b';

          // Map Zone Label
          const zoneName = awMapZones.find(z => z.id === item.zoneId)?.name || item.zoneId || '-';

          return `
                    <div class="asset-card-detail" style="border-left: 3px solid ${statusColor}; border-radius:8px; background:rgba(255,255,255,0.02);">
                      <div class="asset-info-item" style="flex:1;">
                        <span class="asset-info-label">TANGGAL</span>
                        <span class="asset-info-value" style="font-size:11px;">${formatDate(item.tanggalMasuk)}</span>
                      </div>
                      <div class="asset-info-item" style="flex:1.2;">
                        <span class="asset-info-label">KODE ASSET</span>
                        <span class="asset-info-value" style="font-family:monospace; color:var(--accent); font-weight:700;">${item.code || '-'}</span>
                      </div>
                      <div class="asset-info-item" style="flex:1.5;">
                        <span class="asset-info-label">NAMA BARANG</span>
                        <span class="asset-info-value" style="color:#fff;">${item.nama}</span>
                      </div>
                      <div class="asset-info-item" style="flex:1;">
                        <span class="asset-info-label">KATEGORI</span>
                        <span class="asset-info-value" style="color:var(--teal);">${item.kategori || '-'}</span>
                      </div>
                      <div class="asset-info-item" style="flex:1;">
                        <span class="asset-info-label">DIVISI / LOKASI</span>
                        <span class="asset-info-value" style="color:var(--teal);">📍 ${item.divisi || '-'}${item.status && item.status.toLowerCase().includes('rusak') ? ' (Rusak)' : ''}</span>
                      </div>
                      <div class="asset-info-item" style="flex:0 0 50px; text-align:center;">
                        <span class="asset-info-label">QTY</span>
                        <span class="asset-info-value">${item.qty || 1}</span>
                      </div>
                      <div class="asset-info-item" style="flex: 0 0 100px; justify-content: flex-end; flex-direction: row; align-items: center; gap: 8px;">
                          <button class="btn btn-primary" title="Print Tag" style="padding: 6px 10px;" onclick="event.stopPropagation(); printAssetTag('${item.id}')">🖨️</button>
                          <button class="btn btn-secondary" title="Move" style="padding: 6px 10px;" onclick="event.stopPropagation(); openMoveAssetModal('${item.id}')">📦</button>
                          <button class="btn btn-warning" title="Edit" style="padding: 6px 10px;" onclick="event.stopPropagation(); openEditAssetModal('${item.id}')">✏️</button>
                          <button class="btn btn-danger" title="Delete" style="padding: 6px 10px;" onclick="event.stopPropagation(); doDeleteAssetWarehouse('${item.id}')">✕</button>
                      </div>
                    </div>
                  `;
        }).join('')}
              </div>
            </td>
          </tr>
        `;
      });
    }

    function toggleAssetDetail(safeId) {
      const row = document.getElementById('detail-' + safeId);
      const icon = document.getElementById('icon-' + safeId);
      if (!row) return;

      const isActive = row.classList.contains('active');

      // Close all other rows for a cleaner look
      document.querySelectorAll('.asset-detail-row').forEach(r => r.classList.remove('active'));
      document.querySelectorAll('[id^="icon-"]').forEach(i => {
        i.textContent = '+';
        i.style.background = 'rgba(14, 165, 233, 0.1)';
      });

      if (!isActive) {
        row.classList.add('active');
        icon.textContent = '−';
        icon.style.background = 'var(--accent)';
        icon.style.color = '#fff';
      }
    }

    function generateAssetCodeUI(targetId) {
      const now = new Date();
      const datePart = now.getFullYear() + String(now.getMonth() + 1).padStart(2, '0');
      const randomPart = Math.floor(1000 + Math.random() * 9000);
      const code = 'AST-' + datePart + '-' + randomPart;
      setVal(targetId, code);
      toast('Kode unik dibuat: ' + code);
    }

    function syncOldAssetCodes() {
      const btn = document.querySelector('.btn-sync');
      const oldHtml = btn.innerHTML;

      if (!confirm('Berikan Kode Unik otomatis pada data asset yang belum memiliki kode?')) return;

      btn.disabled = true; btn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Sinkronisasi...';

      google.script.run.withSuccessHandler(res => {
        btn.disabled = false; btn.innerHTML = oldHtml;
        if (res.success) {
          toast('✅ Berhasil menyinkronkan ' + res.count + ' kode asset!', 'success');
          loadAssetWarehouse();
        } else toast(res.message, 'error');
      }).bulkSyncAssetCodes();
    }

    function awPopulateZoneSelect(selectId, currentValue) {
      const select = document.getElementById(selectId);
      if (!select) return;
      select.innerHTML = '<option value="">-- Tanpa Zona --</option>';
      awMapZones.forEach(z => {
        const opt = document.createElement('option');
        opt.value = z.id;
        opt.innerText = z.name;
        if (z.id === currentValue || z.name === currentValue) opt.selected = true;
        select.appendChild(opt);
      });
    }

    function openAssetWarehouseModal() {
      resetForm(['awId', 'awNama', 'awCode', 'awDivisi', 'awQty', 'awStatus', 'awKategori', 'awKategoriCustom']);
      document.getElementById('awTanggal').value = new Date().toISOString().split('T')[0];
      setVal('awDivisi', 'Marketplace');
      setVal('awQty', '1');
      setVal('awStatus', 'Aktif');
      setVal('awKategori', '');
      setVal('awKategoriCustom', '');
      document.getElementById('awKategoriCustom').style.display = 'none';
      openModal('modalAssetWarehouse');
    }

    function toggleAwKategoriCustom() {
      const select = document.getElementById('awKategori');
      const custom = document.getElementById('awKategoriCustom');
      if (!select || !custom) return;
      custom.style.display = select.value === 'Lainnya' ? 'block' : 'none';
    }

    function toggleEAwKategoriCustom() {
      const select = document.getElementById('eAwKategori');
      const custom = document.getElementById('eAwKategoriCustom');
      if (!select || !custom) return;
      custom.style.display = select.value === 'Lainnya' ? 'block' : 'none';
    }

    function setAssetWarehouseStatusFilter(status) {
      const select = document.getElementById('filterAssetStatus');
      if (select) select.value = status;
      renderAssetWarehouseTable();
    }

    function submitAssetWarehouse() {
      const nama = v('awNama'), code = v('awCode'), tgl = v('awTanggal'), div = v('awDivisi'), qty = v('awQty'), status = v('awStatus');
      const kategoriSelect = v('awKategori');
      const kategoriCustom = v('awKategoriCustom');
      const kategori = kategoriSelect === 'Lainnya' ? (kategoriCustom || 'Lain-lain') : kategoriSelect;
      const targetDivisi = String(status || '').toLowerCase() === 'rusak' ? '' : div;
      if (!nama || (!targetDivisi && String(status || '').toLowerCase() !== 'rusak')) return toast('Nama dan Divisi wajib diisi', 'error');
      if (!kategori) return toast('Kategori Asset wajib diisi', 'error');
      const btn = document.querySelector('#modalAssetWarehouse .btn-primary');
      if (btn) { btn.disabled = true; btn.textContent = '⏳ Menyimpan...'; }
      google.script.run.withSuccessHandler(res => {
        if (btn) { btn.disabled = false; btn.textContent = '💾 Simpan Asset'; }
        if (res.success) {
          toast(res.message || 'Asset berhasil disimpan ✅');
          closeModal('modalAssetWarehouse');

          setVal('searchAsset', '');
          setVal('filterAssetDivisi', '');

          loadAssetWarehouse();
        } else toast(res.message, 'error');
      }).addAssetWarehouse(code, nama, tgl, targetDivisi, status, currentUser.username, qty || 1, targetDivisi ? '-' : '', kategori);
    }

    function openEditAssetModal(id) {
      const d = assetWarehouseData.find(x => String(x.id) === String(id));
      if (!d) return;
      setVal('eAwId', d.id);
      setVal('eAwNama', d.nama);
      setVal('eAwTanggal', (d.tanggalMasuk || '').split('T')[0]);
      setVal('eAwStatus', d.status || 'Aktif');
      setVal('eAwQty', d.qty || 1);
      const kategori = d.kategori || '';
      if (kategori && Array.from(document.getElementById('eAwKategori').options).some(o => o.value === kategori)) {
        setVal('eAwKategori', kategori);
        setVal('eAwKategoriCustom', '');
        document.getElementById('eAwKategoriCustom').style.display = 'none';
      } else {
        setVal('eAwKategori', 'Lainnya');
        setVal('eAwKategoriCustom', kategori);
        document.getElementById('eAwKategoriCustom').style.display = 'block';
      }
      openModal('modalEditAsset');
    }

    function submitEditAsset() {
      const id = v('eAwId'), nama = v('eAwNama'), tgl = v('eAwTanggal'), status = v('eAwStatus'), qty = v('eAwQty');
      const kategoriSelect = v('eAwKategori');
      const kategoriCustom = v('eAwKategoriCustom');
      const kategori = kategoriSelect === 'Lainnya' ? (kategoriCustom || 'Lain-lain') : kategoriSelect;
      if (!nama || !tgl) return toast('Nama dan Tanggal wajib diisi', 'error');
      if (!kategori) return toast('Kategori Asset wajib diisi', 'error');
      const btn = document.querySelector('#modalEditAsset .btn-primary');
      if (btn) { btn.disabled = true; btn.textContent = '⏳ Menyimpan...'; }
      google.script.run.withSuccessHandler(res => {
        if (btn) { btn.disabled = false; btn.textContent = '💾 Simpan Perubahan'; }
        if (res.success) {
          toast('Asset diperbarui ✅');
          closeModal('modalEditAsset');
          loadAssetWarehouse();
        } else toast(res.message, 'error');
      }).updateAssetWarehouse(id, nama, tgl, status, currentUser.username, qty || 1, '-', kategori);
    }

    function openMoveAssetModal(id) {
      const d = assetWarehouseData.find(x => String(x.id) === String(id));
      if (!d) return;
      setVal('mAwId', d.id);
      setVal('mAwNama', d.nama);
      setVal('mAwDivisiLama', d.divisi || '-');
      setVal('mAwDivisiBaru', d.divisi || '');
      awPopulateZoneSelect('mAwZoneId', d.zoneId || '');
      openModal('modalMoveAsset');
    }

    function clearMoveAssetZone() {
      setVal('mAwZoneId', '');
      toast('Zona lokasi dihapus');
    }

    function submitMoveAsset() {
      const id = v('mAwId'), targetDiv = v('mAwDivisiBaru'), targetZone = v('mAwZoneId');

      const btn = document.querySelector('#modalMoveAsset .btn-primary');
      if (btn) { btn.disabled = true; btn.textContent = '⏳ Memproses...'; }

      google.script.run.withSuccessHandler(res => {
        if (btn) { btn.disabled = false; btn.textContent = '💾 Update Lokasi'; }
        if (res.success) {
          toast('Lokasi asset berhasil diupdate ✅');
          closeModal('modalMoveAsset');
          loadAssetWarehouse();
        } else toast(res.message, 'error');
      }).moveAssetWarehouse(id, targetDiv, targetZone, currentUser.nama);
    }

    function doDeleteAssetWarehouse(id) {
      if (!confirm('Hapus asset ini dari warehouse?')) return;
      toast('🛠️ Menghapus...', 'info');
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          toast('Asset dihapus');
          loadAssetWarehouse();
        } else toast(res.message, 'error');
      }).deleteAssetWarehouse(id);
    }

    function downloadUserExcelTemplate() {
      const data = [
        ["Username", "Password", "Nama Lengkap", "Role"],
        ["staff01", "password123", "Budi Santoso", "Staff Gudang"],
        ["spv01", "spv123", "Siti Aminah", "Supervisor"]
      ];
      const ws = XLSX.utils.aoa_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Template User");
      XLSX.writeFile(wb, "Template_User_FCL.xlsx");
    }

    function downloadAssetWarehouseTemplate() {
      const data = [
        ["CodePrefix (opsional)", "Nama", "Kategori", "TanggalMasuk", "Divisi", "Status", "Qty", "ZoneId"],
        ["Laptop", "Laptop Dell Latitude", "PC AIO", "2026-05-22", "Marketplace", "Aktif", 2, "Z1"],
        ["Meja", "Meja Kerja", "Meja stainless", "2026-05-22", "OFFICE", "Aktif", 1, ""]
      ];
      const ws = XLSX.utils.aoa_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Template Asset Warehouse");
      XLSX.writeFile(wb, "Template_Impor_Asset_Warehouse.xlsx");
    }

    function handleImportAssetWarehouse(input) {
      const file = input.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = function (e) {
        try {
          const isCsv = file.name.toLowerCase().endsWith('.csv');
          let workbook;
          if (isCsv) {
            const text = new TextDecoder('utf-8').decode(e.target.result);
            workbook = XLSX.read(text, { type: 'string' });
          } else {
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, { type: 'array' });
          }
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const json = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });

          if (json.length === 0) return toast('File kosong!', 'error');

          if (!confirm(`Impor ${json.length} baris asset?`)) return;

          showLoading('Sedang mengimpor asset...');
          const userName = (currentUser && (currentUser.username || currentUser.email || currentUser.nama)) ? (currentUser.username || currentUser.email || currentUser.nama) : 'System';
          google.script.run.withSuccessHandler(res => {
            hideLoading();
            if (res.success) {
              toast(res.message, 'success');
              loadAssetWarehouse();
            } else toast(res.message, 'error');
            input.value = '';
          }).importAssetWarehouseBulk(json, userName);
        } catch (err) {
          hideLoading();
          toast('Gagal membaca file: ' + err.message, 'error');
          input.value = '';
        }
      };
      reader.readAsArrayBuffer(file);
    }

    function exportAssetWarehouseCsv() {
      if (!assetWarehouseData || !assetWarehouseData.length) {
        return toast('Tidak ada data asset untuk diekspor', 'error');
      }
      const headers = ['ID', 'Code', 'Nama', 'Kategori', 'TanggalMasuk', 'Divisi', 'Status', 'CreatedBy', 'CreatedAt', 'History', 'Qty', 'ZoneId'];
      const rows = assetWarehouseData.map(item => [
        item.id,
        item.code,
        item.nama,
        item.kategori || '',
        item.tanggalMasuk,
        item.divisi,
        item.status,
        item.createdBy,
        item.createdAt,
        item.history ? item.history.replace(/\n/g, ' / ') : '',
        item.qty,
        item.zoneId
      ]);
      const csv = [headers, ...rows].map(r => r.map(v => `"${String(v || '').replace(/"/g, '""')}"`).join(',')).join('\n');
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const link = document.createElement('a');
      const url = URL.createObjectURL(blob);
      link.setAttribute('href', url);
      link.setAttribute('download', `Asset_Warehouse_${new Date().toISOString().slice(0, 10)}.csv`);
      link.style.display = 'none';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }

    function printAssetTag(id) {
      const file = input.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = function (e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const json = XLSX.utils.sheet_to_json(firstSheet);

          if (json.length === 0) return toast('File kosong!', 'error');

          if (!confirm(`Impor ${json.length} data user?`)) return;

          showLoading('Sedang mengimpor akun...');
          google.script.run.withSuccessHandler(res => {
            hideLoading();
            if (res.success) {
              toast(res.message, 'success');
              loadUsers();
            } else toast(res.message, 'error');
            input.value = '';
          }).importUsersBulk(json);
        } catch (err) {
          hideLoading();
          toast('Gagal membaca file: ' + err.message, 'error');
          input.value = '';
        }
      };
      reader.readAsArrayBuffer(file);
    }

    function printAssetTag(id) {
      const d = assetWarehouseData.find(x => String(x.id) === String(id));
      if (!d) return;

      const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodeURIComponent(d.code)}`;

      const printWin = window.open('', '_blank');
      if (!printWin) return toast('Pop-up terblokir!', 'error');
      printWin.document.write(`
          <html>
            <head>
              <title>Asset Tag - ${d.code}</title>
              <style>
                @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@700;800&display=swap');
                body { font-family:'Plus Jakarta Sans',sans-serif; display:flex; justify-content:center; align-items:center; height:100vh; margin:0; background:#fff }
                .tag-container { border:2px solid #000; width:450px; display: flex; flex-direction: column; overflow: hidden; }
                .company-header { border-bottom:2px solid #000; padding:8px; font-weight:800; text-align:center; font-size: 16px; text-transform: uppercase; }
                .middle-section { display: flex; border-bottom:2px solid #000; height: 130px; }
                .qr-cell { width:130px; border-right:2px solid #000; display:flex; align-items:center; justify-content:center; padding:10px; }
                .qr-cell img { width:110px; height:110px; }
                .info-cell { flex:1; display:flex; flex-direction:column; }
                .info-row { border-bottom:1px solid #000; flex:1; display:flex; align-items:center; padding-left:12px; font-size: 13px; font-weight: 700; }
                .info-row:last-child { border-bottom:none; }
                .info-label { width: 90px; color: #555; font-size: 11px; text-transform: uppercase; }
                .asset-footer { padding:8px; text-align:center; font-weight:800; font-size: 14px; text-transform: uppercase; background: #f9f9f9; }
                @media print { body { height:auto; padding:0; } }
              </style>
            </head>
            <body>
              <div class="tag-container">
                <div class="company-header">Asset Property OF FCL Group</div>
                <div class="middle-section">
                  <div class="qr-cell">
                    <img src="${qrUrl}" alt="QR Asset">
                  </div>
                  <div class="info-cell">
                    <div class="info-row"><span class="info-label">Kode Asset</span>: ${d.code}</div>
                    <div class="info-row"><span class="info-label">Kategori</span>: ${d.kategori || '-'}</div>
                    <div class="info-row"><span class="info-label">Divisi</span>: ${d.divisi}</div>
                    <div class="info-row"><span class="info-label">Tanggal</span>: ${formatDate(d.tanggalMasuk)}</div>
                  </div>
                </div>
                <div class="asset-footer">${d.nama}</div>
              </div>
              \x3Cscript>window.onload=function(){window.print();setTimeout(()=>{window.close()},500)}<\/script>
            </body>
          </html>
        `);
      printWin.document.close();
    }

    function openBulkPrintAssetModal() {
      const currentDiv = document.getElementById('filterAssetDivisi')?.value || '';
      const sel = document.getElementById('bulkPrintAssetDivisi');
      if (sel) sel.value = currentDiv;
      openModal('modalBulkPrintAsset');
    }

    function doBulkPrintAsset() {
      const divFilter = document.getElementById('bulkPrintAssetDivisi')?.value || '';

      const filtered = assetWarehouseData.filter(d => {
        const matchesDivisi = !divFilter || String(d.divisi) === divFilter;
        return matchesDivisi;
      });

      if (filtered.length === 0) return toast('Tidak ada data untuk dicetak pada divisi ini!', 'warning');

      closeModal('modalBulkPrintAsset');

      if (filtered.length > 300) {
        if (!confirm(`Anda akan mencetak ${filtered.length} label. Lanjutkan?`)) return;
      }

      const printWin = window.open('', '_blank');
      if (!printWin) return toast('Pop-up terblokir!', 'error');

      let tagsHtml = '';
      filtered.forEach(d => {
        const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodeURIComponent(d.code)}`;
        tagsHtml += `
          <div class="tag-container">
            <div class="company-header">Asset Property OF FCL Group</div>
            <div class="middle-section">
              <div class="qr-cell">
                <img src="${qrUrl}" alt="QR Asset">
              </div>
              <div class="info-cell">
                <div class="info-row"><span class="info-label">Kode Asset</span>: ${d.code}</div>
                <div class="info-row"><span class="info-label">Kategori</span>: ${d.kategori || '-'}</div>
                <div class="info-row"><span class="info-label">Divisi</span>: ${d.divisi}</div>
                <div class="info-row"><span class="info-label">Tanggal</span>: ${formatDate(d.tanggalMasuk)}</div>
              </div>
            </div>
            <div class="asset-footer">${d.nama}</div>
          </div>
        `;
      });

      printWin.document.write(`
          <html>
            <head>
              <title>Bulk Asset Print</title>
              <style>
                @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@700;800&display=swap');
                body { font-family:'Plus Jakarta Sans',sans-serif; margin:0; padding:20px; background:#fff; display: flex; flex-wrap: wrap; gap: 20px; justify-content: center; }
                .tag-container { border:2px solid #000; width:420px; display: flex; flex-direction: column; overflow: hidden; page-break-inside: avoid; }
                .company-header { border-bottom:2px solid #000; padding:6px; font-weight:800; text-align:center; font-size: 14px; text-transform: uppercase; }
                .middle-section { display: flex; border-bottom:2px solid #000; height: 110px; }
                .qr-cell { width:110px; border-right:2px solid #000; display:flex; align-items:center; justify-content:center; padding:8px; }
                .qr-cell img { width:90px; height:90px; }
                .info-cell { flex:1; display:flex; flex-direction:column; }
                .info-row { border-bottom:1px solid #000; flex:1; display:flex; align-items:center; padding-left:10px; font-size: 11px; font-weight: 700; }
                .info-row:last-child { border-bottom:none; }
                .info-label { width: 75px; color: #555; font-size: 9px; text-transform: uppercase; }
                .asset-footer { padding:6px; text-align:center; font-weight:800; font-size: 12px; text-transform: uppercase; background: #f9f9f9; }
                @media print { 
                  body { padding: 0; }
                  .tag-container { border: 2px solid #000; margin-bottom: 15px; }
                }
              </style>
            </head>
            <body>
              ${tagsHtml}
              \x3Cscript>window.onload=function(){window.print();setTimeout(()=>{window.close()},500)}<\/script>
            </body>
          </html>
        `);
      printWin.document.close();
    }

    function printAllAssets() {
      // Keep this for backward compatibility if needed or as a generic print
      doBulkPrintAsset();
    }

    function printEmployeeCard(id) {
      const d = karyawanData.find(x => String(x.id) === String(id));
      if (!d) return;

      const qrData = encodeURIComponent(`FCL-EMP|${d.nama}|${d.jabatan}|${d.id}`);
      const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${qrData}`;

      const printWin = window.open('', '_blank');
      if (!printWin) return toast('Pop-up terblokir!', 'error');

      printWin.document.write(`
          <html>
            <head>
              <title>ID CARD - ${d.nama}</title>
              <style>
                @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;700;800&display=swap');
                body { font-family:'Plus Jakarta Sans', sans-serif; display:flex; justify-content:center; align-items:center; height:100vh; margin:0; background:#f0f0f0; }
                .card { width:240px; height:380px; background:#fff; border-radius:15px; overflow:hidden; box-shadow:0 10px 20px rgba(0,0,0,0.1); text-align:center; position:relative; border:1px solid #ddd; }
                .card-header { height:80px; background:linear-gradient(135deg, #0a1628, #1a3a5c); color:#fff; display:flex; flex-direction:column; justify-content:center; align-items:center; }
                .card-header h1 { font-size:14px; margin:0; font-weight:800; letter-spacing:1px; }
                .card-header p { font-size:8px; margin:2px 0 0; opacity:0.8; }
                .avatar { width:70px; height:70px; background:#e2e8f0; border-radius:50%; margin:-35px auto 10px; border:4px solid #fff; display:flex; align-items:center; justify-content:center; font-size:30px; }
                .info { padding:10px 15px; }
                .info h2 { font-size:16px; margin:0; font-weight:800; color:#0a1628; }
                .info p { font-size:10px; margin:4px 0 0; color:#64748b; font-weight:600; text-transform:uppercase; }
                .qr-section { margin-top:15px; background:#f9fafb; padding:15px; border-top:1px dashed #ddd; }
                .qr-section img { width:120px; height:120px; }
                .footer { position:absolute; bottom:15px; width:100%; font-size:8px; color:#94a3b8; font-weight:700; }
                @media print { body { background:#fff; height:auto; } .card { box-shadow:none; border:1px solid #000; } }
              </style>
            </head>
            <body>
              <div class="card">
                <div class="card-header">
                  <h1>FCL GROUP</h1>
                  <p>WAREHOUSE MANAGEMENT SYSTEM</p>
                </div>
                <div class="avatar">👤</div>
                <div class="info">
                  <h2>${d.nama}</h2>
                  <p>${d.jabatan}</p>
                </div>
                <div class="qr-section">
                  <img src="${qrUrl}" alt="Employee QR">
                  <div style="font-size:9px; font-weight:800; margin-top:8px; color:#0a1628;">ATTENDANCE QR</div>
                </div>
                <div class="footer">PROPERTY OF FCL GROUP</div>
              </div>
              \x3Cscript>window.onload=function(){window.print();setTimeout(()=>{window.close()},500)}<\/script>
            </body>
          </html>
        `);
      printWin.document.close();
    }

// ================================================================
    // MODUL ABSENSI KARYAWAN + FINGERPRINT X900
    // ================================================================
    var _absensiLaporan = null;
    var _jadwalDataCache = [];

    function getTodayStr() {
      var d = new Date();
      return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    }

    function loadAbsensiKaryawan() {
      var tgl = document.getElementById('absenFilterTgl').value || getTodayStr();
      var div = document.getElementById('absenFilterDiv').value || '';
      document.getElementById('absenStatCards').innerHTML = '<div style="padding:10px;color:var(--gray);">&#9203; Memuat data absensi...</div>';
      document.getElementById('absenTabelBody').innerHTML = '<tr><td colspan="8" class="text-center py-3 text-muted">&#9203; Memuat...</td></tr>';
      document.getElementById('absenDivGrid').innerHTML = '';
      document.getElementById('absenAlfaSection').innerHTML = '';
      google.script.run
        .withSuccessHandler(function (res) {
          try {
            if (!res || !res.success) {
              showToast('Gagal: ' + (res ? res.message : 'error'), 'error');
              document.getElementById('absenTabelBody').innerHTML = '<tr><td colspan="8" class="text-center py-3 text-danger">Gagal memuat data.</td></tr>';
              return;
            }
            _absensiLaporan = res;
            renderAbsensiPage(res);
          } catch (e) {
            console.error('JS Error renderAbsensiPage:', e);
            showToast('Error Tampilan: ' + e.message, 'error');
            document.getElementById('absenTabelBody').innerHTML = '<tr><td colspan="8" class="text-center py-3 text-danger">Error Sistem Tampilan.</td></tr>';
          }
        })
        .withFailureHandler(function (err) {
          showToast('Server Error: ' + err.message, 'error');
          document.getElementById('absenTabelBody').innerHTML = '<tr><td colspan="8" class="text-center py-3 text-danger">Error Koneksi Server.</td></tr>';
        })
        .getLaporanAbsensi(tgl, div, currentUser.username);
    }

    function renderAbsensiPage(res) {
      var terlambat = (res.sudahAbsen || []).filter(function (k) { return k.statusAbsen === 'Terlambat'; }).length;
      var pct = res.totalKaryawan > 0 ? Math.round(res.totalHadir / res.totalKaryawan * 100) : 0;
      document.getElementById('absenStatCards').innerHTML =
        '<div class="stat-card green"><div class="stat-label">&#9989; Total Hadir</div><div class="stat-value green">' + res.totalHadir + '</div><div class="stat-icon">&#128101;</div></div>' +
        '<div class="stat-card red"><div class="stat-label">&#10060; Alfa / Belum Absen</div><div class="stat-value red">' + res.totalAlfa + '</div><div class="stat-icon">&#128683;</div></div>' +
        '<div class="stat-card amber"><div class="stat-label">&#9200; Terlambat</div><div class="stat-value amber">' + terlambat + '</div><div class="stat-icon">&#9888;&#65039;</div></div>' +
        '<div class="stat-card teal"><div class="stat-label">&#128202; Kehadiran</div><div class="stat-value teal">' + pct + '%</div><div class="stat-icon">&#128200;</div></div>';

      var divHtml = '';
      (res.rekapDivisi || []).forEach(function (d) {
        var p = d.total > 0 ? Math.round(d.hadir / d.total * 100) : 0;
        var bc = p >= 80 ? 'var(--green)' : p >= 50 ? 'var(--accent)' : 'var(--red)';
        divHtml += '<div class="absen-div-card">' +
          '<div class="div-name">&#127970; ' + d.divisi + '</div>' +
          '<div class="absen-div-progress"><div class="absen-div-progress-bar" style="width:' + p + '%;background:' + bc + ';"></div></div>' +
          '<div style="font-size:11px;color:var(--gray);margin-bottom:6px;">' + p + '% hadir dari ' + d.total + ' karyawan</div>' +
          '<div class="absen-div-stats">' +
          '<div class="absen-div-stat"><div class="val green">' + d.hadir + '</div><div class="lbl">Hadir</div></div>' +
          '<div class="absen-div-stat"><div class="val amber">' + d.terlambat + '</div><div class="lbl">Terlambat</div></div>' +
          '<div class="absen-div-stat"><div class="val red">' + d.alfa + '</div><div class="lbl">Alfa</div></div>' +
          '</div></div>';
      });
      document.getElementById('absenDivGrid').innerHTML = divHtml || '<div class="text-muted" style="padding:10px;">Tidak ada data divisi.</div>';

      var rows = '';
      (res.sudahAbsen || []).forEach(function (k) {
        var inLog = k.inLog;
        var outLog = k.outLog;
        var jamIn = inLog ? inLog.jam : '-';
        var jamOut = outLog ? outLog.jam : '-';
        var sumber = (inLog && inLog.sumber === 'fingerprint') || (outLog && outLog.sumber === 'fingerprint') ? 'fingerprint' : 'manual';
        var fpId = inLog ? (inLog.fingerprintId || '-') : (outLog ? (outLog.fingerprintId || '-') : '-');
        var logId = inLog ? inLog.id : (outLog ? outLog.id : '');
        var stBadge = getBadgeAbsensi(k.statusAbsen);
        var srcBadge = sumber === 'fingerprint' ? '<span class="badge-fingerprint">FP</span>' : '<span class="badge-manual-src">Manual</span>';

        var nameHtml = k.isUnmapped
          ? '<strong class="text-danger">' + (k.nama || '-') + '</strong> <span class="badge bg-warning text-dark" style="font-size:9px;">UNMAPPED</span>'
          : '<strong>' + (k.nama || '-') + '</strong>';

        rows += '<tr' + (k.isUnmapped ? ' style="background:rgba(220,53,69,0.05);"' : '') + '><td>' + nameHtml + '</td>' +
          '<td><span class="badge" style="background:rgba(26,58,92,0.1); color:var(--primary); font-size:10px;">' + (k.shift || '-') + '</span></td>' +
          '<td style="font-size:12px;color:var(--gray);">' + (k.tanggal || '-') + '</td>' +
          '<td>' + (k.divisi || '-') + '</td>' +
          '<td style="font-size:11px;color:var(--gray);">' + fpId + '</td>' +
          '<td style="font-weight:700;color:var(--teal); font-family:Courier New; font-size:13px;">' + jamIn + '</td>' +
          '<td style="font-weight:700;color:var(--accent); font-family:Courier New; font-size:13px;">' + jamOut + '</td>' +
          '<td>' + stBadge + '</td><td>' + srcBadge + '</td>' +
          '<td><button class="btn btn-danger btn-sm" onclick="hapusAbsensi(\'' + logId + '\')">&#128465;</button></td></tr>';
      });
      document.getElementById('absenTabelBody').innerHTML = rows || '<tr><td colspan="8" class="text-center py-4 text-muted">Tidak ada data absensi untuk tanggal ini.</td></tr>';

      if ((res.belumAbsen || []).length === 0) {
        document.getElementById('absenAlfaSection').innerHTML = '<div class="alfa-section"><h5>&#9989; Semua Karyawan Sudah Absen!</h5></div>';
        return;
      }
      var byDiv = {};
      (res.belumAbsen || []).forEach(function (k) { if (!byDiv[k.divisi]) byDiv[k.divisi] = []; byDiv[k.divisi].push(k); });
      var alfaHtml = '<div class="alfa-section"><h5>&#9888; Belum Absen <span style="font-size:12px;font-weight:600;color:var(--gray);">(' + res.belumAbsen.length + ' karyawan)</span></h5>';
      Object.keys(byDiv).sort().forEach(function (d) {
        alfaHtml += '<div class="alfa-div-group"><div class="alfa-div-title">&#127970; ' + d + ' (' + byDiv[d].length + ' orang)</div><div class="alfa-pill-wrap">';
        byDiv[d].forEach(function (k) {
          alfaHtml += '<span class="alfa-pill">&#128100; ' + k.nama + ' <small style="opacity:0.6;font-size:9px;">(' + (k.shift || '-') + ')</small></span>';
        });
        alfaHtml += '</div></div>';
      });
      alfaHtml += '</div>';
      document.getElementById('absenAlfaSection').innerHTML = alfaHtml;
    }

    function getBadgeAbsensi(st) {
      if (!st || st === 'Hadir') return '<span class="badge-hadir">&#9989; Hadir</span>';
      if (st === 'Terlambat') return '<span class="badge-terlambat">&#9200; Terlambat</span>';
      if (st === 'Alfa') return '<span class="badge-alfa">&#10060; Alfa</span>';
      if (st === 'Pulang Awal') return '<span class="badge-pulang-awal">Pulang Awal</span>';
      if (st === 'Pulang') return '<span class="badge-pulang">Pulang</span>';
      return '<span class="badge-hadir">' + st + '</span>';
    }

    function hapusAbsensi(id) {
      if (!id || !confirm('Hapus data absensi ini?')) return;
      google.script.run.withSuccessHandler(function (res) {
        if (res.success) { showToast('Data absensi dihapus.', 'success'); loadAbsensiKaryawan(); }
        else showToast('Gagal: ' + res.message, 'error');
      }).deleteAbsensiKaryawan(id, currentUser.username);
    }

    function openAbsensiManualModal() {
      var today = getTodayStr(); var now = new Date();
      document.getElementById('amTanggal').value = today;
      document.getElementById('amJam').value = String(now.getHours()).padStart(2, '0') + ':' + String(now.getMinutes()).padStart(2, '0');
      document.getElementById('amKaryawanId').value = '';
      document.getElementById('amNama').value = '';
      document.getElementById('amDivisi').value = '';
      document.getElementById('amJabatan').value = '';
      document.getElementById('amKeterangan').value = '';
      document.getElementById('amTipe').value = 'IN';
      populateAbsenKaryawanSelect();
      openModal('modalAbsensiManual');
    }

    function populateAbsenKaryawanSelect() {
      google.script.run.withSuccessHandler(function (res) {
        var sel = document.getElementById('amKaryawanPicker');
        sel.innerHTML = '<option value="">-- Pilih Karyawan --</option>';
        (res.data || []).forEach(function (k) {
          var opt = document.createElement('option');
          opt.value = k.id;
          opt.setAttribute('data-nama', k.nama || '');
          opt.setAttribute('data-divisi', k.cabang || '');
          opt.setAttribute('data-jabatan', k.jabatan || '');
          opt.textContent = k.nama + ' - ' + (k.jabatan || '');
          sel.appendChild(opt);
        });
      }).getKaryawan();
    }

    function onKaryawanPickerChange() {
      var sel = document.getElementById('amKaryawanPicker');
      var opt = sel.options[sel.selectedIndex];
      if (opt && opt.value) {
        document.getElementById('amKaryawanId').value = opt.value;
        document.getElementById('amNama').value = opt.getAttribute('data-nama') || '';
        document.getElementById('amDivisi').value = opt.getAttribute('data-divisi') || '';
        document.getElementById('amJabatan').value = opt.getAttribute('data-jabatan') || '';
      }
    }

    function simpanAbsensiManual() {
      var kId = document.getElementById('amKaryawanId').value;
      var nama = document.getElementById('amNama').value.trim();
      var div = document.getElementById('amDivisi').value.trim();
      var jab = document.getElementById('amJabatan').value.trim();
      var tipe = document.getElementById('amTipe').value;
      var jam = document.getElementById('amJam').value;
      var tgl = document.getElementById('amTanggal').value;
      var ket = document.getElementById('amKeterangan').value.trim();
      if (!nama || !div) { showToast('Nama dan Divisi wajib diisi!', 'error'); return; }
      var btn = document.getElementById('btnSimpanAbsenManual');
      btn.disabled = true; btn.textContent = 'Menyimpan...';
      google.script.run.withSuccessHandler(function (res) {
        btn.disabled = false; btn.textContent = 'Simpan Absensi';
        if (res.success) { showToast('Absensi disimpan! Status: ' + res.status, 'success'); closeModal('modalAbsensiManual'); loadAbsensiKaryawan(); }
        else showToast('Gagal: ' + res.message, 'error');
      }).addAbsensiKaryawan(kId, nama, div, jab, tipe, jam, tgl, ket, currentUser.username);
    }

    function showFingerprintInfo() {
      google.script.run.withSuccessHandler(function (res) {
        var url = (res && res.url) ? res.url : '[Jalankan setupDatabase() untuk mendapatkan URL deployment]';
        document.getElementById('fpWebhookUrl').textContent = url;
        document.getElementById('fpTestStatus').textContent = '';
      }).getSpreadsheetUrl();

      // RBAC Gating untuk tombol Repair
      const perms = JSON.parse(currentUser.permissions || '[]');
      const btn = document.getElementById('btnRepairAbsensi');
      if (btn) btn.style.display = (currentUser.role === 'admin' || perms.includes('aksesRepairAbsensi')) ? 'inline-block' : 'none';

      openModal('modalFingerprintInfo');
    }

    function copyFpUrl() {
      var txt = document.getElementById('fpWebhookUrl').textContent;
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(txt).then(function () { showToast('URL Webhook tersalin!', 'success'); });
      } else { showToast('Salin URL secara manual.', 'success'); }
    }

    function testFingerprintWebhook() {
      var testRec = [{
        fingerprintId: 'FP-TEST-001', karyawanId: '', nama: 'Test User X900',
        divisi: 'Testing', jabatan: 'QA', jam: '08:00:00', tipe: 'IN', tanggal: getTodayStr()
      }];
      document.getElementById('fpTestStatus').textContent = 'Mengirim test data...';
      google.script.run.withSuccessHandler(function (res) {
        if (res.success) {
          document.getElementById('fpTestStatus').textContent = 'Berhasil! Added: ' + res.added + ', Skipped: ' + res.skipped;
          showToast('Test fingerprint berhasil!', 'success');
          loadAbsensiKaryawan();
        } else {
          document.getElementById('fpTestStatus').textContent = 'Gagal: ' + res.message;
        }
      }).syncFingerprintData(testRec);
    }
    function repairAbsensiDataFront() {
      // RBAC Double Check
      const perms = JSON.parse(currentUser.permissions || '[]');
      if (currentUser.role !== 'admin' && !perms.includes('aksesRepairAbsensi')) {
        Swal.fire('Akses Ditolak', 'Anda tidak memiliki izin untuk melakukan perbaikan data absensi masal.', 'error');
        return;
      }

      if (!confirm('Peringatan: Fungsi ini akan memindai seluruh sejarah absensi dan memperbarui Nama/Divisi yang belum lengkap berdasarkan master Data Karyawan. Lanjutkan?')) return;

      const btn = document.getElementById('btnRepairAbsensi');
      if (btn) { btn.disabled = true; btn.textContent = '⏱️ Memproses...'; }
      document.getElementById('fpTestStatus').textContent = 'Memulai perbaikan data masal...';

      google.script.run.withSuccessHandler(function (res) {
        if (btn) { btn.disabled = false; btn.textContent = '🔄 Sinkronkan Data Lama'; }
        if (res.success) {
          Swal.fire('Berhasil', res.message, 'success');
          document.getElementById('fpTestStatus').textContent = res.message;
          loadAbsensiKaryawan();
        } else {
          Swal.fire('Gagal', res.message, 'error');
          document.getElementById('fpTestStatus').textContent = 'Error: ' + res.message;
        }
      }).repairAbsensiData(currentUser.username);
    }

    // ==========================================
    // JADWAL SHIFT & ROSTER BULANAN
    // ==========================================



    function renderShiftRosterTable(res) {
      const data = res.data || [];
      const days = (res.headers || []).filter(h => !isNaN(h)); // Get 1-31

      if (data.length === 0) {
        document.getElementById('rosterTableContainer').innerHTML = '<div class="p-5 text-center text-muted"><i class="bi bi-info-circle"></i> Belum ada data roster untuk bulan ini. Silakan Impor Excel atau tambah karyawan terlebih dahulu.</div>';
        return;
      }

      // Helper: get shift style
      function getShiftStyle(val) {
        const v = (val || '').trim().toUpperCase();
        if (v === 'PAGI')  return 'background:rgba(16,185,129,0.12); color:var(--green); font-weight:700;';
        if (v === 'MALAM') return 'background:rgba(14,165,233,0.12); color:var(--teal); font-weight:700;';
        if (v === 'SD')    return 'background:rgba(251,191,36,0.15); color:#d97706; font-weight:700;';
        if (v === 'STD')   return 'background:rgba(139,92,246,0.15); color:#7c3aed; font-weight:700;';
        if (v === 'IJIN')  return 'background:rgba(249,115,22,0.15); color:#ea580c; font-weight:700;';
        if (v === 'ALFA')  return 'background:rgba(239,68,68,0.15); color:var(--red); font-weight:700;';
        if (v === 'OFF')   return 'background:rgba(100,116,139,0.15); color:#64748b; font-weight:700;';
        return '';
      }

      let html = '<table class="roster-table"><thead><tr>';
      html += '<th class="sticky-col" style="min-width:160px;">Nama Karyawan</th>';
      html += '<th style="min-width:100px; text-align:center;">Lokasi</th>';

      const dayNames = ['Min', 'Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab'];
      const [year, month] = res.monthYear.split('-').map(Number);

      days.forEach(d => {
        const dateObj = new Date(year, month - 1, parseInt(d));
        const dayName = dayNames[dateObj.getDay()];
        const isSunday = dateObj.getDay() === 0;
        const color = isSunday ? 'color:var(--red);' : '';

        html += `<th style="${color} min-width:52px; text-align:center;">
                  <div style="font-size:10px; opacity:0.7;">${dayName}</div>
                  <div>${d}</div>
                </th>`;
      });
      html += '</tr></thead><tbody>';

      // Build employee lookup from karyawanData (global)
      const empLookup = {};
      if (Array.isArray(karyawanData)) {
        karyawanData.forEach(k => {
          empLookup[k.nama ? k.nama.toLowerCase().trim() : ''] = k;
        });
      }

      data.forEach(row => {
        const nameKey = (row.nama || '').toLowerCase().trim();
        const empData = empLookup[nameKey];
        const lokasi = row.lokasi || (empData ? (empData.cabang || '-') : '-');

        html += '<tr>';
        html += `<td class="sticky-col"><div style="font-weight:600;">${row.nama || '-'}</div></td>`;
        html += `<td style="text-align:center; font-size:11px;"><span style="background:rgba(255,255,255,0.06); padding:2px 8px; border-radius:12px; white-space:nowrap;">${lokasi}</span></td>`;

        days.forEach(d => {
          const val = row[d] || '';
          const shiftStyle = getShiftStyle(val) + ' cursor:pointer;';

          // Indikator Kehadiran
          let attIcon = '';
          const attMap = res.attendanceMap || {};
          const empName = row.nama || '';
          const dayNum = parseInt(d);

          if (attMap[empName] && attMap[empName][dayNum]) {
            const status = attMap[empName][dayNum];
            if (status === 'Terlambat') {
              attIcon = ' <i class="bi bi-exclamation-triangle-fill text-warning" title="Terlambat" style="font-size:9px;"></i>';
            } else {
              attIcon = ' <i class="bi bi-check-circle-fill text-success" title="Hadir" style="font-size:9px;"></i>';
            }
          }

          html += `<td style="${shiftStyle}" onclick="openShiftCellMenu('${res.monthYear}', '${row.nama.replace(/'/g,"\\'")}', '${d}', this)">
                    <div class="d-flex align-items-center justify-content-center gap-1">
                      <span style="font-size:11px;">${val}</span>${attIcon}
                    </div>
                  </td>`;
        });
        html += '</tr>';
      });

      html += '</tbody></table>';
      document.getElementById('rosterTableContainer').innerHTML = html;
      document.getElementById('rosterUpdateStatus').textContent = 'Terakhir diperbarui: ' + new Date().toLocaleTimeString();
      // Re-apply name search filter if active
      if (typeof filterRosterByName === 'function') filterRosterByName();
    }

    function handleImportRoster(input) {
      if (!input.files || !input.files[0]) return;
      if (typeof XLSX === 'undefined') {
        toast('Library Excel (SheetJS) tidak termuat. Periksa koneksi internet Anda atau refresh halaman.', 'error');
        return;
      }
      const file = input.files[0];
      const monthYear = document.getElementById('rosterFilterMonth').value;

      if (!monthYear) {
        showToast('Pilih bulan terlebih dahulu!', 'warning');
        input.value = '';
        return;
      }

      const reader = new FileReader();
      reader.onload = function (e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const sheetName = workbook.SheetNames[0];
          const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

          if (rows.length < 2) {
            showToast('File Excel kosong atau tidak valid!', 'error');
            return;
          }

          // Parsing Header: Kolom A: Nama, Kolom B+: Tanggal 1, 2, ...
          const headers = rows[0];
          const rosterData = [];

          for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            const rawNama = row[0];
            if (!rawNama) continue;

            const obj = { nama: String(rawNama).trim() };
            for (let j = 1; j < headers.length; j++) {
              let dayNum = "";
              const h = headers[j];

              if (h instanceof Date) {
                dayNum = String(h.getDate());
              } else if (h !== null && h !== undefined) {
                // Ambil angka saja dari header (misal "1" atau "1 April" -> "1")
                const match = String(h).match(/\d+/);
                if (match) dayNum = match[0];
              }

              if (dayNum && !isNaN(dayNum)) {
                const cellVal = row[j] ? String(row[j]).trim().toUpperCase() : '';
                if (cellVal !== '') {
                  obj[dayNum] = cellVal;
                }
              }
            }
            rosterData.push(obj);
          }

          if (rosterData.length === 0) {
            showToast('Tidak ada data karyawan ditemukan di Excel!', 'error');
            return;
          }

          // --- Sinkronisasi & Deduplikasi Nama (Client-side) ---
          const uniqueMap = new Map();
          let duplicatesFound = 0;
          rosterData.forEach(item => {
            const nameKey = item.nama.toLowerCase();
            if (uniqueMap.has(nameKey)) {
              // Jika ada nama yang sama dalam satu file, gabungkan datanya (Merge)
              const existing = uniqueMap.get(nameKey);
              Object.assign(existing, item);
              duplicatesFound++;
              console.warn('Duplicate entry found and merged in file:', item.nama);
            } else {
              uniqueMap.set(nameKey, item);
            }
          });
          const finalRosterData = Array.from(uniqueMap.values());

          if (duplicatesFound > 0) {
            showToast(`⚠️ ${duplicatesFound} baris duplikat untuk nama yang sama digabungkan secara otomatis.`, 'warning');
          }
          // -----------------------------------------------------

          showToast('⏳ Mengimpor ' + finalRosterData.length + ' data roster...', 'info');
          google.script.run.withSuccessHandler(function (res) {
            if (res.success) {
              showToast('✅ Berhasil impor roster!', 'success');
              loadShiftRoster();
            } else {
              showToast('Gagal Impor: ' + res.message, 'error');
            }
            input.value = '';
          }).importShiftRoster(monthYear, finalRosterData, currentUser.username);
        } catch (err) {
          showToast('Error Membaca Excel: ' + err.message, 'error');
          input.value = '';
        }
      };
      reader.readAsArrayBuffer(file);
    }

    function downloadRosterTemplate() {
      const monthYear = document.getElementById('rosterFilterMonth').value;
      if (!monthYear) {
        showToast('Pilih bulan terlebih dahulu!', 'warning');
        return;
      }

      const [year, month] = monthYear.split('-');
      const daysInMonth = new Date(year, month, 0).getDate();

      const headers = ['Nama'];
      for (let i = 1; i <= daysInMonth; i++) {
        headers.push(i);
      }

      // Data Template
      const data = [
        headers,
        ['KARYAWAN A (Contoh: PAGI/MALAM/SD/STD/IJIN/ALFA)', 'PAGI', 'PAGI', 'MALAM', 'SD', 'STD', 'IJIN', 'ALFA'],
        ['KARYAWAN B', 'MALAM', 'MALAM', 'PAGI', 'PAGI', 'IJIN']
      ];

      try {
        const ws = XLSX.utils.aoa_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Roster");
        XLSX.writeFile(wb, `Template_Roster_${monthYear}.xlsx`);
        showToast('Template Excel berhasil diunduh.', 'success');
      } catch (err) {
        showToast('Gagal membuat template: ' + err.message, 'error');
      }
    }

    function toggleRosterSettings() {
      const panel = document.getElementById('rosterSettingsPanel');
      if (panel.style.display === 'none') {
        panel.style.display = 'block';
        loadRosterSettings();
        switchRosterSettingsTab('global');
      } else {
        panel.style.display = 'none';
      }
    }

    function switchRosterSettingsTab(tab) {
      document.getElementById('rosterTabGlobal').style.display   = tab === 'global'  ? 'block' : 'none';
      document.getElementById('rosterTabJabatan').style.display  = tab === 'jabatan' ? 'block' : 'none';
      document.getElementById('tabRosterGlobal').className  = 'btn btn-' + (tab === 'global'  ? 'primary' : 'ghost') + ' btn-sm';
      document.getElementById('tabRosterJabatan').className = 'btn btn-' + (tab === 'jabatan' ? 'primary' : 'ghost') + ' btn-sm';
      if (tab === 'jabatan') {
        loadJabatanShiftList();
        populateJabatanDatalist();
      }
    }

    function populateJabatanDatalist() {
      const dl = document.getElementById('jabatanDatalist');
      if (!dl) return;
      const jabs = new Set();
      if (Array.isArray(karyawanData)) {
        karyawanData.forEach(k => { if (k.jabatan) jabs.add(k.jabatan.trim()); });
      }
      dl.innerHTML = '';
      Array.from(jabs).sort().forEach(j => {
        const opt = document.createElement('option');
        opt.value = j;
        dl.appendChild(opt);
      });
    }

    function loadJabatanShiftList() {
      const tb = document.getElementById('jabatanShiftBody');
      tb.innerHTML = '<tr><td colspan="7" class="text-center text-muted py-2"><span class="spinner-border spinner-border-sm"></span> Memuat...</td></tr>';
      google.script.run.withSuccessHandler(function(res) {
        if (!res.success) { tb.innerHTML = '<tr><td colspan="7" class="text-center text-danger">Gagal: ' + res.message + '</td></tr>'; return; }
        const list = res.data || [];
        if (!list.length) {
          tb.innerHTML = '<tr><td colspan="7" class="text-center text-muted py-3">Belum ada pengaturan jam per jabatan. Tambahkan di atas.</td></tr>';
          return;
        }
        tb.innerHTML = list.map(j => {
          const aktifBadge = String(j.aktif).toLowerCase() === 'ya' || j.aktif === true
            ? '<span style="color:var(--green);font-weight:700;">✔ Ya</span>'
            : '<span style="color:var(--red);">✘ Tidak</span>';
          const jamIn  = j.jamMasuk  instanceof Date ? j.jamMasuk.toTimeString().slice(0,5)  : String(j.jamMasuk  || '-');
          const jamOut = j.jamPulang instanceof Date ? j.jamPulang.toTimeString().slice(0,5) : String(j.jamPulang || '-');
          return `<tr>
            <td><strong>${j.namaJadwal || '-'}</strong></td>
            <td><span style="background:rgba(99,102,241,0.12); color:#6366f1; padding:2px 8px; border-radius:8px; font-size:11px;">${j.divisi || '-'}</span></td>
            <td style="font-family:monospace; font-weight:700; color:var(--green);">${jamIn}</td>
            <td style="font-family:monospace; font-weight:700; color:var(--accent);">${jamOut}</td>
            <td>${j.toleransiMenit || 0} mnt</td>
            <td>${aktifBadge}</td>
            <td style="white-space:nowrap;">
              <button class="btn btn-ghost btn-sm py-0" onclick="editJabatanShift('${j.id}','${(j.namaJadwal||'').replace(/'/g,"\\'")}','${(j.divisi||'').replace(/'/g,"\\'")}','${jamIn}','${jamOut}','${j.toleransiMenit||0}','${j.aktif||'Ya'}')">✏️</button>
              <button class="btn btn-danger btn-sm py-0" onclick="delJabatanShift('${j.id}')">🗑</button>
            </td>
          </tr>`;
        }).join('');
      }).getJadwalShift('');
    }

    function saveJabatanShift() {
      const id      = document.getElementById('jabShiftId').value.trim();
      const nama    = document.getElementById('jabShiftNama').value.trim();
      const jabatan = document.getElementById('jabShiftJabatan').value.trim();
      const jamIn   = document.getElementById('jabShiftIn').value;
      const jamOut  = document.getElementById('jabShiftOut').value;
      const tol     = document.getElementById('jabShiftTol').value || '0';
      const aktif   = document.getElementById('jabShiftAktif').value;

      if (!nama || !jabatan || !jamIn || !jamOut) {
        return showToast('Nama jadwal, jabatan, jam masuk & pulang wajib diisi.', 'warning');
      }

      showToast('⏳ Menyimpan...', 'info');
      google.script.run.withSuccessHandler(function(res) {
        if (res.success) {
          showToast('✅ Jadwal jabatan berhasil disimpan!', 'success');
          resetJabatanShiftForm();
          loadJabatanShiftList();
        } else {
          showToast('Gagal: ' + res.message, 'error');
        }
      // namaJadwal=nama, divisi=jabatan (kita simpan jabatan di kolom divisi untuk reuse fungsi yg sama)
      }).saveJadwalShift(id || null, nama, jabatan, 'JABATAN', jamIn, jamOut, tol, aktif);
    }

    function editJabatanShift(id, nama, jabatan, jamIn, jamOut, tol, aktif) {
      document.getElementById('jabShiftId').value      = id;
      document.getElementById('jabShiftNama').value    = nama;
      document.getElementById('jabShiftJabatan').value = jabatan;
      document.getElementById('jabShiftIn').value      = jamIn;
      document.getElementById('jabShiftOut').value     = jamOut;
      document.getElementById('jabShiftTol').value     = tol;
      document.getElementById('jabShiftAktif').value   = (String(aktif).toLowerCase() === 'ya' || aktif === true) ? 'Ya' : 'Tidak';
      document.getElementById('jabShiftNama').focus();
    }

    function resetJabatanShiftForm() {
      document.getElementById('jabShiftId').value      = '';
      document.getElementById('jabShiftNama').value    = '';
      document.getElementById('jabShiftJabatan').value = '';
      document.getElementById('jabShiftIn').value      = '08:00';
      document.getElementById('jabShiftOut').value     = '17:00';
      document.getElementById('jabShiftTol').value     = '0';
      document.getElementById('jabShiftAktif').value   = 'Ya';
    }

    function delJabatanShift(id) {
      if (!confirm('Hapus pengaturan jadwal ini?')) return;
      google.script.run.withSuccessHandler(function(res) {
        if (res.success) { showToast('Dihapus.', 'success'); loadJabatanShiftList(); }
        else showToast('Gagal: ' + res.message, 'error');
      }).deleteJadwalShift(id, currentUser.username);
    }

    function loadRosterSettings() {
      google.script.run.withSuccessHandler(function (res) {
        if (res.success) {
          document.getElementById('setRosterPagiIn').value = res.data.pagiIn || '08:00';
          document.getElementById('setRosterPagiOut').value = res.data.pagiOut || '17:00';
          document.getElementById('setRosterMalamIn').value = res.data.malamIn || '20:00';
          document.getElementById('setRosterMalamOut').value = res.data.malamOut || '05:00';
          document.getElementById('setRosterToleransi').value = res.data.toleransi || 0;
        }
      }).getRosterSettings();
    }

    function saveRosterSettings() {
      const settings = {
        pagiIn: document.getElementById('setRosterPagiIn').value,
        pagiOut: document.getElementById('setRosterPagiOut').value,
        malamIn: document.getElementById('setRosterMalamIn').value,
        malamOut: document.getElementById('setRosterMalamOut').value,
        toleransi: document.getElementById('setRosterToleransi').value
      };

      showToast('⏳ Menyimpan pengaturan roster...', 'info');
      google.script.run.withSuccessHandler(function (res) {
        if (res.success) {
          showToast('✅ Pengaturan roster berhasil disimpan!', 'success');
          document.getElementById('rosterSettingsPanel').style.display = 'none';
        } else {
          showToast('Gagal simpan setelan: ' + res.message, 'error');
        }
      }).saveRosterSettings(settings, currentUser.username);
    }

    let _rosterSpecialDates = {};

    function openRosterSpecialDatesModal() {
      showToast('⏳ Memuat jadwal khusus...', 'info');
      google.script.run.withSuccessHandler(function (res) {
        if (res.success) {
          _rosterSpecialDates = res.data.specialDates || {};
          renderSpecialDateList();
          openModal('modalRosterSpecialDates');
        }
      }).getRosterSettings();
    }

    function renderSpecialDateList() {
      const tb = document.getElementById('specialRosterTableBody');
      tb.innerHTML = '';
      const dates = Object.keys(_rosterSpecialDates).sort((a, b) => new Date(b) - new Date(a));

      if (dates.length === 0) {
        tb.innerHTML = '<tr><td colspan="5" class="text-center text-muted py-3">Belum ada jadwal khusus</td></tr>';
        return;
      }

      dates.forEach(d => {
        const s = _rosterSpecialDates[d];
        tb.innerHTML += `<tr>
          <td><strong>${formatDate(d)}</strong></td>
          <td><span class="badge bg-light text-dark">${s.pagiIn} - ${s.pagiOut}</span></td>
          <td><span class="badge bg-light text-dark">${s.malamIn} - ${s.malamOut}</span></td>
          <td>${s.toleransi}m</td>
          <td>
            <button class="btn btn-danger btn-sm p-0" style="width:24px; height:24px;" onclick="removeSpecialDateRecord('${d}')" title="Hapus">×</button>
          </td>
        </tr>`;
      });
    }

    function addSpecialDateRecord() {
      const date = document.getElementById('specialRosterDate').value;
      if (!date) return showToast('Pilih tanggal!', 'warning');

      _rosterSpecialDates[date] = {
        pagiIn: document.getElementById('specialPagiIn').value,
        pagiOut: document.getElementById('specialPagiOut').value,
        malamIn: document.getElementById('specialMalamIn').value,
        malamOut: document.getElementById('specialMalamOut').value,
        toleransi: document.getElementById('specialToleransi').value
      };

      renderSpecialDateList();
      showToast('Ditambahkan ke daftar (Belum disimpan ke server)', 'info');
    }

    function removeSpecialDateRecord(date) {
      if (confirm('Hapus jadwal khusus untuk tanggal ' + date + '?')) {
        delete _rosterSpecialDates[date];
        renderSpecialDateList();
      }
    }

    function saveSpecialDatesToServer() {
      const btn = document.getElementById('btnSaveSpecialRoster');
      btn.disabled = true; btn.innerHTML = '⏳ Menyimpan...';

      const jsonStr = JSON.stringify(_rosterSpecialDates);
      google.script.run.withSuccessHandler(function (res) {
        btn.disabled = false; btn.innerHTML = '💾 Simpan Semua Perubahan';
        if (res.success) {
          showToast('✅ Berhasil menyimpan jadwal khusus!', 'success');
          closeModal('modalRosterSpecialDates');
        } else {
          showToast('Gagal simpan: ' + res.message, 'error');
        }
      }).saveRosterSpecialDates(jsonStr, currentUser.username);
    }

    // Shift cycle order for click-cycling
    const SHIFT_CYCLE = ['PAGI', 'MALAM', 'SD', 'STD', 'IJIN', 'ALFA', 'OFF', ''];

    function getShiftCellStyle(val) {
      const v = (val || '').trim().toUpperCase();
      if (v === 'PAGI')  return 'background:rgba(16,185,129,0.12); color:var(--green); font-weight:700; cursor:pointer;';
      if (v === 'MALAM') return 'background:rgba(14,165,233,0.12); color:var(--teal); font-weight:700; cursor:pointer;';
      if (v === 'SD')    return 'background:rgba(251,191,36,0.15); color:#d97706; font-weight:700; cursor:pointer;';
      if (v === 'STD')   return 'background:rgba(139,92,246,0.15); color:#7c3aed; font-weight:700; cursor:pointer;';
      if (v === 'IJIN')  return 'background:rgba(249,115,22,0.15); color:#ea580c; font-weight:700; cursor:pointer;';
      if (v === 'ALFA')  return 'background:rgba(239,68,68,0.15); color:var(--red); font-weight:700; cursor:pointer;';
      if (v === 'OFF')   return 'background:rgba(100,116,139,0.15); color:#64748b; font-weight:700; cursor:pointer;';
      return 'cursor:pointer;';
    }

    function openShiftCellMenu(monthYear, name, day, cell) {
      // Remove any existing menu
      const existingMenu = document.getElementById('_shiftCellMenu');
      if (existingMenu) existingMenu.remove();

      const currentVal = (cell.querySelector('span') ? cell.querySelector('span').textContent : cell.textContent).trim().toUpperCase();

      const menu = document.createElement('div');
      menu.id = '_shiftCellMenu';
      menu.style.cssText = 'position:fixed; z-index:9999; background:var(--card-bg,#1e1e2e); border:1px solid var(--border-color,#333); border-radius:10px; box-shadow:0 8px 24px rgba(0,0,0,0.4); padding:6px; min-width:120px;';

      SHIFT_CYCLE.forEach(shift => {
        const opt = document.createElement('div');
        const isActive = shift === currentVal || (shift === '' && currentVal === '');
        opt.style.cssText = `padding:7px 14px; border-radius:7px; cursor:pointer; font-size:12px; font-weight:600; display:flex; align-items:center; gap:6px; ${isActive ? 'background:rgba(255,255,255,0.08);' : ''}`;
        opt.innerHTML = `<span style="${getShiftCellStyle(shift).replace('cursor:pointer;','')} padding:2px 8px; border-radius:10px;">${shift || '(Kosong)'}</span>`;
        opt.onmouseover = () => opt.style.background = 'rgba(255,255,255,0.06)';
        opt.onmouseout  = () => opt.style.background = isActive ? 'rgba(255,255,255,0.08)' : '';
        opt.onclick = (e) => {
          e.stopPropagation();
          menu.remove();
          applyRosterCellValue(monthYear, name, day, shift, cell);
        };
        menu.appendChild(opt);
      });

      // Position near the cell
      const rect = cell.getBoundingClientRect();
      menu.style.top  = Math.min(rect.bottom + 4, window.innerHeight - 260) + 'px';
      menu.style.left = Math.min(rect.left, window.innerWidth - 150) + 'px';

      document.body.appendChild(menu);

      // Close on outside click
      setTimeout(() => {
        document.addEventListener('click', function closeMenu() {
          menu.remove();
          document.removeEventListener('click', closeMenu);
        }, { once: true });
      }, 50);
    }

    function applyRosterCellValue(monthYear, name, day, newVal, cell) {
      const oldStyle = cell.getAttribute('style');
      const spanEl = cell.querySelector('span');
      const oldText = spanEl ? spanEl.textContent : '';

      // Optimistic UI update
      if (spanEl) spanEl.textContent = newVal;
      cell.setAttribute('style', getShiftCellStyle(newVal));

      google.script.run.withSuccessHandler(function (res) {
        if (!res.success) {
          showToast('Gagal update roster: ' + res.message, 'error');
          if (spanEl) spanEl.textContent = oldText;
          cell.setAttribute('style', oldStyle);
        }
      }).updateRosterCell(monthYear, name, day, newVal, currentUser.username);
    }

    function toggleRosterCell(monthYear, name, day, cell) {
      // Legacy – now delegates to openShiftCellMenu
      openShiftCellMenu(monthYear, name, day, cell);
    }


    function loadShiftRoster() {
      const monthYear = document.getElementById('rosterFilterMonth').value;
      if (!monthYear) {
        // Fallback initialization if needed
        const now = new Date();
        const currentMonth = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
        document.getElementById('rosterFilterMonth').value = currentMonth;
        return loadShiftRoster();
      }

      // Apply lokasi filter client-side after loading
      const lokasiFilter = (document.getElementById('rosterFilterLokasi') ? document.getElementById('rosterFilterLokasi').value : '').trim();

      const container = document.getElementById('rosterTableContainer');
      container.innerHTML = '<div class="p-5 text-center text-muted"><div class="spinner-border text-primary mb-3"></div><br>Memuat roster ' + monthYear + '...</div>';

      google.script.run.withSuccessHandler(function (res) {
        if (res.success) {
          // Client-side filter by lokasi (cabang)
          if (lokasiFilter && res.data && res.data.length) {
            // Build employee lookup
            const empLookup = {};
            if (Array.isArray(karyawanData)) {
              karyawanData.forEach(k => {
                empLookup[(k.nama || '').toLowerCase().trim()] = k;
              });
            }
            res.data = res.data.filter(row => {
              const lok = row.lokasi || (empLookup[(row.nama||'').toLowerCase().trim()] ? (empLookup[(row.nama||'').toLowerCase().trim()].cabang || '') : '');
              return (lok || '').toLowerCase() === lokasiFilter.toLowerCase();
            });
          }
          renderShiftRosterTable(res);
        } else {
          container.innerHTML = '<div class="p-5 text-center text-danger">Gagal memuat roster: ' + res.message + '</div>';
        }
      }).getShiftRoster(monthYear, currentUser.username);
    }

    function loadLokasiRosterFilter() {
      const sel = document.getElementById('rosterFilterLokasi');
      if (!sel) return;
      // Collect unique locations from karyawanData
      const locs = new Set();
      if (Array.isArray(karyawanData)) {
        karyawanData.forEach(k => {
          if (k.cabang && k.cabang.trim()) locs.add(k.cabang.trim());
        });
      }
      const currentVal = sel.value;
      sel.innerHTML = '<option value="">Semua Lokasi</option>';
      Array.from(locs).sort().forEach(loc => {
        const opt = document.createElement('option');
        opt.value = loc; opt.textContent = loc;
        if (loc === currentVal) opt.selected = true;
        sel.appendChild(opt);
      });
    }

    function filterRosterByName() {
      const q = (document.getElementById('rosterSearchNama').value || '').toLowerCase().trim();
      const tbody = document.querySelector('#rosterTableContainer table tbody');
      if (!tbody) return;
      Array.from(tbody.querySelectorAll('tr')).forEach(tr => {
        const nameCell = tr.querySelector('td.sticky-col');
        const name = nameCell ? nameCell.textContent.toLowerCase() : '';
        tr.style.display = (!q || name.includes(q)) ? '' : 'none';
      });
    }

    function exportRosterToExcel() {
      if (typeof XLSX === 'undefined') {
        showToast('Library Excel (SheetJS) tidak termuat. Refresh halaman lalu coba lagi.', 'error');
        return;
      }

      const table = document.querySelector('#rosterTableContainer table');
      if (!table) {
        showToast('Tidak ada data roster untuk diekspor.', 'warning');
        return;
      }

      const monthYear = document.getElementById('rosterFilterMonth').value || 'roster';
      const lokasiFilter = (document.getElementById('rosterFilterLokasi').value || '').trim();
      const namaFilter   = (document.getElementById('rosterSearchNama').value || '').trim();

      // Build array-of-arrays from visible table rows
      const aoa = [];

      // Header row
      const headerRow = [];
      table.querySelectorAll('thead tr th').forEach(th => {
        headerRow.push(th.textContent.trim().replace(/\s+/g,' '));
      });
      aoa.push(headerRow);

      // Data rows (only visible ones — respects name & lokasi filter)
      table.querySelectorAll('tbody tr').forEach(tr => {
        if (tr.style.display === 'none') return;
        const row = [];
        tr.querySelectorAll('td').forEach(td => {
          // Get text only (strip attendance icons)
          const span = td.querySelector('span:first-child');
          row.push(span ? span.textContent.trim() : td.textContent.trim());
        });
        aoa.push(row);
      });

      if (aoa.length <= 1) {
        showToast('Tidak ada baris data yang terlihat untuk diekspor.', 'warning');
        return;
      }

      try {
        const ws = XLSX.utils.aoa_to_sheet(aoa);

        // Auto column widths
        const colWidths = headerRow.map((h, i) => {
          const maxLen = aoa.reduce((m, r) => Math.max(m, String(r[i] || '').length), h.length);
          return { wch: Math.min(maxLen + 2, 20) };
        });
        ws['!cols'] = colWidths;

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Roster ' + monthYear);

        let fileName = `Roster_${monthYear}`;
        if (lokasiFilter) fileName += `_${lokasiFilter}`;
        if (namaFilter)   fileName += `_${namaFilter}`;
        fileName += '.xlsx';

        XLSX.writeFile(wb, fileName);
        showToast('✅ Berhasil ekspor roster ke Excel!', 'success');
      } catch (err) {
        showToast('Gagal ekspor: ' + err.message, 'error');
      }
    }

    function loadAbsenDivisiList() {
      google.script.run.withSuccessHandler(function (res) {
        if (!res || !res.success) return;
        var sel = document.getElementById('absenFilterDiv');
        if (!sel) return;
        sel.innerHTML = '<option value="">Semua Divisi</option>';
        (res.data || []).forEach(function (d) {
          var opt = document.createElement('option');
          opt.value = d; opt.textContent = d;
          sel.appendChild(opt);
        });
      }).getDivisiList();
    }

    // Override showPage untuk auto-load modul absensi
    (function () {
      var origSP = window.showPage;
      if (typeof origSP === 'function') {
        window.showPage = function (name) {
          origSP(name);
          if (name === 'absensiKaryawan') {
            var el = document.getElementById('absenFilterTgl');
            if (el && !el.value) el.value = getTodayStr();
            loadAbsensiKaryawan();
            loadAbsenDivisiList();
          }
          if (name === 'jadwalShift') {
            // Auto set current month if empty
            const filter = document.getElementById('rosterFilterMonth');
            if (filter && !filter.value) {
              const now = new Date();
              filter.value = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
            }
            // Populate lokasi filter from karyawan data
            if (!karyawanData || !karyawanData.length) {
              google.script.run.withSuccessHandler(function(res) {
                if (res && res.karyawan && res.karyawan.success) {
                  karyawanData = res.karyawan.data;
                }
                loadLokasiRosterFilter();
                loadShiftRoster();
              }).getKaryawanFullData();
            } else {
              loadLokasiRosterFilter();
              loadShiftRoster();
            }
          }
          if (name === 'handover') {
            loadStockControl();
            loadDashboardStock();
          }
        };
      }
    })();
    // ==========================================
    // STOCK OPNAME ASSET LOGIC
    // ==========================================
    let soAssetSession = null;
    let soAssetLogs = [];
    let soAssetUnscanned = [];
    let soAssetScanner = null;

    function switchAwTab(tab) {
      document.getElementById('tabAwList').classList.remove('active');
      document.getElementById('tabAwOpname').classList.remove('active');
      document.getElementById('awTabListContent').style.display = 'none';
      document.getElementById('awTabOpnameContent').style.display = 'none';

      if (tab === 'list') {
        document.getElementById('tabAwList').classList.add('active');
        document.getElementById('awTabListContent').style.display = 'block';
        renderAssetWarehouseTable();
      } else if (tab === 'opname') {
        document.getElementById('tabAwOpname').classList.add('active');
        document.getElementById('awTabOpnameContent').style.display = 'block';
        if (!document.getElementById('soAssetDate').value) {
          document.getElementById('soAssetDate').value = new Date().toISOString().split('T')[0];
        }
      }
    }

    function startSoAssetSession() {
      const tanggal = document.getElementById('soAssetDate').value;
      const divisi = document.getElementById('soAssetDivisi').value;
      if (!tanggal) return showToast('Pilih tanggal SO', 'warning');

      document.getElementById('btnStartSoAsset').innerHTML = '<span class="spinner-border spinner-border-sm"></span> Memulai...';
      document.getElementById('btnStartSoAsset').disabled = true;

      google.script.run.withSuccessHandler(function (res) {
        if (res.success) {
          soAssetSession = {
            id: res.id,
            tanggal: tanggal,
            divisi: divisi,
            totalAsset: res.totalAsset,
            terscan: 0
          };
          soAssetLogs = [];

          document.getElementById('btnStartSoAsset').style.display = 'none';
          document.getElementById('btnResetSoAsset').style.display = 'inline-block';
          document.getElementById('btnSubmitSoAsset').style.display = 'inline-block';
          document.getElementById('soAssetDivisi').disabled = true;
          document.getElementById('soAssetDate').disabled = true;
          document.getElementById('soAssetActivePanel').style.display = 'block';

          initSoAssetScanner();
          loadSoAssetUnscanned(divisi);
        } else {
          showToast(res.message, 'error');
          document.getElementById('btnStartSoAsset').innerHTML = '▶ Mulai Sesi SO';
          document.getElementById('btnStartSoAsset').disabled = false;
        }
      }).createAssetOpnameSession(tanggal, divisi, currentUser.username);
    }

    function resetSoAssetSession() {
      if (!confirm('Yakin ingin reset sesi ini? Data scan yang belum disubmit akan hilang (dihapus).')) return;
      if (soAssetSession && soAssetSession.id) {
        google.script.run.deleteAssetOpnameSession(soAssetSession.id);
      }
      stopSoAssetScanner();
      soAssetSession = null;
      soAssetLogs = [];
      soAssetUnscanned = [];

      document.getElementById('btnStartSoAsset').style.display = 'inline-block';
      document.getElementById('btnStartSoAsset').innerHTML = '▶ Mulai Sesi SO';
      document.getElementById('btnStartSoAsset').disabled = false;
      document.getElementById('btnResetSoAsset').style.display = 'none';
      document.getElementById('btnSubmitSoAsset').style.display = 'none';
      document.getElementById('btnSubmitSoAsset').disabled = true;
      document.getElementById('soAssetDivisi').disabled = false;
      document.getElementById('soAssetDate').disabled = false;
      document.getElementById('soAssetActivePanel').style.display = 'none';
    }

    function initSoAssetScanner() {
      if (soAssetScanner) return; // already initialized
      try {
        soAssetScanner = new Html5QrcodeScanner("soAssetQrReader", { fps: 10, qrbox: { width: 250, height: 250 } }, false);
        soAssetScanner.render(onSoAssetScanSuccess, onSoAssetScanFailure);
      } catch (e) {
        console.warn("Scanner error:", e);
      }
    }

    function stopSoAssetScanner() {
      if (soAssetScanner) {
        try { soAssetScanner.clear(); } catch (e) { }
        soAssetScanner = null;
      }
    }

    let isScanningSO = false;

    function playBeep() {
      try {
        const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
        const oscillator = audioCtx.createOscillator();
        const gainNode = audioCtx.createGain();
        oscillator.connect(gainNode);
        gainNode.connect(audioCtx.destination);
        oscillator.type = 'sine';
        oscillator.frequency.value = 800;
        gainNode.gain.setValueAtTime(0.1, audioCtx.currentTime);
        oscillator.start();
        setTimeout(function () {
          oscillator.stop();
        }, 150);
      } catch (e) {
        console.warn("AudioContext not supported");
      }
    }

    function onSoAssetScanSuccess(decodedText) {
      if (isScanningSO) return;
      isScanningSO = true;
      playBeep();
      processSoAssetScan(decodedText);
      setTimeout(() => { isScanningSO = false; }, 2000);
    }
    function onSoAssetScanFailure(error) { /* ignore */ }

    function scanImageFileSoAsset(inputElem) {
      if (!inputElem.files || inputElem.files.length === 0) return;
      const file = inputElem.files[0];

      let hiddenDiv = document.getElementById('hiddenSoAssetFileScanner');
      if (!hiddenDiv) {
        hiddenDiv = document.createElement('div');
        hiddenDiv.id = 'hiddenSoAssetFileScanner';
        hiddenDiv.style.display = 'none';
        document.body.appendChild(hiddenDiv);
      }

      const html5QrCode = new Html5Qrcode("hiddenSoAssetFileScanner");
      showToast('Memindai QR dari gambar...', 'info');

      html5QrCode.scanFile(file, true)
        .then(decodedText => {
          inputElem.value = '';
          playBeep();
          processSoAssetScan(decodedText);
        })
        .catch(err => {
          inputElem.value = '';
          showToast('QR Code tidak ditemukan pada gambar', 'warning');
        });
    }

    function loadSoAssetUnscanned(divisi) {
      document.getElementById('soAssetUnscannedList').innerHTML = '<tr><td colspan="5" class="text-center">Memuat daftar asset...</td></tr>';
      google.script.run.withSuccessHandler(function (res) {
        if (res.success) {
          soAssetUnscanned = res.data.filter(a => {
            if (divisi && divisi !== 'Semua' && a.divisi !== divisi) return false;
            return true;
          });
          renderSoAssetLists();
        }
      }).getAssetWarehouseData();
    }

    function processSoAssetScan(code) {
      if (!code) return;
      code = code.trim();
      document.getElementById('soAssetManualCode').value = '';

      if (!soAssetSession) return showToast('Sesi belum dimulai', 'warning');

      google.script.run.withSuccessHandler(function (res) {
        if (res.success) {
          showToast('Asset terscan: ' + res.asset.nama, 'success');
          // Tambah ke scanned list
          soAssetLogs.unshift({
            id: 'temp-' + Date.now(),
            assetCode: res.asset.code,
            assetNama: res.asset.nama,
            divisi: res.asset.divisi,
            qtyFisik: res.asset.qty, // default to system qty
            kondisi: 'Baik',
            scannedAt: new Date().toISOString()
          });
          soAssetSession.terscan++;

          // Hapus dari unscanned list
          soAssetUnscanned = soAssetUnscanned.filter(a => String(a.code).trim().toLowerCase() !== code.toLowerCase());

          document.getElementById('btnSubmitSoAsset').disabled = false;
          renderSoAssetLists();
        } else {
          showToast(res.message, res.alreadyScanned ? 'info' : 'error');
        }
      }).scanAssetForOpname(soAssetSession.id, code, 1, 'Baik', '', currentUser.username);
    }

    function renderSoAssetLists() {
      // Progress
      const total = soAssetSession.totalAsset || 1;
      const scanned = soAssetSession.terscan || 0;
      const pct = Math.min(100, Math.round((scanned / total) * 100));
      document.getElementById('soAssetProgressBar').style.width = pct + '%';
      document.getElementById('soAssetProgressText').textContent = scanned + ' / ' + total + ' Asset Ter-scan';

      // Update counters
      document.getElementById('countSoScanned').textContent = scanned;
      document.getElementById('countSoUnscanned').textContent = soAssetUnscanned.length;

      // Scanned List
      const tbodyS = document.getElementById('soAssetScannedList');
      tbodyS.innerHTML = '';
      if (soAssetLogs.length === 0) {
        tbodyS.innerHTML = '<tr><td colspan="7" class="text-center text-muted">Belum ada asset terscan</td></tr>';
      } else {
        soAssetLogs.forEach((l, i) => {
          const time = new Date(l.scannedAt).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
          tbodyS.innerHTML += `
            <tr>
              <td>${i + 1}</td>
              <td style="font-weight:700; color:var(--teal);">${l.assetCode}</td>
              <td>${l.assetNama}</td>
              <td><span class="badge bg-secondary">${l.divisi}</span></td>
              <td><input type="number" class="form-control form-control-sm" value="${l.qtyFisik}" style="width:70px;" onchange="updateSoQtyFisik('${l.id}', this.value)" readonly title="Qty saat ini (bisa diedit di versi berikutnya jika perlu)"></td>
              <td><span class="badge bg-success">${l.kondisi}</span></td>
              <td>${time}</td>
            </tr>
          `;
        });
      }

      // Unscanned List
      const tbodyU = document.getElementById('soAssetUnscannedList');
      tbodyU.innerHTML = '';
      if (soAssetUnscanned.length === 0) {
        tbodyU.innerHTML = '<tr><td colspan="5" class="text-center text-success">Semua asset telah terscan! 🎉</td></tr>';
      } else {
        soAssetUnscanned.forEach(a => {
          tbodyU.innerHTML += `
            <tr>
              <td style="font-weight:700;">${a.code}</td>
              <td>${a.nama}</td>
              <td>${a.divisi}</td>
              <td>${a.qty}</td>
              <td><span class="badge" style="background:var(--red);">${a.status}</span></td>
            </tr>
          `;
        });
      }
    }

    function updateSoQtyFisik(logId, newVal) {
      // Stub for local update, future enhancement if user wants to change qty after scan
    }

    function submitSoAssetReport() {
      if (!soAssetSession || !soAssetSession.id) return;
      if (!confirm('Ajukan laporan Stock Opname ini untuk approval?')) return;

      const btn = document.getElementById('btnSubmitSoAsset');
      btn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Mengajukan...';
      btn.disabled = true;

      // First, generate a detailed opname report (including negative stock details)
      google.script.run.withSuccessHandler(function (rpt) {
        if (!rpt || !rpt.success) {
          showToast(rpt && rpt.message ? rpt.message : 'Gagal membuat laporan opname', 'error');
          btn.innerHTML = '📋 Buat Laporan & Ajukan';
          btn.disabled = false;
          return;
        }

        // Show report details (negative stocks)
        showOpnameReport(rpt);

        // Then mark the session as submitted (Pending Approval)
        google.script.run.withSuccessHandler(function (res) {
          if (res.success) {
            showToast('Laporan berhasil diajukan!', 'success');
            stopSoAssetScanner();
            // Reset UI
            switchAwTab('list');
            document.getElementById('btnResetSoAsset').style.display = 'none';
            btn.style.display = 'none';
            document.getElementById('btnStartSoAsset').style.display = 'inline-block';
            document.getElementById('btnStartSoAsset').innerHTML = '▶ Mulai Sesi SO';
            document.getElementById('btnStartSoAsset').disabled = false;
            document.getElementById('soAssetDivisi').disabled = false;
            document.getElementById('soAssetDate').disabled = false;
            document.getElementById('soAssetActivePanel').style.display = 'none';
            soAssetSession = null;

            // Close the preview modal, show the Laporan SO list modal, and view the report details directly
            closeModal('modalOpnameReport');
            openModal('modalOpnameReportList');
            google.script.run.withSuccessHandler(function (reportsRes) {
              if (reportsRes && reportsRes.success) {
                auditReportsAll = reportsRes.data || [];
                renderOpnameReports(auditReportsAll);
                // Directly view the report details of the newly created report
                viewOpnameReport(rpt.id);
              } else {
                showToast('Gagal memuat laporan', 'error');
              }
            }).getAuditReports();
          } else {
            showToast(res.message, 'error');
            btn.innerHTML = '📋 Buat Laporan & Ajukan';
            btn.disabled = false;
          }
        }).submitAssetOpname(soAssetSession.id, currentUser.username);

      }).generateOpnameReport(soAssetSession.id, currentUser.username);
    }

    function showOpnameReport(rpt) {
      // rpt: { success:true, id, negative:[], missing:[] }
      if (!rpt || !rpt.success) return;
      const modal = document.getElementById('modalOpnameReport');
      if (!modal) return;
      document.getElementById('opnameReportId').textContent = rpt.id || '-';

      // Render negative stock
      const tblNeg = document.getElementById('opnameNegativeList'); tblNeg.innerHTML = '';
      const neg = rpt.negative || [];
      if (!neg.length) {
        tblNeg.innerHTML = '<tr><td colspan="7" class="text-center text-success">Tidak ada item stock minus 🎉</td></tr>';
      } else {
        neg.forEach((n, idx) => {
          tblNeg.innerHTML += `<tr><td>${idx + 1}</td><td>${n.code}</td><td>${n.nama}</td><td>${n.divisi || '-'}</td><td>${n.systemQty}</td><td>${n.physicalQty}</td><td>${n.difference}</td></tr>`;
        });
      }

      // Render missing assets
      const tblMiss = document.getElementById('opnameMissingList'); tblMiss.innerHTML = '';
      const miss = rpt.missing || [];
      if (!miss.length) {
        tblMiss.innerHTML = '<tr><td colspan="2" class="text-center text-success">Semua asset berhasil terscan! 🚀</td></tr>';
      } else {
        miss.forEach((m, idx) => {
          tblMiss.innerHTML += `<tr><td>${idx + 1}</td><td>${m}</td></tr>`;
        });
      }

      openModal('modalOpnameReport');
    }

    let auditReportsAll = [];
    function openOpnameReportList() {
      openModal('modalOpnameReportList');
      loadOpnameReports();
    }

    function loadOpnameReports() {
      google.script.run.withSuccessHandler(function (res) {
        if (!res || !res.success) return showToast('Gagal memuat laporan', 'error');
        auditReportsAll = res.data || [];
        renderOpnameReports(auditReportsAll);
      }).getAuditReports();
    }

    function renderOpnameReports(list) {
      const tbody = document.getElementById('opnameReportListBody'); tbody.innerHTML = '';
      if (!list || !list.length) {
        tbody.innerHTML = '<tr><td colspan="8" class="text-center text-muted">Belum ada laporan opname</td></tr>';
        return;
      }
      list.sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal)).forEach((r, idx) => {
        tbody.innerHTML += `<tr>
          <td>${idx + 1}</td>
          <td>${r.tanggal || '-'}</td>
          <td>${r.auditor || r.createdBy || '-'}</td>
          <td>${r.totalAsset || 0}</td>
          <td>${r.terscan || 0}</td>
          <td>${r.minus || 0}</td>
          <td>${r.status || '-'}</td>
          <td style="white-space:nowrap;"><button class="btn btn-ghost btn-sm" onclick="viewOpnameReport('${r.id}')">Lihat</button></td>
        </tr>`;
      });
    }

    function viewOpnameReport(id) {
      const rpt = (auditReportsAll || []).find(x => x.id === id);
      if (!rpt) return showToast('Laporan tidak ditemukan', 'error');
      // convert negativeList and missingAssets if they are strings
      try { if (typeof rpt.negativeList === 'string') rpt.negativeList = JSON.parse(rpt.negativeList || '[]'); } catch (e) { }
      try { if (typeof rpt.missingAssets === 'string') rpt.missingAssets = JSON.parse(rpt.missingAssets || '[]'); } catch (e) { }
      // reuse existing modal to show negative details
      showOpnameReport({ success: true, id: rpt.id, negative: rpt.negativeList || [], missing: rpt.missingAssets || [] });
    }
    // ==========================================

(function () {
    // Jalankan setelah DOM ready (bisa sebelum window.onload)
    function emergencyShowLogin() {
      try {
        var ls = document.getElementById('loadingScreen');
        var lp = document.getElementById('loginPage');
        var ap = document.getElementById('app');
        // Hanya paksa jika app tidak tampil (belum login)
        if (ap && ap.style.display === 'block') return;
        if (ls) ls.style.setProperty('display', 'none', 'important');
        if (lp) { lp.style.setProperty('display', 'flex', 'important'); lp.style.zIndex = '9999'; }
        console.info('[EMERGENCY-FAILSAFE] Login page forced visible.');
      } catch (e) { }
    }
    // Coba setelah 8 detik sebagai jaring pengaman terakhir
    setTimeout(emergencyShowLogin, 8000);
  })();