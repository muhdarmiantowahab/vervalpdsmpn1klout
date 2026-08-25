// ========================================
// 1. KONFIGURASI GLOBAL
// ========================================
const CONFIG = {
  SPREADSHEET_ID: '1wO7geWX5DihseZShrVciSmNyNxeB0AMGVA4wmQ7SiaE', // Ganti dengan ID Spreadsheet Anda
  SHEET_PENDAFTAR: 'Pendaftar',
  SHEET_ADMIN: 'Admin',
  SHEET_KONFIG: 'Konfigurasi',    // Konfigurasi Website (Hero, Footer)
  SHEET_BUILDER: 'Config',        // Konfigurasi Form Builder
  FOLDER_DRIVE_ID: '1Sq-DRFjtDp62pSipUhYaV1UC0oCHjSlW', // Ganti dengan ID Folder Drive Anda
  KODE_PREFIX: 'VERVALPD-'
};

// ========================================
// 2. ROUTING UTAMA & TEMPLATING
// ========================================
function doGet(e) {
  // 1. Baca Parameter URL
  e = e || { parameter: {} }; // fallback supaya tidak error
  const page = e.parameter.page || 'index';
  const sessionId = e.parameter.sessionId;
  
  // 2. Ambil Data Konfigurasi dari Spreadsheet (Server-Side)
  // Data ini akan langsung "ditempel" ke HTML sebelum dikirim ke browser
  const konfigurasiResponse = getKonfigurasiAplikasi();
  const konfigurasi = konfigurasiResponse.success ? konfigurasiResponse.data : {};
  
  let template;

  // 3. Logika Routing
  if (page === 'admin') {
    // Cek Session Admin
    const session = getUserSession(sessionId);
    if (!session.success || session.data.role !== 'Admin') {
      // Jika session tidak valid, redirect ke login
      const loginUrl = getScriptUrl() + '?page=login'; // Pastikan logic login di frontend menangani ini atau redirect ke home
      // Untuk SPA (Single Page App) kita biasanya melempar kembali ke index dengan mode guest
       template = HtmlService.createTemplateFromFile('index');
       template.isAdmin = false;
    } else {
       // Session Valid
       template = HtmlService.createTemplateFromFile('index');
       template.isAdmin = true;
    }
  } else {
    // Halaman Publik (Default)
    template = HtmlService.createTemplateFromFile('index');
    template.isAdmin = false;
  }
  
  // 4. INJECTION DATA KE TEMPLATE (KUNCI SSR)
  // Baris ini mengirim objek 'konfigurasi' ke file index.html
  // sehingga bisa dibaca oleh scriptlet <?= konfigurasi.namaApp ?>
  template.konfigurasi = konfigurasi;

  // 5. Render & Return HTML
  return template.evaluate()
    .setTitle(konfigurasi.appName + ' | ' + konfigurasi.footerJudul) // Title Browser Dinamis
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl(konfigurasi.logoUrl || 'https://cdn-icons-png.flaticon.com/512/2920/2920224.png');
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// ========================================
// 3. DYNAMIC FORM BUILDER (LOGIKA BARU)
// ========================================

// A. Ambil Konfigurasi Form untuk Frontend
function getFormConfig() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_BUILDER);
  
  // Jika sheet Config belum ada, buat default melalui initialize (tapi di sini kita handle basic)
  if (!sheet) {
    // Jika belum ada, kembalikan array kosong atau pancing initialize
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  // Return JSON bersih ke frontend
  return data.filter(r => r[0]).map(r => ({
    label: r[0], 
    type: r[1], 
    options: r[2], 
    width: r[3], 
    required: r[4]
  }));
}

// B. Simpan Konfigurasi Form dari Admin
// Simpan Konfigurasi Form Builder - PROTECTED
function saveFormConfig(token, configArray) {
  try {
    // 1. VALIDASI KEAMANAN
    // Memastikan hanya Admin yang sah yang bisa mengubah struktur form
    validateAdminToken(token);

    // 2. SETUP SPREADSHEET
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_BUILDER);
    
    // Jika sheet config belum ada, buat baru
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_BUILDER);
    }
    
    // 3. BERSIHKAN KONFIGURASI LAMA
    sheet.clear(); 
    
    // 4. TULIS KONFIGURASI BARU
    if (configArray && configArray.length > 0) {
      // Mapping object config ke format baris (Array of Arrays)
      const rows = configArray.map(c => [
        c.label, 
        c.type, 
        c.options || "", 
        c.width, 
        c.required
      ]);
      
      // Tulis ke sheet mulai dari baris 1, kolom 1
      sheet.getRange(1, 1, rows.length, 5).setValues(rows);
    }
    
    // 5. UPDATE HEADER SHEET PENDAFTAR
    // Penting: Fungsi ini akan menyesuaikan kolom di sheet database (Pendaftar)
    // agar sesuai dengan inputan baru
    updatePendaftarHeaders(configArray);
    
    return { success: true, message: 'Konfigurasi form berhasil disimpan dan database diperbarui.' };

  } catch (e) {
    // Menangkap error jika token invalid atau gagal akses spreadsheet
    return { success: false, message: e.toString() };
  }
}

// C. Sinkronisasi Header Sheet Pendaftar
function updatePendaftarHeaders(configArray) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_PENDAFTAR);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_PENDAFTAR);
  
  // Kolom Sistem (Kiri)
  const sysLeft = ['Kode Pendaftaran'];
  // Kolom Dinamis (Tengah)
  const dynamicCols = configArray.map(c => c.label);
  // Kolom Sistem (Kanan)
  const sysRight = ['Status', 'Tanggal Daftar', 'Catatan Admin'];
  
  // Gabung Semua
  const allHeaders = [...sysLeft, ...dynamicCols, ...sysRight];

  // Tulis Header Baru di Baris 1
  if (allHeaders.length > 0) {
    // Clear header lama agar bersih (Baris 1 saja)
    sheet.getRange(1, 1, 1, sheet.getMaxColumns()).clearContent();
    
    sheet.getRange(1, 1, 1, allHeaders.length)
         .setValues([allHeaders])
         .setBackground('#1E88E5')
         .setFontColor('#FFFFFF')
         .setFontWeight('bold');
  }
}

// ========================================
// 4. PENYIMPANAN DATA DINAMIS
// ========================================

function generateKodePendaftaran() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_PENDAFTAR);
  if(!sheet) return CONFIG.KODE_PREFIX + Math.floor(100000 + Math.random() * 900000);

  // Ambil semua kode di kolom A
  const lastRow = sheet.getLastRow();
  let existingCodes = [];
  if (lastRow > 1) {
    existingCodes = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  }

  let newCode;
  let isUnique = false;
  while (!isUnique) {
    newCode = CONFIG.KODE_PREFIX + Math.floor(100000 + Math.random() * 900000);
    if (!existingCodes.includes(newCode)) isUnique = true;
  }
  return newCode;
}

function savePendaftaran(formData) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_PENDAFTAR);
    
    if (!sheet) return { success: false, message: 'Database belum siap.' };
    
    // 1. Baca Header Aktual
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // --------------------------------------------------------
    // FITUR CEK NIK GANDA (Super Cepat Menggunakan TextFinder)
    // --------------------------------------------------------
    const nikColumnIndex = headers.indexOf('NIK');
    
    if (nikColumnIndex !== -1 && formData['NIK']) {
      const inputNIK = String(formData['NIK']).trim();
      
      // Ambil seluruh Range Kolom NIK (misal Kolom D:D)
      const colLetter = String.fromCharCode(65 + nikColumnIndex); // Mengubah index ke huruf (0 = A, 3 = D, dst)
      const nikRange = sheet.getRange(`${colLetter}2:${colLetter}`);
      
      // Gunakan TextFinder untuk pencarian kilat di Google Sheets
      const match = nikRange.createTextFinder(inputNIK).matchEntireCell(true).findNext();
      
      if (match) {
        return {
          success: false,
          message: `NIK ${inputNIK} sudah terdaftar dalam sistem! Jika ingin mengubah data, silakan gunakan fitur Revisi Data.`
        };
      }
    }
    // --------------------------------------------------------

    const kodePendaftaran = generateKodePendaftaran();
    const tglDaftar = new Date().toLocaleString('id-ID');
    
    // 2. Mapping Data Form ke Header Spreadsheet
    let rowData = headers.map(header => {
      if (header === 'Kode Pendaftaran') return kodePendaftaran;
      if (header === 'Status') return 'Pending';
      if (header === 'Tanggal Daftar') return tglDaftar;
      if (header === 'Catatan Admin') return '';
      
      let val = formData[header];
      
      if (val && typeof val === 'object' && val.data) {
         try {
           let cleanHeader = header.replace(/[^a-zA-Z0-9]/g, '');
           let fileName = `${cleanHeader}_${kodePendaftaran}`; 
           return uploadFileToDrive(val.data, fileName, val.type);
         } catch(e) {
           return "Error Upload: " + e.message;
         }
      }
      
      return val ? "'" + val : "";
    });

    // 3. Simpan
    sheet.appendRow(rowData);
    
    return {
      success: true,
      kodePendaftaran: kodePendaftaran,
      message: 'Data berhasil disimpan.'
    };
    
  } catch (error) {
    Logger.log(error);
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// Upload File Helper
function uploadFileToDrive(base64Data, fileName, mimeType) {
  try {
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_DRIVE_ID);
    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, mimeType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    // Return URL Viewer
    return file.getUrl();
  } catch (e) {
    throw new Error("Gagal Upload Drive: " + e.message);
  }
}

// Hash password dengan SHA-256 (Utilities bawaan Apps Script, tanpa library luar)
function hashPassword(plainPassword) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(plainPassword));
  return rawHash.map(byte => {
    const val = (byte < 0 ? byte + 256 : byte).toString(16);
    return val.length === 1 ? '0' + val : val;
  }).join('');
}

function validateAdminToken(token) {
  if (!token) {
    throw new Error("⛔ Akses Ditolak: Token autentikasi tidak ditemukan.");
  }

  // Cek apakah token ada di Cache server (Memory Server)
  const cache = CacheService.getUserCache();
  const sessionData = cache.get(token);

  if (!sessionData) {
    throw new Error("⛔ Sesi Kadaluarsa: Silakan login kembali.");
  }

  // Parse data user dari cache
  const user = JSON.parse(sessionData);

  // Pastikan Role-nya benar-benar Admin
  if (user.role !== 'Admin') {
    throw new Error("⛔ Forbidden: Anda tidak memiliki akses Admin.");
  }

  return true; // Token Valid & Aman
}
// ========================================
// 5. ADMIN DATA MANAGEMENT (CRUD)
// ========================================

// Read All Data (Smart Mapping) - PROTECTED
function getAllPendaftar(token) {
  try {
    // 1. VALIDASI KEAMANAN (GATEKEEPER)
    // Fungsi ini akan melempar Error jika token tidak valid/kadaluarsa
    validateAdminToken(token);

    // 2. LOGIKA ORIGINAL (JIKA TOKEN VALID)
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_PENDAFTAR);
    
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: [] };
    
    // Gunakan getDisplayValues agar format tanggal/angka/0 di depan tidak hilang
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0];
    
    // Helper Index
    const getIdx = (name) => headers.indexOf(name);
    const findIdxLike = (part) => headers.findIndex(h => h.toLowerCase().includes(part.toLowerCase()));

    // Identifikasi Kolom Utama (System) untuk keperluan UI
    const idxKode = getIdx('Kode Pendaftaran');
    const idxStatus = getIdx('Status');
    const idxCatatan = getIdx('Catatan Admin');
    const idxNama = findIdxLike('Nama'); 
    const idxNISN = findIdxLike('NISN');

    let result = [];
    
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      let rowObj = {};
      let fileKeys = [];
      let systemKeys = [];

      // A. Masukkan SEMUA Data Kolom Secara Dinamis
      headers.forEach((h, index) => {
        let val = row[index];
        rowObj[h] = val; // Contoh: "Alamat": "Jl. Mawar", "Nama Ayah": "Budi"

        // Deteksi apakah ini kolom File (Link Drive)
        if (val && typeof val === 'string' && (val.includes('drive.google.com') || val.includes('googleusercontent.com'))) {
           fileKeys.push(h);
        }
      });

      // B. Properti Standar untuk Logic Frontend (CamelCase)
      rowObj.kodePendaftaran = idxKode > -1 ? row[idxKode] : '';
      rowObj.status = idxStatus > -1 ? row[idxStatus] : 'Pending';
      rowObj.catatanAdmin = idxCatatan > -1 ? row[idxCatatan] : '';
      rowObj.nama = idxNama > -1 ? row[idxNama] : 'Tanpa Nama';
      rowObj.nisn = idxNISN > -1 ? row[idxNISN] : '-';
      
      // C. Mapping File Spesifik (Ijazah, KK, Foto) untuk UI Profil Kiri
      const findFileKey = (keyword) => fileKeys.find(k => k.toLowerCase().includes(keyword));
      rowObj.urlIjazah = findFileKey('ijazah') ? rowObj[findFileKey('ijazah')] : '';
      rowObj.urlKK = findFileKey('kk') ? rowObj[findFileKey('kk')] : '';
      rowObj.urlFoto = findFileKey('foto') ? rowObj[findFileKey('foto')] : '';

      // D. Catat Header yang sudah tampil di Profil (agar tidak duplikat di Grid)
      if(idxKode > -1) systemKeys.push(headers[idxKode]);
      if(idxStatus > -1) systemKeys.push(headers[idxStatus]);
      if(idxCatatan > -1) systemKeys.push(headers[idxCatatan]);
      if(idxNama > -1) systemKeys.push(headers[idxNama]);
      if(idxNISN > -1) systemKeys.push(headers[idxNISN]);

      // Simpan metadata untuk filtering di Frontend
      rowObj._systemKeys = systemKeys;
      rowObj._fileKeys = fileKeys;

      result.push(rowObj);
    }

    return { success: true, data: result };

  } catch (e) {
    // Jika validateAdminToken gagal, error message akan dikembalikan ke sini
    return { success: false, message: e.toString() };
  }
}

// Cek Status User (Public)
function cekStatus(kode) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_PENDAFTAR);
    
    // Validasi jika sheet tidak ada atau kosong
    if (!sheet || sheet.getLastRow() <= 1) {
       return { success: false, message: 'Data pendaftaran belum tersedia.' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Cari Index Kolom Penting
    const idxKode = headers.indexOf('Kode Pendaftaran');
    // Cari kolom nama yang mengandung kata 'Nama' (case insensitive)
    const idxNama = headers.findIndex(h => h.toLowerCase().includes('nama')); 
    const idxStatus = headers.indexOf('Status');
    const idxCatatan = headers.indexOf('Catatan Admin');
    const idxNISN = headers.findIndex(h => h.toLowerCase().includes('nisn'));

    if (idxKode === -1) {
      return { success: false, message: 'Sistem Error: Kolom Kode Pendaftaran tidak ditemukan.' };
    }

    // Looping pencarian data (Mulai baris ke-2 / index 1)
    for (let i = 1; i < data.length; i++) {
      // Bandingkan kode sebagai string dan trim spasi agar akurat
      if (String(data[i][idxKode]).trim() === String(kode).trim()) {
        
        // Data Ditemukan: Kembalikan hanya data yang aman untuk publik
        return { 
          success: true, 
          data: {
             kodePendaftaran: data[i][idxKode],
             nama: idxNama > -1 ? data[i][idxNama] : 'Calon Siswa',
             status: idxStatus > -1 ? data[i][idxStatus] : 'Pending',
             catatanAdmin: idxCatatan > -1 ? data[i][idxCatatan] : '',
             nisn: idxNISN > -1 ? data[i][idxNISN] : '-'
          }
        };
      }
    }

    // Jika loop selesai dan tidak ketemu
    return { success: false, message: 'Kode Pendaftaran tidak ditemukan.' };

  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

// Update Status
// Update Status & Catatan - PROTECTED
function updateStatus(token, kode, statusBaru, catatanBaru) {
  try {
    // 1. VALIDASI KEAMANAN
    // Cek apakah token valid. Jika tidak, akan throw Error dan loncat ke catch
    validateAdminToken(token);

    // 2. SETUP SPREADSHEET
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_PENDAFTAR);
    
    if (!sheet) {
      return { success: false, message: 'Database Pendaftar tidak ditemukan.' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // 3. CARI INDEX KOLOM
    const idxKode = headers.indexOf('Kode Pendaftaran');
    const idxStatus = headers.indexOf('Status');
    const idxCatatan = headers.indexOf('Catatan Admin'); // Kolom baru untuk catatan

    // Validasi struktur sheet
    if (idxKode === -1 || idxStatus === -1) {
      return { success: false, message: 'Struktur sheet invalid (Kolom Kode/Status hilang).' };
    }

    // 4. LOOPING PENCARIAN DATA
    for (let i = 1; i < data.length; i++) {
      // Pastikan kedua sisi dibandingkan sebagai String agar aman
      if (String(data[i][idxKode]) === String(kode)) {
        
        // A. Update Status (Baris = i + 1 karena array mulai dari 0 tapi sheet mulai dari 1)
        sheet.getRange(i + 1, idxStatus + 1).setValue(statusBaru);

        // B. Update Catatan (Jika kolomnya ada)
        if (idxCatatan > -1) {
            // Pastikan tidak menulis 'undefined' jika catatanBaru kosong
            const catatanFinal = (catatanBaru === undefined || catatanBaru === null) ? '' : catatanBaru;
            sheet.getRange(i + 1, idxCatatan + 1).setValue(catatanFinal);
        }

        return { success: true, message: 'Status dan Catatan berhasil diperbarui.' };
      }
    }

    return { success: false, message: 'Data dengan Kode Pendaftaran tersebut tidak ditemukan.' };

  } catch (e) {
    // Tangkap error validasi token atau error spreadsheet lainnya
    return { success: false, message: e.toString() };
  }
}

// Hapus Pendaftar - PROTECTED
function deletePendaftar(token, kode) {
  try {
    // 1. VALIDASI KEAMANAN
    // Memastikan pemanggil memiliki token admin yang sah
    validateAdminToken(token);

    // 2. SETUP SPREADSHEET
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_PENDAFTAR);
    
    // Cek apakah sheet ada
    if (!sheet) {
      return { success: false, message: 'Database Pendaftar tidak ditemukan.' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxKode = headers.indexOf('Kode Pendaftaran');

    // Cek integritas header
    if (idxKode === -1) {
      return { success: false, message: 'Kolom Kode Pendaftaran tidak ditemukan.' };
    }

    // 3. LOOPING PENCARIAN DATA
    for (let i = 1; i < data.length; i++) {
      // Bandingkan sebagai string untuk akurasi
      if (String(data[i][idxKode]) === String(kode)) {
        
        // Hapus baris (i + 1 karena array mulai 0, sheet mulai 1)
        sheet.deleteRow(i + 1);
        
        return { success: true, message: 'Data berhasil dihapus.' };
      }
    }

    // Jika loop selesai tanpa menemukan data
    return { success: false, message: 'Data tidak ditemukan.' };

  } catch (e) {
    // Menangkap error dari validateAdminToken atau error spreadsheet
    return { success: false, message: e.toString() };
  }
}


// ========================================
// 6. KONFIGURASI WEBSITE (HERO/FOOTER)
// ========================================

function getKonfigurasiAplikasi() {
  const defaultConf = {
    // Basic
    logoUrl: 'https://cdn-icons-png.flaticon.com/512/2920/2920224.png',
    judulSidebar: 'PPDB Online',
    appName: 'PPDB 2025',
    pendaftaranStatus: 'dibuka',
    // Hero
    teksHero: 'Pendaftaran Peserta Didik Baru',
    heroBaris1: 'Masa Depan Cerah',
    heroBaris2: 'Dimulai Di Sini',
    heroSubteks: 'Segera daftarkan diri Anda sebelum kuota penuh.',
    heroImageUrl: 'https://img.freepik.com/free-vector/happy-students-with-backpacks-books_74855-5853.jpg',
    // Footer & Kontak
    footerJudul: 'PPDB Sekolah',
    footerHakCipta: '© 2025 Panitia PPDB',
    footerEmail: 'muhdarmianto@gmail.com',
    footerTelepon: '085215087070',
    footerAlamat: 'Jl. Trans Sulawesi No. 05',
    // Sosmed
    linkFb: '#', linkIg: '#', linkYt: '#',
    // Alur (Step 1-4)
    alur1_judul: 'Isi Formulir', alur1_desc: 'Lengkapi data diri secara online.',
    alur2_judul: 'Upload Berkas', alur2_desc: 'Unggah KK, Ijazah, dan Foto.',
    alur3_judul: 'Verifikasi', alur3_desc: 'Panitia memeriksa data Anda.',
    alur4_judul: 'Pengumuman', alur4_desc: 'Cek kelulusan di website.'
  };

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_KONFIG);
    if (!sheet) return { success: true, data: defaultConf };
    
    // Ambil data sampai kolom 25 (agar muat semua)
    const data = sheet.getRange(2, 1, 1, 25).getValues()[0];

    return {
      success: true,
      data: {
        logoUrl: data[0] || defaultConf.logoUrl,
        judulSidebar: data[1] || defaultConf.judulSidebar,
        appName: data[2] || defaultConf.appName,
        pendaftaranStatus: data[3] || defaultConf.pendaftaranStatus,
        
        teksHero: data[4] || defaultConf.teksHero,
        heroBaris1: data[5] || defaultConf.heroBaris1,
        heroBaris2: data[6] || defaultConf.heroBaris2,
        heroSubteks: data[7] || defaultConf.heroSubteks,
        heroImageUrl: data[8] || defaultConf.heroImageUrl,
        
        footerJudul: data[9] || defaultConf.footerJudul,
        footerHakCipta: data[10] || defaultConf.footerHakCipta,
        footerEmail: data[11] || defaultConf.footerEmail,
        footerTelepon: data[12] || defaultConf.footerTelepon,
        
        // Data Baru (Kolom 14 dst)
        footerAlamat: data[13] || defaultConf.footerAlamat,
        linkFb: data[14] || '',
        linkIg: data[15] || '',
        linkYt: data[16] || '',
        
        alur1_judul: data[17] || defaultConf.alur1_judul,
        alur1_desc: data[18] || defaultConf.alur1_desc,
        alur2_judul: data[19] || defaultConf.alur2_judul,
        alur2_desc: data[20] || defaultConf.alur2_desc,
        alur3_judul: data[21] || defaultConf.alur3_judul,
        alur3_desc: data[22] || defaultConf.alur3_desc,
        alur4_judul: data[23] || defaultConf.alur4_judul,
        alur4_desc: data[24] || defaultConf.alur4_desc
      }
    };
  } catch (e) {
    return { success: true, data: defaultConf };
  }
}

// Simpan Konfigurasi Website (Hero/Footer) - PROTECTED
function saveWebConfig(token, form) {
  try {
    // 1. VALIDASI KEAMANAN
    // Cek apakah token valid. Jika tidak, script akan berhenti di sini.
    validateAdminToken(token);

    // 2. SETUP SPREADSHEET
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_KONFIG);
    
    // Pastikan sheet konfigurasi ada
    if (!sheet) {
        return { success: false, message: 'Sheet Konfigurasi hilang atau belum dibuat.' };
    }

    // 3. MAPPING DATA FORM KE ARRAY BARIS
    // Urutan ini HARUS SESUAI dengan urutan kolom di Sheet Konfigurasi (initialize)
    const rowData = [
      form.logoUrl, 
      form.judulSidebar, 
      form.appName, 
      form.pendaftaranStatus,
      form.teksHero, 
      form.heroBaris1, 
      form.heroBaris2, 
      form.heroSubteks, 
      form.heroImageUrl,
      form.footerJudul, 
      form.footerHakCipta, 
      form.footerEmail, 
      form.footerTelepon,
      form.footerAlamat, 
      form.linkFb, 
      form.linkIg, 
      form.linkYt,
      form.alur1_judul, form.alur1_desc,
      form.alur2_judul, form.alur2_desc,
      form.alur3_judul, form.alur3_desc,
      form.alur4_judul, form.alur4_desc
    ];
    
    // 4. SIMPAN DATA
    // Timpa Baris ke-2 (Baris 1 adalah Header)
    // getRange(row, column, numRows, numColumns)
    sheet.getRange(2, 1, 1, rowData.length).setValues([rowData]);
    
    return { success: true, message: 'Konfigurasi website berhasil disimpan.' };

  } catch(e) {
    // Menangkap error validasi token atau error penulisan sheet
    return { success: false, message: e.toString() };
  }
}

// ========================================
// 7. AUTHENTICATION & SESSION (MODIFIED)
// ========================================

function handleLogin(loginData) {
  if (loginData.loginType === 'admin') {
    // LOGIN MENGGUNAKAN USERNAME (Update parameter)
    return loginAdmin(loginData.username, loginData.password);
  }
  return { success: false, message: 'Invalid Type' };
}

function loginAdmin(username, password) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_ADMIN);
    
    // Jika sheet belum ada, return error agar user menjalankan initialize
    if (!sheet) return { success: false, message: 'Database Admin belum siap. Jalankan fungsi initialize() dulu.' };

    const data = sheet.getDataRange().getValues();
    const inputHash = hashPassword(password);

    // Loop mulai dari baris ke-2 (index 1) karena baris 1 adalah header
    for (let i = 1; i < data.length; i++) {
      // Kolom 0 = Username, Kolom 1 = Password (tersimpan sebagai hash SHA-256)
      const storedPassword = String(data[i][1]);
      const isHashMatch = storedPassword === inputHash;
      // Fallback: dukung akun lama yang passwordnya masih plain-text,
      // lalu otomatis di-upgrade ke hash agar tidak tersimpan plain-text lagi.
      const isLegacyPlainMatch = !isHashMatch && storedPassword === String(password);

      if (String(data[i][0]) === String(username) && (isHashMatch || isLegacyPlainMatch)) {
        if (isLegacyPlainMatch) {
          sheet.getRange(i + 1, 2).setValue(inputHash);
        }

        const adminData = { 
          username: data[i][0], 
          role: data[i][2] || 'Admin',
          nama: data[i][3] || 'Administrator'
        };
        
        const sessionId = Utilities.getUuid();
        CacheService.getUserCache().put(sessionId, JSON.stringify(adminData), 21600); // 6 Jam
        
        return { success: true, sessionId: sessionId, data: adminData };
      }
    }
    return { success: false, message: 'Username atau Password Salah' };
  } catch (e) {
    return { success: false, message: 'Login Error: ' + e.message };
  }
}

function getUserSession(sessionId) {
  if (!sessionId) return { success: false };
  const data = CacheService.getUserCache().get(sessionId);
  return data ? { success: true, data: JSON.parse(data) } : { success: false };
}

// ==========================================
// 8. SETUP OTOMATIS (INITIALIZE)
// ==========================================

/**
 * Fungsi untuk inisialisasi awal.
 * Menyiapkan semua sheet dan data sampel secara otomatis.
 * Jalankan fungsi ini sekali dari Editor Apps Script.
 */
/**
 * Fungsi untuk inisialisasi awal database.
 * Menyiapkan semua sheet, kolom konfigurasi CMS, dan form builder lengkap.
 */
function initialize() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  // =========================================
  // 1. SETUP SHEET ADMIN
  // =========================================
  let sheetAdmin = ss.getSheetByName(CONFIG.SHEET_ADMIN);
  if (!sheetAdmin) {
    sheetAdmin = ss.insertSheet(CONFIG.SHEET_ADMIN);
    sheetAdmin.appendRow(['Username', 'Password', 'Role', 'Nama Lengkap']);
    sheetAdmin.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f3f4f6');
    sheetAdmin.appendRow(['admin', hashPassword('admin123'), 'Admin', 'Super Administrator']);
    Logger.log('Sheet Admin baru dibuat.');
  } else {
    Logger.log('Sheet Admin sudah ada. Melewati.');
  }

  // =========================================
  // 2. SETUP SHEET KONFIGURASI (CMS)
  // =========================================
  let sheetKonfig = ss.getSheetByName(CONFIG.SHEET_KONFIG);
  if (!sheetKonfig) {
    sheetKonfig = ss.insertSheet(CONFIG.SHEET_KONFIG);
    
    const headersConf = [
      'Logo URL', 'Judul Sidebar', 'App Name', 'Status Pendaftaran', 
      'Teks Hero', 'Hero Baris 1', 'Hero Baris 2', 'Hero Subteks', 'Hero Image URL', 
      'Footer Judul', 'Footer Copyright', 'Footer Email', 'Footer Telepon', 'Footer Alamat',
      'Link FB', 'Link IG', 'Link YT',
      'Alur 1 Judul', 'Alur 1 Desc',
      'Alur 2 Judul', 'Alur 2 Desc',
      'Alur 3 Judul', 'Alur 3 Desc',
      'Alur 4 Judul', 'Alur 4 Desc'
    ];
    
    sheetKonfig.appendRow(headersConf);
    sheetKonfig.getRange(1, 1, 1, headersConf.length).setFontWeight('bold').setBackground('#f3f4f6');
    
    sheetKonfig.appendRow([
      'https://upload.wikimedia.org/wikipedia/commons/thumb/9/9c/Logo_of_Ministry_of_Education_and_Culture_of_Republic_of_Indonesia.svg/250px-Logo_of_Ministry_of_Education_and_Culture_of_Republic_of_Indonesia.svg.png', 'SPMB Online', 'SPMB 2026', 'dibuka',
      'Pendaftaran Peserta Didik Baru', 'Pendidikan Terbaik untuk', 'Masa Depan',
      'Bergabunglah dengan komunitas belajar yang inspiratif. Kami mencetak generasi berkarakter, cerdas, dan siap menghadapi tantangan global.',
      'https://i.imgur.com/3B7Wey0.png',
      'SMP Negeri 1 Kolaka Utara', '© 2026 Panitia SPMB', 'admin@sekolah.sch.id', '0812-3456-7890',
      'Jl. Pendidikan No. 1, Jakarta', '#', '#', '#',
      'Isi Formulir', 'Lengkapi data diri calon siswa.',
      'Upload Berkas', 'Unggah dokumen syarat (KK/Ijazah).',
      'Verifikasi', 'Panitia memvalidasi data Anda.',
      'Pengumuman', 'Cek hasil seleksi secara online.'
    ]);
    Logger.log('Sheet Konfigurasi baru dibuat.');
  } else {
    Logger.log('Sheet Konfigurasi sudah ada. Melewati.');
  }

  // =========================================
  // 3. SETUP SHEET BUILDER (FORM CONFIG)
  // =========================================
  let sheetBuilder = ss.getSheetByName(CONFIG.SHEET_BUILDER);
  if (!sheetBuilder) {
    sheetBuilder = ss.insertSheet(CONFIG.SHEET_BUILDER);
    
    // Konfigurasi Default yang LENGKAP
    const defaults = [
        ["Nama Lengkap", "text", "", "12", true],
        ["NISN", "number", "", "6", true],
        ["NIK", "number", "", "6", true],
        ["NIS", "number", "", "6", false],
        ["Tempat Lahir", "text", "", "6", true],
        ["Tanggal Lahir", "date", "", "6", true],
        ["Jenis Kelamin", "dropdown", "Laki-laki, Perempuan", "6", true],
        ["Agama", "dropdown", "Islam, Kristen, Katolik, Hindu, Buddha", "6", true],
        ["Golongan Darah", "dropdown", "A, B, AB, O, -", "one-fifth", false],
        ["Tinggi Badan (cm)", "number", "", "one-fifth", false],
        ["Berat Badan (kg)", "number", "", "one-fifth", false],
        ["Alamat Domisili", "textarea", "", "12", true],
        
        ["Nama Ayah", "text", "", "6", true],
        ["Pekerjaan Ayah", "text", "", "6", false],
        ["Nama Ibu", "text", "", "6", true],
        ["Pekerjaan Ibu", "text", "", "6", false],
        ["No Telfon Orang Tua", "number", "", "6", true],
        ["Alamat Orang Tua", "textarea", "", "6", false],
        
        ["Kelas", "text", "", "6", true],
        ["Tahun Masuk", "number", "", "6", true],
        ["Riwayat Pendidikan", "textarea", "", "12", false],
        ["Prestasi", "textarea", "", "12", false],
        ["Hobi", "text", "", "12", false],
        ["Catatan Penting", "textarea", "", "12", false],
        
        ["Pas Foto", "file", "", "12", true]
    ];
    sheetBuilder.getRange(1, 1, defaults.length, 5).setValues(defaults);
    Logger.log('Sheet Builder baru dibuat.');
  } else {
    Logger.log('Sheet Builder sudah ada. Melewati.');
  }

  // =========================================
  // 4. SETUP SHEET PENDAFTAR (5 SAMPEL LENGKAP)
  // =========================================
  let sheetPendaftar = ss.getSheetByName(CONFIG.SHEET_PENDAFTAR);
  if (!sheetPendaftar) {
    sheetPendaftar = ss.insertSheet(CONFIG.SHEET_PENDAFTAR);
    
    // Ambil Config Column dari sheetBuilder
    const builderData = sheetBuilder.getDataRange().getValues();
    // Filter baris kosong untuk mendapatkan label kolom yang valid
    const currentConfig = builderData.filter(r => r[0] && r[0] !== '').map(r => ({label: r[0]}));

    // Setup Header Gabungan
    const sysLeft = ['Kode Pendaftaran'];
    const dynamicCols = currentConfig.map(c => c.label);
    const sysRight = ['Status', 'Tanggal Daftar', 'Catatan Admin'];
    const allHeaders = [...sysLeft, ...dynamicCols, ...sysRight];

    // Tulis Header & Styling
    sheetPendaftar.getRange(1, 1, 1, allHeaders.length)
         .setValues([allHeaders])
         .setBackground('#1E88E5')
         .setFontColor('#FFFFFF')
         .setFontWeight('bold');

    // --- DATA SAMPEL YANG LENGKAP UNTUK 5 ORANG ---
    const tglNow = new Date().toLocaleString('id-ID');
    
    // Array Data Dummy Lengkap
    const sampleDetails = [
      {
        nama: "Ahmad Fikri", nisn: "0012345678", nik: "3201012010100001", nis: "2023001",
        tmpLahir: "Jakarta", tglLahir: "2010-05-12", jk: "Laki-laki", agama: "Islam",
        goldar: "O", tinggi: "155", berat: "45", alamat: "Jl. Merpati No. 10, Jakarta Selatan",
        ayah: "Budi Santoso", jobAyah: "Wiraswasta", ibu: "Siti Aminah", jobIbu: "Ibu Rumah Tangga",
        telpOrtu: "081234567890", alamatOrtu: "Jl. Merpati No. 10, Jakarta Selatan",
        kelas: "7", thnMasuk: "2025", riwayat: "SDN 01 Jakarta", prestasi: "Juara 1 Lomba Matematika",
        hobi: "Membaca", catatan: "Alergi kacang", status: "Pending", catatanAdmin: ""
      },
      {
        nama: "Citra Kirana", nisn: "0012345679", nik: "3201012010100002", nis: "2023002",
        tmpLahir: "Bandung", tglLahir: "2010-08-20", jk: "Perempuan", agama: "Islam",
        goldar: "A", tinggi: "150", berat: "40", alamat: "Jl. Kenanga No. 5, Bandung",
        ayah: "Dadang Supriatna", jobAyah: "PNS", ibu: "Rina Marlina", jobIbu: "Guru",
        telpOrtu: "081234567891", alamatOrtu: "Jl. Kenanga No. 5, Bandung",
        kelas: "7", thnMasuk: "2025", riwayat: "SDN 03 Bandung", prestasi: "Juara 2 Lomba Melukis",
        hobi: "Menari", catatan: "-", status: "Diterima", catatanAdmin: "Berkas lengkap"
      },
      {
        nama: "Eko Prasetyo", nisn: "0012345680", nik: "3201012010100003", nis: "2023003",
        tmpLahir: "Surabaya", tglLahir: "2009-12-01", jk: "Laki-laki", agama: "Kristen",
        goldar: "B", tinggi: "160", berat: "50", alamat: "Jl. Pahlawan No. 2, Surabaya",
        ayah: "Joko Susilo", jobAyah: "Karyawan Swasta", ibu: "Maria Ulfa", jobIbu: "Ibu Rumah Tangga",
        telpOrtu: "081234567892", alamatOrtu: "Jl. Pahlawan No. 2, Surabaya",
        kelas: "7", thnMasuk: "2025", riwayat: "SD Katolik Surabaya", prestasi: "Juara Harapan Catur",
        hobi: "Catur", catatan: "-", status: "Perbaikan", catatanAdmin: "Foto ijazah kurang jelas"
      },
      {
        nama: "Dewi Sartika", nisn: "0012345681", nik: "3201012010100004", nis: "2023004",
        tmpLahir: "Yogyakarta", tglLahir: "2010-02-14", jk: "Perempuan", agama: "Islam",
        goldar: "AB", tinggi: "152", berat: "42", alamat: "Jl. Malioboro No. 1, Yogyakarta",
        ayah: "Slamet Rahardjo", jobAyah: "Seniman", ibu: "Sri Wahyuni", jobIbu: "Pedagang",
        telpOrtu: "081234567893", alamatOrtu: "Jl. Malioboro No. 1, Yogyakarta",
        kelas: "7", thnMasuk: "2025", riwayat: "SDN 1 Yogyakarta", prestasi: "-",
        hobi: "Menyanyi", catatan: "-", status: "Pending", catatanAdmin: ""
      },
      {
        nama: "Fajar Nugraha", nisn: "0012345682", nik: "3201012010100005", nis: "2023005",
        tmpLahir: "Medan", tglLahir: "2010-06-10", jk: "Laki-laki", agama: "Islam",
        goldar: "O", tinggi: "158", berat: "55", alamat: "Jl. Merdeka No. 8, Medan",
        ayah: "Andi Siregar", jobAyah: "Polisi", ibu: "Nurhasanah", jobIbu: "Bidan",
        telpOrtu: "081234567894", alamatOrtu: "Jl. Merdeka No. 8, Medan",
        kelas: "7", thnMasuk: "2025", riwayat: "SDN 5 Medan", prestasi: "Juara Futsal Antar Sekolah",
        hobi: "Sepak Bola", catatan: "Riwayat asma", status: "Tidak Diterima", catatanAdmin: "Kuota penuh"
      }
    ];

    // Logic Mapping Data ke Kolom Spreadsheet
    // (Mencocokkan nama kolom header dengan data di object sampleDetails)
    const rows = sampleDetails.map((d, i) => {
      return allHeaders.map(header => {
        const h = header.toLowerCase();
        
        // --- 1. Kolom Sistem ---
        if (h === 'kode pendaftaran') return CONFIG.KODE_PREFIX + '00' + (i + 1);
        if (h === 'status') return d.status;
        if (h === 'tanggal daftar') return tglNow;
        if (h === 'catatan admin') return d.catatanAdmin;
        
        // --- 2. Kolom Data Diri ---
        if (h.includes('nama lengkap')) return d.nama;
        if (h.includes('nisn')) return "'" + d.nisn; // Pakai petik biar jadi string
        if (h.includes('nik')) return "'" + d.nik;
        if (h.includes('nis')) return "'" + d.nis;
        if (h.includes('tempat lahir')) return d.tmpLahir;
        if (h.includes('tanggal lahir')) return d.tglLahir;
        if (h.includes('jenis kelamin')) return d.jk;
        if (h.includes('agama')) return d.agama;
        if (h.includes('golongan darah')) return d.goldar;
        if (h.includes('tinggi badan')) return d.tinggi;
        if (h.includes('berat badan')) return d.berat;
        if (h.includes('alamat domisili')) return d.alamat;
        
        // --- 3. Kolom Orang Tua ---
        if (h.includes('nama ayah')) return d.ayah;
        if (h.includes('pekerjaan ayah')) return d.jobAyah;
        if (h.includes('nama ibu')) return d.ibu;
        if (h.includes('pekerjaan ibu')) return d.jobIbu;
        if (h.includes('no telfon')) return "'" + d.telpOrtu;
        if (h.includes('alamat orang tua')) return d.alamatOrtu;
        
        // --- 4. Kolom Akademik & Tambahan ---
        if (h.includes('kelas')) return d.kelas;
        if (h.includes('tahun masuk')) return d.thnMasuk;
        if (h.includes('riwayat pendidikan')) return d.riwayat;
        if (h.includes('prestasi')) return d.prestasi;
        if (h.includes('hobi')) return d.hobi;
        if (h.includes('catatan penting')) return d.catatan;
        
        // --- 5. File (Kosongkan atau beri dummy link) ---
        if (h.includes('foto') || h.includes('file')) return '';

        return ''; // Default kosong jika tidak ada yang cocok
      });
    });

    // Tulis Data Sampel ke Spreadsheet
    sheetPendaftar.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    Logger.log('Sheet Pendaftar baru dibuat dengan 5 sampel data LENGKAP.');
  } else {
    Logger.log('Sheet Pendaftar sudah ada. Melewati.');
  }
  
  return "Inisialisasi selesai. Struktur database siap digunakan.";
}





// ========================================
// 9. FITUR PERBAIKAN DATA (USER)
// ========================================

// Ambil data siswa untuk diedit (Hanya jika status = Perbaikan)
function getDataForEdit(kode) {
  try {
    // 1. Setup Spreadsheet
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_PENDAFTAR);
    
    // Cek jika sheet ada
    if (!sheet) return { success: false, message: 'Database tidak ditemukan.' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0]; // Baris pertama adalah nama kolom (Header)
    
    const idxKode = headers.indexOf('Kode Pendaftaran');
    const idxStatus = headers.indexOf('Status');
    
    // Validasi kolom penting
    if (idxKode === -1 || idxStatus === -1) {
      return { success: false, message: 'Struktur database tidak valid (Kolom Kode/Status hilang).' };
    }

    // 2. Loop mencari data
    for (let i = 1; i < data.length; i++) {
      // Pastikan kode cocok (String comparison & Trim spasi)
      if (String(data[i][idxKode]).trim() === String(kode).trim()) {
        
        // 3. KEAMANAN: Pastikan status memang 'Perbaikan'
        const currentStatus = String(data[i][idxStatus]).toLowerCase();
        if (currentStatus !== 'perbaikan') {
          return { success: false, message: 'Status data bukan Perbaikan. Akses edit ditolak.' };
        }
        
        // 4. MAPPING DATA: Header -> Isi Data
        let rowData = {};
        headers.forEach((headerName, colIndex) => {
          if(['Tanggal Daftar', 'Catatan Admin', 'Status'].includes(headerName)) return;
          
          let cellValue = data[i][colIndex];
          
          if (cellValue instanceof Date) {
             try {
               let yyyy = cellValue.getFullYear();
               let mm = String(cellValue.getMonth() + 1).padStart(2, '0');
               let dd = String(cellValue.getDate()).padStart(2, '0');
               cellValue = `${yyyy}-${mm}-${dd}`;
             } catch(e) {
               // Biarkan nilai aslinya
             }
          } else if (typeof cellValue === 'number') {
             cellValue = String(cellValue);
          }
          
          rowData[headerName] = cellValue;
        });
        
        return { success: true, data: rowData };
      }
    }
    
    return { success: false, message: 'Kode Pendaftaran tidak ditemukan.' };

  } catch (e) {
    return { success: false, message: 'Error Server: ' + e.toString() };
  }
}

// Menyimpan data hasil revisi (dengan validasi NIK/Identitas Unik agar tidak bentrok)
function submitPerbaikan(formData) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_PENDAFTAR);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getDataRange().getValues();
    
    const idxKode = headers.indexOf('Kode Pendaftaran');
    const idxNIK = headers.indexOf('NIK'); // ⚠️ Sesuaikan nama kolom ini jika di spreadsheet pakai nama lain
    
    let rowIndex = -1;
    
    // 1. CARI BARIS DATA LAMA
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxKode]).trim() === String(formData.kodePendaftaran).trim()) {
        rowIndex = i + 1; 
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: 'Data tidak ditemukan untuk diperbarui.' };
    }
    
    // 2. CEK DUPLIKASI NIK (PENCEGAHAN GANDA)
    if (idxNIK !== -1 && formData['NIK']) {
      const nikInput = String(formData['NIK']).trim();
      
      for (let i = 1; i < data.length; i++) {
        // Skip baris milik siswa itu sendiri yang sedang diedit
        if (i + 1 === rowIndex) continue; 
        
        const nikDiSheet = String(data[i][idxNIK]).trim();
        if (nikDiSheet === nikInput) {
          return { 
            success: false, 
            message: `Gagal memperbarui! NIK "${nikInput}" NIK sudah terdaftar.` 
          };
        }
      }
    }
    
    // 3. LOOP HEADER UNTUK UPDATE KOLOM
    headers.forEach((h, colIndex) => {
      if (['Kode Pendaftaran', 'Tanggal Daftar', 'Catatan Admin'].includes(h)) return;
      
      if (h === 'Status') {
         sheet.getRange(rowIndex, colIndex + 1).setValue('Pending');
         return;
      }
      
      if (formData[h] && typeof formData[h] === 'object' && formData[h].data) {
         try {
           let cleanHeader = h.replace(/[^a-zA-Z0-9]/g, '');
           let fileName = `${cleanHeader}_${formData.kodePendaftaran}_REV`; 
           let url = uploadFileToDrive(formData[h].data, fileName, formData[h].type);
           sheet.getRange(rowIndex, colIndex + 1).setValue(url);
         } catch(e) {}
      } 
      else if (formData.hasOwnProperty(h)) {
         let val = formData[h];
         
         let isFileCol = h.toLowerCase().match(/(foto|file|ijazah|kk|dokumen|lampiran)/);
         if (val === "" && isFileCol) {
            return; 
         }
         
         sheet.getRange(rowIndex, colIndex + 1).setValue("'" + val);
      }
    });
    
    return { success: true, message: 'Data revisi berhasil dikirim. Menunggu verifikasi admin.' };
    
  } catch (e) {
    return { success: false, message: 'Error Server: ' + e.toString() };
  }
}

function logoutServer(token) {
  try {
    // Cek jika token ada
    if (!token) return { success: false, message: 'Token kosong' };

    // Hapus token dari Cache Server
    CacheService.getUserCache().remove(token);
    
    return { success: true, message: 'Sesi server berhasil dihapus' };
  } catch (e) {
    // Abaikan error jika cache sudah tidak ada, tetap return true agar frontend logout
    return { success: true, message: 'Logout (Error Handler): ' + e.message };
  }
}

// ========================================
// 10. FITUR EXPORT EXCEL (PROFESSIONAL)
// ========================================

// ========================================
// 10. FITUR EXPORT EXCEL (PROFESSIONAL - FIXED FORMAT)
// ========================================

function generateExcelReport(token) {
  try {
    // 1. VALIDASI KEAMANAN
    validateAdminToken(token);

    // 2. AMBIL DATA DARI SHEET UTAMA
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_PENDAFTAR);
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: false, message: 'Tidak ada data untuk diexport.' };
    }

    // Ambil semua data sebagai String (Display Values)
    const dataRange = sheet.getDataRange();
    const rawData = dataRange.getDisplayValues(); 

    // --- BAGIAN BARU: FORMATTING ANGKA ---
    // Loop data untuk menambahkan tanda petik (') pada string angka
    const formattedData = rawData.map((row, rowIndex) => {
      // Biarkan baris Header (baris pertama/index 0) apa adanya
      if (rowIndex === 0) return row; 

      return row.map(cell => {
        // Cek apakah cell berisi angka murni (misal: "08123456", "3201123456789000")
        // Regex /^\d+$/ artinya: string yang isinya HANYA angka dari awal sampai akhir
        // Kita juga cek apakah cell tidak kosong
        if (cell && String(cell).trim().match(/^\d+$/)) {
           return "'" + cell; // Tambahkan petik satu di depan
        }
        return cell;
      });
    });
    // -------------------------------------

    // 3. BUAT SPREADSHEET SEMENTARA (TEMP)
    const tempSS = SpreadsheetApp.create("Temp_Export_" + new Date().getTime());
    const tempSheet = tempSS.getSheets()[0];
    const tempId = tempSS.getId();

    // 4. TULIS DATA YANG SUDAH DIFORMAT KE SHEET TEMP
    tempSheet.getRange(1, 1, formattedData.length, formattedData[0].length).setValues(formattedData);

    // Styling Header (Baris 1)
    const headerRange = tempSheet.getRange(1, 1, 1, formattedData[0].length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1E88E5'); 
    headerRange.setFontColor('#FFFFFF');
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');
    
    // Auto Resize Kolom
    tempSheet.autoResizeColumns(1, formattedData[0].length);

    // Flush wajib dilakukan agar data tersimpan sebelum dikonversi
    SpreadsheetApp.flush();

    // 5. KONVERSI KE BLOB EXCEL (.XLSX)
    const url = "https://docs.google.com/spreadsheets/d/" + tempId + "/export?format=xlsx";
    const params = {
      method: "GET",
      headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    };
    
    const blob = UrlFetchApp.fetch(url, params).getBlob();
    
    // Penamaan File
    const timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd_HHmm");
    const fileName = `Data_Pendaftar_PPDB_${timestamp}.xlsx`;
    blob.setName(fileName);

    // 6. BERSIHKAN FILE SEMENTARA
    DriveApp.getFileById(tempId).setTrashed(true);

    // 7. RETURN HASIL
    return {
      success: true,
      data: Utilities.base64Encode(blob.getBytes()),
      filename: fileName,
      message: 'Excel berhasil dibuat.'
    };

  } catch (e) {
    return { success: false, message: 'Gagal Export: ' + e.toString() };
  }
}

function pancinganIzin() {
  // Fungsi ini hanya dummy untuk memicu pop-up otorisasi
  UrlFetchApp.fetch("https://www.google.com");
  DriveApp.getRootFolder();
  SpreadsheetApp.getActiveSpreadsheet();
}








//pdf 
function generatePDF(kodePendaftaran) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Pendaftar"); 
    
    if (!sheet) {
      return { success: false, message: "Sheet 'Pendaftar' tidak ditemukan!" };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: false, message: "Sheet 'Pendaftar' masih kosong." };
    }

    // 1. Bersihkan semua nama Header dari Karakter Enter (\n) dan Spasi Berlebih
    const rawHeaders = data[0];
    const headers = rawHeaders.map(h => 
      String(h).replace(/[\r\n]+/g, " ").replace(/\s+/g, " ").trim()
    );
    
    // Cari index kolom Kode Pendaftaran
    const kodeIndex = headers.indexOf("Kode Pendaftaran");
    if (kodeIndex === -1) {
      return { success: false, message: "Kolom 'Kode Pendaftaran' tidak ditemukan!" };
    }

    // Cari baris data berdasarkan Kode Pendaftaran
    let rowData = null;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][kodeIndex]).trim() === String(kodePendaftaran).trim()) {
        rowData = data[i];
        break;
      }
    }
    
    if (!rowData) {
      return { success: false, message: "Data pendaftaran tidak ditemukan." };
    }

    // 2. Pemetaan Data ke Objek Pendaftar
    let pendaftar = {};
    headers.forEach((header, index) => {
      let val = rowData[index];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
      }
      pendaftar[header] = (val !== null && val !== undefined && String(val).trim() !== "") ? String(val).trim() : "-";
    });

    // Template HTML PDF Menurun Compact (Fit 1 Halaman A4)
    let htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          @page { size: A4; margin: 6mm 12mm; }
          body { 
            font-family: Arial, Helvetica, sans-serif; 
            padding: 0; 
            margin: 0; 
            color: #1e293b; 
            line-height: 1.15; 
            font-size: 11px; 
          }
          
          .header { text-align: center; margin-bottom: 4px; }
          .header h3 { margin: 0; font-size: 12px; color: #0f172a; font-weight: bold; text-transform: uppercase; }
          .header p.sub-title { margin: 1px 0 0 0; font-size: 8.5px; color: #475569; }
          .header p.print-date { margin: 1px 0 0 0; font-size: 8px; color: #64748b; }
          
          .divider-line { border-bottom: 1px solid #334155; margin-bottom: 4px; }

          .kode-box { 
            text-align: center; 
            font-size: 9px; 
            font-weight: bold; 
            margin-bottom: 4px; 
            padding: 2px 4px; 
            background-color: #ffffff; 
            border: 1px dashed #64748b; 
          }
          
          .section-title { 
            color: #475569; 
            font-weight: bold; 
            font-size: 8px; 
            margin-top: 4px; 
            margin-bottom: 2px; 
            text-transform: uppercase; 
          }
          
          table.data-table { width: 100%; border-collapse: collapse; margin-bottom: 3px; }
          table.data-table td { padding: 2px 4px; font-size: 8.5px; vertical-align: middle; border: 1px solid #cbd5e1; }
          table.data-table td.label { font-weight: bold; width: 35%; color: #0f172a; background-color: #ffffff; }
          table.data-table td.val { color: #1e293b; width: 65%; }
          
          .ttd-container { width: 100%; margin-top: 8px; border-collapse: collapse; page-break-inside: avoid; }
          .ttd-container td.empty-side { width: 65%; }
          .ttd-container td.sign-side { width: 35%; text-align: center; font-size: 8.5px; vertical-align: top; border: none; }
          .space-ttd { height: 28px; }
        </style>
      </head>
      <body>

        <div class="header">
          <h3>BUKTI VERVAL DATA PESERTA DIDIK</h3>
          <p class="sub-title">Sistem Informasi Data Pokok Pendidikan (DAPODIK)</p>
          <p class="print-date">Dokumen ini dicetak otomatis dari Sistem VervalPD pada: ${Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "d/M/yyyy")}</p>
        </div>

        <div class="divider-line"></div>
        <div class="kode-box">KODE PENDAFTARAN: ${pendaftar['Kode Pendaftaran'] || kodePendaftaran}</div>

        <!-- A. IDENTITAS PESERTA DIDIK -->
        <div class="section-title">A. IDENTITAS PESERTA DIDIK</div>
        <table class="data-table">
          <tr><td class="label">Nama Lengkap Sesuai KK</td><td class="val">${pendaftar['Nama Lengkap Sesuai KK']}</td></tr>
          <tr><td class="label">NIK</td><td class="val">${pendaftar['NIK']}</td></tr>
          <tr><td class="label">No. KK</td><td class="val">${pendaftar['No KK']}</td></tr>
          <tr><td class="label">Nomor Akta Lahir</td><td class="val">${pendaftar['Nomor Akta Lahir']}</td></tr>
          <tr><td class="label">NISN</td><td class="val">${pendaftar['NISN']}</td></tr>
          <tr><td class="label">Jenis Kelamin</td><td class="val">${pendaftar['Jenis Kelamin']}</td></tr>
          <tr><td class="label">Tempat Lahir</td><td class="val">${pendaftar['Tempat Lahir']}</td></tr>
          <tr><td class="label">Tanggal Lahir</td><td class="val">${pendaftar['Tanggal Lahir']}</td></tr>
          <tr><td class="label">Agama</td><td class="val">${pendaftar['Agama']}</td></tr>
          <tr><td class="label">Anak Ke</td><td class="val">${pendaftar['Anak Ke []'] || pendaftar['Anak Ke']}</td></tr>
          <tr><td class="label">Hobby</td><td class="val">${pendaftar['Hobby']}</td></tr>
          <tr><td class="label">Cita-Cita</td><td class="val">${pendaftar['Cita-Cita']}</td></tr>
        </table>

        <!-- B. DATA PENDIDIKAN -->
        <div class="section-title">B. DATA PENDIDIKAN</div>
        <table class="data-table">
        <tr><td class="label">Kelas Saat Ini</td><td class="val">${pendaftar['Kelas Saat Ini']}</td></tr>
          <tr><td class="label">Apakah Pernah PAUD (TK)</td><td class="val">${pendaftar['Apakah Pernah PAUD (TK)'] || pendaftar['Apakah Pernah PAUD [TK]']}</td></tr>
          <tr><td class="label">Asal Sekolah</td><td class="val">${pendaftar['Asal Sekolah SD/MI'] || pendaftar['Asal Sekolah']}</td></tr>
          <tr><td class="label">Tahun Lulus</td><td class="val">${pendaftar['Tahun Lulus']}</td></tr>
        </table>

        <!-- C. DATA TEMPAT TINGGAL -->
        <div class="section-title">C. DATA TEMPAT TINGGAL</div>
        <table class="data-table">
          <tr><td class="label">Alamat Peserta Didik</td><td class="val">${pendaftar['Alamat Peserta Didik']}</td></tr>
          <tr><td class="label">Tempat Tinggal</td><td class="val">${pendaftar['Tempat Tinggal']}</td></tr>
          <tr><td class="label">Mode Transportasi</td><td class="val">${pendaftar['Mode Transportasi']}</td></tr>
          <tr><td class="label">No. HP Peserta Didik</td><td class="val">${pendaftar['No HP Peserta Didik']}</td></tr>
        </table>

        <!-- D. DATA AYAH -->
        <div class="section-title">D. DATA AYAH</div>
        <table class="data-table">
          <tr><td class="label">Nama Lengkap Ayah Sesuai KK</td><td class="val">${pendaftar['Nama Lengkap Ayah Sesuai KK']}</td></tr>
          <tr><td class="label">NIK Ayah</td><td class="val">${pendaftar['NIK Ayah']}</td></tr>
          <tr><td class="label">Tahun Lahir Ayah</td><td class="val">${pendaftar['Tahun Lahir Ayah']}</td></tr>
          <tr><td class="label">Pendidikan Ayah</td><td class="val">${pendaftar['Pendidikan Ayah']}</td></tr>
          <tr><td class="label">Pekerjaan Ayah</td><td class="val">${pendaftar['Pekerjaan Ayah']}</td></tr>
          <tr><td class="label">Penghasilan Ayah</td><td class="val">${pendaftar['Penghasilan Ayah']}</td></tr>
        </table>

        <!-- E. DATA IBU -->
        <div class="section-title">E. DATA IBU</div>
        <table class="data-table">
          <tr><td class="label">Nama Ibu Sesuai KK</td><td class="val">${pendaftar['Nama Ibu Sesuai KK']}</td></tr>
          <tr><td class="label">NIK Ibu</td><td class="val">${pendaftar['NIK Ibu']}</td></tr>
          <tr><td class="label">Tahun Lahir Ibu</td><td class="val">${pendaftar['Tahun Lahir Ibu']}</td></tr>
          <tr><td class="label">Pendidikan Ibu</td><td class="val">${pendaftar['Pendidikan Ibu']}</td></tr>
          <tr><td class="label">Pekerjaan Ibu</td><td class="val">${pendaftar['Pekerjaan Ibu']}</td></tr>
          <tr><td class="label">Penghasilan Ibu</td><td class="val">${pendaftar['Penghasilan Ibu']}</td></tr>
        </table>

        <!-- F. KONTAK ORANG TUA -->
        <div class="section-title">F. KONTAK ORANG TUA / WALI</div>
        <table class="data-table">
          <tr><td class="label">No HP Orang Tua</td><td class="val">${pendaftar['No HP Orang Tua']}</td></tr>
        </table>

        <!-- G. DATA WALI -->
        <div class="section-title">G. DATA WALI</div>
        <table class="data-table">
          <tr><td class="label">Nama Wali Sesuai KK</td><td class="val">${pendaftar['Nama Wali Sesuai KK']}</td></tr>
          <tr><td class="label">NIK Wali</td><td class="val">${pendaftar['NIK Wali']}</td></tr>
          <tr><td class="label">Tahun Lahir Wali</td><td class="val">${pendaftar['Tahun Lahir Wali']}</td></tr>
          <tr><td class="label">Pendidikan Wali</td><td class="val">${pendaftar['Pendidikan Wali']}</td></tr>
          <tr><td class="label">Pekerjaan Wali</td><td class="val">${pendaftar['Pekerjaan Wali']}</td></tr>
          <tr><td class="label">Penghasilan Wali</td><td class="val">${pendaftar['Penghasilan Wali']}</td></tr>
        </table>

        <!-- TANDA TANGAN -->
        <table class="ttd-container">
          <tr>
            <td class="empty-side"></td>
            <td class="sign-side">
              <p>Orang Tua / Wali Murid,</p>
              <div class="space-ttd"></div>
              <p><b>( ............................................ )</b></p>
            </td>
          </tr>
        </table>

      </body>
      </html>
    `;

    const blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf');
    blob.setName(`Bukti_VervalPD_${kodePendaftaran}.pdf`);
    
    return {
      success: true,
      filename: `Bukti_VervalPD_${kodePendaftaran}.pdf`,
      bytes: Utilities.base64Encode(blob.getBytes())
    };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
