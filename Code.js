// File: Code.gs (Pengontrol Utama)

/**
 * GASPUL Hadir - Aplikasi Absensi dengan Google Apps Script
 * Berbasis Google Spreadsheet
 * Versi: 2.0
 */

// ======== FUNGSI UTAMA DAN ROUTING ========

/**
 * Fungsi utama untuk menangani request HTTP GET
 * Routing antar halaman aplikasi
 */
function doGet(e) {
  // Setup basic logging
  Logger.log("Access time (UTC): " + new Date().toISOString());
  
  try {
    // Validasi parameter e dan e.parameter
    e = e || {};
    e.parameter = e.parameter || {};
    
    Logger.log("Query parameters: " + JSON.stringify(e.parameter));
    
    // Get page parameter with default to 'login' with safe navigation
    const page = (e.parameter?.page || 'login').toLowerCase();
    Logger.log("Requested page: " + page);

    // Get user session
    const userSession = getUserSession();
    Logger.log("User session exists: " + (userSession !== null));

    // Define allowed pages and their configurations
    const pages = {
      login: {
        title: 'Login - GASPUL Hadir',
        template: 'Login',
        public: true
      },
      home: {
        title: 'Beranda - GASPUL Hadir',
        template: 'Home',
        requireAuth: true
      },
      history: {
        title: 'Riwayat - GASPUL Hadir',
        template: 'History',
        requireAuth: true
      },
      profile: {
        title: 'Profil - GASPUL Hadir',
        template: 'Profile',
        requireAuth: true
      }
    };

    // Get page configuration with fallback to login
    const pageConfig = pages[page] || pages.login;

    // Check authentication for protected pages
    if (pageConfig.requireAuth && !userSession) {
      Logger.log("Auth required but user not logged in. Redirecting to login.");
      return createHtmlOutput(pages.login);
    }

    // If user is logged in but accessing login page, redirect to home
    if (userSession && page === 'login') {
      Logger.log("Logged in user accessing login page. Redirecting to home.");
      return createHtmlOutput(pages.home);
    }

    // Refresh session if user is logged in
    if (userSession) {
      refreshUserSession();
    }

    // Return requested page
    return createHtmlOutput(pageConfig);

  } catch (error) {
    Logger.log("Error in doGet: " + error.toString());
    console.error("doGet error:", error);

    // Return error page with more details
    return HtmlService.createTemplate(
      `<div class="error-container">
        <h1>Terjadi Kesalahan</h1>
        <p>${error.message || 'Unknown error occurred'}</p>
        <p class="error-details">Time: ${new Date().toISOString()}</p>
        <button onclick="window.location.href='?page=login'">Kembali ke Login</button>
       </div>
       <style>
         .error-container {
           text-align: center;
           padding: 2rem;
           max-width: 600px;
           margin: 2rem auto;
           font-family: Arial, sans-serif;
         }
         .error-details {
           color: #666;
           font-size: 0.9rem;
           margin-top: 1rem;
         }
         button {
           padding: 0.5rem 1rem;
           margin-top: 1rem;
           background: #007bff;
           color: white;
           border: none;
           border-radius: 4px;
           cursor: pointer;
           font-size: 1rem;
         }
         button:hover {
           background: #0056b3;
         }
       </style>`
    )
    .evaluate()
    .setTitle('Error - GASPUL Hadir')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// Helper function untuk createHtmlOutput
function createHtmlOutput(pageConfig) {
  try {
    const template = HtmlService.createTemplateFromFile(pageConfig.template);
    
    // Add user data to template if available
    const userSession = getUserSession();
    if (userSession) {
      template.user = userSession;
    }
    
    return template
      .evaluate()
      .setTitle(pageConfig.title)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    Logger.log("Error in createHtmlOutput: " + error.toString());
    throw error; // Re-throw untuk ditangkap oleh error handler di doGet
  }
}

// Helper function untuk refresh session
function refreshUserSession() {
  try {
    const userSession = getUserSession();
    if (userSession) {
      userSession.lastActivity = new Date().getTime();
      PropertiesService.getUserProperties().setProperty(
        'userSession',
        JSON.stringify(userSession)
      );
    }
  } catch (error) {
    Logger.log("Error refreshing session: " + error.toString());
    // Don't throw error here, just log it
  }
}

// Helper function to create HTML output
function createHtmlOutput(pageConfig) {
  const template = HtmlService.createTemplateFromFile(pageConfig.template);
  
  // Add user data to template if available
  const userSession = getUserSession();
  if (userSession) {
    template.user = userSession;
  }
  
  return template
    .evaluate()
    .setTitle(pageConfig.title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Helper function to get user session
function getUserSession() {
  try {
    const userProps = PropertiesService.getUserProperties();
    const sessionData = userProps.getProperty('userSession');
    return sessionData ? JSON.parse(sessionData) : null;
  } catch (error) {
    Logger.log("Error getting user session: " + error.toString());
    return null;
  }
}

// Helper function to refresh user session
function refreshUserSession() {
  try {
    const userSession = getUserSession();
    if (userSession) {
      userSession.lastActivity = new Date().getTime();
      PropertiesService.getUserProperties().setProperty(
        'userSession',
        JSON.stringify(userSession)
      );
    }
  } catch (error) {
    Logger.log("Error refreshing session: " + error.toString());
  }
}

// Helper function to check if session is expired
function isSessionExpired(session) {
  if (!session || !session.lastActivity) return true;
  
  const sessionTimeout = 30 * 60 * 1000; // 30 minutes
  const currentTime = new Date().getTime();
  return (currentTime - session.lastActivity) > sessionTimeout;
}

/**
 * Fungsi untuk mengambil konten file HTML
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    Logger.log("Error including file " + filename + ": " + e.message);
    return "<!-- Error loading " + filename + " -->";
  }
}

/**
 * Mendapatkan URL script untuk redirects
 */
function getScriptURL() {
  return ScriptApp.getService().getUrl();
}

/**
 * Memperbarui sesi user untuk memperpanjang masa aktif
 */
function refreshUserSession() {
  try {
    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('sessionTimestamp', new Date().getTime().toString());
    return true;
  } catch (error) {
    Logger.log("Error pada refreshUserSession: " + error.message);
    return false;
  }
}

/**
 * Periksa apakah user sudah login
 */
function isUserLoggedIn() {
  try {
    var userProperties = PropertiesService.getUserProperties();
    var isLoggedIn = userProperties.getProperty('isLoggedIn');
    var timestamp = userProperties.getProperty('sessionTimestamp');
    
    // Jika tidak ada timestamp atau tidak login, return false
    if (!isLoggedIn || isLoggedIn !== 'true' || !timestamp) {
      return false;
    }
    
    // Cek apakah session lebih dari 8 jam
    var now = new Date().getTime();
    var sessionTime = parseInt(timestamp);
    var eightHours = 8 * 60 * 60 * 1000;
    
    if (now - sessionTime > eightHours) {
      // Session expired, logout
      userProperties.deleteAllProperties();
      return false;
    }
    
    return true;
  } catch (error) {
    Logger.log("Error pada fungsi isUserLoggedIn: " + error.message);
    return false;
  }
}

// ======== FUNGSI AUTENTIKASI ========

/**
 * Fungsi untuk login user
 */
function login(nip, password) {
  try {
    // Validasi input
    if (!nip || !password) {
      return { success: false, message: "NIP dan password tidak boleh kosong" };
    }
    
    // Cari user berdasarkan NIP
    var user = findUserByNIP(nip);
    
    if (!user) {
      return { success: false, message: "User dengan NIP tersebut tidak ditemukan" };
    }
    
    // Verifikasi password (di implementasi sebenarnya, gunakan hash)
    if (user.password === password) {
      // Set session
      var userProperties = PropertiesService.getUserProperties();
      userProperties.setProperty('userNIP', user.nip);
      userProperties.setProperty('userName', user.name);
      userProperties.setProperty('userEmail', user.email);
      userProperties.setProperty('userPosition', user.position || "");
      userProperties.setProperty('userUnit', user.unit || "");
      userProperties.setProperty('isLoggedIn', 'true');
      userProperties.setProperty('sessionTimestamp', new Date().getTime().toString());
      
      Logger.log("Login berhasil, session disimpan untuk user: " + user.name);
      
      return { 
        success: true, 
        message: "Login berhasil", 
        user: {
          nip: user.nip,
          name: user.name,
          email: user.email
        }
      };
    } else {
      return { success: false, message: "Password salah" };
    }
  } catch (error) {
    Logger.log("Error pada fungsi login: " + error.message);
    return { success: false, message: "Terjadi kesalahan: " + error.message };
  }
}

/**
 * Mencari user berdasarkan NIP
 */
function findUserByNIP(nip) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    
    if (!sheet) {
      Logger.log("Sheet 'Users' tidak ditemukan!");
      return null;
    }
    
    var data = sheet.getDataRange().getValues();
    
    // Mulai dari baris 1 (baris setelah header)
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == nip) {
        return {
          nip: data[i][0],
          name: data[i][1],
          email: data[i][2],
          password: data[i][3],
          position: data[i][4],
          unit: data[i][5],
          registerDate: data[i][6],
          qrCodeID: data[i][7],
          avatar: data[i][8]
        };
      }
    }
    
    return null;
  } catch (error) {
    Logger.log("Error pada fungsi findUserByNIP: " + error.message);
    return null;
  }
}

/**
 * Mendapatkan data user dari session
 */
function getUserSession() {
  try {
    var userProperties = PropertiesService.getUserProperties();
    var nip = userProperties.getProperty('userNIP');
    
    if (!nip) {
      return null;
    }
    
    return {
      nip: nip,
      name: userProperties.getProperty('userName'),
      email: userProperties.getProperty('userEmail'),
      position: userProperties.getProperty('userPosition'),
      unit: userProperties.getProperty('userUnit')
    };
  } catch (error) {
    Logger.log("Error pada fungsi getUserSession: " + error.message);
    return null;
  }
}

/**
 * Menghapus session (logout)
 */
function logout() {
  try {
    var userProperties = PropertiesService.getUserProperties();
    userProperties.deleteAllProperties();
    return { success: true, message: "Logout berhasil" };
  } catch (error) {
    Logger.log("Error pada fungsi logout: " + error.message);
    return { success: false, message: "Terjadi kesalahan saat logout: " + error.message };
  }
}

// ======== FUNGSI PENGATURAN WAKTU KERJA ========

/**
 * Mendapatkan waktu mulai yang dikonfigurasi untuk hari tertentu dalam seminggu
 * dengan fitur caching dan penanganan hari khusus/libur
 * 
 * @param {number} hariDalamMinggu - Hari dalam seminggu (0 = Minggu, 1 = Senin, dll.)
 * @param {string} tanggal - Tanggal dalam format "YYYY-MM-DD", opsional untuk memeriksa hari khusus
 * @return {object} Konfigurasi waktu mulai {jam, menit, status}
 */
function getWaktuMulaiUntukHari(hariDalamMinggu, tanggal) {
  try {
    // Cek cache terlebih dahulu untuk performa lebih baik
    var cacheKey = "waktuMulai_" + hariDalamMinggu;
    var cache = CacheService.getScriptCache();
    var cachedData = cache.get(cacheKey);
    
    if (cachedData !== null) {
      try {
        return JSON.parse(cachedData);
      } catch (e) {
        // Jika terjadi kesalahan parsing, lanjutkan dengan membaca dari spreadsheet
        Logger.log("Error parsing cached data: " + e.message);
      }
    }
    
    // Cek apakah tanggal yang diberikan adalah hari libur/khusus
    if (tanggal) {
      var statusHariKhusus = cekHariKhusus(tanggal);
      if (statusHariKhusus) {
        return statusHariKhusus.waktuMulai;
      }
    }
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PengaturanWaktu');
    
    if (!sheet) {
      var defaultSettings = buatPengaturanWaktuDefault();
      return defaultSettings.waktuMulai;
    }
    
    var data = sheet.getDataRange().getValues();
    
    // Validasi data sheet
    if (data.length < 8 || data[0].length < 5) { // Minimal header + 7 hari & minimal 5 kolom
      var defaultSettings = buatPengaturanWaktuDefault();
      return defaultSettings.waktuMulai;
    }
    
    // Hari dalam minggu + 1 (untuk memperhitungkan baris header)
    var row = hariDalamMinggu + 1;
    
    // Cek apakah row berada dalam range valid
    if (row >= data.length) {
      Logger.log("Indeks hari tidak valid: " + hariDalamMinggu);
      var defaultSettings = buatPengaturanWaktuDefault();
      return defaultSettings.waktuMulai;
    }
    
    // Parse waktu mulai
    var waktuMulaiStr = data[row][1];
    
    // Validasi format waktu
    if (!waktuMulaiStr || typeof waktuMulaiStr !== 'string' || !waktuMulaiStr.includes(':')) {
      Logger.log("Format waktu mulai tidak valid: " + waktuMulaiStr);
      return { jam: 8, menit: 0, status: 'default' };
    }
    
    var bagianWaktuMulai = waktuMulaiStr.split(':');
    
    // Cek status hari (Normal/Libur/Setengah Hari)
    var statusHari = data[row][4] || "Normal";
    
    var result = {
      jam: parseInt(bagianWaktuMulai[0], 10),
      menit: parseInt(bagianWaktuMulai[1], 10),
      status: statusHari
    };
    
    // Simpan ke cache untuk 6 jam
    cache.put(cacheKey, JSON.stringify(result), 21600);
    
    return result;
  } catch (error) {
    Logger.log("Error mendapatkan waktu mulai untuk hari: " + error.message);
    return { jam: 8, menit: 0, status: 'error' }; // Default ke 8:00 pagi
  }
}

/**
 * Mendapatkan waktu selesai yang dikonfigurasi untuk hari tertentu dalam seminggu
 * dengan fitur caching dan penanganan hari khusus/libur
 * 
 * @param {number} hariDalamMinggu - Hari dalam seminggu (0 = Minggu, 1 = Senin, dll.)
 * @param {string} tanggal - Tanggal dalam format "YYYY-MM-DD", opsional untuk memeriksa hari khusus
 * @return {object} Konfigurasi waktu selesai {jam, menit, status}
 */
function getWaktuSelesaiUntukHari(hariDalamMinggu, tanggal) {
  try {
    // Cek cache terlebih dahulu untuk performa lebih baik
    var cacheKey = "waktuSelesai_" + hariDalamMinggu;
    var cache = CacheService.getScriptCache();
    var cachedData = cache.get(cacheKey);
    
    if (cachedData !== null) {
      try {
        return JSON.parse(cachedData);
      } catch (e) {
        // Jika terjadi kesalahan parsing, lanjutkan dengan membaca dari spreadsheet
        Logger.log("Error parsing cached data: " + e.message);
      }
    }
    
    // Cek apakah tanggal yang diberikan adalah hari libur/khusus
    if (tanggal) {
      var statusHariKhusus = cekHariKhusus(tanggal);
      if (statusHariKhusus) {
        return statusHariKhusus.waktuSelesai;
      }
    }
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PengaturanWaktu');
    
    if (!sheet) {
      var defaultSettings = buatPengaturanWaktuDefault();
      return defaultSettings.waktuSelesai;
    }
    
    var data = sheet.getDataRange().getValues();
    
    // Validasi data sheet
    if (data.length < 8 || data[0].length < 5) { // Minimal header + 7 hari & minimal 5 kolom
      var defaultSettings = buatPengaturanWaktuDefault();
      return defaultSettings.waktuSelesai;
    }
    
    // Hari dalam minggu + 1 (untuk memperhitungkan baris header)
    var row = hariDalamMinggu + 1;
    
    // Cek apakah row berada dalam range valid
    if (row >= data.length) {
      Logger.log("Indeks hari tidak valid: " + hariDalamMinggu);
      var defaultSettings = buatPengaturanWaktuDefault();
      return defaultSettings.waktuSelesai;
    }
    
    // Parse waktu selesai
    var waktuSelesaiStr = data[row][2];
    
    // Validasi format waktu
    if (!waktuSelesaiStr || typeof waktuSelesaiStr !== 'string' || !waktuSelesaiStr.includes(':')) {
      Logger.log("Format waktu selesai tidak valid: " + waktuSelesaiStr);
      return { jam: 17, menit: 0, status: 'default' };
    }
    
    var bagianWaktuSelesai = waktuSelesaiStr.split(':');
    
    // Cek status hari (Normal/Libur/Setengah Hari)
    var statusHari = data[row][4] || "Normal";
    
    // Cek toleransi pulang awal (dalam menit)
    var toleransiPulangAwal = data[row][5] ? parseInt(data[row][5], 10) : 0;
    
    var result = {
      jam: parseInt(bagianWaktuSelesai[0], 10),
      menit: parseInt(bagianWaktuSelesai[1], 10),
      status: statusHari,
      toleransiPulangAwal: toleransiPulangAwal
    };
    
    // Simpan ke cache untuk 6 jam
    cache.put(cacheKey, JSON.stringify(result), 21600);
    
    return result;
  } catch (error) {
    Logger.log("Error mendapatkan waktu selesai untuk hari: " + error.message);
    return { jam: 17, menit: 0, status: 'error' }; // Default ke 5:00 sore
  }
}

/**
 * Buat sheet pengaturan waktu default dengan kemampuan untuk hari khusus
 * 
 * @return {object} Pengaturan waktu default {waktuMulai, waktuSelesai, batasTerlambat}
 */
function buatPengaturanWaktuDefault() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Buat sheet PengaturanWaktu jika belum ada
    var pengaturanSheet = ss.getSheetByName('PengaturanWaktu');
    if (!pengaturanSheet) {
      pengaturanSheet = ss.insertSheet('PengaturanWaktu');
      
      // Buat header dengan kolom tambahan
      pengaturanSheet.appendRow([
        'Hari', 
        'Waktu Mulai', 
        'Waktu Selesai', 
        'Batas Terlambat (menit)',
        'Status Hari',
        'Toleransi Pulang Awal (menit)',
        'Persentase Kehadiran (%)'
      ]);
      
      // Tambahkan waktu default untuk setiap hari
      var hari = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
      var statusHari = ['Libur', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Setengah Hari'];
      
      for (var i = 0; i < hari.length; i++) {
        // Default untuk hari kerja: 08:00-17:00, hari Sabtu: 08:00-13:00, Minggu: libur
        var waktuMulai = '08:00';
        var waktuSelesai = (i === 0) ? '00:00' : (i === 6) ? '13:00' : '17:00';
        var batasTerlambat = (i === 0) ? '0' : '15';
        var toleransiPulangAwal = (i === 0) ? '0' : (i === 6) ? '30' : '15';
        var persentaseKehadiran = (i === 0) ? '0' : '100';
        
        pengaturanSheet.appendRow([
          hari[i], 
          waktuMulai, 
          waktuSelesai, 
          batasTerlambat,
          statusHari[i],
          toleransiPulangAwal,
          persentaseKehadiran
        ]);
      }
      
      // Format tampilan
      pengaturanSheet.getRange(1, 1, 1, 7).setFontWeight('bold');
      pengaturanSheet.getRange(1, 1, 8, 7).setHorizontalAlignment('center');
      pengaturanSheet.setColumnWidth(1, 150);
      pengaturanSheet.setColumnWidth(2, 120);
      pengaturanSheet.setColumnWidth(3, 120);
      pengaturanSheet.setColumnWidth(4, 150);
      pengaturanSheet.setColumnWidth(5, 120);
      pengaturanSheet.setColumnWidth(6, 180);
      pengaturanSheet.setColumnWidth(7, 180);
      
      // Warna header
      pengaturanSheet.getRange(1, 1, 1, 7).setBackground('#4285f4').setFontColor('white');
      
      // Warna baris berselang-seling
      for (var i = 2; i <= 8; i++) {
        var color = (i % 2 === 0) ? '#f3f3f3' : '#ffffff';
        pengaturanSheet.getRange(i, 1, 1, 7).setBackground(color);
      }
      
      // Highlight untuk hari libur dan setengah hari
      pengaturanSheet.getRange(2, 1, 1, 7).setBackground('#ffcdd2'); // Minggu
      pengaturanSheet.getRange(8, 1, 1, 7).setBackground('#fff9c4'); // Sabtu
    }
    
    // 2. Buat sheet HariKhusus jika belum ada
    var hariKhususSheet = ss.getSheetByName('HariKhusus');
    if (!hariKhususSheet) {
      hariKhususSheet = ss.insertSheet('HariKhusus');
      
      // Buat header
      hariKhususSheet.appendRow([
        'Tanggal', 
        'Keterangan', 
        'Status', 
        'Waktu Mulai', 
        'Waktu Selesai', 
        'Batas Terlambat (menit)',
        'Toleransi Pulang Awal (menit)'
      ]);
      
      // Tambahkan contoh hari libur nasional
      hariKhususSheet.appendRow([
        '2025-01-01', 
        'Tahun Baru', 
        'Libur', 
        '00:00', 
        '00:00', 
        '0',
        '0'
      ]);
      
      hariKhususSheet.appendRow([
        '2025-05-01', 
        'Hari Buruh', 
        'Libur', 
        '00:00', 
        '00:00', 
        '0',
        '0'
      ]);
      
      hariKhususSheet.appendRow([
        '2025-05-09', 
        'Cuti Bersama', 
        'Libur', 
        '00:00', 
        '00:00', 
        '0',
        '0'
      ]);
      
      hariKhususSheet.appendRow([
        '2025-08-17', 
        'Hari Kemerdekaan', 
        'Libur', 
        '00:00', 
        '00:00', 
        '0',
        '0'
      ]);
      
      hariKhususSheet.appendRow([
        '2025-12-24', 
        'Malam Natal', 
        'Setengah Hari', 
        '08:00', 
        '13:00', 
        '15',
        '30'
      ]);
      
      hariKhususSheet.appendRow([
        '2025-12-25', 
        'Hari Natal', 
        'Libur', 
        '00:00', 
        '00:00', 
        '0',
        '0'
      ]);
      
      // Format tampilan
      hariKhususSheet.getRange(1, 1, 1, 7).setFontWeight('bold');
      hariKhususSheet.getRange(1, 1, 7, 7).setHorizontalAlignment('center');
      hariKhususSheet.setColumnWidth(1, 150);
      hariKhususSheet.setColumnWidth(2, 200);
      hariKhususSheet.setColumnWidth(3, 120);
      hariKhususSheet.setColumnWidth(4, 120);
      hariKhususSheet.setColumnWidth(5, 120);
      hariKhususSheet.setColumnWidth(6, 180);
      hariKhususSheet.setColumnWidth(7, 180);
      
      // Warna header
      hariKhususSheet.getRange(1, 1, 1, 7).setBackground('#4285f4').setFontColor('white');
      
      // Warna baris berselang-seling dan berdasarkan status
      for (var i = 2; i <= 7; i++) {
        var color = '#ffffff';
        var status = hariKhususSheet.getRange(i, 3).getValue();
        
        if (status === 'Libur') {
          color = '#ffcdd2'; // Merah muda untuk hari libur
        } else if (status === 'Setengah Hari') {
          color = '#fff9c4'; // Kuning muda untuk setengah hari
        } else {
          color = (i % 2 === 0) ? '#f3f3f3' : '#ffffff';
        }
        
        hariKhususSheet.getRange(i, 1, 1, 7).setBackground(color);
      }
    }
    
    // Kembalikan pengaturan default
    return {
      waktuMulai: { jam: 8, menit: 0, status: 'default' },
      waktuSelesai: { jam: 17, menit: 0, status: 'default' },
      batasTerlambat: 15
    };
  } catch (error) {
    Logger.log("Error membuat pengaturan waktu default: " + error.message);
    return {
      waktuMulai: { jam: 8, menit: 0, status: 'error' },
      waktuSelesai: { jam: 17, menit: 0, status: 'error' },
      batasTerlambat: 15
    };
  }
}

/**
 * Cek apakah tanggal tertentu adalah hari khusus/libur
 * 
 * @param {string} tanggal - Tanggal dalam format "YYYY-MM-DD"
 * @return {object|null} Informasi hari khusus atau null jika bukan hari khusus
 */
function cekHariKhusus(tanggal) {
  try {
    // Cek cache terlebih dahulu
    var cacheKey = "hariKhusus_" + tanggal;
    var cache = CacheService.getScriptCache();
    var cachedData = cache.get(cacheKey);
    
    if (cachedData !== null) {
      try {
        var result = JSON.parse(cachedData);
        if (result.isCached) { // Mencegah hasil false positif
          return result;
        }
      } catch (e) {
        // Lanjutkan jika parsing gagal
        Logger.log("Error parsing cached data: " + e.message);
      }
    }
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HariKhusus');
    
    if (!sheet) {
      return null; // Tidak ada sheet hari khusus
    }
    
    var data = sheet.getDataRange().getValues();
    
    // Jika hanya header, tidak ada data
    if (data.length < 2) {
      return null;
    }
    
    // Format tanggal untuk perbandingan
    var tanggalDicari = formatTanggal(tanggal);
    
    // Cari tanggal di data
    for (var i = 1; i < data.length; i++) {
      var tanggalData = formatTanggal(data[i][0]);
      
      if (tanggalDicari === tanggalData) {
        // Temukan hari khusus
        var keterangan = data[i][1];
        var status = data[i][2];
        var waktuMulai = data[i][3].split(':');
        var waktuSelesai = data[i][4].split(':');
        var batasTerlambat = parseInt(data[i][5], 10) || 0;
        var toleransiPulangAwal = parseInt(data[i][6], 10) || 0;
        
        var result = {
          tanggal: tanggalDicari,
          keterangan: keterangan,
          status: status,
          waktuMulai: {
            jam: parseInt(waktuMulai[0], 10),
            menit: parseInt(waktuMulai[1], 10),
            status: status
          },
          waktuSelesai: {
            jam: parseInt(waktuSelesai[0], 10),
            menit: parseInt(waktuSelesai[1], 10),
            status: status,
            toleransiPulangAwal: toleransiPulangAwal
          },
          batasTerlambat: batasTerlambat,
          isCached: true // Penanda untuk cache
        };
        
        // Simpan ke cache selama 24 jam (hasil hari khusus jarang berubah)
        cache.put(cacheKey, JSON.stringify(result), 86400);
        
        return result;
      }
    }
    
    // Jika tidak ditemukan, cache hasil null selama 24 jam
    cache.put(cacheKey, JSON.stringify({isCached: true, isNull: true}), 86400);
    
    return null;
  } catch (error) {
    Logger.log("Error memeriksa hari khusus: " + error.message);
    return null;
  }
}

/**
 * Format tanggal ke format standar YYYY-MM-DD untuk perbandingan
 * 
 * @param {string|Date} tanggal - Tanggal yang akan diformat
 * @return {string} Tanggal dalam format YYYY-MM-DD
 */
function formatTanggal(tanggal) {
  try {
    var date;
    
    if (tanggal instanceof Date) {
      date = tanggal;
    } else if (typeof tanggal === 'string') {
      // Coba parse tanggal dari string
      if (tanggal.includes('-')) {
        var parts = tanggal.split('-');
        if (parts.length === 3) {
          // Sudah dalam format YYYY-MM-DD
          return tanggal;
        }
      }
      
      date = new Date(tanggal);
    } else {
      return ''; // Format tidak valid
    }
    
    // Cek validitas tanggal
    if (isNaN(date.getTime())) {
      return '';
    }
    
    // Format tanggal ke YYYY-MM-DD
    var year = date.getFullYear();
    var month = (date.getMonth() + 1).toString().padStart(2, '0');
    var day = date.getDate().toString().padStart(2, '0');
    
    return year + '-' + month + '-' + day;
  } catch (error) {
    Logger.log("Error memformat tanggal: " + error.message);
    return '';
  }
}

/**
 * Mendapatkan semua pengaturan jadwal untuk hari tertentu
 * Fungsi yang lebih komprehensif untuk mendapatkan semua informasi jadwal sekaligus
 * 
 * @param {Date} tanggal - Tanggal yang akan dicek
 * @return {object} Informasi lengkap pengaturan jadwal kerja
 */
// Deklarasi variabel cache di luar fungsi
let cachedSchedule = null;
let lastCacheTime = null;
const CACHE_DURATION = 5 * 60 * 1000; // 5 menit dalam milliseconds

function getJadwalKerja(tanggal) {
  try {
    if (!waktuMulai || typeof waktuMulai.jam === 'undefined' || typeof waktuMulai.menit === 'undefined') {
        waktuMulai = { jam: 8, menit: 0, status: 'default' };
      }

    if (!waktuSelesai || typeof waktuSelesai.jam === 'undefined' || typeof waktuSelesai.menit === 'undefined') {
        waktuSelesai = { jam: 17, menit: 0, status: 'default' };
      }
    // Cek cache terlebih dahulu
    if (cachedSchedule && lastCacheTime && (new Date() - lastCacheTime) < CACHE_DURATION) {
      return cachedSchedule;
    }

    // Format tanggal untuk pengecekan hari khusus
    var tanggalStr = Utilities.formatDate(tanggal, "GMT+7", "yyyy-MM-dd");
    var hariDalamMinggu = tanggal.getDay();
    
    // Cek apakah ini hari khusus
    var hariKhususInfo = cekHariKhusus(tanggalStr);
    
    let result; // Deklarasi variable untuk hasil
    
    if (hariKhususInfo) {
      result = {
        tanggal: tanggalStr,
        hariDalamMinggu: hariDalamMinggu,
        keterangan: hariKhususInfo.keterangan,
        status: hariKhususInfo.status,
        waktuMulai: hariKhususInfo.waktuMulai,
        waktuSelesai: hariKhususInfo.waktuSelesai,
        batasTerlambat: hariKhususInfo.batasTerlambat,
        isHariKhusus: true
      };
    } else {
      // Jika bukan hari khusus, dapatkan pengaturan reguler
      var waktuMulai = getWaktuMulaiUntukHari(hariDalamMinggu);
      var waktuSelesai = getWaktuSelesaiUntukHari(hariDalamMinggu);
      
      // Dapatkan batas terlambat dari sheet
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PengaturanWaktu');
      var batasTerlambat = 15; // Default
      
      if (sheet) {
        var data = sheet.getDataRange().getValues();
        if (data.length > hariDalamMinggu + 1 && data[0].length > 3) {
          batasTerlambat = parseInt(data[hariDalamMinggu + 1][3], 10) || 15;
        }
      }
      
      // Nama hari dalam Bahasa Indonesia
      var namaHari = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
      
      result = {
        tanggal: tanggalStr,
        hariDalamMinggu: hariDalamMinggu,
        namaHari: namaHari[hariDalamMinggu],
        status: waktuMulai.status,
        waktuMulai: waktuMulai,
        waktuSelesai: waktuSelesai,
        batasTerlambat: batasTerlambat,
        isHariKhusus: false
      };
    }

    // Simpan ke cache
    cachedSchedule = result;
    lastCacheTime = new Date();
    
    // Logging untuk debugging
    Logger.log("Jadwal berhasil diambil untuk tanggal: " + tanggalStr);
    Logger.log("Cache time: " + lastCacheTime.toISOString());
    
    return result;

  } catch (error) {
    Logger.log("Error mendapatkan jadwal kerja: " + error.message);
    console.error("Error detail:", error);
    
    // Kembalikan default
    const defaultResult = {
      tanggal: Utilities.formatDate(tanggal, "GMT+7", "yyyy-MM-dd"),
      hariDalamMinggu: tanggal.getDay(),
      status: "Error",
      waktuMulai: { jam: 8, menit: 0, status: 'error' },
      waktuSelesai: { jam: 17, menit: 0, status: 'error' },
      batasTerlambat: 15,
      isHariKhusus: false,
      error: error.message
    };

    // Simpan error ke cache untuk menghindari request berlebihan
    cachedSchedule = defaultResult;
    lastCacheTime = new Date();
    
    return defaultResult;
  }
}

// Fungsi helper untuk reset cache jika diperlukan
function resetJadwalCache() {
  cachedSchedule = null;
  lastCacheTime = null;
  Logger.log("Cache jadwal direset");
}

function logAttendanceError(error, context = {}) {
  try {
    const errorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ErrorLogs');
    
    // Validasi error object
    const errorMessage = error?.message || String(error) || "Unknown error";
    const errorStack = error?.stack || "No stack trace available";
    const timestamp = new Date();
    const contextStr = JSON.stringify(context || {});
    
    // Log ke console terlebih dahulu
    Logger.log("Attendance Error at " + timestamp.toISOString());
    Logger.log("Error message: " + errorMessage);
    Logger.log("Context: " + contextStr);
    Logger.log("Stack trace: " + errorStack);
    
    // Jika sheet ErrorLogs ada, tambahkan error log
    if (errorSheet) {
      errorSheet.appendRow([
        timestamp,
        errorMessage,
        contextStr,
        errorStack,
        Session.getActiveUser().getEmail(),
        ScriptApp.getScriptId()
      ]);
    }
    
    // Return error object untuk chaining
    return {
      timestamp: timestamp,
      message: errorMessage,
      context: context,
      stack: errorStack
    };
    
  } catch (loggingError) {
    // Fallback jika terjadi error saat logging
    Logger.log("Error in logAttendanceError: " + loggingError);
    console.error("Failed to log error:", loggingError);
    return null;
  }
}

// Fungsi helper untuk format error
function formatError(error) {
  if (typeof error === 'string') {
    return {
      message: error,
      stack: new Error().stack
    };
  }
  if (error instanceof Error) {
    return error;
  }
  return new Error(String(error));
}

// Contoh penggunaan dalam getAttendanceHistory
function getAttendanceHistory(startDate, endDate) {
  try {
    // Validasi input
    if (!startDate || !endDate) {
      throw new Error("Tanggal tidak valid");
    }
    
    const user = getUserSession();
    if (!user) {
      throw new Error("Sesi telah berakhir");
    }
    
    // Log akses
    Logger.log(`Accessing attendance history for user ${user.username}`);
    Logger.log(`Date range: ${startDate} to ${endDate}`);
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance');
    if (!sheet) {
      throw new Error("Sheet Attendance tidak ditemukan");
    }
    
    // Proses data...
    const data = sheet.getDataRange().getValues();
    
    // Filter dan proses data...
    const attendance = processAttendanceData(data, user, startDate, endDate);
    
    return {
      success: true,
      data: attendance
    };
    
  } catch (error) {
    // Format error dan log
    const formattedError = formatError(error);
    logAttendanceError(formattedError, {
      startDate: startDate,
      endDate: endDate,
      functionName: 'getAttendanceHistory'
    });
    
    return {
      success: false,
      message: "Gagal memuat data: " + formattedError.message
    };
  }
}

// Fungsi helper untuk memproses data absensi
function processAttendanceData(data, user, startDate, endDate) {
  try {
    const headers = data[0];
    const startDateTime = new Date(startDate);
    const endDateTime = new Date(endDate);
    
    // Filter data
    const filteredData = data.slice(1).filter(row => {
      try {
        const rowDate = new Date(row[2]);
        return row[1] === user.nip &&
               rowDate >= startDateTime &&
               rowDate <= endDateTime;
      } catch (e) {
        logAttendanceError(e, {
          row: row,
          functionName: 'processAttendanceData/filter'
        });
        return false;
      }
    });
    
    // Map data ke format yang diinginkan
    const attendance = filteredData.map(row => {
      try {
        return {
          date: formatDate(new Date(row[2])),
          checkIn: row[3] || null,
          checkOut: row[4] || null,
          status: row[5] || "Tidak Ada Status",
          keterangan: row[6] || "",
          checkInPhoto: row[7] || null,
          location: row[9] || null,
          checkOutPhoto: row[10] || null,
          totalHours: calculateWorkHours(row[3], row[4])
        };
      } catch (e) {
        logAttendanceError(e, {
          row: row,
          functionName: 'processAttendanceData/map'
        });
        return null;
      }
    }).filter(item => item !== null);
    
    return {
      attendance: attendance,
      summary: calculateSummary(attendance),
      workHours: calculateWorkHoursChart(attendance)
    };
    
  } catch (error) {
    logAttendanceError(error, {
      functionName: 'processAttendanceData'
    });
    throw error;
  }
}

// Helper function untuk menghitung jam kerja
function calculateWorkHours(checkIn, checkOut) {
  if (!checkIn || !checkOut) return 0;
  
  try {
    const [checkInHour, checkInMin] = checkIn.split(':').map(Number);
    const [checkOutHour, checkOutMin] = checkOut.split(':').map(Number);
    
    const totalMinutes = (checkOutHour * 60 + checkOutMin) - 
                        (checkInHour * 60 + checkInMin);
    
    return Math.max(0, totalMinutes / 60);
  } catch (e) {
    logAttendanceError(e, {
      checkIn: checkIn,
      checkOut: checkOut,
      functionName: 'calculateWorkHours'
    });
    return 0;
  }
}

/**
 * Fungsi utilitas untuk mengubah objek waktu ke string format HH:MM
 * 
 * @param {object} waktuObj - Objek waktu {jam, menit}
 * @return {string} String waktu dalam format HH:MM
 */
function formatWaktuKeString(waktuObj) {
  if (!waktuObj || typeof waktuObj.jam !== 'number' || typeof waktuObj.menit !== 'number') {
    return "00:00";
  }
  
  var jam = waktuObj.jam.toString().padStart(2, '0');
  var menit = waktuObj.menit.toString().padStart(2, '0');
  
  return jam + ":" + menit;
}

/**
 * Tampilkan status kehadiran berdasarkan waktu absen dan jadwal
 * 
 * @param {Date} waktuAbsen - Waktu absen
 * @param {object} jadwalKerja - Objek jadwal kerja
 * @param {string} tipeAbsen - 'masuk' atau 'pulang'
 * @return {object} Informasi status kehadiran
 */
function hitungStatusKehadiran(waktuAbsen, jadwalKerja, tipeAbsen) {
  try {
    // Validasi parameter
    if (!waktuAbsen || !(waktuAbsen instanceof Date)) {
      throw new Error("Parameter waktuAbsen tidak valid");
    }
    
    if (!jadwalKerja || typeof jadwalKerja !== 'object') {
      throw new Error("Parameter jadwalKerja tidak valid");
    }
    
    if (!tipeAbsen || !['masuk', 'pulang'].includes(tipeAbsen)) {
      throw new Error("Parameter tipeAbsen tidak valid");
    }
    
    // Cek status hari libur
    if (jadwalKerja.status === "Libur") {
      return {
        status: "Lembur",
        keterangan: jadwalKerja.isHariKhusus ? 
          `Lembur pada ${jadwalKerja.keterangan}` : 
          "Lembur pada hari libur",
        selisihMenit: 0,
        timestamp: new Date().toISOString()
      };
    }
    
    var jamAbsen = waktuAbsen.getHours();
    var menitAbsen = waktuAbsen.getMinutes();
    
    // Validasi waktu jadwal
    if (tipeAbsen === 'masuk') {
      if (!jadwalKerja.waktuMulai || 
          typeof jadwalKerja.waktuMulai.jam !== 'number' || 
          typeof jadwalKerja.waktuMulai.menit !== 'number') {
        throw new Error("Waktu mulai tidak valid dalam jadwal");
      }
      
      var waktuMulai = jadwalKerja.waktuMulai;
      var batasTerlambat = jadwalKerja.batasTerlambat || 15; // Default 15 menit
      
      // Konversi ke menit untuk perbandingan
      var totalMenitJadwal = waktuMulai.jam * 60 + waktuMulai.menit;
      var totalMenitAbsen = jamAbsen * 60 + menitAbsen;
      var selisihMenit = totalMenitAbsen - totalMenitJadwal;
      
      // Format waktu untuk keterangan
      var waktuMulaiStr = `${String(waktuMulai.jam).padStart(2, '0')}:${String(waktuMulai.menit).padStart(2, '0')}`;
      
      if (selisihMenit <= 0) {
        return {
          status: "Hadir",
          keterangan: `Tepat waktu (${waktuMulaiStr})`,
          selisihMenit: selisihMenit,
          timestamp: new Date().toISOString()
        };
      } else if (selisihMenit <= batasTerlambat) {
        return {
          status: "Hadir",
          keterangan: `Masih dalam toleransi (Terlambat ${selisihMenit} menit)`,
          selisihMenit: selisihMenit,
          timestamp: new Date().toISOString()
        };
      } else {
        return {
          status: "Terlambat",
          keterangan: `Terlambat ${selisihMenit} menit (Jadwal: ${waktuMulaiStr})`,
          selisihMenit: selisihMenit,
          timestamp: new Date().toISOString()
        };
      }
    } else if (tipeAbsen === 'pulang') {
      if (!jadwalKerja.waktuSelesai || 
          typeof jadwalKerja.waktuSelesai.jam !== 'number' || 
          typeof jadwalKerja.waktuSelesai.menit !== 'number') {
        throw new Error("Waktu selesai tidak valid dalam jadwal");
      }
      
      var waktuSelesai = jadwalKerja.waktuSelesai;
      var toleransiPulangAwal = jadwalKerja.toleransiPulangAwal || 30; // Default 30 menit
      
      // Konversi ke menit untuk perbandingan
      var totalMenitJadwal = waktuSelesai.jam * 60 + waktuSelesai.menit;
      var totalMenitAbsen = jamAbsen * 60 + menitAbsen;
      var selisihMenit = totalMenitAbsen - totalMenitJadwal;
      
      // Format waktu untuk keterangan
      var waktuSelesaiStr = `${String(waktuSelesai.jam).padStart(2, '0')}:${String(waktuSelesai.menit).padStart(2, '0')}`;
      
      if (selisihMenit >= 0) {
        let status = selisihMenit > 60 ? "Lembur" : "Pulang";
        let keterangan = selisihMenit > 60 ? 
          `Lembur ${Math.floor(selisihMenit/60)} jam ${selisihMenit%60} menit` : 
          `Tepat waktu (${waktuSelesaiStr})`;
        
        return {
          status: status,
          keterangan: keterangan,
          selisihMenit: selisihMenit,
          timestamp: new Date().toISOString()
        };
      } else if (selisihMenit >= -toleransiPulangAwal) {
        return {
          status: "Pulang",
          keterangan: `Dalam toleransi pulang awal (${Math.abs(selisihMenit)} menit)`,
          selisihMenit: selisihMenit,
          timestamp: new Date().toISOString()
        };
      } else {
        return {
          status: "Pulang Awal",
          keterangan: `Pulang ${Math.abs(selisihMenit)} menit lebih awal (Jadwal: ${waktuSelesaiStr})`,
          selisihMenit: selisihMenit,
          timestamp: new Date().toISOString()
        };
      }
    }
    
  } catch (error) {
    Logger.log("Error menghitung status kehadiran: " + error.message);
    console.error("Detail error:", error);
    
    return {
      status: "Error",
      keterangan: "Terjadi kesalahan: " + error.message,
      selisihMenit: 0,
      timestamp: new Date().toISOString(),
      error: error
    };
  }
}

// ======== FUNGSI UTILITAS WAKTU ========

/**
 * Mengubah string tanggal dan waktu menjadi objek Date
 * 
 * @param {string} dateStr - String tanggal dalam format 'YYYY-MM-DD'
 * @param {string} timeStr - String waktu dalam format 'HH:MM:SS' atau 'HH:MM'
 * @return {Date} Objek Date yang berisi waktu yang ditentukan
 */
function parseTime(dateStr, timeStr) {
  try {
    // Validasi input
    if (!dateStr || !timeStr || typeof dateStr !== 'string' || typeof timeStr !== 'string') {
      Logger.log("Format tanggal atau waktu tidak valid");
      return new Date(); // Return waktu saat ini jika format tidak valid
    }
    
    // Pastikan format tanggal benar
    if (!dateStr.includes('-') || dateStr.split('-').length !== 3) {
      Logger.log("Format tanggal harus YYYY-MM-DD");
      return new Date();
    }
    
    // Pastikan format waktu benar
    if (!timeStr.includes(':')) {
      Logger.log("Format waktu harus HH:MM:SS atau HH:MM");
      return new Date();
    }
    
    // Parse tanggal
    var [year, month, day] = dateStr.split('-').map(Number);
    
    // Parse waktu, mendukung format HH:MM:SS atau HH:MM
    var timeParts = timeStr.split(':').map(Number);
    var hour = timeParts[0] || 0;
    var minute = timeParts[1] || 0;
    var second = timeParts[2] || 0;
    
    // Validasi nilai komponen waktu
    if (isNaN(year) || isNaN(month) || isNaN(day) || 
        isNaN(hour) || isNaN(minute) || isNaN(second)) {
      Logger.log("Komponen tanggal/waktu tidak valid");
      return new Date();
    }
    
    // Validasi rentang
    if (year < 2000 || year > 2100 || month < 1 || month > 12 || day < 1 || day > 31 ||
        hour < 0 || hour > 23 || minute < 0 || minute > 59 || second < 0 || second > 59) {
      Logger.log("Nilai komponen tanggal/waktu di luar rentang valid");
      return new Date();
    }
    
    // Bulan di JavaScript dimulai dari 0 (Januari = 0)
    var result = new Date(year, month-1, day, hour, minute, second);
    
    // Cek validitas hasil
    if (isNaN(result.getTime())) {
      Logger.log("Hasil konversi tanggal tidak valid");
      return new Date();
    }
    
    return result;
  } catch (error) {
    Logger.log("Error memparsing waktu: " + error.message);
    return new Date();  // Return waktu saat ini jika parsing gagal
  }
}

/**
 * Hitung durasi antara dua string waktu
 * 
 * @param {string} waktuMulai - Waktu mulai (format: "JJ:MM:DD")
 * @param {string} waktuSelesai - Waktu selesai (format: "JJ:MM:DD")
 * @return {string} Durasi dalam format "J:MM"
 */
function hitungDurasi(waktuMulai, waktuSelesai) {
  if (!waktuMulai || !waktuSelesai) return "-";
  
  try {
    // Parse jam dan menit
    var bagianMulai = waktuMulai.split(':');
    var bagianSelesai = waktuSelesai.split(':');
    
    var jamMulai = parseInt(bagianMulai[0], 10);
    var menitMulai = parseInt(bagianMulai[1], 10);
    
    var jamSelesai = parseInt(bagianSelesai[0], 10);
    var menitSelesai = parseInt(bagianSelesai[1], 10);
    
    // Hitung total menit
    var totalMenitMulai = (jamMulai * 60) + menitMulai;
    var totalMenitSelesai = (jamSelesai * 60) + menitSelesai;
    
    // Jika waktu selesai lebih awal (hari berikutnya), tambahkan 24 jam
    if (totalMenitSelesai < totalMenitMulai) {
      totalMenitSelesai += 24 * 60;
    }
    
    var durasiMenit = totalMenitSelesai - totalMenitMulai;
    var durasiJam = Math.floor(durasiMenit / 60);
    var sisaMenit = durasiMenit % 60;
    
    // Format sebagai "J:MM"
    return durasiJam + ":" + (sisaMenit < 10 ? "0" : "") + sisaMenit;
  } catch (error) {
    Logger.log("Error menghitung durasi: " + error.message);
    return "-";
  }
}

// ======== FUNGSI PENGELOLAAN FOTO ========

/**
 * Kompres dan simpan foto ke Google Drive dengan penamaan standar
 * Menyertakan informasi hari dan waktu dalam pengelolaan foto
 * 
 * @param {string} photoDataUrl - Data URL dari foto
 * @param {string} nip - NIP pegawai
 * @param {string} tipeAbsen - Tipe absen ('masuk' atau 'pulang')
 * @param {Date} waktuAbsen - Waktu absen dilakukan
 * @return {string} URL foto yang disimpan
 */
function kompresidanSimpanFoto(photoDataUrl, nip, tipeAbsen, waktuAbsen) {
  try {
    // Hapus awalan data URL untuk mendapatkan data base64 saja
    var base64Data = photoDataUrl.split(',')[1];
    
    // Format tanggal untuk nama file dan folder
    var tanggal = Utilities.formatDate(waktuAbsen, "GMT+7", "yyyy-MM-dd");
    var hariDalamMinggu = waktuAbsen.getDay(); // 0=Minggu, 1=Senin, dst
    var waktu = Utilities.formatDate(waktuAbsen, "GMT+7", "HH-mm-ss");
    
    // Nama-nama hari dalam bahasa Indonesia untuk folder
    var namaHari = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
    var hariFull = namaHari[hariDalamMinggu];
    
    // Buat nama file yang mencakup semua informasi penting
    var fileName = nip + "_" + tipeAbsen + "_" + tanggal + "_" + waktu + ".jpg";
    
    // Kompres gambar dengan kualitas yang lebih rendah (80%)
    var kompresiByte = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(kompresiByte, 'image/jpeg', fileName);
    
    // Dapatkan pengaturan jam kerja untuk hari ini
    var jamKerja = getWaktuMulaiUntukHari(hariDalamMinggu);
    var statusJam = "reguler";
    
    // Cek apakah absen dilakukan di luar jam reguler
    if (tipeAbsen === 'masuk') {
      var waktuMulaiKerja = new Date(waktuAbsen);
      waktuMulaiKerja.setHours(jamKerja.jam, jamKerja.menit, 0, 0);
      
      // Jika absen masuk 30 menit lebih awal atau lebih lambat dari jadwal
      if (Math.abs(waktuAbsen.getTime() - waktuMulaiKerja.getTime()) > 30 * 60 * 1000) {
        statusJam = "nonreguler";
      }
    } else if (tipeAbsen === 'pulang') {
      var jamPulang = getWaktuSelesaiUntukHari(hariDalamMinggu);
      var waktuPulangKerja = new Date(waktuAbsen);
      waktuPulangKerja.setHours(jamPulang.jam, jamPulang.menit, 0, 0);
      
      // Jika absen pulang 30 menit lebih awal atau lebih dari 2 jam lebih lambat
      if (waktuAbsen.getTime() < waktuPulangKerja.getTime() - 30 * 60 * 1000 || 
          waktuAbsen.getTime() > waktuPulangKerja.getTime() + 120 * 60 * 1000) {
        statusJam = "nonreguler";
      }
    }
    
    // Buat struktur folder yang lebih terorganisir
    var rootFolderName = "GASPUL_Hadir_Foto";
    var rootFolders = DriveApp.getFoldersByName(rootFolderName);
    var rootFolder;
    
    // Buat root folder jika belum ada
    if (rootFolders.hasNext()) {
      rootFolder = rootFolders.next();
    } else {
      rootFolder = DriveApp.createFolder(rootFolderName);
    }
    
    // Buat subfolder berdasarkan tahun-bulan
    var tahunBulan = tanggal.substring(0, 7); // "yyyy-MM"
    var tahunBulanFolders = rootFolder.getFoldersByName(tahunBulan);
    var tahunBulanFolder;
    
    if (tahunBulanFolders.hasNext()) {
      tahunBulanFolder = tahunBulanFolders.next();
    } else {
      tahunBulanFolder = rootFolder.createFolder(tahunBulan);
    }
    
    // Buat subfolder untuk NIP jika belum ada
    var nipFolders = tahunBulanFolder.getFoldersByName(nip);
    var nipFolder;
    
    if (nipFolders.hasNext()) {
      nipFolder = nipFolders.next();
    } else {
      nipFolder = tahunBulanFolder.createFolder(nip);
    }
    
    // Cek apakah file sudah ada dengan nama yang sama dan hapus
    var existingFiles = nipFolder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }
    
    // Kompres gambar sebelum disimpan dengan mengatur ukuran (kompresi tambahan)
    var compressedBlob;
    try {
      // Mencoba kompresi gambar lebih lanjut dengan Apps Script
      var imageBlob = blob.copyBlob();
      
      // Simpan dalam kualitas lebih rendah jika memungkinkan
      // Namun Apps Script terbatas dalam kemampuan kompresi gambar
      compressedBlob = imageBlob;
    } catch (e) {
      Logger.log("Error saat mengkompres tambahan: " + e.message);
      compressedBlob = blob; // Gunakan blob asli jika gagal
    }
    
    // Simpan file ke folder yang tepat
    var file = nipFolder.createFile(compressedBlob);
    
    // Tambahkan properti khusus ke file untuk memudahkan pencarian
    file.setDescription("Absensi " + tipeAbsen + " oleh " + nip + " pada " + tanggal + 
                       " (" + hariFull + "), status: " + statusJam);
    
    // Set properti custom untuk metadata
    file.setProperty("nip", nip);
    file.setProperty("tanggal", tanggal);
    file.setProperty("hari", hariFull);
    file.setProperty("tipe_absen", tipeAbsen);
    file.setProperty("status_jam", statusJam);
    
    // Kembalikan URL file
    return file.getUrl();
  } catch (error) {
    Logger.log("Error saat mengkompresi dan menyimpan foto: " + error.message);
    return photoDataUrl; // Kembalikan yang asli jika terjadi error
  }
}

// ======== FUNGSI ABSENSI ========

/**
 * Mendapatkan data user dan absensi
 */
function getUserDataWithAttendance() {
  try {
    // Dapatkan data user dari session
    var user = getUserSession();
    if (!user) {
      return { success: false, message: "Sesi telah berakhir" };
    }
    
    // Dapatkan data absensi hari ini
    var today = new Date();
    var date = Utilities.formatDate(today, "GMT+7", "yyyy-MM-dd");
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance');
    var data = sheet.getDataRange().getValues();
    var todayAttendance = null;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] == user.nip && data[i][2] == date) {
        todayAttendance = {
          id: data[i][0],
          checkIn: data[i][3],
          checkOut: data[i][4],
          status: data[i][5],
          notes: data[i][6],
          photo: data[i][7] ? true : false,
          gps: data[i][8],
          location: data[i][9]
        };
        break;
      }
    }
    
    // Dapatkan statistik kehadiran
    var stats = getUserStats(user.nip);
    
    return {
      success: true,
      user: user,
      attendance: todayAttendance,
      stats: stats
    };
  } catch (error) {
    Logger.log("Error pada fungsi getUserDataWithAttendance: " + error.message);
    return { success: false, message: "Terjadi kesalahan: " + error.message };
  }
}


/**
 * Fungsi untuk melakukan check-in dengan foto
 * Menyimpan data masuk karyawan, foto, lokasi, dan informasi waktu
 * 
 * @param {string} photoDataUrl - Data URL dari foto yang diambil
 * @param {string} gps - Koordinat GPS dalam format "latitude,longitude"
 * @param {string} location - Alamat lokasi dalam teks
 * @return {Object} Objek hasil operasi dengan status keberhasilan
 */
function checkInWithPhoto(photoDataUrl, gps, location) {
  try {
    // Dapatkan data user dari session
    var userSession = getUserSession();
    if (!userSession) {
      return { success: false, message: "Sesi telah berakhir" };
    }
    
    var nip = userSession.nip;
    
    // Validasi input
    if (!photoDataUrl) {
      return { success: false, message: "Foto tidak tersedia" };
    }
    
    // Validasi GPS
    if (!gps || gps.trim() === "") {
      return { 
        success: false, 
        message: "Data GPS tidak tersedia. Pastikan lokasi diizinkan." 
      };
    }
    
    // Validasi format GPS
    var gpsCoords = gps.split(',');
    if (gpsCoords.length !== 2 || isNaN(gpsCoords[0]) || isNaN(gpsCoords[1])) {
      return { 
        success: false, 
        message: "Format GPS tidak valid" 
      };
    }
    
    // Jika lokasi tidak ada, coba dapatkan dari koordinat
    if (!location || location.trim() === "") {
      location = getLocationAddress(gpsCoords[0], gpsCoords[1]);
    }
    
    // Dapatkan informasi waktu dan tanggal saat ini
    var now = new Date();
    var date = Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd");
    var time = Utilities.formatDate(now, "GMT+7", "HH:mm:ss");
    
    // Dapatkan jadwal kerja untuk hari ini
    var jadwalKerja = getJadwalKerja(now);
    if (!jadwalKerja) {
      return { 
        success: false, 
        message: "Gagal mendapatkan jadwal kerja" 
      };
    }
    
    // Hitung status kehadiran
    var statusKehadiran = hitungStatusKehadiran(now, jadwalKerja, 'masuk');
    
    // Cek apakah sudah absen masuk hari ini
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance');
    if (!sheet) {
      throw new Error("Sheet Attendance tidak ditemukan");
    }
    
    var data = sheet.getDataRange().getValues();
    var existingRow = -1;
    
    // Cari data absensi hari ini
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === nip && data[i][2] === date) {
        existingRow = i + 1;
        // Cek jika sudah absen masuk
        if (data[i][3] && data[i][3] !== "") {
          return { 
            success: false, 
            message: "Anda sudah melakukan absen masuk hari ini" 
          };
        }
        break;
      }
    }
    
    // Kompres dan simpan foto
    var photoUrl = kompresidanSimpanFoto(photoDataUrl, nip, "masuk", now);
    if (!photoUrl) {
      throw new Error("Gagal menyimpan foto");
    }
    
    if (existingRow > 0) {
      // Update baris yang sudah ada
      var range = sheet.getRange(existingRow, 4, 1, 7);
      range.setValues([[
        time,                    // Jam Masuk
        statusKehadiran.status,  // Status
        statusKehadiran.keterangan, // Keterangan
        photoUrl,                // Foto Masuk
        gps,                     // GPS
        location,                // Lokasi
        jadwalKerja.namaHari    // Nama Hari
      ]]);
    } else {
      // Buat entri baru
      var newId = sheet.getLastRow() + 1;
      sheet.appendRow([
        newId,                    // ID
        nip,                      // NIP
        date,                     // Tanggal
        time,                     // Jam Masuk
        "",                       // Jam Pulang
        statusKehadiran.status,   // Status
        statusKehadiran.keterangan, // Keterangan
        photoUrl,                 // Foto Masuk
        gps,                      // GPS
        location,                 // Lokasi
        "",                       // Foto Pulang
        jadwalKerja.namaHari,    // Nama Hari
        ""                        // Total Jam Kerja
      ]);
    }
    
    // Log aktivitas
    Logger.log("Absen masuk berhasil untuk NIP: " + nip + " pada: " + time);
    
    return { 
      success: true, 
      message: "Absen masuk berhasil", 
      time: time, 
      status: statusKehadiran.status 
    };
    
  } catch (error) {
    Logger.log("Error pada fungsi checkInWithPhoto: " + error.message);
    return { 
      success: false, 
      message: "Terjadi kesalahan saat absen masuk: " + error.message 
    };
  }
}

/**
 * Fungsi untuk melakukan check-out dengan foto
 * Menyimpan data pulang karyawan, foto, dan menghitung total jam kerja
 * 
 * @param {string} photoDataUrl - Data URL dari foto yang diambil
 * @param {string} gps - Koordinat GPS dalam format "latitude,longitude"
 * @param {string} location - Alamat lokasi dalam teks
 * @return {Object} Objek hasil operasi dengan status keberhasilan
 */
function checkOutWithPhoto(photoDataUrl, gps, location) {
  try {
    // Dapatkan user dari session
    var userSession = getUserSession();
    if (!userSession) {
      return { success: false, message: "Sesi telah berakhir" };
    }
    
    var nip = userSession.nip;
    
    // Validasi input
    if (!photoDataUrl) {
      return { success: false, message: "Foto tidak tersedia" };
    }
    
    // Validasi GPS
    if (!gps || gps.trim() === "") {
      return { 
        success: false, 
        message: "Data GPS tidak tersedia. Pastikan lokasi diizinkan." 
      };
    }
    
    // Validasi format GPS
    var gpsCoords = gps.split(',');
    if (gpsCoords.length !== 2 || isNaN(gpsCoords[0]) || isNaN(gpsCoords[1])) {
      return { 
        success: false, 
        message: "Format GPS tidak valid" 
      };
    }
    
    // Jika lokasi tidak ada, coba dapatkan dari koordinat
    if (!location || location.trim() === "") {
      location = getLocationAddress(gpsCoords[0], gpsCoords[1]);
    }
    
    // Dapatkan informasi waktu saat ini
    var now = new Date();
    var date = Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd");
    var time = Utilities.formatDate(now, "GMT+7", "HH:mm:ss");
    
    // Dapatkan jadwal kerja untuk hari ini
    var jadwalKerja = getJadwalKerja(now);
    if (!jadwalKerja) {
      return { 
        success: false, 
        message: "Gagal mendapatkan jadwal kerja" 
      };
    }
    
    // Hitung status kehadiran
    var statusKehadiran = hitungStatusKehadiran(now, jadwalKerja, 'pulang');
    
    // Dapatkan sheet attendance
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance');
    if (!sheet) {
      throw new Error("Sheet Attendance tidak ditemukan");
    }
    
    // Cek apakah sudah absen masuk hari ini
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    // Cari data absensi hari ini
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === nip && data[i][2] === date) {
        rowIndex = i + 1;
        
        // Cek jika sudah absen pulang
        if (data[i][4] && data[i][4] !== "") {
          return { 
            success: false, 
            message: "Anda sudah melakukan absen pulang hari ini" 
          };
        }
        
        // Pastikan sudah absen masuk
        if (!data[i][3] || data[i][3] === "") {
          return { 
            success: false, 
            message: "Anda belum absen masuk hari ini" 
          };
        }
        break;
      }
    }
    
    // Kompres dan simpan foto
    var photoUrl = kompresidanSimpanFoto(photoDataUrl, nip, "pulang", now);
    if (!photoUrl) {
      throw new Error("Gagal menyimpan foto");
    }
    
    if (rowIndex > 0) {
      // Update data absensi yang sudah ada
      var waktuMasuk = data[rowIndex-1][3]; // Jam masuk
      var totalJam = hitungDurasi(waktuMasuk, time);
      var keterangan = data[rowIndex-1][6];
      
      if (statusKehadiran.status === "Pulang Awal") {
        keterangan = keterangan + " | " + statusKehadiran.keterangan;
      }
      
      var range = sheet.getRange(rowIndex, 5, 1, 9);
      range.setValues([[
        time,           // Jam Pulang
        gps,            // GPS
        location,       // Lokasi
        photoUrl,       // Foto Pulang
        totalJam,       // Total Jam Kerja
        keterangan,     // Update Keterangan
        statusKehadiran.status, // Status Pulang
        jadwalKerja.namaHari,   // Nama Hari
        ""             // Kolom tambahan jika ada
      ]]);
      
    } else {
      // Buat entri baru untuk absen pulang saja
      var newId = sheet.getLastRow() + 1;
      sheet.appendRow([
        newId,                    // ID
        nip,                      // NIP
        date,                     // Tanggal
        "",                       // Jam Masuk
        time,                     // Jam Pulang
        "Pulang Tanpa Absen Masuk", // Status
        "Hanya absen pulang",     // Keterangan
        "",                       // Foto Masuk
        gps,                      // GPS
        location,                 // Lokasi
        photoUrl,                 // Foto Pulang
        jadwalKerja.namaHari,    // Nama Hari
        ""                        // Total Jam Kerja
      ]);
    }
    
    // Log aktivitas
    Logger.log("Absen pulang berhasil untuk NIP: " + nip + " pada: " + time);
    
    return { 
      success: true, 
      message: rowIndex > 0 ? "Absen pulang berhasil" : "Absen pulang berhasil (tanpa absen masuk)", 
      time: time,
      status: statusKehadiran.status 
    };
    
  } catch (error) {
    Logger.log("Error pada fungsi checkOutWithPhoto: " + error.message);
    return { 
      success: false, 
      message: "Terjadi kesalahan saat absen pulang: " + error.message 
    };
  }
}

// Tambahkan fungsi ini untuk memeriksa dan meminta izin lokasi
function requestLocationPermission() {
  if (navigator.geolocation) {
    navigator.geolocation.getCurrentPosition(
      function(position) {
        // Sukses mendapatkan lokasi
        return {
          success: true,
          coords: {
            latitude: position.coords.latitude,
            longitude: position.coords.longitude
          }
        };
      },
      function(error) {
        // Error mendapatkan lokasi
        return {
          success: false,
          error: error.message
        };
      }
    );
  } else {
    return {
      success: false,
      error: "Geolocation tidak didukung di browser ini"
    };
  }
}

function getLocationAddress(latitude, longitude) {
  try {
    var mapsApiKey = getMapsApiKey();
    var response = UrlFetchApp.fetch(
      `https://maps.googleapis.com/maps/api/geocode/json?latlng=${latitude},${longitude}&key=${mapsApiKey}`
    );
    var data = JSON.parse(response.getContentText());
    
    if (data.status === "OK" && data.results && data.results[0]) {
      return data.results[0].formatted_address;
    }
    return "Lokasi tidak ditemukan";
  } catch (error) {
    Logger.log("Error getting location address: " + error.message);
    return "Error: " + error.message;
  }
}

/**
 * Mendapatkan riwayat absensi
 * @param {string} nip - NIP karyawan (opsional, akan menggunakan dari session jika kosong)
 * @param {number} month - Bulan (1-12)
 * @param {number} year - Tahun
 * @return {Array} Array objek riwayat absensi
 */

function getAttendanceHistory(startDate, endDate) {
  try {
    // Log awal eksekusi
    Logger.log("Starting getAttendanceHistory execution");
    Logger.log(`Parameters - startDate: ${startDate}, endDate: ${endDate}`);

    // Validasi input
    if (!startDate || !endDate) {
      throw new Error("Parameter tanggal tidak lengkap");
    }

    // Get user session
    const userSession = getUserSession();
    if (!userSession || !userSession.nip) {
      throw new Error("Sesi user tidak valid atau NIP tidak ditemukan");
    }
    Logger.log(`User session found - NIP: ${userSession.nip}`);

    // Get spreadsheet and sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Attendance');
    if (!sheet) {
      throw new Error("Sheet 'Attendance' tidak ditemukan");
    }
    Logger.log("Sheet 'Attendance' found");

    // Get data range
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    if (lastRow <= 1) {
      Logger.log("No attendance data found");
      return {
        success: true,
        data: {
          attendance: [],
          summary: createEmptySummary(),
          workHours: { dates: [], hours: [] }
        }
      };
    }

    // Get all data
    const range = sheet.getRange(1, 1, lastRow, lastColumn);
    const values = range.getValues();
    Logger.log(`Retrieved ${values.length} rows of data`);

    // Parse dates
    const startDateTime = new Date(startDate);
    const endDateTime = new Date(endDate);
    startDateTime.setHours(0, 0, 0, 0);
    endDateTime.setHours(23, 59, 59, 999);

    // Filter and process data
    const attendance = processAttendanceRecords(values, userSession.nip, startDateTime, endDateTime);
    Logger.log(`Processed ${attendance.length} attendance records`);

    // Calculate summary and work hours
    const summary = calculateAttendanceSummary(attendance);
    const workHours = calculateWorkHoursData(attendance);

    return {
      success: true,
      data: {
        attendance: attendance,
        summary: summary,
        workHours: workHours
      }
    };

  } catch (error) {
    const errorMessage = error.message || "Unknown error occurred";
    Logger.log(`Error in getAttendanceHistory: ${errorMessage}`);
    console.error("Full error details:", error);

    return {
      success: false,
      message: `Gagal memuat data: ${errorMessage}`
    };
  }
}

function processAttendanceRecords(values, nip, startDate, endDate) {
  const headers = values[0];
  const records = [];

  for (let i = 1; i < values.length; i++) {
    try {
      const row = values[i];
      const recordDate = new Date(row[2]); // Assuming date is in column 3
      
      if (row[1] === nip && recordDate >= startDate && recordDate <= endDate) {
        records.push({
          date: formatDate(recordDate),
          checkIn: row[3] || null,
          checkOut: row[4] || null,
          status: row[5] || "Tidak Ada Status",
          keterangan: row[6] || "",
          checkInPhoto: row[7] || null,
          location: row[9] || null,
          checkOutPhoto: row[10] || null,
          totalHours: calculateTotalHours(row[3], row[4])
        });
      }
    } catch (error) {
      Logger.log(`Error processing row ${i}: ${error.message}`);
      // Continue to next record
    }
  }

  return records;
}

function calculateAttendanceSummary(attendance) {
  try {
    const totalWorkDays = attendance.length;
    const onTimeCount = attendance.filter(r => r.status === "Hadir" || r.status === "Tepat Waktu").length;
    const lateCount = attendance.filter(r => r.status === "Terlambat").length;
    
    let totalHours = 0;
    let validDaysCount = 0;

    attendance.forEach(record => {
      if (record.totalHours && !isNaN(record.totalHours)) {
        totalHours += record.totalHours;
        validDaysCount++;
      }
    });

    return {
      totalWorkDays: totalWorkDays,
      onTimeCount: onTimeCount,
      lateCount: lateCount,
      avgWorkHours: validDaysCount > 0 ? (totalHours / validDaysCount) : 0
    };
  } catch (error) {
    Logger.log(`Error calculating summary: ${error.message}`);
    return createEmptySummary();
  }
}

function calculateWorkHoursData(attendance) {
  try {
    const dates = [];
    const hours = [];

    attendance.forEach(record => {
      dates.push(formatDisplayDate(new Date(record.date)));
      hours.push(record.totalHours || 0);
    });

    return { dates, hours };
  } catch (error) {
    Logger.log(`Error calculating work hours data: ${error.message}`);
    return { dates: [], hours: [] };
  }
}

function calculateTotalHours(checkIn, checkOut) {
  if (!checkIn || !checkOut) return 0;
  
  try {
    const [inHours, inMinutes] = checkIn.split(':').map(Number);
    const [outHours, outMinutes] = checkOut.split(':').map(Number);
    
    const totalMinutes = (outHours * 60 + outMinutes) - (inHours * 60 + inMinutes);
    return Math.max(0, totalMinutes / 60);
  } catch (error) {
    Logger.log(`Error calculating total hours: ${error.message}`);
    return 0;
  }
}

function createEmptySummary() {
  return {
    totalWorkDays: 0,
    onTimeCount: 0,
    lateCount: 0,
    avgWorkHours: 0
  };
}

function formatDate(date) {
  try {
    return Utilities.formatDate(date, "GMT+7", "yyyy-MM-dd");
  } catch (error) {
    Logger.log(`Error formatting date: ${error.message}`);
    return date.toISOString().split('T')[0];
  }
}

function formatDisplayDate(date) {
  try {
    return Utilities.formatDate(date, "GMT+7", "dd/MM");
  } catch (error) {
    Logger.log(`Error formatting display date: ${error.message}`);
    return date.toLocaleDateString();
  }
}

function calculateSummary(attendance) {
  try {
    const totalWorkDays = attendance.length;
    
    const onTimeCount = attendance.filter(record => 
      record.status === "Hadir" || record.status === "Tepat Waktu").length;
    
    const lateCount = attendance.filter(record => 
      record.status === "Terlambat").length;
    
    let totalHours = 0;
    let validDaysCount = 0;
    
    attendance.forEach(record => {
      if (record.totalHours) {
        const hours = parseFloat(record.totalHours);
        if (!isNaN(hours)) {
          totalHours += hours;
          validDaysCount++;
        }
      }
    });
    
    const avgWorkHours = validDaysCount > 0 ? 
      totalHours / validDaysCount : 0;
    
    return {
      totalWorkDays,
      onTimeCount,
      lateCount,
      avgWorkHours
    };
  } catch (error) {
    Logger.log("Error in calculateSummary: " + error.message);
    return {
      totalWorkDays: 0,
      onTimeCount: 0,
      lateCount: 0,
      avgWorkHours: 0
    };
  }
}

function calculateWorkHours(attendance) {
  try {
    const dates = [];
    const hours = [];
    
    attendance.forEach(record => {
      const date = record.date;
      const totalHours = record.totalHours ? 
        parseFloat(record.totalHours) : 0;
      
      if (!isNaN(totalHours)) {
        dates.push(formatDisplayDate(new Date(date)));
        hours.push(totalHours);
      }
    });
    
    return { dates, hours };
  } catch (error) {
    Logger.log("Error in calculateWorkHours: " + error.message);
    return { dates: [], hours: [] };
  }
}

function formatDate(date) {
  try {
    return Utilities.formatDate(date, "GMT+7", "yyyy-MM-dd");
  } catch (error) {
    Logger.log("Error formatting date: " + error.message);
    return date.toISOString().split('T')[0];
  }
}

function formatDisplayDate(date) {
  try {
    return Utilities.formatDate(date, "GMT+7", "dd/MM");
  } catch (error) {
    Logger.log("Error formatting display date: " + error.message);
    return date.toLocaleDateString();
  }
}

function getAttendanceHistory(nip, month, year) {
  try {
    // Jika tidak ada nip, ambil dari session
    if (!nip) {
      var user = getUserSession();
      if (!user) {
        return [];
      }
      nip = user.nip;
    }
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance');
    var data = sheet.getDataRange().getValues();
    var history = [];
    
    // Cek parameter
    month = month ? parseInt(month) : null;
    year = year ? parseInt(year) : null;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] == nip) {
        var rowDate = new Date(data[i][2]);
        
        if ((month == null || rowDate.getMonth() + 1 == month) && 
            (year == null || rowDate.getFullYear() == year)) {
          
          // Pastikan semua properti didefinisikan dengan benar
          var attendanceRecord = {
            id: data[i][0],
            date: data[i][2],
            checkIn: data[i][3],
            checkOut: data[i][4],
            status: data[i][5],
            notes: data[i][6],
            photo: data[i][7] ? true : false,
            gps: data[i][8],
            location: data[i][9],
            photoOut: data[i][10] ? true : false,
            dayName: data[i][11],
            totalHours: data[i][12]
          };
          
          history.push(attendanceRecord);
        }
      }
    }
    
    return history;
  } catch (error) {
    Logger.log("Error pada fungsi getAttendanceHistory: " + error.message);
    return [];
  }
}

/**
 * Mendapatkan data riwayat absensi dengan ringkasan
 */
function getAttendanceHistoryWithSummary(month, year) {
  try {
    // Dapatkan user dari session
    var user = getUserSession();
    if (!user) {
      return { success: false, message: "Sesi telah berakhir" };
    }
    
    var nip = user.nip;
    
    // Konversi parameter ke angka
    month = parseInt(month);
    year = parseInt(year);
    
    // Dapatkan history
    var history = getAttendanceHistory(nip, month, year);
    
    // Hitung ringkasan
    var summary = {
      totalDays: history.length,
      presentDays: 0,
      lateDays: 0,
      leaveDays: 0,
      totalWorkHours: "0:00"
    };
    
    var totalMinutes = 0;
    
    history.forEach(function(day) {
      if (day.status === "Hadir") {
        summary.presentDays++;
      } else if (day.status === "Terlambat") {
        summary.lateDays++;
        summary.presentDays++; // Terlambat juga dihitung hadir
      } else {
        summary.leaveDays++;
      }
      
      // Hitung total jam kerja
      if (day.totalHours) {
        var hoursMinutes = day.totalHours.split(':');
        if (hoursMinutes.length >= 2) {
          var hours = parseInt(hoursMinutes[0], 10) || 0;
          var minutes = parseInt(hoursMinutes[1], 10) || 0;
          totalMinutes += (hours * 60) + minutes;
        }
      }
    });
    
    // Format total jam kerja
    var totalHours = Math.floor(totalMinutes / 60);
    var remainingMinutes = totalMinutes % 60;
    summary.totalWorkHours = totalHours + ":" + (remainingMinutes < 10 ? "0" : "") + remainingMinutes;
    
    return {
      success: true,
      history: history,
      summary: summary
    };
  } catch (error) {
    Logger.log("Error pada fungsi getAttendanceHistoryWithSummary: " + error.message);
    return { success: false, message: "Terjadi kesalahan: " + error.message };
  }
}

/**
 * Mendapatkan riwayat absensi dengan detail lengkap
 */
function getExtendedHistoryData(month, year) {
  try {
    // Dapatkan hasil dari getAttendanceHistoryWithSummary
    var result = getAttendanceHistoryWithSummary(month, year);
    if (!result.success) {
      return result;
    }
    
    // Tambahkan data analisis durasi kerja per hari
    var dailyWorkHoursData = {
      dailyHours: {0: "0:00", 1: "0:00", 2: "0:00", 3: "0:00", 4: "0:00", 5: "0:00", 6: "0:00"}
    };
    
    // Counter untuk setiap hari
    var dayCounts = [0, 0, 0, 0, 0, 0, 0];
    var dayTotalMinutes = [0, 0, 0, 0, 0, 0, 0];
    
    result.history.forEach(function(record) {
      if (record.totalHours && record.totalHours !== "-") {
        var date = new Date(record.date);
        var dayOfWeek = date.getDay();
        
        var hourMinute = record.totalHours.split(':');
        if (hourMinute.length >= 2) {
          var hours = parseInt(hourMinute[0], 10) || 0;
          var minutes = parseInt(hourMinute[1], 10) || 0;
          var totalMinutes = (hours * 60) + minutes;
          
          dayTotalMinutes[dayOfWeek] += totalMinutes;
          dayCounts[dayOfWeek]++;
        }
      }
    });
    
    // Hitung rata-rata untuk setiap hari
    for (var i = 0; i < 7; i++) {
      if (dayCounts[i] > 0) {
        var avgMinutes = Math.round(dayTotalMinutes[i] / dayCounts[i]);
        var avgHours = Math.floor(avgMinutes / 60);
        var remainingMinutes = avgMinutes % 60;
        
        dailyWorkHoursData.dailyHours[i] = avgHours + ":" + (remainingMinutes < 10 ? "0" : "") + remainingMinutes;
      }
    }
    
    // Tambahkan data ke hasil
    result.workHoursData = dailyWorkHoursData;
    
    return result;
  } catch (error) {
    Logger.log("Error pada fungsi getExtendedHistoryData: " + error.message);
    return { success: false, message: "Terjadi kesalahan: " + error.message };
  }
}

/**
 * Mendapatkan statistik kehadiran user
 */
function getUserStats(nip) {
  try {
    // Jika tidak ada nip, ambil dari session
    if (!nip) {
      var user = getUserSession();
      if (!user) {
        return {
          presentDays: 0,
          leaveDays: 0,
          attendancePercentage: 0
        };
      }
      nip = user.nip;
    }
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance');
    var data = sheet.getDataRange().getValues();
    
    var presentDays = 0;
    var leaveDays = 0;
    var totalDays = 0;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] == nip) {
        totalDays++;
        
        if (data[i][5] == "Hadir" || data[i][5] == "Terlambat") {
          presentDays++;
        } else {
          leaveDays++;
        }
      }
    }
    
    var attendancePercentage = totalDays > 0 ? Math.round((presentDays / totalDays) * 100) : 0;
    
    return {
      presentDays: presentDays,
      leaveDays: leaveDays,
      attendancePercentage: attendancePercentage
    };
  } catch (error) {
    Logger.log("Error pada fungsi getUserStats: " + error.message);
    return {
      presentDays: 0,
      leaveDays: 0,
      attendancePercentage: 0
    };
  }
}

/**
 * Mendapatkan statistik jam kerja
 * Menghitung rata-rata jam kerja, total jam kerja, dan hari terlama
 * Juga menghitung rata-rata jam kerja per hari dalam seminggu
 * 
 * @return {Object} Statistik jam kerja
 */
function getWorkHoursStats() {
  try {
    // Dapatkan user dari session
    var user = getUserSession();
    if (!user) {
      return null;
    }
    
    var nip = user.nip;
    
    // Dapatkan bulan dan tahun saat ini
    var now = new Date();
    var currentMonth = now.getMonth() + 1;
    var currentYear = now.getFullYear();
    
    // Dapatkan history untuk bulan ini
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance');
    var data = sheet.getDataRange().getValues();
    var workHoursData = [];
    var dailyHours = [0, 0, 0, 0, 0, 0, 0]; // Minggu, Senin, ..., Sabtu
    var dailyCounts = [0, 0, 0, 0, 0, 0, 0]; // Counter untuk setiap hari
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] == nip) {
        var rowDate = new Date(data[i][2]);
        
        // Filter untuk bulan dan tahun ini
        if (rowDate.getMonth() + 1 == currentMonth && rowDate.getFullYear() == currentYear) {
          // Ambil total jam kerja jika ada
          if (data[i].length > 12 && data[i][12]) {
            var timeStr = data[i][12];
            
            // Parse jam:menit ke menit total
            var timeParts = timeStr.split(':');
            var hours = parseInt(timeParts[0], 10);
            var minutes = parseInt(timeParts[1], 10);
            var totalMinutes = hours * 60 + minutes;
            
            workHoursData.push({
              date: rowDate,
              day: rowDate.getDay(), // 0 = Minggu, 1 = Senin, dst.
              totalMinutes: totalMinutes,
              totalHoursStr: timeStr
            });
            
            // Tambahkan ke rata-rata per hari
            dailyHours[rowDate.getDay()] += totalMinutes;
            dailyCounts[rowDate.getDay()]++;
          }
        }
      }
    }
    
    if (workHoursData.length === 0) {
      // Tidak ada data, kembalikan nilai default
      return {
        averageHours: "0:00",
        totalHours: "0:00",
        longestDay: "0:00",
        dailyAverage: {
          0: "0:00", // Minggu
          1: "0:00", // Senin
          2: "0:00", // Selasa
          3: "0:00", // Rabu
          4: "0:00", // Kamis
          5: "0:00", // Jumat
          6: "0:00"  // Sabtu
        }
      };
    }
    
    // Hitung total menit
    var totalMinutes = workHoursData.reduce(function(sum, item) {
      return sum + item.totalMinutes;
    }, 0);
    
    // Hitung rata-rata menit per hari
    var avgMinutes = Math.round(totalMinutes / workHoursData.length);
    
    // Temukan hari terlama
    var longestDay = workHoursData.reduce(function(max, item) {
      return item.totalMinutes > max.totalMinutes ? item : max;
    }, { totalMinutes: 0, totalHoursStr: "0:00" });
    
    // Format total jam kerja bulan ini
    var totalHours = Math.floor(totalMinutes / 60);
    var totalMins = totalMinutes % 60;
    var totalHoursStr = totalHours + ":" + (totalMins < 10 ? "0" : "") + totalMins;
    
    // Format rata-rata jam kerja
    var avgHours = Math.floor(avgMinutes / 60);
    var avgMins = avgMinutes % 60;
    var avgHoursStr = avgHours + ":" + (avgMins < 10 ? "0" : "") + avgMins;
    
    // Format rata-rata per hari dalam seminggu
    var dailyAverageObj = {};
    for (var day = 0; day < 7; day++) {
      if (dailyCounts[day] > 0) {
        var dayAvgMinutes = Math.round(dailyHours[day] / dailyCounts[day]);
        var dayAvgHours = Math.floor(dayAvgMinutes / 60);
        var dayAvgMins = dayAvgMinutes % 60;
        dailyAverageObj[day] = dayAvgHours + ":" + (dayAvgMins < 10 ? "0" : "") + dayAvgMins;
      } else {
        dailyAverageObj[day] = "0:00";
      }
    }
    
    return {
      averageHours: avgHoursStr,
      totalHours: totalHoursStr,
      longestDay: longestDay.totalHoursStr,
      dailyAverage: dailyAverageObj
    };
  } catch (error) {
    Logger.log("Error pada fungsi getWorkHoursStats: " + error.message);
    return null;
  }
}

// ======== FUNGSI EKSPOR DAN LAPORAN ========

/**
 * Fungsi untuk mengekspor data absensi ke PDF
 */
function exportAttendanceToPdf(month, year) {
  try {
    // Dapatkan user dari session
    var user = getUserSession();
    if (!user) {
      return { success: false, message: "Sesi telah berakhir" };
    }
    
    var nip = user.nip;
    
    // Konversi parameter ke angka
    month = parseInt(month);
    year = parseInt(year);
    
    // Nama bulan dan tahun
    var monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
                      "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    var monthName = monthNames[month - 1];
    
    // Dapatkan data absensi
    var history = getAttendanceHistory(nip, month, year);
    
    if (history.length === 0) {
      return { success: false, message: "Tidak ada data untuk periode yang dipilih" };
    }
    
    // Buat dokumen HTML
    var htmlContent = "<html><head><style>";
    htmlContent += "body { font-family: Arial, sans-serif; margin: 20px; }";
    htmlContent += "h1, h2 { color: #4285f4; }";
    htmlContent += "table { border-collapse: collapse; width: 100%; margin: 20px 0; }";
    htmlContent += "th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }";
    htmlContent += "th { background-color: #f2f2f2; }";
    htmlContent += "tr:nth-child(even) { background-color: #f9f9f9; }";
    htmlContent += ".hadir { color: #34a853; }";
    htmlContent += ".terlambat { color: #fbbc05; }";
    htmlContent += ".izin { color: #4285f4; }";
    htmlContent += ".footer { margin-top: 30px; font-size: 0.8em; color: #666; }";
    htmlContent += "</style></head><body>";
    
    // Header
    htmlContent += "<h1>Laporan Absensi</h1>";
    htmlContent += "<h2>" + user.name + " - " + monthName + " " + year + "</h2>";
    
    // Info Pengguna
    htmlContent += "<p><strong>NIP:</strong> " + user.nip + "</p>";
    htmlContent += "<p><strong>Jabatan:</strong> " + (user.position || "-") + "</p>";
    htmlContent += "<p><strong>Unit:</strong> " + (user.unit || "-") + "</p>";
    
    // Tabel Absensi
    htmlContent += "<table>";
    htmlContent += "<tr><th>Tanggal</th><th>Jam Masuk</th><th>Jam Pulang</th><th>Status</th><th>Lokasi</th><th>Durasi</th><th>Keterangan</th></tr>";
    
    history.forEach(function(item) {
      var statusClass = "";
      if (item.status === "Hadir") statusClass = "hadir";
      else if (item.status === "Terlambat") statusClass = "terlambat";
      else statusClass = "izin";
      
      htmlContent += "<tr>";
      htmlContent += "<td>" + (item.dayName ? item.dayName + ", " : "") + item.date + "</td>";
      htmlContent += "<td>" + (item.checkIn || "-") + "</td>";
      htmlContent += "<td>" + (item.checkOut || "-") + "</td>";
      htmlContent += "<td class='" + statusClass + "'>" + item.status + "</td>";
      htmlContent += "<td>" + (item.location ? item.location.split(',')[0] : "-") + "</td>";
      htmlContent += "<td>" + (item.totalHours || "-") + "</td>";
      htmlContent += "<td>" + (item.notes || "-") + "</td>";
      htmlContent += "</tr>";
    });
    
    htmlContent += "</table>";
    
    // Ringkasan
    var summary = {
      totalDays: history.length,
      presentDays: 0,
      lateDays: 0,
      leaveDays: 0,
      totalWorkHours: "0:00"
    };
    
    var totalMinutes = 0;
    
    history.forEach(function(day) {
      if (day.status === "Hadir") {
        summary.presentDays++;
      } else if (day.status === "Terlambat") {
        summary.lateDays++;
        summary.presentDays++; // Terlambat juga dihitung hadir
      } else {
        summary.leaveDays++;
      }
      
      // Hitung total jam kerja
      if (day.totalHours) {
        var hoursMinutes = day.totalHours.split(':');
        if (hoursMinutes.length >= 2) {
          var hours = parseInt(hoursMinutes[0], 10) || 0;
          var minutes = parseInt(hoursMinutes[1], 10) || 0;
          totalMinutes += (hours * 60) + minutes;
        }
      }
    });
    
    // Format total jam kerja
    var totalHours = Math.floor(totalMinutes / 60);
    var remainingMinutes = totalMinutes % 60;
    summary.totalWorkHours = totalHours + ":" + (remainingMinutes < 10 ? "0" : "") + remainingMinutes;
    
    htmlContent += "<h2>Ringkasan</h2>";
    htmlContent += "<p><strong>Total Hari Kerja:</strong> " + summary.totalDays + "</p>";
    htmlContent += "<p><strong>Hadir:</strong> " + summary.presentDays + "</p>";
    htmlContent += "<p><strong>Terlambat:</strong> " + summary.lateDays + "</p>";
    htmlContent += "<p><strong>Izin/Sakit/Cuti:</strong> " + summary.leaveDays + "</p>";
    htmlContent += "<p><strong>Total Jam Kerja:</strong> " + summary.totalWorkHours + "</p>";
    
    // Footer
    htmlContent += "<div class='footer'>";
    htmlContent += "<p>Laporan ini dibuat otomatis oleh aplikasi GASPUL Hadir pada " + new Date().toLocaleString() + "</p>";
    htmlContent += "</div>";
    
    htmlContent += "</body></html>";
    
    // Konversi HTML ke PDF
    var fileName = "Laporan Absensi " + user.name + " - " + monthName + " " + year + ".pdf";
    
    // Buat blob dari HTML
    var blob = Utilities.newBlob(htmlContent, "text/html", fileName);
    
    // Buat file di Google Drive
    var folder = DriveApp.getRootFolder();
    var file = folder.createFile(blob);
    
    // Konversi ke PDF menggunakan Google Drive API
    var pdfFile = DriveApp.getFileById(file.getId()).getAs('application/pdf');
    var newFile = folder.createFile(pdfFile).setName(fileName);
    
    // Hapus file HTML sementara
    file.setTrashed(true);
    
    return { 
      success: true, 
      message: "PDF berhasil dibuat dan tersedia di Google Drive Anda",
      fileId: newFile.getId(),
      fileName: fileName
    };
  } catch (error) {
    Logger.log("Error pada fungsi exportAttendanceToPdf: " + error.message);
    return { success: false, message: "Terjadi kesalahan: " + error.message };
  }
}

/**
 * Fungsi untuk mengekspor data absensi ke Excel
 */
function exportAttendanceToExcel(month, year) {
  try {
    // Dapatkan user dari session
    var user = getUserSession();
    if (!user) {
      return { success: false, message: "Sesi telah berakhir" };
    }
    
    var nip = user.nip;
    
    // Konversi parameter ke angka
    month = parseInt(month);
    year = parseInt(year);
    
    // Nama bulan dan tahun
    var monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
                      "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    var monthName = monthNames[month - 1];
    
    // Dapatkan data absensi
    var history = getAttendanceHistory(nip, month, year);
    
    if (history.length === 0) {
      return { success: false, message: "Tidak ada data untuk periode yang dipilih" };
    }
    
    // Buat spreadsheet baru
    var ss = SpreadsheetApp.create("Laporan Absensi " + user.name + " - " + monthName + " " + year);
    var sheet = ss.getActiveSheet();
    
    // Tambahkan header
    sheet.appendRow([
      "Tanggal", 
      "Hari", 
      "Jam Masuk", 
      "Jam Pulang", 
      "Status", 
      "Lokasi", 
      "Durasi", 
      "Keterangan"
    ]);
    
    // Tambahkan data
    history.forEach(function(item) {
      sheet.appendRow([
        item.date,
        item.dayName || "",
        item.checkIn || "-",
        item.checkOut || "-",
        item.status,
        item.location ? item.location.split(',')[0] : "-",
        item.totalHours || "-",
        item.notes || "-"
      ]);
    });
    
    // Format header
    sheet.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#f2f2f2");
    
    // Auto-resize kolom
    sheet.autoResizeColumns(1, 8);
    
    // Tambahkan info user
    sheet.getRange(history.length + 3, 1).setValue("Informasi Pegawai:");
    sheet.getRange(history.length + 4, 1).setValue("Nama:");
    sheet.getRange(history.length + 4, 2).setValue(user.name);
    sheet.getRange(history.length + 5, 1).setValue("NIP:");
    sheet.getRange(history.length + 5, 2).setValue(nip);
    sheet.getRange(history.length + 6, 1).setValue("Jabatan:");
    sheet.getRange(history.length + 6, 2).setValue(user.position || "-");
    sheet.getRange(history.length + 7, 1).setValue("Unit:");
    sheet.getRange(history.length + 7, 2).setValue(user.unit || "-");
    
    // Hitung ringkasan
    var summary = {
      totalDays: history.length,
      presentDays: 0,
      lateDays: 0,
      leaveDays: 0,
      totalWorkHours: "0:00"
    };
    
    var totalMinutes = 0;
    
    history.forEach(function(day) {
      if (day.status === "Hadir") {
        summary.presentDays++;
      } else if (day.status === "Terlambat") {
        summary.lateDays++;
        summary.presentDays++; // Terlambat juga dihitung hadir
      } else {
        summary.leaveDays++;
      }
      
      // Hitung total jam kerja
      if (day.totalHours) {
        var hoursMinutes = day.totalHours.split(':');
        if (hoursMinutes.length >= 2) {
          var hours = parseInt(hoursMinutes[0], 10) || 0;
          var minutes = parseInt(hoursMinutes[1], 10) || 0;
          totalMinutes += (hours * 60) + minutes;
        }
      }
    });
    
    // Format total jam kerja
    var totalHours = Math.floor(totalMinutes / 60);
    var remainingMinutes = totalMinutes % 60;
    summary.totalWorkHours = totalHours + ":" + (remainingMinutes < 10 ? "0" : "") + remainingMinutes;
    
    // Tambahkan ringkasan
    sheet.getRange(history.length + 9, 1).setValue("Ringkasan:");
    sheet.getRange(history.length + 10, 1).setValue("Total Hari Kerja:");
    sheet.getRange(history.length + 10, 2).setValue(summary.totalDays);
    sheet.getRange(history.length + 11, 1).setValue("Hadir:");
    sheet.getRange(history.length + 11, 2).setValue(summary.presentDays);
    sheet.getRange(history.length + 12, 1).setValue("Terlambat:");
    sheet.getRange(history.length + 12, 2).setValue(summary.lateDays);
    sheet.getRange(history.length + 13, 1).setValue("Izin/Sakit/Cuti:");
    sheet.getRange(history.length + 13, 2).setValue(summary.leaveDays);
    sheet.getRange(history.length + 14, 1).setValue("Total Jam Kerja:");
    sheet.getRange(history.length + 14, 2).setValue(summary.totalWorkHours);
    
    // Format file
    sheet.getRange(history.length + 3, 1).setFontWeight("bold");
    sheet.getRange(history.length + 9, 1).setFontWeight("bold");
    
    // Simpan dan dapatkan URL
    var fileUrl = ss.getUrl();
    
    return { 
      success: true, 
      message: "File Excel berhasil dibuat dan tersedia di Google Drive Anda",
      fileUrl: fileUrl,
      fileName: ss.getName()
    };
  } catch (error) {
    Logger.log("Error pada fungsi exportAttendanceToExcel: " + error.message);
    return { success: false, message: "Terjadi kesalahan: " + error.message };
  }
}

/**
 * Mengirim laporan absensi melalui email
 */
function sendReport(month, year) {
  try {
    // Dapatkan user dari session
    var user = getUserSession();
    if (!user) {
      return { success: false, message: "Sesi telah berakhir" };
    }
    
    var nip = user.nip;
    var email = user.email;
    
    // Jika month dan year tidak didefinisikan, gunakan bulan dan tahun saat ini
    if (!month || !year) {
      var now = new Date();
      month = now.getMonth() + 1;
      year = now.getFullYear();
    }
    
    // Konversi parameter ke angka
    month = parseInt(month);
    year = parseInt(year);
    
    // Dapatkan history
    var history = getAttendanceHistory(nip, month, year);
    
    // Jika tidak ada data
    if (history.length === 0) {
      return { success: false, message: "Tidak ada data absensi untuk periode tersebut" };
    }
    
    // Format nama bulan
    var monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
                     "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    
    var subject = "Laporan Absensi " + user.name + " - " + monthNames[month-1] + " " + year;
    
    // Create email body (HTML)
    var body = "<h2>Laporan Absensi</h2>" +
               "<p><strong>Nama:</strong> " + user.name + "</p>" +
               "<p><strong>NIP:</strong> " + user.nip + "</p>" +
               "<p><strong>Jabatan:</strong> " + (user.position || "-") + "</p>" +
               "<p><strong>Unit:</strong> " + (user.unit || "-") + "</p>" +
               "<p><strong>Periode:</strong> " + monthNames[month-1] + " " + year + "</p><br>" +
               "<table border='1' cellpadding='5' style='border-collapse: collapse; width: 100%;'>" +
               "<tr style='background-color: #f2f2f2;'>" +
               "<th>Tanggal</th>" +
               "<th>Jam Masuk</th>" +
               "<th>Jam Pulang</th>" +
               "<th>Status</th>" +
               "<th>Lokasi</th>" +
               "<th>Durasi</th>" +
               "<th>Keterangan</th>" +
               "</tr>";
    
    for (var i = 0; i < history.length; i++) {
      var rowColor = i % 2 === 0 ? "#ffffff" : "#f9f9f9";
      var statusColor = history[i].status === "Hadir" ? "#34a853" : 
                  (history[i].status === "Terlambat" ? "#fbbc05" : "#4285f4");
      
      body += "<tr style='background-color: " + rowColor + ";'>" +
          "<td>" + (history[i].dayName ? history[i].dayName + ", " : "") + history[i].date + "</td>" +
          "<td>" + (history[i].checkIn || "-") + "</td>" +
          "<td>" + (history[i].checkOut || "-") + "</td>" +
          "<td style='color: " + statusColor + "; font-weight: bold;'>" + history[i].status + "</td>" +
          "<td>" + (history[i].location ? history[i].location.split(',')[0] : "-") + "</td>" +
          "<td>" + (history[i].totalHours || "-") + "</td>" +
          "<td>" + (history[i].notes || "-") + "</td>" +
          "</tr>";
    }
    
    // Add statistics
    var stats = getUserStats(nip);
    
    body += "</table><br>" +
            "<h3>Ringkasan Kehadiran</h3>" +
            "<p><strong>Hari Hadir:</strong> " + stats.presentDays + "</p>" +
            "<p><strong>Izin/Sakit:</strong> " + stats.leaveDays + "</p>" +
            "<p><strong>Persentase Kehadiran:</strong> " + stats.attendancePercentage + "%</p>" +
            "<p><em>Laporan ini dikirim otomatis dari aplikasi GASPUL Hadir</em></p>";
    
    // Send email
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: body
    });
    
    return { success: true, message: "Laporan berhasil dikirim ke " + email };
  } catch (error) {
    Logger.log("Error pada fungsi sendReport: " + error.message);
    return { success: false, message: "Terjadi kesalahan saat mengirim laporan: " + error.message };
  }
}

/**
 * Mendapatkan data profil pengguna lengkap
 */
function getUserProfile() {
  try {
    // Dapatkan data dari session
    var user = getUserSession();
    if (!user) {
      return null;
    }
    
    // Dapatkan data lengkap dari database
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    var data = sheet.getDataRange().getValues();
    
    var userProfile = null;
    
    // Cari user berdasarkan NIP
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == user.nip) {
        userProfile = {
          nip: data[i][0],
          name: data[i][1],
          email: data[i][2],
          position: data[i][4] || "",
          unit: data[i][5] || "",
          registerDate: data[i][6],
          avatar: data[i][8] || ""
        };
        break;
      }
    }
    
    if (!userProfile) {
      return null;
    }
    
    // Tambahkan statistik kehadiran
    userProfile.stats = getUserStats(user.nip);
    
    // Tambahan statistik jam kerja
    userProfile.workHours = getWorkHoursStats();
    
    return userProfile;
  } catch (error) {
    Logger.log("Error pada fungsi getUserProfile: " + error.message);
    return null;
  }
}

/**
 * Mengubah password pengguna
 */
function changePassword(currentPassword, newPassword) {
  try {
    // Dapatkan user dari session
    var user = getUserSession();
    if (!user) {
      return { success: false, message: "Sesi telah berakhir" };
    }
    
    var nip = user.nip;
    
    // Cari user di database
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == nip) {
        // Verifikasi password saat ini
        if (data[i][3] != currentPassword) {
          return { success: false, message: "Password saat ini tidak sesuai" };
        }
        
        // Validasi password baru
        if (newPassword.length < 6) {
          return { success: false, message: "Password baru minimal 6 karakter" };
        }
        
        // Update password baru
        sheet.getRange(i+1, 4).setValue(newPassword);
        
        return { success: true, message: "Password berhasil diubah" };
      }
    }
    
    return { success: false, message: "User tidak ditemukan" };
  } catch (error) {
    Logger.log("Error pada fungsi changePassword: " + error.message);
    return { success: false, message: "Terjadi kesalahan: " + error.message };
  }
}

/**
 * Ekspor laporan profil pengguna ke PDF
 * Termasuk statistik kehadiran dan jam kerja
 * 
 * @return {object} Status ekspor
 */
function exportProfileReport() {
  try {
    // Dapatkan user dari session
    var user = getUserSession();
    if (!user) {
      return { success: false, message: "Sesi telah berakhir" };
    }
    
    // Dapatkan statistik kehadiran
    var stats = getUserStats(user.nip);
    
    // Dapatkan statistik jam kerja
    var workHoursStats = getWorkHoursStats();
    
    // Dapatkan bulan dan tahun saat ini
    var now = new Date();
    var currentMonth = now.getMonth();
    var months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    var currentMonthName = months[currentMonth];
    var currentYear = now.getFullYear();
    
    // Buat dokumen HTML
    var htmlContent = "<html><head><style>";
    htmlContent += "body { font-family: Arial, sans-serif; margin: 20px; }";
    htmlContent += "h1, h2, h3 { color: #4285f4; }";
    htmlContent += "table { border-collapse: collapse; width: 100%; margin: 20px 0; }";
    htmlContent += "th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }";
    htmlContent += "th { background-color: #f2f2f2; }";
    htmlContent += ".profile-header { text-align: center; margin-bottom: 30px; }";
    htmlContent += ".stats-section { margin: 20px 0; padding: 15px; background-color: #f9f9f9; border-radius: 8px; }";
    htmlContent += ".stats-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-top: 15px; }";
    htmlContent += ".stat-item { background-color: white; padding: 15px; border-radius: 4px; text-align: center; }";
    htmlContent += ".stat-value { font-size: 24px; font-weight: bold; color: #4285f4; margin-bottom: 5px; }";
    htmlContent += ".stat-label { color: #666; }";
    htmlContent += ".footer { margin-top: 30px; font-size: 0.8em; color: #666; text-align: center; }";
    htmlContent += "</style></head><body>";
    
    // Header
    htmlContent += "<div class='profile-header'>";
    htmlContent += "<h1>Laporan Profil Pengguna</h1>";
    htmlContent += "<h2>" + user.name + "</h2>";
    htmlContent += "<p>Periode: " + currentMonthName + " " + currentYear + "</p>";
    htmlContent += "</div>";
    
    // Informasi Pengguna
    htmlContent += "<h2>Informasi Pribadi</h2>";
    htmlContent += "<table>";
    htmlContent += "<tr><th>NIP</th><td>" + user.nip + "</td></tr>";
    htmlContent += "<tr><th>Nama</th><td>" + user.name + "</td></tr>";
    htmlContent += "<tr><th>Email</th><td>" + user.email + "</td></tr>";
    htmlContent += "<tr><th>Jabatan</th><td>" + (user.position || "-") + "</td></tr>";
    htmlContent += "<tr><th>Unit</th><td>" + (user.unit || "-") + "</td></tr>";
    htmlContent += "</table>";
    
    // Statistik Kehadiran
    htmlContent += "<div class='stats-section'>";
    htmlContent += "<h2>Statistik Kehadiran</h2>";
    htmlContent += "<div class='stats-grid'>";
    htmlContent += "<div class='stat-item'><div class='stat-value'>" + stats.presentDays + "</div><div class='stat-label'>Hari Hadir</div></div>";
    htmlContent += "<div class='stat-item'><div class='stat-value'>" + stats.leaveDays + "</div><div class='stat-label'>Izin/Sakit</div></div>";
    htmlContent += "<div class='stat-item'><div class='stat-value'>" + stats.attendancePercentage + "%</div><div class='stat-label'>Kehadiran</div></div>";
    htmlContent += "</div>";
    htmlContent += "</div>";

    // Statistik Jam Kerja
    if (workHoursStats) {
      htmlContent += "<div class='stats-section'>";
      htmlContent += "<h2>Statistik Jam Kerja</h2>";
      htmlContent += "<div class='stats-grid'>";
      htmlContent += "<div class='stat-item'><div class='stat-value'>" + workHoursStats.averageHours + "</div><div class='stat-label'>Rata-rata Jam Kerja</div></div>";
      htmlContent += "<div class='stat-item'><div class='stat-value'>" + workHoursStats.totalHours + "</div><div class='stat-label'>Total Jam Kerja</div></div>";
      htmlContent += "<div class='stat-item'><div class='stat-value'>" + workHoursStats.longestDay + "</div><div class='stat-label'>Hari Terlama</div></div>";
      htmlContent += "</div>";
      
      // Rata-rata per hari
      htmlContent += "<h3>Rata-rata Jam Kerja per Hari</h3>";
      htmlContent += "<table>";
      htmlContent += "<tr><th>Minggu</th><th>Senin</th><th>Selasa</th><th>Rabu</th><th>Kamis</th><th>Jumat</th><th>Sabtu</th></tr>";
      htmlContent += "<tr>";
      for (var day = 0; day < 7; day++) {
        htmlContent += "<td>" + (workHoursStats.dailyAverage[day] || "0:00") + "</td>";
      }
      htmlContent += "</tr>";
      htmlContent += "</table>";
      htmlContent += "</div>";
    }
    
    // Footer
    htmlContent += "<div class='footer'>";
    htmlContent += "<p>Laporan ini dibuat otomatis oleh aplikasi GASPUL Hadir pada " + new Date().toLocaleString() + "</p>";
    htmlContent += "</div>";
    
    htmlContent += "</body></html>";
    
    // Konversi HTML ke PDF
    var fileName = "Profil " + user.name + " - " + currentMonthName + " " + currentYear + ".pdf";
    
    // Buat blob dari HTML
    var blob = Utilities.newBlob(htmlContent, "text/html", fileName);
    
    // Buat file di Google Drive
    var folder = DriveApp.getRootFolder();
    var file = folder.createFile(blob);
    
    // Konversi ke PDF menggunakan Google Drive API
    var pdfFile = DriveApp.getFileById(file.getId()).getAs('application/pdf');
    var newFile = folder.createFile(pdfFile).setName(fileName);
    
    // Hapus file HTML sementara
    file.setTrashed(true);
    
    return { 
      success: true, 
      message: "Laporan profil berhasil diekspor",
      fileId: newFile.getId(),
      fileName: fileName
    };
  } catch (error) {
    Logger.log("Error pada fungsi exportProfileReport: " + error.message);
    return { success: false, message: "Terjadi kesalahan: " + error.message };
  }
}

/**
 * Menyiapkan spreadsheet dengan data contoh
 * dan struktur yang diperlukan untuk aplikasi
 */
function setupSpreadsheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Setup Users sheet
    var usersSheet = ss.getSheetByName('Users');
    if (!usersSheet) {
      usersSheet = ss.insertSheet('Users');
      usersSheet.appendRow(['NIP', 'Nama', 'Email', 'Password', 'Jabatan', 'Unit Kerja', 'Tanggal Registrasi', 'QR Code ID', 'Avatar URL']);
      
      // Format header
      usersSheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
      
      // Set lebar kolom
      usersSheet.setColumnWidth(1, 150);  // NIP
      usersSheet.setColumnWidth(2, 200);  // Nama
      usersSheet.setColumnWidth(3, 200);  // Email
      usersSheet.setColumnWidth(4, 150);  // Password
      usersSheet.setColumnWidth(5, 150);  // Jabatan
      usersSheet.setColumnWidth(6, 200);  // Unit Kerja
      usersSheet.setColumnWidth(7, 150);  // Tanggal Registrasi
      usersSheet.setColumnWidth(8, 200);  // QR Code ID
      usersSheet.setColumnWidth(9, 250);  // Avatar URL
      
      // Tambahkan user contoh
      var now = new Date();
      usersSheet.appendRow([
        '198501262008011013', 
        'Aditya Rahman', 
        'aditya@example.com', 
        'password123', 
        'Staf Administrasi', 
        'Bidang Madrasah', 
        now, 
        Utilities.getUuid(), 
        ''
      ]);
    }
    
    // Setup Attendance sheet dengan kolom tambahan
    var attendanceSheet = ss.getSheetByName('Attendance');
    if (!attendanceSheet) {
      attendanceSheet = ss.insertSheet('Attendance');
      attendanceSheet.appendRow([
        'ID', 'NIP', 'Tanggal', 'Jam Masuk', 'Jam Pulang', 'Status', 'Keterangan', 
        'Foto Masuk', 'GPS', 'Lokasi', 'Foto Pulang', 'Hari Kerja', 'Total Jam Kerja'
      ]);
      
      // Format header
      attendanceSheet.getRange(1, 1, 1, 13).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
      
      // Set lebar kolom
      attendanceSheet.setColumnWidth(1, 50);   // ID
      attendanceSheet.setColumnWidth(2, 150);  // NIP
      attendanceSheet.setColumnWidth(3, 120);  // Tanggal
      attendanceSheet.setColumnWidth(4, 100);  // Jam Masuk
      attendanceSheet.setColumnWidth(5, 100);  // Jam Pulang
      attendanceSheet.setColumnWidth(6, 100);  // Status
      attendanceSheet.setColumnWidth(7, 200);  // Keterangan
      attendanceSheet.setColumnWidth(8, 300);  // Foto Masuk
      attendanceSheet.setColumnWidth(9, 200);  // GPS
      attendanceSheet.setColumnWidth(10, 250); // Lokasi
      attendanceSheet.setColumnWidth(11, 300); // Foto Pulang
      attendanceSheet.setColumnWidth(12, 100); // Hari Kerja
      attendanceSheet.setColumnWidth(13, 120); // Total Jam Kerja
      
      // Tambahkan contoh absensi
      var today = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
      var namaHari = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
      var dayName = namaHari[new Date().getDay()];
      
      attendanceSheet.appendRow([
        1, 
        '198501262008011013', 
        today, 
        '08:00:00', 
        '17:00:00', 
        'Hadir', 
        'Normal',
        '', // Foto Masuk (kosong untuk contoh)
        '-2.5983,122.5148', // GPS (koordinat contoh)
        'Makassar, Sulawesi Selatan, Indonesia', // Lokasi
        '', // Foto Pulang
        dayName, // Hari Kerja
        '9:00'  // Total Jam Kerja
      ]);
    } else {
      // Periksa apakah perlu menambah kolom baru
      var headers = attendanceSheet.getRange(1, 1, 1, attendanceSheet.getLastColumn()).getValues()[0];
      
      // Jika belum ada kolom yang diperlukan, tambahkan
      var requiredColumns = [
        { index: 7, name: 'Keterangan' },
        { index: 8, name: 'Foto Masuk' },
        { index: 9, name: 'GPS' },
        { index: 10, name: 'Lokasi' },
        { index: 11, name: 'Foto Pulang' },
        { index: 12, name: 'Hari Kerja' },
        { index: 13, name: 'Total Jam Kerja' }
      ];
      
      for (var i = 0; i < requiredColumns.length; i++) {
        var col = requiredColumns[i];
        if (headers.length < col.index || headers[col.index - 1] !== col.name) {
          // Jika kolom tidak ada atau nama tidak sesuai
          if (headers.length < col.index) {
            // Tambahkan kolom baru jika belum ada
            while (headers.length < col.index) {
              attendanceSheet.insertColumnAfter(headers.length);
              headers.push("");
            }
          }
          // Set nama kolom
          attendanceSheet.getRange(1, col.index).setValue(col.name);
        }
      }
    }
    
    // Setup PengaturanWaktu sheet
    buatPengaturanWaktuDefault();
    
    // Setup HariKhusus sheet jika belum ada
    cekHariKhusus("2025-01-01"); // Fungsi ini akan membuat sheet jika belum ada
    
    return { success: true, message: "Spreadsheet setup berhasil" };
  } catch (error) {
    Logger.log("Error pada fungsi setupSpreadsheet: " + error.message);
    return { success: false, message: "Terjadi kesalahan saat setup spreadsheet: " + error.message };
  }
}

/**
 * Fungsi untuk mendapatkan kunci API Maps
 */
function getMapsApiKey() {
  return "AIzaSyAN6pFsaw9Xl6VOS1btzfM4_MrieeFUJl0"; // Ganti dengan API key yang Anda dapatkan
}
