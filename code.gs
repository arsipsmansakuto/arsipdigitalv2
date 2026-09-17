

// --- KONFIGURASI UTAMA ---
const ADMIN_DEFAULT = { 
  user: 'admin', 
  pass: '123', 
  nama: 'Administrator Sistem', 
  role: 'Admin',
  email: 'javaknight7@gmail.com' // Ubah dengan email admin asli
};

// Nama Folder di Google Drive tempat file disimpan
const DRIVE_FOLDER_NAME = "Arsip Digital Uploads";

/**
 * 1. HTTP GET HANDLER
 * Entry point aplikasi web.
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Arsip Digital - Manajemen Dokumen')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 2. INITIALIZATION
 * Membuat database (Sheet) jika belum ada.
 */
function initializeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // A. Setup Sheet Users
  let userSheet = ss.getSheetByName('Users');
  if (!userSheet) {
    userSheet = ss.insertSheet('Users');
    // Header: [0]Username, [1]Password, [2]Nama, [3]Role, [4]Status, [5]Token, [6]Email, [7]OTP
    userSheet.appendRow(['Username', 'Password', 'Nama Lengkap', 'Role', 'Status', 'Token', 'Email', 'OTP']);
    // Buat User Default
    userSheet.appendRow([
      ADMIN_DEFAULT.user, 
      ADMIN_DEFAULT.pass, 
      ADMIN_DEFAULT.nama, 
      ADMIN_DEFAULT.role, 
      'Active', 
      '',
      ADMIN_DEFAULT.email,
      ''
    ]);
    userSheet.setFrozenRows(1);
  }

  // B. Setup Sheet Arsip (UPDATED: Tambah Nomor & Perihal)
  let arsipSheet = ss.getSheetByName('Arsip');
  if (!arsipSheet) {
    arsipSheet = ss.insertSheet('Arsip');
    // Header Baru: 
    // [0]ID, [1]Nomor, [2]Nama, [3]Perihal, [4]Kategori, [5]Jenis, [6]Link, [7]Tanggal, [8]Uploader, [9]Shared
    arsipSheet.appendRow(['ID', 'Nomor Arsip', 'Nama Arsip', 'Perihal', 'Kategori', 'Jenis File', 'Link File', 'Tanggal Upload', 'Pengupload', 'SharedWith']);
    arsipSheet.setFrozenRows(1);
  }

  // C. Setup Sheet ActivityLog
  let logSheet = ss.getSheetByName('ActivityLog');
  if (!logSheet) {
    logSheet = ss.insertSheet('ActivityLog');
    logSheet.appendRow(['Waktu', 'User', 'Aktivitas', 'Detail']);
    logSheet.setFrozenRows(1);
  }

  // D. Setup Sheet Settings (Master Data)
  let setSheet = ss.getSheetByName('Settings');
  if (!setSheet) {
    setSheet = ss.insertSheet('Settings');
    setSheet.appendRow(['Type', 'Value']);
    const defaults = [
      ['Category', 'Surat Masuk'], 
      ['Category', 'Surat Keluar'], 
      ['Category', 'Laporan'], 
      ['Category', 'SK / Peraturan'], 
      ['Extension', 'PDF'], 
      ['Extension', 'DOCX'], 
      ['Extension', 'XLSX'], 
      ['Extension', 'JPG']
    ];
    defaults.forEach(row => setSheet.appendRow(row));
    setSheet.setFrozenRows(1);
  }

  // E. Setup Sheet Notifications
  let notifSheet = ss.getSheetByName('Notifications');
  if (!notifSheet) {
    notifSheet = ss.insertSheet('Notifications');
    // Header: [0]ID, [1]ToUser, [2]Message, [3]Type, [4]IsRead, [5]RelatedId, [6]Timestamp, [7]FromUser
    notifSheet.appendRow(['ID', 'ToUser', 'Message', 'Type', 'IsRead', 'RelatedId', 'Timestamp', 'FromUser']);
    notifSheet.setFrozenRows(1);
  }
  
  return "Database berhasil diinisialisasi dan siap digunakan.";
}

/**
 * 3. API ROUTER (Dispatcher)
 * Mengarahkan request dari Frontend ke fungsi Backend yang sesuai.
 */
function apiHandler(action, payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    switch (action) {
      // --- Authentication ---
      case 'login': return loginUser(ss, payload);
      case 'logout': return logoutUser(ss, payload);
      
      // --- Password Reset ---
      case 'requestOtp': return requestOtp(ss, payload);
      case 'resetPassword': return resetPassword(ss, payload);

      // --- User Profile ---
      case 'updateProfile': return updateUserProfile(ss, payload);
      
      // --- Dashboard Analytics ---
      case 'getStats': return getDashboardStats(ss, payload);
      
      // --- Core Data Operations ---
      case 'getData': return getData(ss, payload);
      case 'saveData': return saveData(ss, payload);
      case 'deleteData': return deleteData(ss, payload);
      
      // --- Settings & Master Data ---
      case 'getSettings': return getSettings(ss, payload);
      case 'saveSetting': return saveSetting(ss, payload);
      case 'deleteSetting': return deleteSetting(ss, payload);
      
      // --- Reporting ---
      case 'generateReport': return generateReport(ss, payload);
      
      // --- Notifications ---
      case 'getNotifications': return getNotifications(ss, payload);
      case 'markRead': return markRead(ss, payload);
      
      default: 
        throw new Error("Action tidak dikenal: " + action);
    }
  } catch (e) {
    console.error("API Error [" + action + "]: " + e.message);
    return { status: 'error', message: e.message };
  }
}

// ==========================================
// LOGIKA BISNIS (CORE FUNCTIONS)
// ==========================================

// --- 1. AUTENTIKASI & RESET PASSWORD ---

function loginUser(ss, { username, password }) {
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == username && data[i][1] == password) {
      if (data[i][4] !== 'Active') throw new Error("Akun Anda dinonaktifkan/suspend.");
      
      const token = Utilities.getUuid();
      sheet.getRange(i + 1, 6).setValue(token); // Simpan token sesi
      
      logActivity(ss, username, 'Login', 'User berhasil login');
      return {
        status: 'success',
        user: { 
          username: data[i][0], 
          nama_lengkap: data[i][2], 
          role: data[i][3], 
          token: token,
          email: data[i][6] || ''
        }
      };
    }
  }
  throw new Error("Username atau Password salah.");
}

function requestOtp(ss, { email }) {
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  let found = false;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] && data[i][6].toString().toLowerCase() === email.toLowerCase()) {
      const otp = Math.floor(100000 + Math.random() * 900000).toString();
      sheet.getRange(i + 1, 8).setValue(otp);
      
      try {
        MailApp.sendEmail({
          to: email,
          subject: "[Arsip Digital] Kode Reset Password",
          htmlBody: `
            <h3>Permintaan Reset Password</h3>
            <p>Halo ${data[i][2]},</p>
            <p>Kami menerima permintaan untuk mereset password akun Anda. Gunakan kode OTP berikut:</p>
            <h2 style="background: #eee; padding: 10px; display: inline-block; letter-spacing: 5px;">${otp}</h2>
            <p>Jika Anda tidak meminta ini, abaikan email ini.</p>
          `
        });
      } catch (e) {
        throw new Error("Gagal mengirim email (Kuota habis atau email tidak valid).");
      }
      
      found = true;
      break;
    }
  }

  if (!found) throw new Error("Email tidak terdaftar dalam sistem.");
  return { status: 'success', message: 'Kode OTP telah dikirim ke email Anda.' };
}

function resetPassword(ss, { email, otp, newPassword }) {
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] && data[i][6].toString().toLowerCase() === email.toLowerCase()) {
      if (String(data[i][7]) === String(otp)) {
        sheet.getRange(i + 1, 2).setValue(newPassword);
        sheet.getRange(i + 1, 8).setValue(""); // Hapus OTP
        logActivity(ss, data[i][0], 'Reset Password', 'Sukses reset password via OTP');
        return { status: 'success', message: 'Password berhasil diubah. Silakan login kembali.' };
      } else {
        throw new Error("Kode OTP salah atau kadaluarsa.");
      }
    }
  }
  throw new Error("User tidak ditemukan.");
}

function logoutUser(ss, { token }) {
  const user = validateToken(ss, token);
  if (user) {
    const sheet = ss.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++){
       if(data[i][0] === user.username) {
         sheet.getRange(i+1, 6).setValue("");
         break;
       }
    }
    logActivity(ss, user.username, 'Logout', 'User logout');
  }
  return { status: 'success' };
}

// --- 2. UPDATE PROFILE MANDIRI ---

function updateUserProfile(ss, { token, nama_lengkap, password_lama, password_baru }) {
  const user = validateToken(ss, token);
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === user.username) {
      if (String(data[i][1]) !== String(password_lama)) {
        throw new Error("Password lama yang Anda masukkan salah.");
      }

      sheet.getRange(i + 1, 3).setValue(nama_lengkap);
      
      if (password_baru && password_baru.trim() !== "") {
        sheet.getRange(i + 1, 2).setValue(password_baru);
      }

      logActivity(ss, user.username, 'Update Profil', 'User memperbarui profil/password mandiri');
      
      return {
        status: 'success',
        user: { 
          username: user.username, 
          nama_lengkap: nama_lengkap, 
          role: user.role, 
          token: token,
          email: data[i][6]
        }
      };
    }
  }
  throw new Error("Data user tidak ditemukan.");
}

// --- 3. DASHBOARD STATS (UPDATED INDEXES) ---

function getDashboardStats(ss, { token }) {
  const user = validateToken(ss, token);
  const arsipSheet = ss.getSheetByName('Arsip');
  const userSheet = ss.getSheetByName('Users');
  const logSheet = ss.getSheetByName('ActivityLog');
  
  const totalUsers = Math.max(0, userSheet.getLastRow() - 1);
  let myArsip = 0;
  let sharedWithMe = 0;
  let userVisibleTotal = 0;
  let categoryCounts = {};

  const arsipData = arsipSheet.getDataRange().getValues();
  for(let i=1; i<arsipData.length; i++){
    const row = arsipData[i];
    // Index bergeser: Uploader=8, Shared=9, Kategori=4
    const uploader = row[8];
    const sharedVal = row[9] ? row[9].toString() : '';
    const category = row[4] || 'Tanpa Kategori';
    
    const isOwner = uploader === user.username;
    const isShared = sharedVal.includes(user.username) || sharedVal === 'Public';
    
    if (isOwner) myArsip++;
    if (isShared && !isOwner) sharedWithMe++;
    
    if (user.role === 'Admin' || isOwner || isShared) {
        userVisibleTotal++;
        categoryCounts[category] = (categoryCounts[category] || 0) + 1;
    }
  }

  let totalArsipDisplay = user.role === 'Admin' ? Math.max(0, arsipSheet.getLastRow() - 1) : userVisibleTotal;

  // Filter Recent Logs
  let recentLogs = [];
  const logData = logSheet.getDataRange().getValues();
  for(let i = logData.length - 1; i > 0; i--) {
     if (recentLogs.length >= 5) break;
     
     const logRow = logData[i];
     const logUser = logRow[1];

     if (user.role === 'Admin' || logUser === user.username) {
         recentLogs.push({
           waktu: formatDate(logRow[0]),
           user: logRow[1],
           aksi: logRow[2],
           detail: logRow[3]
         });
     }
  }

  return {
    status: 'success',
    data: { 
        totalArsip: totalArsipDisplay, 
        totalUsers, 
        myArsip, 
        sharedWithMe, 
        recentLogs, 
        chartData: categoryCounts,
        role: user.role 
    }
  };
}

// --- 4. GET DATA (READ - UPDATED INDEXES) ---

function getData(ss, { token, type }) {
  const user = validateToken(ss, token);
  const sheet = ss.getSheetByName(type === 'users' ? 'Users' : (type === 'logs' ? 'ActivityLog' : 'Arsip'));
  if (!sheet) throw new Error("Sheet tidak ditemukan.");

  let rawData = sheet.getDataRange().getValues();
  rawData.shift(); // Hapus header
  let result = [];
  
  if (type === 'users') {
    if (user.role !== 'Admin') throw new Error("Akses Ditolak. Hanya Admin yang bisa melihat data user.");
    rawData.forEach(row => {
      result.push({ 
        username: row[0], 
        password: row[1], 
        nama_lengkap: row[2],
        role: row[3], 
        status: row[4],
        email: row[6] || '' 
      });
    });
  } 
  else if (type === 'archives') {
    rawData.forEach(row => {
      // Index Baru: [0]ID, [1]Nomor, [2]Nama, [3]Perihal, [4]Kategori, [5]Jenis, [6]Link, [7]Tanggal, [8]Uploader, [9]Shared
      let isOwner = row[8] === user.username;
      let sharedVal = row[9] ? row[9].toString() : '';
      let isShared = sharedVal.includes(user.username) || sharedVal === 'Public';
      
      if (user.role === 'Admin' || isOwner || isShared) {
        result.push({
          id: row[0], 
          nomor: row[1], // New
          nama: row[2], 
          perihal: row[3], // New
          kategori: row[4], 
          jenis: row[5], 
          link: row[6], 
          tanggal: formatDate(row[7]), 
          uploader: row[8], 
          shared: row[9]
        });
      }
    });
  } 
  else if (type === 'logs') {
    const logs = rawData.reverse().slice(0, 1000);
    logs.forEach(row => {
      if (user.role === 'Admin' || row[1] === user.username) {
        result.push({ 
          waktu: formatDate(row[0]), 
          user: row[1], 
          aksi: row[2], 
          detail: row[3],
          timestamp: new Date(row[0]).getTime()
        });
      }
    });
  } 
  else {
    throw new Error("Tipe data tidak dikenali.");
  }

  return { status: 'success', data: result };
}

// --- 5. SAVE DATA (CREATE / UPDATE - UPDATED INDEXES) ---

function saveData(ss, { token, type, data }) {
  const user = validateToken(ss, token);
  const sheet = ss.getSheetByName(type === 'users' ? 'Users' : 'Arsip');
  
  // A. SAVE ARSIP
  if (type === 'archives') {
    const id = data.id || Utilities.getUuid();
    const timestamp = new Date();
    
    // --- Upload Handler (Struktur Folder) ---
    let fileUrl = data.link; // Default pakai link lama
    let isNewFileUploaded = false; // Penanda apakah ada file baru

    if (data.fileObj && data.fileObj.base64) {
        try {
            // Mengirim kategori & nama_lengkap agar masuk folder yang sesuai
            fileUrl = uploadToDrive(data.fileObj, data.nama, data.kategori, user.nama_lengkap);
            isNewFileUploaded = true;
        } catch (e) {
            throw new Error("Gagal upload file: " + e.message);
        }
    }

    if (data.id) { // Edit Mode
       const allData = sheet.getDataRange().getValues();
       let found = false;
       for(let i=1; i<allData.length; i++){
         if(allData[i][0] == data.id){
           
           // --- VALIDASI HAK AKSES (PROTEKSI PRIVATE) ---
           const owner = allData[i][8];        // Kolom Pengupload (Index 8)
           const sharedStatus = allData[i][9]; // Kolom Shared (Index 9)
           const isOwner = owner === user.username;
           const isPublic = sharedStatus === 'Public';

           // Jika user yang mengedit BUKAN pemilik asli
           if (!isOwner) {
             if (user.role === 'Admin') {
               // Admin HANYA boleh edit jika statusnya Public
               if (!isPublic) {
                 throw new Error("Akses Ditolak: Admin tidak boleh mengedit dokumen Private milik user lain.");
               }
             } else {
               // User biasa tidak boleh edit punya orang lain
               throw new Error("Anda tidak memiliki izin mengedit file ini.");
             }
           }
           // ---------------------------------------------
           
           // --- LOGIKA GANTI FILE & HAPUS FILE LAMA ---
           let finalLink = allData[i][6]; // Ambil link lama
           
           if (isNewFileUploaded) {
             // 1. Hapus file fisik lama di Drive agar tidak menumpuk
             deleteFileFromDrive(allData[i][6]);
             // 2. Gunakan link file baru
             finalLink = fileUrl;
           }
           
           // Update Row
           // [0]ID, [1]Nomor, [2]Nama, [3]Perihal, [4]Kat, [5]Jenis, [6]Link, [7]Tgl, [8]Up, [9]Shared
           sheet.getRange(i+1, 1, 1, 10).setValues([[
             data.id, 
             data.nomor,      
             data.nama, 
             data.perihal,    
             data.kategori, 
             data.jenis, 
             finalLink,       // Link baru atau lama
             allData[i][7],   // Tanggal upload awal tetap
             allData[i][8],   // Uploader tetap
             data.shared || 'Public'
           ]]);
           found = true; 
           break;
         }
       }
       if(!found) throw new Error("Data tidak ditemukan.");
       logActivity(ss, user.username, 'Edit Arsip', `Mengubah arsip: ${data.nama}`);
    } else { // New Mode
      if (!fileUrl) throw new Error("File wajib diupload.");
      sheet.appendRow([
        id, 
        data.nomor,       
        data.nama, 
        data.perihal,     
        data.kategori, 
        data.jenis, 
        fileUrl, 
        timestamp, 
        user.username, 
        data.shared || 'Public'
      ]);
      logActivity(ss, user.username, 'Upload Arsip', `Menambah arsip: ${data.nama}`);
    }

    // --- TRIGGER NOTIFICATION ---
    if (data.shared && data.shared !== 'Public') {
        const recipients = data.shared.split(',').map(s => s.trim());
        recipients.forEach(targetUser => {
            if (targetUser !== user.username) {
                createNotification(ss, {
                    to: targetUser,
                    from: user.username,
                    msg: `membagikan dokumen "${data.nama}" kepada Anda.`,
                    type: 'share',
                    relId: id
                });
            }
        });
    }

  } 
  // B. SAVE USERS (ADMIN ONLY)
  else if (type === 'users' && user.role === 'Admin') {
    if(!data.isEdit) { // New User
       const users = sheet.getDataRange().getValues();
       if(users.some(u => u[0] === data.username)) throw new Error("Username sudah ada.");
       sheet.appendRow([
         data.username, 
         data.password, 
         data.nama_lengkap, 
         data.role, 
         data.status, 
         '', 
         data.email || '', 
         ''
       ]);
       logActivity(ss, user.username, 'Add User', `Menambah user: ${data.username}`);
    } else { // Edit User
       const allData = sheet.getDataRange().getValues();
       for(let i=1; i<allData.length; i++){
         if(allData[i][0] === data.username) { 
            // Update Pass, Nama, Role, Status
            sheet.getRange(i+1, 2, 1, 4).setValues([[data.password, data.nama_lengkap, data.role, data.status]]);
            // Update Email
            sheet.getRange(i+1, 7).setValue(data.email || '');
            break;
         }
       }
       logActivity(ss, user.username, 'Edit User', `Mengubah user: ${data.username}`);
    }
  } 
  else {
    throw new Error("Operasi penyimpanan tidak valid.");
  }

  return { status: 'success' };
}

// --- 6. DELETE DATA (UPDATED INDEXES) ---

function deleteData(ss, { token, type, id }) {
  const user = validateToken(ss, token);
  const sheet = ss.getSheetByName(type === 'users' ? 'Users' : 'Arsip');
  const data = sheet.getDataRange().getValues();
  let deleted = false;
  let deletedName = '';

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      // Cek Permission Arsip
      if(type === 'archives' && user.role !== 'Admin' && data[i][8] !== user.username) {
        throw new Error("Anda tidak berhak menghapus file ini.");
      }
      
      deletedName = type === 'archives' ? data[i][2] : data[i][0]; 
      
      // --- UPDATE BARU: Hapus file fisik di Drive jika tipe archives ---
      if (type === 'archives') {
        const fileUrl = data[i][6]; // Ambil Link File dari kolom index 6
        deleteFileFromDrive(fileUrl);
      }
      // ---------------------------------------------------------------

      sheet.deleteRow(i + 1);
      deleted = true;
      break;
    }
  }
  
  if(deleted) {
      logActivity(ss, user.username, 'Hapus Data', `Menghapus ${type}: ${deletedName}`);
      return { status: 'success' };
  } else {
      throw new Error("Data tidak ditemukan.");
  }
}

// --- 7. SETTINGS & USER DIRECTORY ---

function getSettings(ss, { token }) {
  const sheet = ss.getSheetByName('Settings');
  const userSheet = ss.getSheetByName('Users');
  
  let categories = [];
  let extensions = [];
  let nomors = []; // Array baru untuk menyimpan Nomor Arsip

  if(sheet) {
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++){
      if(data[i][0] === 'Category') categories.push(data[i][1]);
      if(data[i][0] === 'Extension') extensions.push(data[i][1]);
      // Membaca data jika Tipe-nya adalah 'Nomor'
      if(data[i][0] === 'Nomor') nomors.push(data[i][1]); 
    }
  }

  let userDirectory = [];
  if(userSheet) {
    const users = userSheet.getDataRange().getValues();
    for(let i=1; i<users.length; i++){
      if(users[i][4] === 'Active') {
         userDirectory.push({
           username: users[i][0],
           fullname: users[i][2] || users[i][0]
         });
      }
    }
  }
  
  // Sertakan 'nomors' dalam return data
  return { status: 'success', data: { categories, extensions, nomors, userDirectory } };
}

function saveSetting(ss, { token, type, value }) {
  const user = validateToken(ss, token);
  if(user.role !== 'Admin') throw new Error("Akses Ditolak.");
  
  const sheet = ss.getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++){
    if(data[i][0] === type && data[i][1].toLowerCase() === value.toLowerCase()){
      throw new Error("Data master sudah ada.");
    }
  }
  
  sheet.appendRow([type, value]);
  logActivity(ss, user.username, 'Update Setting', `Menambah ${type}: ${value}`);
  return { status: 'success' };
}

function deleteSetting(ss, { token, type, value }) {
  const user = validateToken(ss, token);
  if(user.role !== 'Admin') throw new Error("Akses Ditolak.");
  
  const sheet = ss.getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++){
    if(data[i][0] === type && data[i][1] === value){
      sheet.deleteRow(i+1);
      logActivity(ss, user.username, 'Update Setting', `Menghapus ${type}: ${value}`);
      return { status: 'success' };
    }
  }
  throw new Error("Data tidak ditemukan.");
}

// --- 8. REPORTING (UPDATED INDEXES & COLUMNS) ---

function generateReport(ss, { token, startDate, endDate }) {
  const user = validateToken(ss, token);
  const sheet = ss.getSheetByName('Arsip');
  const data = sheet.getDataRange().getValues();
  
  // Header Laporan
  const displayHeaders = ['ID Arsip', 'Nomor Arsip', 'Nama Dokumen', 'Perihal', 'Kategori', 'Ekstensi', 'Link File', 'Tanggal Upload', 'Pengupload', 'Status Share'];
  const start = new Date(startDate); start.setHours(0,0,0,0);
  const end = new Date(endDate); end.setHours(23,59,59,999);
  
  let filteredData = [];
  for(let i=1; i<data.length; i++){
    const row = data[i];
    // Index: [0]ID, [1]Nomor, [2]Nama, [3]Perihal, [4]Kat, [5]Jenis, [6]Link, [7]Tgl, [8]Up, [9]Share
    const uploader = row[8];
    const sharedVal = row[9] ? row[9].toString() : '';
    const isOwner = uploader === user.username;
    const isShared = sharedVal.includes(user.username) || sharedVal === 'Public';
    
    const hasAccess = (user.role === 'Admin' || isOwner || isShared);
    const rowDate = new Date(row[7]); // Tanggal di index 7
    
    if (hasAccess && rowDate >= start && rowDate <= end) {
      filteredData.push([
        row[0], // ID
        row[1], // Nomor
        row[2], // Nama
        row[3], // Perihal
        row[4], // Kategori
        row[5], // Jenis
        row[6], // Link
        Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'), 
        row[8], // Uploader
        row[9] === 'Public' ? 'Publik' : (row[9] ? 'Terbatas' : 'Pribadi') 
      ]);
    }
  }

  if (filteredData.length === 0) throw new Error("Tidak ada data pada rentang tanggal tersebut.");
  
  const tempSS = SpreadsheetApp.create("Temp_Report");
  const tempSheet = tempSS.getSheets()[0];
  
  const titleRange = tempSheet.getRange("A1:J1");
  titleRange.merge().setValue("LAPORAN ARSIP DIGITAL INSTANSI").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(14);
    
  const periodRange = tempSheet.getRange("A2:J2");
  periodRange.merge().setValue(`Periode: ${Utilities.formatDate(start, Session.getScriptTimeZone(), 'dd MMM yyyy')} - ${Utilities.formatDate(end, Session.getScriptTimeZone(), 'dd MMM yyyy')}`).setHorizontalAlignment("center").setFontStyle("italic");

  const headerRowIdx = 4;
  tempSheet.getRange(headerRowIdx, 1, 1, displayHeaders.length).setValues([displayHeaders]).setBackground("#2c3e50").setFontColor("#ffffff").setFontWeight("bold");
  
  if (filteredData.length > 0) {
    const dataRange = tempSheet.getRange(headerRowIdx + 1, 1, filteredData.length, displayHeaders.length);
    dataRange.setValues(filteredData).setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  }
  
  tempSheet.autoResizeColumns(1, displayHeaders.length);
  SpreadsheetApp.flush();
  
  const url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + tempSS.getId() + "&exportFormat=xlsx";
  const params = { method: "get", headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
  const blob = UrlFetchApp.fetch(url, params).getBlob();
  blob.setName(`Laporan_Arsip.xlsx`);
  const base64 = Utilities.base64Encode(blob.getBytes());
  
  DriveApp.getFileById(tempSS.getId()).setTrashed(true);
  
  return { status: 'success', file: base64, filename: blob.getName() };
}

// --- 9. NOTIFICATION SYSTEM ---

function createNotification(ss, { to, from, msg, type, relId }) {
  let sheet = ss.getSheetByName('Notifications');
  if (!sheet) {
    sheet = ss.insertSheet('Notifications');
    sheet.appendRow(['ID', 'ToUser', 'Message', 'Type', 'IsRead', 'RelatedId', 'Timestamp', 'FromUser']);
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    Utilities.getUuid(), 
    to, 
    msg, 
    type, 
    'FALSE', 
    relId, 
    new Date(),
    from
  ]);
}

function getNotifications(ss, { token }) {
  const user = validateToken(ss, token);
  const sheet = ss.getSheetByName('Notifications');
  if(!sheet) return { status: 'success', data: [] };

  const data = sheet.getDataRange().getValues();
  let notifs = [];
  let count = 0;
  // Ambil 20 notifikasi terbaru
  for (let i = data.length - 1; i > 0; i--) {
    if (count >= 20) break;
    if (data[i][1] === user.username) {
      notifs.push({
        id: data[i][0],
        message: data[i][2],
        type: data[i][3],
        isRead: data[i][4] === true || data[i][4] === 'TRUE',
        relatedId: data[i][5],
        timestamp: formatDate(data[i][6]),
        from: data[i][7]
      });
      count++;
    }
  }
  
  return { status: 'success', data: notifs };
}

function markRead(ss, { token, notifId }) {
  const user = validateToken(ss, token);
  const sheet = ss.getSheetByName('Notifications');
  const data = sheet.getDataRange().getValues();

  if (notifId === 'all') {
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === user.username && data[i][4] !== 'TRUE') {
        sheet.getRange(i + 1, 5).setValue('TRUE');
      }
    }
  } else {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == notifId && data[i][1] === user.username) {
        sheet.getRange(i + 1, 5).setValue('TRUE');
        break;
      }
    }
  }
  return { status: 'success' };
}

// ==========================================
// HELPER FUNCTIONS
// ==========================================

function validateToken(ss, token) {
  if (!token) throw new Error("Sesi tidak valid.");
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === token) {
        if(data[i][4] === 'Active') {
            return { 
                username: data[i][0], 
                nama_lengkap: data[i][2], 
                role: data[i][3], 
                token: token,
                email: data[i][6] || ''
            };
        } else {
            throw new Error("Akun dinonaktifkan.");
        }
    }
  }
  throw new Error("Sesi kadaluarsa.");
}

// --- UPDATED HELPER FUNCTIONS (FOLDER STRUCTURE) ---

// Fungsi Helper untuk Cek/Buat Folder
function getOrCreateSubFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(folderName);
  }
}

function uploadToDrive(fileObj, fileName, category, uploaderFullName) {
  // 1. Folder Induk (ARSIP DIGITAL)
  const rootIterator = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  let rootFolder;
  if (rootIterator.hasNext()) {
    rootFolder = rootIterator.next();
  } else {
    rootFolder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
  }

  const now = new Date();

  // 2. Folder Tahun (Contoh: 2024)
  const yearStr = now.getFullYear().toString();
  const yearFolder = getOrCreateSubFolder(rootFolder, yearStr);

  // 3. Folder Bulan (Contoh: 02 - Februari)
  const months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  const monthIndex = now.getMonth(); 
  // Menambah angka 0 di depan jika bulan < 10 agar urutan folder rapi
  const monthPrefix = (monthIndex + 1).toString().padStart(2, '0'); 
  const monthName = `${monthPrefix} - ${months[monthIndex]}`;
  const monthFolder = getOrCreateSubFolder(yearFolder, monthName);

  // 4. Folder Kategori (Contoh: Surat Masuk)
  // Jika kategori kosong/undefined, masukkan ke folder "Umum"
  const catName = category ? category.trim() : "Umum"; 
  const categoryFolder = getOrCreateSubFolder(monthFolder, catName);

  // 5. Folder Nama Pengupload (Contoh: Budi Santoso)
  const uploaderName = uploaderFullName ? uploaderFullName.trim() : "Anonim";
  const finalFolder = getOrCreateSubFolder(categoryFolder, uploaderName);

  // 6. Proses Pembuatan File
  const decoded = Utilities.base64Decode(fileObj.base64);
  const blob = Utilities.newBlob(decoded, fileObj.mimeType, fileName);
  const file = finalFolder.createFile(blob);
  
  // Set permission agar bisa diakses via link (VIEW Only)
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return file.getUrl();
}

function logActivity(ss, user, action, detail) {
  const sheet = ss.getSheetByName('ActivityLog');
  sheet.appendRow([new Date(), user, action, detail]);
}

function formatDate(date) {
  if (!date) return '-';
  try {
      return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  } catch (e) {
      return date.toString();
  }
}

// --- HELPER: HAPUS FILE FISIK DI DRIVE ---
function deleteFileFromDrive(fileUrl) {
  if (!fileUrl) return;
  
  try {
    // Ekstrak ID dari URL Google Drive
    // Pola URL: https://drive.google.com/file/d/FILE_ID/view...
    const idMatch = fileUrl.match(/[-\w]{25,}/);
    if (idMatch) {
      const fileId = idMatch[0];
      const file = DriveApp.getFileById(fileId);
      file.setTrashed(true); // Pindahkan ke Sampah (Trash)
    }
  } catch (e) {
    // Abaikan error jika file sudah tidak ada atau URL salah
    console.warn("Gagal menghapus file fisik: " + e.message);
  }
}

function requestPermissions() {
  // 1. Memancing izin https://www.googleapis.com/auth/spreadsheets (Create & Read/Write All)
  // Ini diperlukan untuk SpreadsheetApp.create()
  const tempSheet = SpreadsheetApp.create("Dummy_Permission_Trigger");
  
  // 2. Memancing izin Drive (sudah ada, tapi dipastikan ulang untuk penghapusan file)
  DriveApp.getRootFolder();
  
  // 3. Memancing izin UrlFetch (untuk download blob export)
  UrlFetchApp.fetch("https://www.google.com");

  // Bersihkan file dummy yang baru dibuat agar tidak nyampah di Drive
  const tempId = tempSheet.getId();
  DriveApp.getFileById(tempId).setTrashed(true);
  
  console.log("Izin berhasil diperbarui! Anda bisa menghapus fungsi ini sekarang.");
}