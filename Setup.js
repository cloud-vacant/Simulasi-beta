// setup.gs
// Fungsi untuk setup dan validasi struktur spreadsheet

const SETUP_CONFIG = {
  // Sheet yang harus ada di MASTER SPREADSHEET (Admin)
  MASTER_SHEETS: [
    {
      name: "Users",
      headers: ["Timestamp", "Username", "Email", "PasswordHash", "TelegramToken", "SpreadsheetID", "Status", "ChatID", "Role", "FullName", "LastLogin"]
    },
    {
      name: "Logs",
      headers: ["Timestamp", "UserEmail", "Action", "Details", "IP", "UserAgent"]
    },
    {
      name: "Activity",
      headers: ["Timestamp", "UserEmail", "ActivityType", "Description", "SpreadsheetID"]
    }
  ],
  
  // Sheet yang harus ada di USER SPREADSHEET (salin dari template)
  USER_SHEETS: [
    "Formulir", "LisensiMS", "Reset Device", "WhitelistApp", "Logs"
  ]
};

/**
 * Fungsi untuk menangani transisi akhir bulan
 * Diimplementasikan sebagai time-based trigger (setiap akhir bulan)
 */
function handleEndOfMonthTransition() {
  try {
    Logger.log('=== END OF MONTH TRANSITION START ===');
    
    // Dapatkan semua user yang aktif
    const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    const usersSheet = masterSS.getSheetByName('Users');
    
    if (!usersSheet) {
      Logger.log('Users sheet not found');
      return;
    }
    
    const data = usersSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Cari indeks kolom yang diperlukan
    const emailCol = findColumnIndex(headers, ['Email', 'email', 'EMAIL']);
    const statusCol = findColumnIndex(headers, ['Status', 'status', 'STATUS']);
    
    if (emailCol === -1) {
      Logger.log('Email column not found');
      return;
    }
    
    // Proses setiap user aktif
    for (let i = 1; i < data.length; i++) {
      const userEmail = data[i][emailCol];
      const userStatus = data[i][statusCol] || 'active';
      
      if (userEmail && userStatus === 'active') {
        // Kirim notifikasi untuk membuat PDF bulanan
        sendMonthlyPdfNotification(userEmail);
        
        // Log aktivitas
        logActivity(userEmail, 'END_OF_MONTH_NOTIFICATION', 'Monthly PDF reminder sent');
        
        Logger.log(`Monthly PDF notification sent to ${userEmail}`);
      }
    }
    
    Logger.log('=== END OF MONTH TRANSITION COMPLETE ===');
    
    // Bersihkan log lama (opsional)
    cleanupOldLogs(30); // Simpan log 90 hari terakhir
    
  } catch (error) {
    Logger.log('Error in handleEndOfMonthTransition: ' + error.toString());
  }
}

/**
 * Kirim notifikasi ke pengguna untuk membuat PDF bulanan
 * @param {string} userEmail
 */
function sendMonthlyPdfNotification(userEmail) {
  try {
    // Dapatkan data user
    const userData = getUserFromMasterSheet(userEmail);
    if (!userData) {
      Logger.log(`User data not found for ${userEmail}`);
      return;
    }
    
    const fullName = userData.fullName || userEmail.split('@')[0];
    
    // Kirim notifikasi via Telegram jika terhubung
    if (userData.chatId) {
      const message = `📊 *Pengingat Akhir Bulan*\n\nHalo ${fullName},\n\nBulan ini akan segera berakhir. Jangan lupa untuk membuat PDF laporan bulanan Anda sebelum data dibersihkan otomatis.\n\nGunakan fitur "Generate PDF" di website untuk membuat laporan bulanan ini.`;
      
      // Kirim pesan ke Telegram
      const telegramUrl = `https://api.telegram.org/bot${CONFIG.TELEGRAM_TOKEN}`;
      const payload = {
        chat_id: userData.chatId,
        text: message,
        parse_mode: 'HTML'
      };
      
      UrlFetchApp.fetch(`${telegramUrl}/sendMessage`, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload)
      });
      
      Logger.log(`Telegram notification sent to ${userEmail} via Chat ID: ${userData.chatId}`);
    }
    
    // Kirim notifikasi email (opsional)
    sendEmailNotification(userEmail, fullName);
    
  } catch (error) {
    Logger.log('Error in sendMonthlyPdfNotification: ' + error.toString());
  }
}

/**
 * Kirim notifikasi email ke pengguna
 * @param {string} userEmail
 * @param {string} fullName
 */
function sendEmailNotification(userEmail, fullName) {
  try {
    const subject = 'Pengingat Akhir Bulan - IT Helpdesk';
    const body = `
      Halo ${fullName},
      
      Bulan ini akan segera berakhir. Jangan lupa untuk membuat PDF laporan bulanan Anda sebelum data dibersihkan otomatis.
      
      Anda dapat membuat PDF dengan:
      1. Login ke website IT Helpdesk
      2. Pilih menu "Generate PDF"
      3. Pilih "Laporan Harian" atau "Laporan Reset Device"
      4. Klik "Cetak PDF"
      
      Terima kasih.
    `;
    
    MailApp.sendEmail(userEmail, subject, body);
    Logger.log(`Email notification sent to ${userEmail}`);
  } catch (error) {
    Logger.log('Error in sendEmailNotification: ' + error.toString());
  }
}

/**
 * Bersihkan data lama dari sheet pengguna
 * @param {number} daysToKeep - Jumlah hari data yang akan disimpan
 */
function cleanupOldData(daysToKeep = 60) {
  try {
    Logger.log(`=== CLEANUP OLD DATA START (Keeping ${daysToKeep} days) ===`);
    
    // Dapatkan semua user yang aktif
    const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    const usersSheet = masterSS.getSheetByName('Users');
    
    if (!usersSheet) {
      Logger.log('Users sheet not found');
      return;
    }
    
    const data = usersSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Cari indeks kolom yang diperlukan
    const emailCol = findColumnIndex(headers, ['Email', 'email', 'EMAIL']);
    const statusCol = findColumnIndex(headers, ['Status', 'status', 'STATUS']);
    const spreadsheetIdCol = findColumnIndex(headers, ['SpreadsheetID', 'SpreadsheetId', 'spreadsheetid']);
    
    if (emailCol === -1 || spreadsheetIdCol === -1) {
      Logger.log('Required columns not found');
      return;
    }
    
    // Hitung tanggal cutoff
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);
    cutoffDate.setHours(0, 0, 0, 0);
    
    // Proses setiap user
    let totalRowsCleaned = 0;
    for (let i = 1; i < data.length; i++) {
      const userEmail = data[i][emailCol];
      const userStatus = data[i][statusCol] || 'active';
      const spreadsheetId = data[i][spreadsheetIdCol];
      
      if (userEmail && userStatus === 'active' && spreadsheetId) {
        try {
          const userSS = SpreadsheetApp.openById(spreadsheetId);
          
          // Bersihkan sheet "Formulir"
          cleanUpSheet(userSS, 'Formulir', cutoffDate);
          
          // Bersihkan sheet "Reset Device"
          cleanUpSheet(userSS, 'ResetDevice', cutoffDate);
          
          totalRowsCleaned++;
          
          Logger.log(`Cleaned up data for user: ${userEmail}`);
        } catch (error) {
          Logger.log(`Error cleaning up data for ${userEmail}: ${error.toString()}`);
        }
      }
    }
    
    Logger.log(`=== CLEANUP OLD DATA COMPLETE (Cleaned ${totalRowsCleaned} users) ===`);
    
  } catch (error) {
    Logger.log('Error in cleanupOldData: ' + error.toString());
  }
}

/**
 * Bersihkan data lama dari sheet tertentu
 * @param {Spreadsheet} spreadsheet - Spreadsheet object
 * @param {string} sheetName - Nama sheet yang akan dibersihkan
 * @param {Date} cutoffDate - Tanggal cutoff
 */
function cleanUpSheet(spreadsheet, sheetName, cutoffDate) {
  try {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`Sheet ${sheetName} not found`);
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      Logger.log(`No data to clean in sheet ${sheetName}`);
      return;
    }
    
    // Hapus baris yang lebih lama dari cutoff date
    let rowsToDelete = [];
    for (let i = 1; i < data.length; i++) { // Mulai dari baris kedua (lewati header)
      const dateValue = data[i][1]; // Kolom kedua adalah tanggal
      
      if (dateValue) {
        const rowDate = new Date(dateValue);
        if (rowDate < cutoffDate) {
          rowsToDelete.push(i + 1); // +1 karena spreadsheet dimulai dari 1
        }
      }
    }
    
    // Hapus baris dari bawah ke atas untuk tidak mengubah indeks
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(rowsToDelete[i]);
    }
    
    Logger.log(`Cleaned ${rowsToDelete.length} old rows from sheet ${sheetName}`);
  } catch (error) {
    Logger.log(`Error cleaning up sheet ${sheetName}: ${error.toString()}`);
  }
}

/**
 * Fungsi untuk mengirim notifikasi PDF bulanan ke pengguna
 * @param {string} userEmail
 * @param {string} bulan - Format "YYYY-MM"
 * @param {string} tahun - Format "YYYY"
 */
function sendMonthlyPdfReminder(userEmail, bulan, tahun) {
  try {
    // Dapatkan data user
    const userData = getUserFromMasterSheet(userEmail);
    if (!userData || !userData.chatId) {
      Logger.log(`User ${userEmail} not connected to Telegram or not found`);
      return;
    }
    
    const fullName = userData.FullName || userEmail.split('@')[0];
    
    // Format tanggal untuk tampilan
    const monthNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    const monthIndex = parseInt(bulan) - 1;
    const monthName = monthNames[monthIndex];
    
    const message = `📊 *Pengingat PDF Bulanan*\n\nHalo ${fullName},\n\nIni adalah pengingat untuk membuat PDF laporan bulanan Anda:\n\n📅 *Bulan: ${monthName} ${tahun}*\n\nSilakan buat PDF laporan bulanan ini sebelum akhir bulan tiba.\n\nAnda dapat membuat PDF dengan:\n1. Login ke website IT Helpdesk\n2. Pilih menu "Generate PDF"\n3. Pilih "Laporan Harian" atau "Laporan Reset Device"\n4. Pilih periode tanggal dari ${bulan} ${tahun}\n5. Klik "Cetak PDF"\n\n⚠️ *Penting: Data lama akan dibersihkan otomatis setelah akhir bulan!*\n\nTerima kasih.`;
    
    // Kirim pesan ke Telegram
    const telegramUrl = `https://api.telegram.org/bot${CONFIG.TELEGRAM_TOKEN}`;
    const payload = {
      chat_id: userData.chatId,
      text: message,
      parse_mode: 'HTML'
    };
    
    UrlFetchApp.fetch(`${telegramUrl}/sendMessage`, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    });
    
    Logger.log(`Monthly PDF reminder sent to ${userEmail} for ${bulan} ${tahun}`);
    
    // Log aktivitas
    logActivity(userEmail, 'MONTHLY_PDF_REMINDER', `Reminder sent for ${bulan} ${tahun}`);
    
    return {
      success: true,
      message: 'Reminder sent successfully'
    };
    
  } catch (error) {
    Logger.log('Error in sendMonthlyPdfReminder: ' + error.toString());
    return {
      success: false,
      message: 'Error: ' + error.message
    };
  }
}

/**
 * Fungsi untuk mengirim notifikasi saat memasuki bulan baru
 * @param {string} userEmail
 */
function sendNewMonthNotification(userEmail) {
  try {
    // Dapatkan data user
    const userData = getUserFromMasterSheet(userEmail);
    if (!userData || !userData.chatId || userData.chatId.toString().trim() === "") {
      Logger.log(`User ${userEmail} not connected to Telegram or not found`);
      return;
    }
    
    const fullName = userData.fullName || userEmail.split('@')[0];
    
    // Dapatkan bulan dan tahun sebelumnya
    const now = new Date();
    const prevMonth = new Date(now.getFullYear(), now.getMonth(), 0); // Hari pertama bulan sekarang
    const monthNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    const prevMonthName = monthNames[prevMonth.getMonth()];
    
    const message = `📅 *Bulan Baru Dimulai*\n\nHalo ${fullName},\n\nBulan ${prevMonthName} telah dimulai. Jangan lupa untuk membuat PDF laporan bulan lama Anda jika belum sempat membuatnya.\n\nAnda dapat membuat PDF dengan:\n1. Login ke website IT Helpdesk\n2. Pilih menu "Generate PDF"\n3. Pilih "Laporan Harian" atau "Laporan Reset Device"\n4. Pilih periode tanggal dari ${prevMonthName}\n5. Klik "Cetak PDF"\n\n⚠️ *Data lama akan dibersihkan otomatis!*\n\nTerima kasih.`;
    
    // Kirim pesan ke Telegram
    const telegramUrl = `https://api.telegram.org/bot${CONFIG.TELEGRAM_TOKEN}`;
    const payload = {
      chat_id: userData.chatId,
      text: message,
      parse_mode: 'HTML'
    };
    
    UrlFetchApp.fetch(`${telegramUrl}/sendMessage`, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    });
    
    Logger.log(`New month notification sent to ${userEmail} for ${prevMonthName}`);
    
    // Log aktivitas
    logActivity(userEmail, 'NEW_MONTH_NOTIFICATION', `Reminder for ${prevMonthName}`);
    
    return {
      success: true,
      message: 'Notification sent successfully'
    };
    
  } catch (error) {
    Logger.log('Error in sendNewMonthNotification: ' + error.toString());
    return {
      success: false,
      message: 'Error: ' + error.message
    };
  }
}

/**
 * Fungsi scheduler untuk notifikasi PDF bulanan
 * Diimplementasikan sebagai time-based trigger (setiap hari)
 */
function scheduleMonthlyPdfReminders() {
  try {
    Logger.log('=== SCHEDULE MONTHLY PDF REMINDERS START ===');
    
    const now = new Date();
    const currentMonth = now.getMonth() + 1; // +1 karena getMonth() dimulai dari 0
    const currentYear = now.getFullYear();
    
    // Jika kita sudah di akhir bulan, tidak perlu mengirim reminder
    const lastDayOfMonth = new Date(currentYear, currentMonth, 0).getDate();
    const daysUntilEnd = lastDayOfMonth - now.getDate();
    
    if (daysUntilEnd <= 7 && daysUntilEnd > 0) {
      // Kirim reminder untuk bulan ini
      const monthNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
      const monthName = monthNames[currentMonth - 1];
      const formattedMonth = (currentMonth < 10 ? '0' : '') + currentMonth;
      
      // Dapatkan semua user yang aktif
      const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
      const usersSheet = masterSS.getSheetByName('Users');
      
      if (usersSheet) {
        const data = usersSheet.getDataRange().getValues();
        const headers = data[0];
        
        // Cari indeks kolom yang diperlukan
        const emailCol = findColumnIndex(headers, ['Email', 'email', 'EMAIL']);
        const statusCol = findColumnIndex(headers, ['Status', 'status', 'STATUS']);
        const telegramIdCol = findColumnIndex(headers, ['ChatID', 'Chat Id', 'chatid']);
        
        if (emailCol !== -1 && telegramIdCol !== -1) {
          // Proses setiap user aktif
          for (let i = 1; i < data.length; i++) {
            const userEmail = data[i][emailCol];
            const userStatus = data[i][statusCol] || 'active';
            const chatId = data[i][telegramIdCol];
            
            if (userEmail && userStatus === 'active' && chatId) {
              // Kirim reminder
              sendMonthlyPdfReminder(userEmail, formattedMonth, currentYear.toString());
            }
          }
        }
      }
    }
    
    Logger.log('=== SCHEDULE MONTHLY PDF REMINDERS COMPLETE ===');
    
  } catch (error) {
    Logger.log('Error in scheduleMonthlyPdfReminders: ' + error.toString());
  }
}

/**
 * Fungsi scheduler untuk notifikasi bulan baru
 * Diimplementasikan sebagai time-based trigger (setiap hari)
 */
function checkForNewMonth() {
  try {
    Logger.log('=== CHECK FOR NEW MONTH START ===');
    
    const now = new Date();
    const currentDay = now.getDate();
    
    // Jika hari pertama bulan, kirim notifikasi
    if (currentDay === 1) {
      // Dapatkan semua user yang aktif
      const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
      const usersSheet = masterSS.getSheetByName('Users');
      
      if (usersSheet) {
        const data = usersSheet.getDataRange().getValues();
        const headers = data[0];
        
        // Cari indeks kolom yang diperlukan
        const emailCol = findColumnIndex(headers, ['Email', 'email', 'EMAIL']);
        const statusCol = findColumnIndex(headers, ['Status', 'status', 'STATUS']);
        const telegramIdCol = findColumnIndex(headers, ['ChatID', 'Chat Id', 'chatid']);
        
        if (emailCol !== -1 && telegramIdCol !== -1) {
          // Proses setiap user aktif
          for (let i = 1; i < data.length; i++) {
            const userEmail = data[i][emailCol];
            const userStatus = data[i][statusCol] || 'active';
            const chatId = data[i][telegramIdCol];
            
            if (userEmail && userStatus === 'active' && chatId) {
              // Kirim notifikasi bulan baru
              sendNewMonthNotification(userEmail);
            }
          }
        }
      }
    }
    
    Logger.log('=== CHECK FOR NEW MONTH COMPLETE ===');
    
  } catch (error) {
    Logger.log('Error in checkForNewMonth: ' + error.toString());
  }
}