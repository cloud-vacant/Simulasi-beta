// ============================================
// UNIVERSAL LOGGER
// ============================================

/**
 * Fungsi logging universal yang bisa menulis ke spreadsheet dan sheet mana pun
 * @param {string} spreadsheetId - ID spreadsheet tujuan
 * @param {string} sheetName - Nama sheet tujuan
 * @param {string} userEmail - Email user yang melakukan aksi
 * @param {string} action - Jenis aksi (contoh: 'TELEGRAM_INPUT_CASE')
 * @param {string} details - Detail aksi
 * @param {string} source - Sumber aksi (contoh: 'Bot Telegram')
 */
function logToSheet(spreadsheetId, sheetName, userEmail, action, details, source) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    let sheet = ss.getSheetByName(sheetName);
    
    // Jika sheet tidak ada, buat sheet baru
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      
      // Buat header jika sheet baru
      const headers = ['Timestamp', 'UserEmail', 'Action', 'Details', 'Source'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight('bold')
        .setBackground('#4a86e8')
        .setFontColor('#ffffff');
    }
    
    // Siapkan data log
    const timestamp = new Date();
    const logData = [timestamp, userEmail, action, details, source];
    
    // Tambahkan baris baru ke sheet
    sheet.appendRow(logData);
    
    // Auto-resize kolom untuk keterbacaan
    sheet.autoResizeColumns(1, 5);
    
  } catch (error) {
    Logger.log(`Error logging to ${sheetName}: ${error.toString()}`);
  }
}

/**
 * Fungsi khusus untuk logging ke master spreadsheet
 * @param {string} userEmail - Email user
 * @param {string} action - Jenis aksi
 * @param {string} details - Detail aksi
 * @param {string} source - Sumber aksi
 */
function logToMasterSheet(userEmail, action, details, source) {
  logToSheet(CONFIG.MASTER_SPREADSHEET_ID, 'Logs', userEmail, action, details, source);
}

/**
 * Fungsi khusus untuk logging ke spreadsheet user
 * @param {string} userEmail - Email user
 * @param {string} action - Jenis aksi
 * @param {string} details - Detail aksi
 * @param {string} source - Sumber aksi
 */
function logToUserSheet(userEmail, action, details, source) {
  try {
    // Dapatkan spreadsheet ID user
    const userData = getUserFromMasterSheet(userEmail);
    if (!userData || !userData.spreadsheetId) {
      Logger.log(`Cannot log to user sheet for ${userEmail}: User data not found`);
      return;
    }
    
    logToSheet(userData.spreadsheetId, 'Logs', userEmail, action, details, source);
  } catch (error) {
    Logger.log(`Error logging to user sheet for ${userEmail}: ${error.toString()}`);
  }
}

/**
 * Fungsi untuk dijalankan oleh scheduler - mengisi master logs
 * Fungsi ini akan dipanggil oleh time-based trigger
 */
function populateMasterLogs() {
  try {
    Logger.log('Starting scheduled master logs population...');
    
    // Contoh: Anda bisa menambahkan log rutin di sini
    const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    const activitySheet = masterSS.getSheetByName('Activity');
    
    if (activitySheet) {
      const logData = [
        new Date(),
        'SYSTEM',
        'SCHEDULED_LOG_POPULATION',
        'Scheduler successfully populated master logs',
        'Google Apps Script Scheduler'
      ];
      
      activitySheet.appendRow(logData);
      Logger.log('Successfully logged to Activity sheet');
    }
    
    // Anda bisa menambahkan log rutin lainnya di sini
    // Misalnya: membersihkan cache, mengecek user tidak aktif, dll
    
    Logger.log('Scheduled master logs population completed');
  } catch (error) {
    Logger.log(`Error in populateMasterLogs: ${error.toString()}`);
  }
}

/**
 * Fungsi untuk membersihkan log lama (opsional)
 * Bisa dijalankan oleh scheduler untuk menjaga agar log tidak terlalu besar
 * @param {number} daysToKeep - Jumlah hari log yang akan disimpan
 */
function cleanupOldLogs(daysToKeep = 30) {
  try {
    const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    const logsSheet = masterSS.getSheetByName('Logs');
    
    if (!logsSheet) return;
    
    const data = logsSheet.getDataRange().getValues();
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);
    
    let rowsToDelete = [];
    for (let i = 1; i < data.length; i++) { // Mulai dari 1 untuk lewati header
      const logDate = new Date(data[i][0]);
      if (logDate < cutoffDate) {
        rowsToDelete.push(i + 1); // +1 karena spreadsheet dimulai dari 1
      }
    }
    
    // Hapus baris dari bawah ke atas untuk tidak mengubah indeks
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      logsSheet.deleteRow(rowsToDelete[i]);
    }
    
    Logger.log(`Cleaned up ${rowsToDelete.length} old log entries`);
  } catch (error) {
    Logger.log(`Error in cleanupOldLogs: ${error.toString()}`);
  }
}

function logActivity(userEmail, action, details) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
      sheet.appendRow(['Timestamp', 'UserEmail', 'Action', 'Details']);
    }
    const timestamp = new Date();
    sheet.appendRow([timestamp, userEmail, action, details]);
  } catch (e) {
    console.error('Gagal menulis log: ', e);
  }
}