// ============================================
// DATABASE FUNCTIONS (User Spreadsheet Operations)
// ============================================

function getLaporanData(email, page = 1, pageSize = 10) {
  console.log('=== getLaporanData dipanggil ===');
  console.log('Email:', email, 'Page:', page, 'PageSize:', pageSize);
  
  try {
    if (!email || !email.includes('@')) {
      console.error('Email tidak valid');
      return { success: false, data: [], totalRows: 0, message: 'Email tidak valid' };
    }
    
    const userData = getUserFromMasterSheet(email);
    if (!userData || !userData.spreadsheetId) {
      console.error('User atau spreadsheet ID tidak ditemukan');
      return { success: false, data: [], totalRows: 0, message: 'User data not found' };
    }
    
    const spreadsheetId = userData.spreadsheetId;
    const userSS = SpreadsheetApp.openById(spreadsheetId);
    const sheet = userSS.getSheetByName('Formulir');
    
    if (!sheet) {
      console.error('Sheet "Formulir" tidak ditemukan');
      return { success: false, data: [], totalRows: 0, message: 'Sheet Formulir tidak ditemukan' };
    }
    
    const lastRow = sheet.getLastRow();
    const totalRows = (lastRow > 1) ? lastRow - 1 : 0; // Kurangi header
    
    if (totalRows === 0) {
      return { success: true, data: [], totalRows: 0 };
    }

    // Hitung indeks awal dan akhir untuk data yang diambil
    const startIndex = (page - 1) * pageSize + 2; // +2 karena header di baris 1 dan data dimulai baris 2
    const numRowsToGet = Math.min(pageSize, totalRows - (page - 1) * pageSize);

    // Ambil data range yang relevan
    const dataRange = sheet.getRange(startIndex, 1, numRowsToGet, 8); // Ambil 8 kolom
    const data = dataRange.getValues();
    
    console.log(`Total Rows: ${totalRows}, Fetching from row ${startIndex} for ${numRowsToGet} rows.`);
    
    // Proses data dengan format khusus untuk merge D-E
    const formattedData = data.map((row, rowIndex) => {
      const formattedRow = [
        row[0] || '',                            // A: No
        formatDateForDisplay(row[1]),            // B: Date
        row[2] || '',                            // C: Project
        row[3] || row[4] || '',                  // D: New Task (ambil dari D atau E jika D kosong)
        formatProgressValue(row[5]),             // F: Progress
        formatDateForDisplay(row[6]),            // G: Target
        row[7] || ''                             // H: Note
      ];
      return formattedRow;
    });
    
    return {
      success: true,
      data: formattedData,
      totalRows: totalRows
    };
    
  } catch (error) {
    console.error('=== ERROR di getLaporanData ===');
    console.error('Error:', error.toString());
    return {
      success: false,
      data: [],
      totalRows: 0,
      message: 'Error: ' + error.message
    };
  }
}

// Helper function untuk format tanggal
function formatDateForDisplay(dateValue) {
  if (!dateValue) return '';
  
  try {
    // Jika sudah string, return as-is
    if (typeof dateValue === 'string') {
      return dateValue;
    }
    
    // Jika Date object, format ke dd/MM/yyyy
    if (dateValue instanceof Date) {
      return Utilities.formatDate(dateValue, 'GMT+7', 'dd/MM/yyyy');
    }
    
    return String(dateValue);
  } catch (error) {
    console.error('Error format date:', error);
    return String(dateValue);
  }
}

// Tambahkan fungsi baru untuk format progress
function formatProgressValue(progress) {
  if (progress === null || progress === undefined || progress === '') {
    return '';
  }
  
  // Konversi ke string
  const progressStr = String(progress);
  
  // Jika sudah mengandung %, biarkan
  if (progressStr.includes('%')) {
    return progressStr;
  }
  
  // Jika angka, tambahkan %
  if (!isNaN(parseFloat(progressStr))) {
    return progressStr + '%';
  }
  
  return progressStr;
}

/**
 * Write data to user's Formulir sheet
 * @param {string} userEmail
 * @param {array} data - array of values [no, date, project, task, progress, target, note]
 * @return {object} {success: boolean, message: string}
 */
function writeLaporanData(userEmail, data) {
  try {
    const userSS = openUserSpreadsheet(userEmail);
    const sheet = userSS.getSheetByName('Formulir');
    if (!sheet) throw new Error('Sheet Formulir tidak ditemukan');

    sheet.appendRow(data);
    logActivity(userEmail, 'WRITE_LAPORAN', `Data ditambahkan: ${data[0]}`);
    return {success: true, message: 'Data berhasil ditambahkan'};
  } catch (error) {
    logActivity(userEmail, 'WRITE_LAPORAN_ERROR', error.toString());
    return {success: false, message: error.message};
  }
}

/**
 * Mengambil data dari sheet 'Reset Device' untuk pengguna tertentu.
 * @param {string} email Email pengguna.
 * @returns {Array<Array>} Data dari sheet Reset Device, atau array kosong jika tidak ada.
 */
function getResetDeviceData(email, page = 1, pageSize = 10) {
  console.log('=== getResetDeviceData dipanggil ===');
  console.log('Email:', email, 'Page:', page, 'PageSize:', pageSize);
  
  try {
    if (!email || !email.includes('@')) {
      console.error('Email tidak valid');
      return { success: false, data: [], totalRows: 0, message: 'Email tidak valid' };
    }
    
    const userData = getUserFromMasterSheet(email);
    if (!userData || !userData.spreadsheetId) {
      console.error('User atau spreadsheet ID tidak ditemukan');
      return { success: false, data: [], totalRows: 0, message: 'User data not found' };
    }
    
    const spreadsheetId = userData.spreadsheetId;
    const userSS = SpreadsheetApp.openById(spreadsheetId);
    const sheet = userSS.getSheetByName('ResetDevice');
    
    if (!sheet) {
      console.error('Sheet "ResetDevice" tidak ditemukan');
      return { success: false, data: [], totalRows: 0, message: 'Sheet ResetDevice tidak ditemukan' };
    }
    
    const lastRow = sheet.getLastRow();
    const totalRows = (lastRow > 1) ? lastRow - 1 : 0; // Kurangi header
    
    if (totalRows === 0) {
      return { success: true, data: [], totalRows: 0 };
    }
    
    // Hitung indeks awal dan akhir
    const startIndex = (page - 1) * pageSize + 2;
    const numRowsToGet = Math.min(pageSize, totalRows - (page - 1) * pageSize);

    // Ambil data range yang relevan
    const dataRange = sheet.getRange(startIndex, 1, numRowsToGet, 7); // Ambil 7 kolom
    const data = dataRange.getValues();
    
    console.log(`Total Rows: ${totalRows}, Fetching from row ${startIndex} for ${numRowsToGet} rows.`);
    
    // Proses data
    const formattedData = data.map((row, rowIndex) => {
      const formattedRow = [
        row[0] || '',                            // A: No
        row[1] || '',                            // B: Name
        formatDateForDisplay(row[2]),            // C: Date
        row[3] || '',                            // D: Divisi
        row[4] || '',                            // E: Dev Model
        row[5] || '',                            // F: Serial Number
        row[6] || ''                             // G: Action By
      ];
      return formattedRow;
    });
    
    return {
      success: true,
      data: formattedData,
      totalRows: totalRows
    };
    
  } catch (error) {
    console.error('=== ERROR di getResetDeviceData ===');
    console.error('Error:', error.toString());
    return {
      success: false,
      data: [],
      totalRows: 0,
      message: 'Error: ' + error.message
    };
  }
}

// Di file DatabaseFunction.txt, ganti fungsi getAnalyticsData dengan yang berikut:

function getAnalyticsData(email) {
  try {
    const userData = getUserFromMasterSheet(email);
    if (!userData || !userData.spreadsheetId) {
      return { success: false, message: 'User data not found' };
    }
    
    const userSS = SpreadsheetApp.openById(userData.spreadsheetId);
    const sheet = userSS.getSheetByName('Formulir');
    if (!sheet) {
      return { success: false, message: 'Formulir sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Gunakan findColumnIndex untuk pencarian kolom yang lebih fleksibel
    const dateCol = findColumnIndex(headers, ['Date', 'Tanggal', 'Tanggal Laporan']);
    const progressCol = findColumnIndex(headers, ['Progress', 'Progress Status (%)', 'Kemajuan', '% Progress']);
    const projectCol = findColumnIndex(headers, ['Project', 'Project Name', 'Proyek', 'Nama Proyek']);
    
    // Logging untuk debugging
    Logger.log(`Analytics for ${email}: Headers found -> ${headers.join(', ')}`);
    Logger.log(`Analytics for ${email}: Column indices -> Date=${dateCol}, Progress=${progressCol}, Project=${projectCol}`);
    
    if (dateCol === -1 || progressCol === -1 || projectCol === -1) {
      return { 
        success: false, 
        message: 'Kolom yang diperlukan (Tanggal, Progress/Kemajuan, Proyek) tidak ditemukan di sheet Formulir. Periksa kembali nama header kolom Anda.' 
      };
    }
    
    // Proses data (sisa kode fungsi tidak berubah)
    const monthlyData = {};
    const projectData = {};
    let totalReports = 0;
    let completedTasks = 0;
    let pendingTasks = 0;
    let todayActivity = 0;
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const date = row[dateCol];
      const progress = row[progressCol];
      const project = row[projectCol];
      
      if (!date) continue;
      
      totalReports++;
      
      // Check if today's activity
      if (date instanceof Date && date >= today) {
        todayActivity++;
      }
      
      // Count completed/pending tasks
      if (progress && progress.toString().includes('100')) {
        completedTasks++;
      } else {
        pendingTasks++;
      }
      
      // Count by month
      if (date instanceof Date) {
        const monthYear = `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}`;
        monthlyData[monthYear] = (monthlyData[monthYear] || 0) + 1;
      }
      
      // Count by project
      if (project) {
        projectData[project] = (projectData[project] || 0) + 1;
      }
    }
    
    // Calculate completion rate
    const completionRate = totalReports > 0 ? Math.round((completedTasks / totalReports) * 100) : 0;
    const pendingRate = totalReports > 0 ? Math.round((pendingTasks / totalReports) * 100) : 0;
    
    // Calculate growth (compare with last month)
    const currentMonth = `${today.getFullYear()}-${(today.getMonth() + 1).toString().padStart(2, '0')}`;
    const lastMonth = `${today.getFullYear()}-${today.getMonth().toString().padStart(2, '0')}`;
    const currentMonthCount = monthlyData[currentMonth] || 0;
    const lastMonthCount = monthlyData[lastMonth] || 0;
    const growth = lastMonthCount > 0 ? Math.round(((currentMonthCount - lastMonthCount) / lastMonthCount) * 100) : 0;
    
    // Determine activity status
    let activityStatus = 'Normal';
    if (todayActivity > 5) activityStatus = 'Sangat Aktif';
    else if (todayActivity > 2) activityStatus = 'Aktif';
    else if (todayActivity > 0) activityStatus = 'Sedikit Aktif';
    else activityStatus = 'Tidak Aktif';
    
    return {
      success: true,
      data: {
        totalReports,
        completedTasks,
        pendingTasks,
        todayActivity,
        completionRate,
        pendingRate,
        growth,
        activityStatus,
        monthlyData,
        projectData
      }
    };
  } catch (error) {
    console.error('Error in getAnalyticsData:', error);
    return { success: false, message: error.toString() };
  }
}

// Di backend, tambahkan fungsi untuk task management
function getTasks(email) {
  try {
    const userData = getUserFromMasterSheet(email);
    if (!userData || !userData.spreadsheetId) {
      return { success: false, message: 'User data not found' };
    }
    
    const userSS = SpreadsheetApp.openById(userData.spreadsheetId);
    let taskSheet = userSS.getSheetByName('Tasks');
    
    // Create Tasks sheet if it doesn't exist
    if (!taskSheet) {
      taskSheet = userSS.insertSheet('Tasks');
      taskSheet.appendRow(['ID', 'Task Name', 'Project', 'Due Date', 'Priority', 'Status', 'Description', 'Created At', 'Updated At']);
      taskSheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#4361ee').setFontColor('white');
    }
    
    const data = taskSheet.getDataRange().getValues();
    const tasks = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      tasks.push({
        id: row[0],
        name: row[1],
        project: row[2],
        dueDate: row[3] instanceof Date ? Utilities.formatDate(row[3], 'GMT+7', 'yyyy-MM-dd') : row[3],
        priority: row[4],
        status: row[5],
        description: row[6],
        createdAt: row[7] instanceof Date ? Utilities.formatDate(row[7], 'GMT+7', 'yyyy-MM-dd HH:mm:ss') : row[7],
        updatedAt: row[8] instanceof Date ? Utilities.formatDate(row[8], 'GMT+7', 'yyyy-MM-dd HH:mm:ss') : row[8]
      });
    }
    
    return { success: true, data: tasks };
  } catch (error) {
    console.error('Error in getTasks:', error);
    return { success: false, message: error.toString() };
  }
}

function saveTask(email, task) {
  try {
    const userData = getUserFromMasterSheet(email);
    if (!userData || !userData.spreadsheetId) {
      return { success: false, message: 'User data not found' };
    }
    
    const userSS = SpreadsheetApp.openById(userData.spreadsheetId);
    let taskSheet = userSS.getSheetByName('Tasks');
    
    // Create Tasks sheet if it doesn't exist
    if (!taskSheet) {
      taskSheet = userSS.insertSheet('Tasks');
      taskSheet.appendRow(['ID', 'Task Name', 'Project', 'Due Date', 'Priority', 'Status', 'Description', 'Created At', 'Updated At']);
      taskSheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#4361ee').setFontColor('white');
    }
    
    const data = taskSheet.getDataRange().getValues();
    const now = new Date();
    
    if (task.id) {
      // Update existing task
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === task.id) {
          taskSheet.getRange(i + 1, 2).setValue(task.name);
          taskSheet.getRange(i + 1, 3).setValue(task.project);
          taskSheet.getRange(i + 1, 4).setValue(new Date(task.dueDate));
          taskSheet.getRange(i + 1, 5).setValue(task.priority);
          taskSheet.getRange(i + 1, 6).setValue(task.status);
          taskSheet.getRange(i + 1, 7).setValue(task.description);
          taskSheet.getRange(i + 1, 9).setValue(now);
          break;
        }
      }
    } else {
      // Add new task
      const newId = Utilities.getUuid();
      taskSheet.appendRow([
        newId,
        task.name,
        task.project,
        new Date(task.dueDate),
        task.priority,
        task.status,
        task.description,
        now,
        now
      ]);
    }
    
    return { success: true, message: 'Task saved successfully' };
  } catch (error) {
    console.error('Error in saveTask:', error);
    return { success: false, message: error.toString() };
  }
}

function deleteTask(email, taskId) {
  try {
    const userData = getUserFromMasterSheet(email);
    if (!userData || !userData.spreadsheetId) {
      return { success: false, message: 'User data not found' };
    }
    
    const userSS = SpreadsheetApp.openById(userData.spreadsheetId);
    const taskSheet = userSS.getSheetByName('Tasks');
    
    if (!taskSheet) {
      return { success: false, message: 'Tasks sheet not found' };
    }
    
    const data = taskSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === taskId) {
        taskSheet.deleteRow(i + 1);
        break;
      }
    }
    
    return { success: true, message: 'Task deleted successfully' };
  } catch (error) {
    console.error('Error in deleteTask:', error);
    return { success: false, message: error.toString() };
  }
}

// Di backend, tambahkan fungsi untuk backup/restore
function createBackup(email, name, description) {
  try {
    const userData = getUserFromMasterSheet(email);
    if (!userData || !userData.spreadsheetId) {
      return { success: false, message: 'User data not found' };
    }
    
    const userSS = SpreadsheetApp.openById(userData.spreadsheetId);
    
    // Get data from all sheets
    const backupData = {
      metadata: {
        name: name,
        description: description,
        createdAt: new Date().toISOString(),
        version: '1.0'
      },
      sheets: {}
    };
    
    // Get all sheet names
    const sheets = userSS.getSheets();
    
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const data = sheet.getDataRange().getValues();
      
      backupData.sheets[sheetName] = {
        data: data,
        properties: {
          name: sheetName,
          index: sheet.getIndex()
        }
      };
    });
    
    // Create backup file in Drive
    const fileName = `Backup_${email}_${new Date().getTime()}.json`;
    const blob = Utilities.newBlob(JSON.stringify(backupData), 'application/json', fileName);
    const file = DriveApp.createFile(blob);
    
    // Share with user
    file.addViewer(email);
    
    // Log backup
    logActivity(email, 'CREATE_BACKUP', `File: ${file.getId()}`);
    
    return {
      success: true,
      message: 'Backup created successfully',
      fileId: file.getId(),
      fileName: fileName,
      downloadUrl: file.getDownloadUrl()
    };
  } catch (error) {
    console.error('Error in createBackup:', error);
    return { success: false, message: error.toString() };
  }
}

function getBackupHistory(email) {
  try {
    // Get all files with "Backup_" prefix
    const files = DriveApp.searchFiles(`title contains 'Backup_${email}' and trashed=false`);
    
    const backups = [];
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const match = fileName.match(/Backup_([^_]+)_(\d+)\.json/);
      
      if (match) {
        const timestamp = parseInt(match[2]);
        const createdAt = new Date(timestamp);
        
        backups.push({
          id: file.getId(),
          name: fileName,
          createdAt: createdAt.toISOString(),
          size: file.getSize(),
          downloadUrl: file.getDownloadUrl()
        });
      }
    }
    
    // Sort by creation date (newest first)
    backups.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
    
    return { success: true, data: backups };
  } catch (error) {
    console.error('Error in getBackupHistory:', error);
    return { success: false, message: error.toString() };
  }
}

function restoreFromBackup(email, fileId) {
  try {
    const userData = getUserFromMasterSheet(email);
    if (!userData || !userData.spreadsheetId) {
      return { success: false, message: 'User data not found' };
    }
    
    // Get backup file
    const file = DriveApp.getFileById(fileId);
    const content = file.getBlob().getDataAsString();
    const backupData = JSON.parse(content);
    
    // Validate backup data
    if (!backupData.sheets) {
      return { success: false, message: 'Invalid backup file format' };
    }
    
    const userSS = SpreadsheetApp.openById(userData.spreadsheetId);
    
    // Clear all sheets (except the first one to avoid errors)
    const sheets = userSS.getSheets();
    for (let i = 1; i < sheets.length; i++) {
      userSS.deleteSheet(sheets[i]);
    }
    
    // Restore sheets
    Object.keys(backupData.sheets).forEach(sheetName => {
      const sheetData = backupData.sheets[sheetName];
      
      // Create or get sheet
      let sheet;
      try {
        sheet = userSS.getSheetByName(sheetName);
        if (!sheet) {
          sheet = userSS.insertSheet(sheetName);
        }
      } catch (e) {
        // If sheet name is invalid, use a generic name
        sheet = userSS.insertSheet(`Sheet_${userSS.getSheets().length + 1}`);
      }
      
      // Clear sheet
      sheet.clear();
      
      // Restore data
      if (sheetData.data && sheetData.data.length > 0) {
        sheet.getRange(1, 1, sheetData.data.length, sheetData.data[0].length).setValues(sheetData.data);
      }
    });
    
    // Log restore
    logActivity(email, 'RESTORE_BACKUP', `File: ${fileId}`);
    
    return {
      success: true,
      message: 'Data restored successfully'
    };
  } catch (error) {
    console.error('Error in restoreFromBackup:', error);
    return { success: false, message: error.toString() };
  }
}

// Fungsi untuk menampilkan status koneksi Telegram di website
function getTelegramConnectionStatus(email) {
  try {
    const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    const userSheet = masterSS.getSheetByName('Users');
    
    if (!userSheet) {
      return { connected: false, message: 'Sheet Users tidak ditemukan' };
    }
    
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Cari indeks kolom
    const emailCol = headers.indexOf('Email');
    const chatIdCol = headers.indexOf('ChatID');
    const telegramTokenCol = headers.indexOf('TelegramToken');
    
    if (emailCol === -1 || chatIdCol === -1 || telegramTokenCol === -1) {
      return { connected: false, message: 'Struktur sheet tidak valid' };
    }
    
    // Cari user dengan email yang sesuai
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailCol] === email) {
        const chatId = data[i][chatIdCol];
        const token = data[i][telegramTokenCol];
        
        if (chatId) {
          return { 
            connected: true, 
            message: 'Telegram terhubung',
            token: token
          };
        } else if (token) {
          return { 
            connected: false, 
            message: 'Token tersedia, tetapi belum terhubung dengan Telegram',
            token: token
          };
        } else {
          return { 
            connected: false, 
            message: 'Token belum dibuat',
            token: null
          };
        }
      }
    }
    
    return { connected: false, message: 'User tidak ditemukan' };
    
  } catch (error) {
    Logger.log('Error in getTelegramConnectionStatus: ' + error.toString());
    return { connected: false, message: 'Error: ' + error.message };
  }
}

/**
 * Update data laporan di sheet Formulir
 * @param {string} userEmail
 * @param {number} rowIndex - Index baris yang akan diupdate (dimulai dari 1 untuk data pertama)
 * @param {object} updatedData - Object berisi data baru
 * @return {object} {success: boolean, message: string}
 */
function updateLaporanData(userEmail, globalIndex, updatedData) {
  try {
    const userSS = openUserSpreadsheet(userEmail);
    const sheet = userSS.getSheetByName('Formulir');
    if (!sheet) {
      return { success: false, message: 'Sheet Formulir tidak ditemukan' };
    }

    // Hitung nomor baris sebenarnya di spreadsheet
    // Header di baris 1, data pertama di baris 2
    const rowIndex = globalIndex + 2;

    // Konversi data ke format array sesuai kolom
    const rowData = [
      sheet.getRange(rowIndex, 1).getValue(), // No (tidak diubah)
      new Date(updatedData.date),             // Tanggal
      updatedData.project,                     // Project
      updatedData.task,                        // Task
      '',                                      // Task
      updatedData.progress,                    // Progress
      updatedData.target ? new Date(updatedData.target) : null, // Target
      updatedData.note                         // Note
    ];

    // Tulis data ke baris yang ditentukan
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    
    logActivity(userEmail, 'UPDATE_LAPORAN', `Updated row ${rowIndex}`);
    return { success: true, message: 'Data berhasil diperbarui' };
    
  } catch (error) {
    logActivity(userEmail, 'UPDATE_LAPORAN_ERROR', error.toString());
    return { success: false, message: 'Error: ' + error.message };
  }
}

/**
 * Hapus data laporan di sheet Formulir
 * @param {string} userEmail
 * @param {number} rowIndex - Index baris yang akan dihapus (dimulai dari 1 untuk data pertama)
 * @return {object} {success: boolean, message: string}
 */
function deleteLaporanData(userEmail, globalIndex) {
  try {
    const userSS = openUserSpreadsheet(userEmail);
    const sheet = userSS.getSheetByName('Formulir');
    if (!sheet) {
      return { success: false, message: 'Sheet Formulir tidak ditemukan' };
    }
    
    // Hitung nomor baris sebenarnya di spreadsheet
    const rowIndex = globalIndex + 2;
    
    // Hapus baris
    sheet.deleteRow(rowIndex);
    
    logActivity(userEmail, 'DELETE_LAPORAN', `Deleted row ${rowIndex}`);
    return { success: true, message: 'Data berhasil dihapus' };
    
  } catch (error) {
    logActivity(userEmail, 'DELETE_LAPORAN_ERROR', error.toString());
    return { success: false, message: 'Error: ' + error.message };
  }
}

/**
 * Get a list of generated PDF files for the user from Google Drive.
 * @param {string} userEmail
 * @return {object} {success: boolean, data: array, message: string}
 */
function getPdfHistory(userEmail) {
  try {
    const userData = getUserFromMasterSheet(userEmail);
    if (!userData || !userData.spreadsheetId) {
      return { success: false, message: 'User data not found' };
    }
    
    const searchPattern = `
    (
      title contains 'Laporan_${userData.fullName}_'
      or title contains 'Laporan_Harian_${userData.fullName}_'
      or title contains 'Laporan_ResetDevice_${userData.fullName}_'
    )
    and mimeType = 'application/pdf'
  `;
    
    const files = DriveApp.searchFiles(searchPattern);
    const pdfHistory = [];
    
    while (files.hasNext()) {
      const file = files.next();
      const createdDate = file.getDateCreated();
      
      pdfHistory.push({
        id: file.getId(),
        name: file.getName(),
        createdAt: Utilities.formatDate(createdDate, 'GMT+7', 'dd/MM/yyyy HH:mm:ss'),
        size: formatFileSize(file.getSize()),
        downloadUrl: file.getDownloadUrl()
      });
    }
    
    // Urutkan dari yang terbaru
    pdfHistory.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
    
    logActivity(userEmail, 'GET_PDF_HISTORY', `Found ${pdfHistory.length} PDF files.`);
    
    return {
      success: true,
      data: pdfHistory,
      message: 'PDF history loaded successfully'
    };
    
  } catch (error) {
    logActivity(userEmail, 'GET_PDF_HISTORY_ERROR', error.toString());
    return {
      success: false,
      message: 'Error loading PDF history: ' + error.message,
      data: []
    };
  }
}

/**
 * Helper function untuk memformat ukuran file.
 * @param {number} bytes
 * @return {string} Formatted size string (e.g., "125.5 KB")
 */
function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}