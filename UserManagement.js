// ============================================
// USER SPREADSHEET MANAGEMENT
// ============================================

/**
 * Create a new spreadsheet for user based on template
 * @param {string} userEmail
 * @return {string} spreadsheetId
 */
function createUserSpreadsheet(userEmail) {
  try {
    logActivity('SYSTEM', 'CREATE_USER_SPREADSHEET', `Creating for: ${userEmail}`);

    // Buka template
    const template = SpreadsheetApp.openById(CONFIG.TEMPLATE_SPREADSHEET_ID);

    // Buat nama untuk spreadsheet baru
    const userName = userEmail.split('@')[0];
    const timestamp = new Date().getTime();
    const newFileName = `User_${userName}_${timestamp}`;

    // Salin template
    const newFile = DriveApp.getFileById(CONFIG.TEMPLATE_SPREADSHEET_ID).makeCopy(newFileName);
    const newSpreadsheetId = newFile.getId();

    // Beri akses view dan edit ke user (edit karena user akan menulis data)
    newFile.addEditor(userEmail);

    // Simpan mapping di ScriptProperties
    const props = PropertiesService.getScriptProperties();
    props.setProperty(`USER_SS_${userEmail}`, newSpreadsheetId);

    logActivity(userEmail, 'SPREADSHEET_CREATED', `ID: ${newSpreadsheetId}`);
    return newSpreadsheetId;

  } catch (error) {
    logActivity('SYSTEM', 'CREATE_USER_SPREADSHEET_ERROR', error.toString());
    throw new Error('Gagal membuat spreadsheet: ' + error.message);
  }
}

/**
 * Get user's profile data
 * @param {string} userEmail
 * @return {object} user profile
 */
function getUserProfile(userEmail) {
  const userData = getUserFromMasterSheet(userEmail);
  if (!userData) throw new Error('User tidak ditemukan');

  return {
    email: userData.Email,
    fullName: userData.FullName,
    telegramId: userData.TelegramId,
    role: userData.Role,
    status: userData.Status
  };
}

function getUserFromMasterSheet(email) {
  console.log('=== getUserFromMasterSheet ===');
  console.log('Mencari:', email);
  
  try {
    // 1. Buka spreadsheet master
    let masterSS;
    try {
      masterSS = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
      console.log('Master spreadsheet ditemukan:', masterSS.getName());
    } catch (error) {
      console.error('Tidak bisa membuka master spreadsheet:', error);
      return null;
    }
    
    // 2. Dapatkan sheet Users
    const userSheet = masterSS.getSheetByName('Users');
    if (!userSheet) {
      console.error('Sheet "Users" tidak ditemukan di master spreadsheet');
      
      // Cek semua sheets yang ada
      const allSheets = masterSS.getSheets().map(s => s.getName());
      console.log('Sheets yang ada:', allSheets);
      
      return null;
    }
    
    console.log('Sheet Users ditemukan');
    
    // 3. Ambil semua data
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    
    console.log('Total baris:', data.length);
    console.log('Headers:', headers);
    
    // 4. Cari kolom yang diperlukan
    const emailCol = findColumnIndex(headers, ['Email', 'email', 'EMAIL']);
    const ssIdCol = findColumnIndex(headers, ['SpreadsheetID', 'SpreadsheetId', 'Spreadsheet ID', 'spreadsheetId']);
    const roleCol = findColumnIndex(headers, ['Role', 'role', 'ROLE']);
    const fullNameCol = findColumnIndex(headers, ['FullName', 'Full Name', 'fullName', 'Nama']);
    const telegramCol = findColumnIndex(headers, ['TelegramToken', 'Telegram Token', 'TelegramID', 'telegramId']);
    const chatidCol = findColumnIndex(headers, ['ChatID', 'ChatId', 'Chat_ID', 'Chat_Id']);
    
    console.log('Kolom ditemukan:', {
      email: emailCol,
      spreadsheetId: ssIdCol,
      role: roleCol,
      fullName: fullNameCol,
      telegram: telegramCol,
      chatId: chatidCol
    });
    
    if (emailCol === -1) {
      console.error('Kolom Email tidak ditemukan di sheet Users');
      return null;
    }
    
    // 5. Cari user berdasarkan email (case-insensitive)
    for (let i = 1; i < data.length; i++) {
      const rowEmail = data[i][emailCol];
      console.log('nilai data email:',rowEmail); 
      if (rowEmail && rowEmail.toString().toLowerCase() === email.toLowerCase()) {
        console.log('User ditemukan di baris:', i + 1);
        
        const userData = {
          email: rowEmail,
          spreadsheetId: ssIdCol !== -1 ? data[i][ssIdCol] : null,
          role: roleCol !== -1 ? data[i][roleCol] : 'user',
          fullName: fullNameCol !== -1 ? data[i][fullNameCol] : rowEmail.split('@')[0],
          telegramId: telegramCol !== -1 ? data[i][telegramCol] : '',
          chatId: chatidCol !== -1 ? data[i][chatidCol] : '',
          rowIndex: i + 1
        };
        
        console.log('Data user:', userData);
        return userData;
      }
    }
    
    console.log('User tidak ditemukan di sheet Users');
    return null;
    
  } catch (error) {
    console.error('Error getUserFromMasterSheet:', error);
    return null;
  }
}

// Helper function untuk mencari kolom dengan berbagai kemungkinan nama
function findColumnIndex(headers, possibleNames) {
  for (const name of possibleNames) {
    const index = headers.indexOf(name);
    if (index !== -1) return index;
  }
  return -1;
}

/**
 * Get user's spreadsheet ID
 * @param {string} userEmail
 * @return {string|null} spreadsheetId
 */
function getUserSpreadsheetId(userEmail) {
  const props = PropertiesService.getScriptProperties();
  const ssId = props.getProperty(`USER_SS_${userEmail}`);
  if (ssId) return ssId;

  // Jika tidak ada di Properties, cari di sheet Users
  const userData = getUserFromMasterSheet(userEmail);
  if (userData && userData.spreadsheetId) {
    return userData.spreadsheetId;
  }

  return null;
}

/**
 * Open user's spreadsheet
 * @param {string} userEmail
 * @return {Spreadsheet} spreadsheet object
 */
function openUserSpreadsheet(userEmail) {
  // userEmail='dbcloudvacant@gmail.com';
  const ssId = getUserSpreadsheetId(userEmail);
  console.log('mendapatkan nilai ssID :', ssId);
  if (!ssId) throw new Error('Spreadsheet user tidak ditemukan');
  return SpreadsheetApp.openById(ssId);
}

/**
 * Get all users for admin
 * @return {array} array of user data
 */
function getAllUsers() {
  console.log('=== getAllUsers dipanggil ===');
  
  try {
    const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    if (!masterSS) {
      console.error('Tidak bisa membuka spreadsheet master');
      return [];
    }
    
    const userSheet = masterSS.getSheetByName('Users');
    if (!userSheet) {
      console.error('Sheet Users tidak ditemukan');
      return [];
    }
    
    const data = userSheet.getDataRange().getValues();
    console.log('Total baris di sheet Users:', data.length);
    
    if (data.length <= 1) {
      console.log('Hanya ada header, tidak ada data user');
      return [];
    }
    
    const headers = data[0];
    console.log('Headers ditemukan:', headers);
    
    // Mapping header ke lowercase untuk case-insensitive
    const headerMap = {};
    headers.forEach((header, index) => {
      if (header && typeof header === 'string') {
        headerMap[header.toString().trim().toLowerCase()] = index;
      }
    });
    
    console.log('Header mapping:', headerMap);
    
    const users = [];
    
    // Proses data dengan error handling per baris
    for (let i = 1; i < data.length; i++) {
      try {
        const row = data[i];
        const user = {};
        
        // Pastikan setiap properti ada dan valid
        user.rowNumber = i + 1;
        user.Timestamp = safeGetValue(row, headerMap, 'timestamp');
        user.Username = safeGetValue(row, headerMap, 'username') || '';
        user.Email = safeGetValue(row, headerMap, 'email') || '';
        user.PasswordHash = safeGetValue(row, headerMap, ['passwordhash', 'password']) || '';
        user.TelegramToken = safeGetValue(row, headerMap, ['telegramtoken', 'telegram id', 'telegram']) || '';
        user.SpreadsheetID = safeGetValue(row, headerMap, ['spreadsheetid', 'spreadsheet id']) || '';
        user.Status = safeGetValue(row, headerMap, 'status') || 'active';
        user.ChatID = safeGetValue(row, headerMap, ['chatid', 'chat id','ChatID','ChatId']) || '';
        user.Role = safeGetValue(row, headerMap, 'role') || 'user';
        user.FullName = safeGetValue(row, headerMap, ['fullname', 'full name', 'nama']) || '';
        user.LastLogin = safeGetValue(row, headerMap, ['lastlogin', 'last login']) || '';
        
        // Hanya tambahkan jika email valid
        if (user.Email && user.Email.includes('@')) {
          users.push(user);
        }
      } catch (rowError) {
        console.error(`Error memproses baris ${i + 1}:`, rowError);
        // Lanjut ke baris berikutnya
      }
    }
    
    console.log(`Berhasil memproses ${users.length} user`);
    
    // Sederhanakan data untuk menghindari masalah serialisasi
    const simplifiedUsers = users.map(user => {
      return {
        Email: user.Email,
        FullName: user.FullName,
        Role: user.Role,
        Status: user.Status,
        TelegramToken: user.TelegramToken,
        SpreadsheetID: user.SpreadsheetID,
        Username: user.Username,
        ChatID: user.ChatID
      };
    });
    
    return simplifiedUsers;
    
  } catch (error) {
    console.error('Error kritis di getAllUsers:', error);
    // Return array kosong agar frontend tidak error
    return [];
  }
}

function safeGetValue(row, headerMap, possibleHeaders) {
  try {
    if (!row || !headerMap) return '';
    
    const headers = Array.isArray(possibleHeaders) ? possibleHeaders : [possibleHeaders];
    
    for (const header of headers) {
      const index = headerMap[header.toLowerCase()];
      if (index !== undefined && row[index] !== undefined && row[index] !== null && row[index] !== '') {
        // Format tanggal jika perlu
        const value = row[index];
        if (value instanceof Date) {
          return value.toISOString();
        }
        return value.toString();
      }
    }
    
    return '';
  } catch (error) {
    console.error('Error safeGetValue:', error);
    return '';
  }
}

/**
 * Reset user data (admin only)
 * @param {string} adminEmail
 * @param {string} targetEmail
 * @param {string} newPassword
 * @param {string} newTelegramId
 * @param {string} newStatus
 * @param {string} newChatId
 * @return {object} {success: boolean, message: string}
 */
function resetUserData(adminEmail, targetEmail, newPassword, newTelegramId, newStatus, newChatId) {
  console.log('=== resetUserData ===');
  console.log('Admin:', adminEmail);
  console.log('Target:', targetEmail);
  console.log('Updates:', { newPassword, newTelegramId, newStatus, newChatId });
  
  try {
    // 1. Verifikasi admin
    const adminData = getUserFromMasterSheet(adminEmail);
    if (!adminData || adminData.role !== 'admin') {
      throw new Error('Hanya admin yang bisa reset user');
    }
    
    // 2. Dapatkan data user target
    const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    const userSheet = masterSS.getSheetByName('Users');
    
    if (!userSheet) {
      throw new Error('Sheet Users tidak ditemukan');
    }
    
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Cari indeks kolom
    const emailCol = headers.indexOf('Email');
    const passCol = findColumnIndex(headers, ['Password', 'PasswordHash', 'Passwordhash']);
    const telegramCol = headers.indexOf('TelegramToken');
    const statusCol = headers.indexOf('Status');
    const chatIdCol = headers.indexOf('ChatID');
    
    if (emailCol === -1) {
      throw new Error('Kolom Email tidak ditemukan');
    }
    
    // 3. Cari dan update user
    let userFound = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailCol] === targetEmail) {
        userFound = true;

        if (newPassword && newPassword.trim() !== '') {
          if (passCol === -1) {
            throw new Error('Kolom Password tidak ditemukan');
          }
          const newHash = hashPassword(newPassword);
          userSheet.getRange(i + 1, passCol + 1).setValue(newHash);
        }
        
        // Update telegram ID
        if (newTelegramId && telegramCol !== -1) {
          userSheet.getRange(i + 1, telegramCol + 1).setValue(newTelegramId);
        }
        
        // Update status
        if (newStatus && statusCol !== -1) {
          userSheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
        }
        
        // Update chat ID
        if (newChatId && chatIdCol !== -1) {
          userSheet.getRange(i + 1, chatIdCol + 1).setValue(newChatId);
        }
        
        break;
      }
    }
    
    if (!userFound) {
      throw new Error(`User ${targetEmail} tidak ditemukan`);
    }
    
    // 4. Update cache jika perlu
    const props = PropertiesService.getScriptProperties();
    const userKey = `USER_${targetEmail.toLowerCase()}`;
    const cachedUser = props.getProperty(userKey);
    
    if (cachedUser) {
      const userData = JSON.parse(cachedUser);
      if (newPassword) userData.passwordHash = hashPassword(newPassword);
      if (newTelegramId) userData.telegramId = newTelegramId;
      if (newStatus) userData.status = newStatus;
      
      props.setProperty(userKey, JSON.stringify(userData));
    }
    
    console.log('Reset user berhasil untuk:', targetEmail);
    return {
      success: true,
      message: `User ${targetEmail} berhasil direset`
    };
    
  } catch (error) {
    console.error('Error resetUserData:', error);
    return {
      success: false,
      message: 'Error: ' + error.message
    };
  }
}

// ========== FUNGSI BANTUAN ==========
function saveUserToSheet(email, passwordHash, fullName, telegramId, spreadsheetId) {
  try {
    console.log('Opening master spreadsheet...');
    console.log('Master ID:', CONFIG.MASTER_SPREADSHEET_ID);
    
    // Coba buka spreadsheet
    let masterSS;
    try {
      masterSS = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
      console.log('Spreadsheet opened:', masterSS.getName());
    } catch (openError) {
      console.error('Cannot open spreadsheet:', openError);
      
      // Fallback: Coba buat spreadsheet baru jika tidak ada
      console.log('Trying to create new spreadsheet...');
      masterSS = SpreadsheetApp.create('MyApp_Master_Database');
      console.log('Created new spreadsheet:', masterSS.getId());
      
      // Update config dengan ID baru
      CONFIG.MASTER_SPREADSHEET_ID = masterSS.getId();
    }
    
    // Dapatkan atau buat sheet Users
    let userSheet = masterSS.getSheetByName('Users');
    if (!userSheet) {
      console.log('Creating Users sheet...');
      userSheet = masterSS.insertSheet('Users');
      
      // Buat header
      const headers = [
        'Timestamp', 'Username', 'Email', 'Password', 
        'TelegramToken', 'SpreadsheetID', 'Status', 
        'ChatID', 'Role', 'FullName'
      ];
      userSheet.appendRow(headers);
      
      // Format header
      userSheet.getRange(1, 1, 1, headers.length)
        .setBackground('#4CAF50')
        .setFontColor('white')
        .setFontWeight('bold');
      
      console.log('Users sheet created with headers');
    }
    
    // Tambahkan user baru
    const timestamp = new Date();
    const username = email.split('@')[0];
    const role = (email.toLowerCase() === CONFIG.ADMIN_EMAIL.toLowerCase()) ? 'admin' : 'user';
    
    const newRow = [
      timestamp,
      username,
      email,
      passwordHash,
      telegramId || '',
      spreadsheetId,
      'active',
      '',
      role,
      fullName
    ];
    
    userSheet.appendRow(newRow);
    console.log('User saved to sheet at row:', userSheet.getLastRow());
    
    return {success: true};
    
  } catch (error) {
    console.error('Error saving to sheet:', error);
    return {
      success: false,
      message: 'Error saving to sheet: ' + error.message
    };
  }
}