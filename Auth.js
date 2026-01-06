// ============================================
// AUTHENTICATION FUNCTIONS
// ============================================

/**
 * Register new user
 * @param {string} email
 * @param {string} password
 * @param {string} fullName
 * @param {string} telegramId (optional)
 * @return {object} {success: boolean, message: string, data: object}
 */
function registerUser(email, password, fullName, telegramId = '') {
  console.log('=== REGISTER USER ===');
  console.log('Email:', email);
  console.log('Full Name:', fullName);
  
  try {
    // 1. Cek apakah user sudah ada
    console.log('Checking existing user...');
    const existingUser = getUserFromSheet(email);
    if (existingUser) {
      console.log('User already exists');
      return {success: false, message: 'Email sudah terdaftar'};
    }
    
    // 2. Hash password
    console.log('Hashing password...');
    const passwordHash = hashPassword(password);
    console.log('Password hash:', passwordHash);
    
    // 3. Buat spreadsheet untuk user
    console.log('Creating user spreadsheet...');
    const spreadsheetId = createUserSpreadsheet(email);
    console.log('Spreadsheet created with ID:', spreadsheetId);
    
    // 4. Simpan ke sheet Users di spreadsheet admin
    console.log('Saving to Users sheet...');
    const result = saveUserToSheet(email, passwordHash, fullName, telegramId, spreadsheetId);
    
    if (!result.success) {
      return result;
    }
    
    console.log('=== REGISTRATION SUCCESSFUL ===');
    return {
      success: true,
      message: 'Registrasi berhasil! Silakan login.',
      data: {
        email: email,
        fullName: fullName,
        role: (email.toLowerCase() === CONFIG.ADMIN_EMAIL.toLowerCase()) ? 'admin' : 'user'
      }
    };
    
  } catch (error) {
    console.error('=== REGISTRATION ERROR ===');
    console.error('Error:', error.toString());
    console.error('Stack:', error.stack);
    
    return {
      success: false, 
      message: 'Error registrasi: ' + error.message
    };
  }
}

// Helper function: Cari user di sheet
function getUserFromSheet(email) {
  try {
    const masterSS = SpreadsheetApp.openById('180H7RpJnQ5-VgmDs7o7QU061EvrFQDqYXYrqSS0hTU4');
    const userSheet = masterSS.getSheetByName('Users');
    
    if (!userSheet) return null;
    
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    
    if (emailCol === -1) return null;
    
    for (let i = 1; i < data.length; i++) {
      const rowEmail = data[i][emailCol];
      if (rowEmail && rowEmail.toString().toLowerCase() === email.toLowerCase()) {
        return {
          email: data[i][emailCol],
          passwordHash: data[i][headers.indexOf('Password')],
          role: data[i][headers.indexOf('Role')] || 'user',
          fullName: data[i][headers.indexOf('FullName')] || email.split('@')[0],
          telegramId: data[i][headers.indexOf('TelegramToken')] || '',
          spreadsheetId: data[i][headers.indexOf('SpreadsheetID')] || ''
        };
      }
    }
    
    return null;
  } catch (error) {
    console.error('Error getUserFromSheet:', error);
    return null;
  }
}

/**
 * Login user
 * @param {string} email
 * @param {string} password
 * @return {object} {success: boolean, message: string, data: object}
 */
function loginUser(email, password) {
  console.log('Login attempt for:', email);
  
  try {
    // Langsung cek di spreadsheet Users
    console.log('Checking in spreadsheet...');
    const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    const userSheet = masterSS.getSheetByName('Users');
    
    if (!userSheet) {
      console.log('Users sheet not found');
      return {success: false, message: 'Sheet Users tidak ditemukan'};
    }
    
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Cari indeks kolom dengan cara yang lebih aman
    const emailCol = findColumnIndex(headers, ['Email', 'email', 'EMAIL']);
    const passCol = findColumnIndex(headers, ['Password', 'PasswordHash', 'password', 'passwordhash']);
    const roleCol = findColumnIndex(headers, ['Role', 'role', 'ROLE']);
    const nameCol = findColumnIndex(headers, ['FullName', 'Full Name', 'fullName', 'Nama']);
    const telegramCol = findColumnIndex(headers, ['TelegramToken', 'Telegram Token', 'TelegramID', 'telegramId']);
    const ssIdCol = findColumnIndex(headers, ['SpreadsheetID', 'SpreadsheetId', 'Spreadsheet id', 'spreadsheetid']);
    const lastLoginCol = findColumnIndex(headers, ['LastLogin', 'Last Login', 'lastlogin']);
    
    if (emailCol === -1) {
      console.log('Email column not found');
      return {success: false, message: 'Kolom Email tidak ditemukan di sheet Users'};
    }
    
    // Cari user berdasarkan email
    for (let i = 1; i < data.length; i++) {
      const rowEmail = data[i][emailCol];
      
      if (rowEmail && rowEmail.toString().toLowerCase() === email.toLowerCase()) {
        console.log('Found user in row:', i);
        
        // Jika password kosong di sheet, buat dari input
        const storedHash = passCol !== -1 ? data[i][passCol] : '';
        const inputHash = hashPassword(password);
        
        console.log('Row password:', storedHash);
        console.log('Input hash:', inputHash);
        
        // Jika password di sheet kosong, set password baru
        if (!storedHash || storedHash === '' || storedHash === 'pending') {
          console.log('Password is empty, setting new password');
          if (passCol !== -1) {
            userSheet.getRange(i + 1, passCol + 1).setValue(inputHash);
          }
          
          // Update last login
          const now = new Date();
          if (lastLoginCol !== -1) {
            userSheet.getRange(i + 1, lastLoginCol + 1).setValue(now);
          }
          
          const userData = {
            email: email,
            passwordHash: inputHash,
            role: roleCol !== -1 ? data[i][roleCol] : 'user',
            fullName: nameCol !== -1 ? data[i][nameCol] : email.split('@')[0],
            telegramId: telegramCol !== -1 ? data[i][telegramCol] : '',
            spreadsheetId: ssIdCol !== -1 ? data[i][ssIdCol] : ''
          };
          
          return {
            success: true,
            data: userData,
            message: 'Login berhasil (password di-set)'
          };
        }
        
        // Bandingkan hash
        if (storedHash === inputHash) {
          console.log('Password matches!');
          
          const userData = {
            email: email,
            passwordHash: storedHash,
            role: roleCol !== -1 ? data[i][roleCol] : 'user',
            fullName: nameCol !== -1 ? data[i][nameCol] : email.split('@')[0],
            telegramId: telegramCol !== -1 ? data[i][telegramCol] : '',
            spreadsheetId: ssIdCol !== -1 ? data[i][ssIdCol] : ''
          };
          
          // Update last login
          const now = new Date();
          if (lastLoginCol !== -1) {
            userSheet.getRange(i + 1, lastLoginCol + 1).setValue(now);
          }
          
          return {
            success: true,
            data: userData,
            message: 'Login berhasil'
          };
        } else {
          console.log('Password mismatch');
          return {success: false, message: 'Password salah'};
        }
      }
    }
    
    console.log('User not found in sheet');
    return {success: false, message: 'User tidak ditemukan'};
    
  } catch (error) {
    console.error('Login error:', error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Fungsi hash password yang konsisten
function hashPassword(password) {
  if (!password) return '';
  
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, 
    password
  );
  
  // Convert byte array to hex string
  let hash = '';
  for (let i = 0; i < digest.length; i++) {
    const byte = digest[i];
    if (byte < 0) {
      hash += (byte + 256).toString(16).padStart(2, '0');
    } else {
      hash += byte.toString(16).padStart(2, '0');
    }
  }
  
  return hash;
}

/**
 * Get user data from master sheet
 * @param {string} email
 * @return {object|null} user object
 */
function getUserFromMasterSheet(email) {
  const ss = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.USER_SHEET_NAME);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[2] === email) { // Kolom C adalah Email
      const user = {};
      headers.forEach((header, index) => {
        user[header] = row[index];
      });
      return user;
    }
  }
  return null;
}

/**
 * Update user profile
 * @param {string} email
 * @param {object} updates - fields to update
 * @return {object} {success: boolean, message: string}
 */
function updateProfile(email, updates) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.USER_SHEET_NAME);
    if (!sheet) return {success: false, message: 'Sheet Users tidak ditemukan'};

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    let updated = false;

    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === email) { // Kolom C = Email
        // Update each field
        if (updates.newPassword) {
          const passwordHash = hashPassword(updates.newPassword);
          const colIndex = headers.indexOf('PasswordHash');
          sheet.getRange(i + 1, colIndex + 1).setValue(passwordHash);
        }
        if (updates.newFullName) {
          const colIndex = headers.indexOf('FullName');
          sheet.getRange(i + 1, colIndex + 1).setValue(updates.newFullName);
        }
        if (updates.newTelegramId) {
          const colIndex = headers.indexOf('TelegramId');
          sheet.getRange(i + 1, colIndex + 1).setValue(updates.newTelegramId);
        }
        updated = true;
        break;
      }
    }

    if (updated) {
      logActivity(email, 'UPDATE_PROFILE', 'Profile updated');
      return {success: true, message: 'Profile updated successfully'};
    } else {
      return {success: false, message: 'User not found'};
    }
  } catch (error) {
    logActivity(email, 'UPDATE_PROFILE_ERROR', error.toString());
    return {success: false, message: 'Error: ' + error.message};
  }
}

/**
 * Logout user (client-side hanya hapus session, ini untuk log)
 * @param {string} email
 * @return {object} {success: boolean, message: string}
 */
function logoutUser(email) {
  logActivity(email, 'LOGOUT', 'User logged out');
  return {success: true, message: 'Logout berhasil'};
}

/**
 * Get user's profile data
 * @param {string} userEmail
 * @return {object} user profile
 */
function getUserDetail(email) {
  console.log('=== getUserDetail ===', email);
  
  try {
    const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    const userSheet = masterSS.getSheetByName('Users');
    
    if (!userSheet) {
      return null;
    }
    
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Cari indeks kolom
    const emailCol = headers.indexOf('Email');
    const fullNameCol = headers.indexOf('FullName');
    const roleCol = headers.indexOf('Role');
    const telegramTokenCol = headers.indexOf('TelegramToken');
    const chatIdCol = headers.indexOf('ChatID');

    console.log('=== getData ===', data);
    console.log('=== getHeader ===', headers);
    
    if (emailCol === -1) {
      return null;
    }
    
    // Cari user berdasarkan email
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailCol] === email) {
        return {
          Email: data[i][emailCol],
          FullName: data[i][fullNameCol] || email.split('@')[0],
          Role: data[i][roleCol] || 'user',
          TelegramToken: data[i][telegramTokenCol] || '',
          ChatID: data[i][chatIdCol] || ''
        };
      }
      console.log('=== getDataIF ===', data);
    }
    
    return null;
    
  } catch (error) {
    console.error('Error getUserDetail:', error);
    return null;
  }
}