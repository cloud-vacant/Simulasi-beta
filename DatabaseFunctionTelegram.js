// === FORM INPUT FUNCTIONS ===
function tulisNamaKeSheetBot(nama, keter, chatId) {
  try {
    // Dapatkan email user berdasarkan Chat ID
    const userEmail = getUserEmailByChatId(chatId);
    
    if (!userEmail) {
      return "❌ Akun Anda belum terhubung. Silakan hubungkan terlebih dahulu dengan menggunakan token dari website.";
    }
    
    // Dapatkan spreadsheet user
    const userSS = openUserSpreadsheet(userEmail);
    const sheet = userSS.getSheetByName("Formulir");

    const info = getAutoIncrementNumberBot(chatId);
    var project = 'Remote';
    var status = '100';

    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;

    // sheet.getRange(newRow, 1, 1, 7).setValues([
    //   [info.number, info.date, project, nama, status, '', keter]
    // ]);

      // ==== TULIS DATA SATU-PER-SATU (AMAN WALAU ADA MERGE) ====
    sheet.getRange(newRow, 1).setValue(info.number);  // No
    sheet.getRange(newRow, 2).setValue(info.date);    // Date
    sheet.getRange(newRow, 3).setValue(project);      // Project
    sheet.getRange(newRow, 4).setValue(nama);         // Item Task (merged D–E)
    sheet.getRange(newRow, 6).setValue(status);       // Progress
    sheet.getRange(newRow, 7).setValue("");           // Target Date
    sheet.getRange(newRow, 8).setValue(keter);        // Note
  // =========================================================

    // Formatting cells
    const kolomNo = sheet.getRange(2, 1, newRow - 1, 1);
    const kolomTanggal = sheet.getRange(2, 2, newRow - 1, 1);
    const kolom3 = sheet.getRange(2, 3, newRow - 1, 1);
    const kolom5 = sheet.getRange(2, 5, newRow - 1, 1);

    kolomNo.setHorizontalAlignment("center").setVerticalAlignment("middle");
    kolomTanggal.setHorizontalAlignment("center").setVerticalAlignment("middle");
    kolom3.setHorizontalAlignment("center").setVerticalAlignment("middle");
    kolom5.setHorizontalAlignment("center").setVerticalAlignment("middle");

    const rangeDuaKolom = sheet.getRange(2, 1, lastRow, 8);

    rangeDuaKolom.setBorder (
      true, // Batas Atas
      true, // Batas Kiri
      true, // Batas Bawah
      true, // Batas Kanan
      true, // Batas Vertikal (Garis batas antara Kolom A & B)
      true  // Batas Horizontal (Garis batas antara baris)
    );

  // Log aktivitas
  logActivity(userEmail, 'TELEGRAM_INPUT_CASE', `Case: ${nama}, Keterangan: ${keter}`);

    return "✅ Data "+ nama +" berhasil dikirim ke sheet Formulir";
  } catch (error) {
    return "❌ Error: " + error.message;
  }
}

function tulisLisensiBot(emAil, hostName, pswd, expired, keter, chatId) {
  try {
    // Dapatkan email user berdasarkan Chat ID
    const userEmail = getUserEmailByChatId(chatId);
    
    if (!userEmail) {
      return "❌ Akun Anda belum terhubung. Silakan hubungkan terlebih dahulu dengan menggunakan token dari website.";
    }
    
    // Dapatkan spreadsheet user
    const userSS = openUserSpreadsheet(userEmail);
    const sheet = userSS.getSheetByName("LisensiMS");
    
    if (!sheet) {
      return "❌ Sheet 'LisensiMS' tidak ditemukan!";
    }

    var tanggal = new Date();
    var mypassword = hashPasswordBot(pswd);

    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    const newNumber = lastRow;

    sheet.getRange(newRow, 1, 1, 7).setValues([
      [newNumber, emAil, hostName, tanggal, expired, mypassword, keter]
    ]);

    // Log aktivitas
    logActivity(userEmail, 'TELEGRAM_INPUT_LICENSE', `Email: ${emAil}, Hostname: ${hostName}`);

    return "✅ Data " + emAil + " berhasil disimpan ke sheet LisensiMS";
  } catch (error) {
    return "❌ Error: " + error.message;
  }
}

function hashPasswordBot(password) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  var hash = '';
  for (var i = 0; i < digest.length; i++) {
    var hex = (digest[i] & 0xFF).toString(16);
    hash += (hex.length === 1 ? '0' + hex : hex);
  }
  return hash;
}

// === PASSWORD GENERATOR ===
function generatePasswordBot(length, useNumbers, useMixCase, useSymbols) {
  var lowerChars = "abcdefghijklmnopqrstuvwxyz";
  var upperChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  var numbers = "0123456789";
  var symbols = "!@#$&*-_=+?";
  
  var chars = lowerChars;
  if (useMixCase) chars += upperChars;
  if (useNumbers) chars += numbers;
  if (useSymbols) chars += symbols;
  
  var password = "";
  var symbolCount = 0;

  for (var i = 0; i < length; i++) {
    var char;
    var valid = false;
    
    while (!valid) {
      var randomIndex = Math.floor(Math.random() * chars.length);
      char = chars.charAt(randomIndex);
      
      if (symbols.indexOf(char) !== -1) {
        if (symbolCount < 3) {
          symbolCount++;
          valid = true;
        } else {
          valid = false;
        }
      } else {
        valid = true;
      }
    }
    password += char;
  }

  return password;
}

// === WHITELIST SEARCH ===
function searchDataByColumnBot(column, keyword) {
  try {

    // Dapatkan spreadsheet user
    const userSS = SpreadsheetApp.openById(CONFIG.TEMPLATE_SPREADSHEET_ID);
    const sheet = userSS.getSheetByName("WhitelistApp");
    
    if (!sheet) {
      throw new Error("Sheet 'WhitelistApp' tidak ditemukan!");
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    let colIndex = headers.indexOf(column);
    if (colIndex === -1) {
      throw new Error("Kolom tidak ditemukan: " + column);
    }

    let results = [];
    for (let i = 1; i < data.length; i++) {
      let cellValue = String(data[i][colIndex]).toLowerCase();
      if (cellValue.includes(keyword.toLowerCase())) {
        let rowObj = {};
        headers.forEach((h, j) => {
          rowObj[h] = data[i][j];
        });
        results.push(rowObj);
      }
    }
    return results;
  } catch (error) {
    throw new Error("Gagal mencari data: " + error.message);
  }
}

// === AUTO INCREMENT NUMBER ===
function getAutoIncrementNumberBot(chatId) {
  try {
    // Dapatkan email user berdasarkan Chat ID
    const userEmail = getUserEmailByChatId(chatId);
    
    if (!userEmail) {
      return "❌ Akun Anda belum terhubung. Silakan hubungkan terlebih dahulu dengan menggunakan token dari website.";
    }
    
    // Dapatkan spreadsheet user
    const userSS = openUserSpreadsheet(userEmail);
    const sheet = userSS.getSheetByName("Formulir");
    
    if (!sheet) {
      throw new Error("Sheet 'Formulir' tidak ditemukan!");
    }
    
    const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let uniqueDates = [];
    data.forEach(row => {
      const cell = row[0];
      if (cell) {
        const d = new Date(cell);
        d.setHours(0, 0, 0, 0);
        if (!uniqueDates.some(u => u.getTime() === d.getTime())) {
          uniqueDates.push(d);
        }
      }
    });

    const alreadyExists = uniqueDates.some(d => d.getTime() === today.getTime());

    let newNumber;
    if (alreadyExists) {
      newNumber = "";
    } else {
      newNumber = uniqueDates.length + 1;
    }

    return { number: newNumber, date: today };
  } catch (error) {
    // Fallback jika error
    return { number: "1", date: new Date() };
  }
}

// === RESET DEVICE FUNCTIONS ===
function tulisDataResetDeviceBot(data, chatId) {
  try {
    // Dapatkan email user berdasarkan Chat ID
    const userEmail = getUserEmailByChatId(chatId);
    
    if (!userEmail) {
      return "❌ Akun Anda belum terhubung. Silakan hubungkan terlebih dahulu dengan menggunakan token dari website.";
    }
    
    // Dapatkan spreadsheet user
    const userSS = openUserSpreadsheet(userEmail);
    const sheet = userSS.getSheetByName('ResetDevice');
    
    if (!sheet) {
      throw new Error('Sheet "Reset Device" tidak ditemukan!');
    }

    const lastRow = sheet.getLastRow();
    let newNumber = 1;
    if (lastRow >= 2) {
      const lastNumber = sheet.getRange(lastRow, 1).getValue();
      newNumber = (isNaN(lastNumber) ? 0 : Number(lastNumber)) + 1;
    }

    const tanggal = new Date();

    const newRow = [
      newNumber,
      data.nama,
      tanggal,
      data.divisi,
      data.device,
      data.serial,
      data.approved || 'Pending'
    ];

    sheet.appendRow(newRow);
    sheet.getRange(lastRow + 1, 3).setNumberFormat("dd/MM/yyyy");
    sheet.getRange(lastRow + 1, 1, 1, 7)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    const rangeDuaKolom = sheet.getRange(2, 1, lastRow, 7);

    rangeDuaKolom.setBorder (
      true, // Batas Atas
      true, // Batas Kiri
      true, // Batas Bawah
      true, // Batas Kanan
      true, // Batas Vertikal (Garis batas antara Kolom A & B)
      true  // Batas Horizontal (Garis batas antara baris)
    );

    // Log aktivitas
    logActivity(userEmail, 'TELEGRAM_INPUT_RESET_DEVICE', `Name: ${data.nama}, Device: ${data.device}`);

    return `✅ Data ${data.nama} berhasil disimpan ke sheet Reset Device`;
  } catch (error) {
    throw new Error("Gagal menyimpan data reset device: " + error.message);
  }
}

// === RESET DEVICE LIST HANDLER ===
function handleCommandListResetBot(chatId) {
  try {
    // Dapatkan email user berdasarkan Chat ID
    const userEmail = getUserEmailByChatId(chatId);
    
    if (!userEmail) {
      return "❌ Akun Anda belum terhubung. Silakan hubungkan terlebih dahulu dengan menggunakan token dari website.";
    }
    
    // Dapatkan spreadsheet user
    const userSS = openUserSpreadsheet(userEmail);
    const sheet = userSS.getSheetByName('ResetDevice');
    
    if (!sheet) {
      kirimPesanBot(chatId, '❌ Sheet "ResetDevice" tidak ditemukan!');
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      kirimPesanBot(chatId, '📋 Tidak ada data reset device.');
      return;
    }
    
    const headers = data.shift();

    const lastFive = data.slice(-5);
    let count = 0;

    lastFive.forEach((row) => {
      const [no, nama, tanggal, divisi, device, serial, approved] = row;

      const isNotApprovedYet = !approved || approved.trim() === '' || approved.trim() === 'Pending';

      if (isNotApprovedYet) {
        count++;

        const text =
          `🆔 No: ${no}\n` +
          `👤 Nama: ${nama}\n` +
          `🏢 Divisi: ${divisi}\n` +
          `💻 Device: ${device}\n` +
          `🔢 Serial: ${serial}\n` +
          `✅ Status: ${approved || 'Belum di-approve'}`;

        const inlineKeyboard = {
          inline_keyboard: [
            [
              { text: '✅ Approve', callback_data: `approve_${no}` },
              { text: '❌ Reject', callback_data: `reject_${no}` },
            ],
          ],
        };

        kirimPesanDenganKeyboardBot(chatId, text, inlineKeyboard);
      }
    });

    if (count === 0) {
      kirimPesanBot(chatId, '✅ Tidak ada data baru untuk Anda.');
    }
  } catch (error) {
    kirimPesanBot(chatId, '❌ Error: ' + error.message);
  }
}

function updateApprovalBot(no, status, userName,chatId) {
  try {
    // Dapatkan email user berdasarkan Chat ID
    const userEmail = getUserEmailByChatId(chatId);
    
    if (!userEmail) {
      return "❌ Akun Anda belum terhubung. Silakan hubungkan terlebih dahulu dengan menggunakan token dari website.";
    }

    // Logging ke master sheet
    logToMasterSheet(userEmail, 'TELEGRAM_APPROVAL', `No: ${no}, Status: ${status}, By: ${userName}`, 'Bot Telegram');
    
    // Logging ke user sheet
    logToUserSheet(userEmail, 'TELEGRAM_APPROVAL', `No: ${no}, Status: ${status}, By: ${userName}`, 'Bot Telegram');

    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ResetDevice');
    
    if (!sheet) {
      throw new Error('Sheet "Reset Device" tidak ditemukan!');
    }
    
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === no) {
        const approvedCol = 7;
        sheet.getRange(i + 1, approvedCol).setValue(`${userName}`);
        break;
      }
    }
  } catch (error) {
    throw new Error("Gagal update approval: " + error.message);
  }
}

function kirimPesanDenganPilihanDivisiBot(chatId) {
  const divisions = [
    { label: '🧾 BO', code: 'BO' },
    { label: '🏬 BSO', code: 'BSO' },
    { label: '⚙️ DMO', code: 'DMO' },
    { label: '🔧 KMO', code: 'KMO' },
    { label: '🖥️ MSO', code: 'MSO' },
    { label: '📦 PDO', code: 'PDO' },
    { label: '🧪 QC', code: 'QC' },
    { label: '🚛 PMO', code: 'PMO' },
    { label: '📡 SDO', code: 'SDO' },
    { label: '📝 Other', code: 'OTHER' },
  ];

  const keyboard = {
    inline_keyboard: divisions.map(d => [{ text: d.label, callback_data: `divisi_${d.code}` }]),
  };

  kirimPesanDenganKeyboardBot(chatId, '🏢 Silakan pilih divisi kamu:', keyboard);
}

// === UTILITY FUNCTION UNTUK KIRIM PESAN ===
function kirimPesanDenganKeyboardBot(chatId, text, keyboard) {
  const telegramUrl = `https://api.telegram.org/bot${TELEGRAM_TOKEN}`;
  const payload = {
    chat_id: chatId,
    text: text,
    reply_markup: keyboard,
    parse_mode: 'HTML'
  };
  
  UrlFetchApp.fetch(telegramUrl + '/sendMessage', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  });
}

// Fungsi untuk mendapatkan email user berdasarkan Chat ID
function getUserEmailByChatId(chatId) {
  try {
    // Coba ambil dari cache dulu
    const cache = CacheService.getUserCache();
    const userEmail = cache.get(`CHAT_${chatId}`);
    
    if (userEmail) {
      return userEmail;
    }
    
    // Jika tidak ada di cache, cari di sheet Users
    const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    const userSheet = masterSS.getSheetByName('Users');
    
    if (!userSheet) {
      return null;
    }
    
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Cari indeks kolom
    const emailCol = headers.indexOf('Email');
    const chatIdCol = headers.indexOf('ChatID');
    
    if (emailCol === -1 || chatIdCol === -1) {
      return null;
    }
    
    // Cari user dengan Chat ID yang sesuai
    for (let i = 1; i < data.length; i++) {
      if (data[i][chatIdCol] == chatId) {
        const userEmail = data[i][emailCol];
        
        // Simpan ke cache untuk akses lebih cepat di masa depan
        cache.put(`CHAT_${chatId}`, userEmail, 86400); // Cache for 24 hours
        
        return userEmail;
      }
    }
    
    return null;
    
  } catch (error) {
    Logger.log('Error in getUserEmailByChatId: ' + error.toString());
    return null;
  }
}