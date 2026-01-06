// ============================================
// TELEGRAM BOT FUNCTIONS
// ============================================

const TELEGRAM_TOKEN = '8312352125:AAFgCzZ16n9uQJiL6woitfFRA14YwPQQ324';
const LOG_SHEET_NAME = 'Logs';
const telegramUrl = `https://api.telegram.org/bot${TELEGRAM_TOKEN}`;

// Fungsi untuk memeriksa apakah user sudah terhubung dengan bot
function isUserConnected(chatId) {
  const userEmail = getUserEmailByChatId(chatId);
  return userEmail !== null;
}

// === MODIFIED doPost UNTUK HANDLE CANCEL ===
function doPost(e) {
  try {
    const update = JSON.parse(e.postData.contents);
    let chatId, userName;

    // Logging
    if (update.message) {
      chatId = update.message.chat.id;
      userName = update.message.from.first_name;
      logToMasterSheet(chatId, userName, '-', 'ReceivedMessage', update.message.text);
    } else if (update.callback_query) {
      chatId = update.callback_query.message.chat.id;
      userName = update.callback_query.from.first_name;
      logToMasterSheet(chatId, userName, '-', 'CallbackQuery', update.callback_query.data);
    }

    if (update.callback_query) {
      handleCallbackQueryBot(update.callback_query);
      return;
    }

    const message = update.message;
    const text = message.text;

    if (!text) return;

    
    // Handle connect command
    if (text.startsWith('/connect ')) {
      // Gunakan .trim() untuk menghapus spasi di awal dan akhir token
      const token = text.substring(9).trim(); 
      
      // Tambahkan log untuk debugging
      Logger.log(`[CONNECT] Received token: "${token}" from Chat ID: ${chatId}`);
      
      handleConnectCommand(chatId, token);
      return;
    }

    // Handle cancel command pertama
    if (text === '/cancel' || text === '❌ Batalkan') {
      if (handleCancelCommand(chatId)) {
        return;
      }
    }

    // Handle commands dan menu keyboard
    switch(text) {
      case '/start':
        showMainMenuBot(chatId);
        break;
      case '/resetdevice':
        startResetDeviceBot(chatId);
        break;
      case '/inputcase':
        startInputCaseBot(chatId);
        break;
      case '/lisensi':
        startInputLisensiBot(chatId);
        break;
      case '/generatepassword':
        showPasswordGeneratorBot(chatId);
        break;
      case '/whitelist':
        showWhitelistSearchBot(chatId);
        break;
      case '/listreset':
        handleCommandListResetBot(chatId);
        break;
      // Handle menu keyboard
      case '📱 Reset Device':
        startResetDeviceBot(chatId);
        break;
      case '📝 Input Case':
        startInputCaseBot(chatId);
        break;
      case '🔑 Input Lisensi':
        startInputLisensiBot(chatId);
        break;
      case '🎲 Generate Password':
        showPasswordGeneratorBot(chatId);
        break;
      case '📋 Cek Whitelist':
        showWhitelistSearchBot(chatId);
        break;
      case '📊 List Reset Terbaru':
        handleCommandListResetBot(chatId);
        break;
      case '❌ Batalkan':
        // Already handled above
        break;
      default:
        // Cek jika user dalam state tertentu
        handleUserStateBot(chatId, text, userName);
    }

  } catch (err) {
    Logger.log('Error: ' + err.stack);
    kirimPesanBot(chatId || '', '⚠️ Terjadi kesalahan sistem, coba lagi ya.');
  }
}

// Modifikasi showMainMenuBot untuk menampilkan status koneksi
function showMainMenuBot(chatId) {
  const isConnected = isUserConnected(chatId);
  let welcomeMessage = isConnected ? 
    '👋 Selamat datang kembali! Pilih menu yang diinginkan:' : 
    '👋 Halo! Pilih menu yang diinginkan:';
    
  if (!isConnected) {
    welcomeMessage += '\n\n⚠️ Akun Anda belum terhubung. Gunakan /connect <token> untuk menghubungkan akun Anda.';
  }

  const keyboard = {
    keyboard: [
      ['📱 Reset Device', '📝 Input Case'],
      ['🔑 Input Lisensi', '🎲 Generate Password'],
      ['📋 Cek Whitelist', '📊 List Reset Terbaru']
    ],
    resize_keyboard: true,
    one_time_keyboard: false
  };

  kirimPesanBot(chatId, welcomeMessage, keyboard);
}

// Modifikasi fungsi start untuk mengecek koneksi
function startResetDeviceBot(chatId) {
  if (!isUserConnected(chatId)) {
    kirimPesanBot(chatId, '❌ Akun Anda belum terhubung. Silakan hubungkan terlebih dahulu dengan menggunakan token dari website.\n\nGunakan /connect <token> untuk menghubungkan akun Anda.');
    return;
  }
  
  clearUserStateBot(chatId);
  
  const keyboard = {
    keyboard: [
      ['❌ Batalkan']
    ],
    resize_keyboard: true,
    one_time_keyboard: true
  };
  
  kirimPesanBot(chatId, '📝 Masukkan Nama:\n\nGunakan "❌ Batalkan" untuk membatalkan input.', keyboard);
  setUserStateBot(chatId, 'ASK_NAMA_RESET', {});
}

// Modifikasi fungsi start lainnya dengan cara yang sama
function startInputCaseBot(chatId) {
  if (!isUserConnected(chatId)) {
    kirimPesanBot(chatId, '❌ Akun Anda belum terhubung. Silakan hubungkan terlebih dahulu dengan menggunakan token dari website.\n\nGunakan /connect <token> untuk menghubungkan akun Anda.');
    return;
  }
  
  clearUserStateBot(chatId);
  const keyboard = {
    inline_keyboard: [
      [
        { text: '📱 Monitoring Telegram', callback_data: 'case_telegram' },
        { text: '🚫 Blokir Spam', callback_data: 'case_spam' }
      ],
      [
        { text: '📧 Monitoring Email', callback_data: 'case_email' },
        { text: '✏️ Input Manual', callback_data: 'case_manual' }
      ]
    ]
  };

  kirimPesanDenganKeyboardBot(chatId, '📋 Pilih jenis case:', keyboard);
}

function startInputLisensiBot(chatId) {
  clearUserStateBot(chatId);
  
  const keyboard = {
    keyboard: [
      ['❌ Batalkan']
    ],
    resize_keyboard: true,
    one_time_keyboard: true
  };
  
  kirimPesanBot(chatId, '📧 Masukkan Email:\n\nGunakan "❌ Batalkan" untuk membatalkan input.', keyboard);
  setUserStateBot(chatId, 'ASK_EMAIL_LISENSI', {});
}

// === PASSWORD GENERATOR ===
function showPasswordGeneratorBot(chatId) {
  clearUserStateBot(chatId);
  const keyboard = {
    inline_keyboard: [
      [
        { text: '🔢 12 Karakter (Standard)', callback_data: 'pwd_12' },
        { text: '🔢 16 Karakter (Strong)', callback_data: 'pwd_16' }
      ],
      [
        { text: '🎲 Custom Length', callback_data: 'pwd_custom' },
        { text: '⚡ Auto Strong (13 chars)', callback_data: 'pwd_strong' }
      ]
    ]
  };

  kirimPesanDenganKeyboardBot(chatId, '🎲 Pilih jenis password:', keyboard);
}

// === WHITELIST SEARCH ===
function showWhitelistSearchBot(chatId) {
  clearUserStateBot(chatId);
  const keyboard = {
    inline_keyboard: [
      [
        { text: '🔍 Cari by Nama', callback_data: 'search_nama' },
        { text: '🔍 Cari by Informations', callback_data: 'search_info' }
      ]
    ]
  };

  kirimPesanDenganKeyboardBot(chatId, '🔍 Pilih jenis pencarian:', keyboard);
}

// Modifikasi handleUserStateBot untuk menyertakan chatId
function handleUserStateBot(chatId, text, userName) {
  const stateObj = getUserStateBot(chatId);
  
  if (!stateObj) {
    kirimPesanBot(chatId, '❌ Sesi tidak aktif. Silakan pilih menu dari keyboard atau ketik /start');
    showMainMenuBot(chatId);
    return;
  }

  const { state, data = {} } = stateObj;
  
  // Cancel keyboard untuk semua state
  const cancelKeyboard = {
    keyboard: [
      ['❌ Batalkan']
    ],
    resize_keyboard: true,
    one_time_keyboard: true
  };

  switch(state) {
    case 'ASK_NAMA_RESET':
      data.nama = text.toUpperCase();
      kirimPesanBot(chatId, `✅ Nama: ${text}\n\n🏢 Silakan pilih divisi:`, cancelKeyboard);
      kirimPesanDenganPilihanDivisiBot(chatId);
      setUserStateBot(chatId, 'ASK_DIVISI_RESET', data);
      break;

    case 'ASK_DIVISI_RESET':
      data.divisi = text.toUpperCase();
      kirimPesanBot(chatId, `✅ Divisi: ${text}\n\n💻 Masukkan Device Model:`, cancelKeyboard);
      setUserStateBot(chatId, 'ASK_DEVICE_RESET', data);
      break;

    case 'ASK_DEVICE_RESET':
      data.device = text.toUpperCase();
      kirimPesanBot(chatId, `✅ Device: ${text}\n\n🔢 Masukkan Serial Number:`, cancelKeyboard);
      setUserStateBot(chatId, 'ASK_SERIAL_RESET', data);
      break;

    case 'ASK_SERIAL_RESET':
      data.serial = text.toUpperCase();
      kirimPesanBot(chatId, `✅ Serial: ${text}\n\nMenyelesaikan input...`, cancelKeyboard);
      handleApprovalQuestionBot(chatId, data);
      break;

    case 'ASK_CASE_MANUAL':
      data.case = text;
      kirimPesanBot(chatId, `✅ Case: ${text}\n\n📝 Masukkan Keterangan:`, cancelKeyboard);
      setUserStateBot(chatId, 'ASK_KETERANGAN_CASE', data);
      break;

    case 'ASK_KETERANGAN_CASE':
      data.keterangan = text;
      if (data.keterangan == '-') {
        data.keterangan='';
      }
      kirimPesanBot(chatId, `✅ Keterangan: ${text}\n\nMenyimpan data...`);
      submitCaseDataBot(chatId, data);
      showMainMenuBot(chatId);
      break;

    case 'ASK_EMAIL_LISENSI':
      data.email = text;
      kirimPesanBot(chatId, `✅ Email: ${text}\n\n🖥️ Masukkan Hostname:`, cancelKeyboard);
      setUserStateBot(chatId, 'ASK_HOSTNAME_LISENSI', data);
      break;

    case 'ASK_HOSTNAME_LISENSI':
      data.hostname = text.toUpperCase();
      kirimPesanBot(chatId, `✅ Hostname: ${text}\n\n🔐 Masukkan Password:`, cancelKeyboard);
      setUserStateBot(chatId, 'ASK_PASSWORD_LISENSI', data);
      break;

    case 'ASK_PASSWORD_LISENSI':
      data.password = text;
      kirimPesanBot(chatId, `✅ Password: ***\n\n📅 Masukkan Tanggal Expired (YYYY-MM-DD):`, cancelKeyboard);
      setUserStateBot(chatId, 'ASK_EXPIRED_LISENSI', data);
      break;

    case 'ASK_EXPIRED_LISENSI':
      data.expired = text;
      kirimPesanBot(chatId, `✅ Expired: ${text}\n\n📝 Masukkan Keterangan:`, cancelKeyboard);
      setUserStateBot(chatId, 'ASK_KETERANGAN_LISENSI', data);
      break;

    case 'ASK_KETERANGAN_LISENSI':
      data.keterangan = text;
      kirimPesanBot(chatId, `✅ Keterangan: ${text}\n\nMenyimpan data...`, cancelKeyboard);
      submitLisensiDataBot(chatId, data);
      break;

    case 'SEARCH_WHITELIST':
      searchAndSendResultsBot(chatId, data.searchType, text);
      break;

    case 'ASK_PASSWORD_LENGTH':
      const length = parseInt(text);
      if (length >= 4 && length <= 50) {
        generateAndSendPasswordBot(chatId, length, true, true, true);
        clearUserStateBot(chatId);
      } else {
        kirimPesanBot(chatId, '❌ Panjang password harus antara 4-50 karakter.', cancelKeyboard);
      }
      break;

    default:
      kirimPesanBot(chatId, '❌ State tidak dikenali. Silakan mulai ulang dengan /start');
      clearUserStateBot(chatId);
  }
}

// === CALLBACK HANDLER ===
function handleCallbackQueryBot(callback) {

  const chatId = callback.message.chat.id;
  const data = callback.data;
  const callbackId = callback.id;
  const userName = callback.from.first_name;

  // Handle existing callbacks from reset device
  if (data.startsWith('approve_') && /\d+$/.test(data)) {
    const no = parseInt(data.split('_')[1]);
    updateApprovalBot(no, 'Approved ✅', userName,chatId);
    kirimPesanBot(chatId, `✅ Data nomor ${no} telah Approved oleh ${userName}.`);
    jawabCallbackQueryBot(callbackId);
    return;
  }

  if (data.startsWith('reject_') && /\d+$/.test(data)) {
    const no = parseInt(data.split('_')[1]);
    updateApprovalBot(no, 'Rejected ❌', userName,chatId);
    kirimPesanBot(chatId, `❌ Data nomor ${no} telah Rejected oleh ${userName}.`);
    jawabCallbackQueryBot(callbackId);
    return;
  }

  if (data.startsWith('divisi_')) {
    const selectedDivisi = data.split('_')[1];
    const userState = getUserStateBot(chatId);
    if (!userState) {
      kirimPesanBot(chatId, '❌ Sesi tidak aktif. Silakan mulai ulang dengan /resetdevice');
      jawabCallbackQueryBot(callbackId);
      return;
    }

    const currentData = userState.data;
    if (selectedDivisi === 'OTHER') {
      kirimPesanBot(chatId, '✏️ Silakan ketik nama divisi lain secara manual:');
      setUserStateBot(chatId, 'ASK_DIVISI_RESET', currentData);
    } else {
      currentData.divisi = selectedDivisi.toUpperCase();
      kirimPesanBot(chatId, '💻 Masukkan device model:');
      setUserStateBot(chatId, 'ASK_DEVICE_RESET', currentData);
    }
    jawabCallbackQueryBot(callbackId);
    return;
  }

  // Handle callback untuk approve dari form reset device
  if (data === 'approve_YA' || data === 'approve_TIDAK') {
    const userState = getUserStateBot(chatId);
    if (!userState) {
      kirimPesanBot(chatId, '❌ Sesi tidak aktif. Silakan mulai ulang dengan /resetdevice');
      jawabCallbackQueryBot(callbackId);
      return;
    }

    const currentData = userState.data;
    currentData.approved = data === 'approve_YA' ? userName : 'Pending';

    try {
      const userEmail = getUserEmailByChatId(chatId);
      
      if (!userEmail) {
        kirimPesanBot(chatId, "❌ Akun Anda belum terhubung. Silakan hubungkan terlebih dahulu dengan menggunakan token dari website.");
        return;
      }
      
      // Logging ke master sheet
      logToMasterSheet(userEmail, 'TELEGRAM_INPUT_RESET_DEVICE', `Name: ${data.nama}, Device: ${data.device}`, 'Bot Telegram');
      
      // Logging ke user sheet
      logToUserSheet(userEmail, 'TELEGRAM_INPUT_RESET_DEVICE', `Name: ${data.nama}, Device: ${data.device}`, 'Bot Telegram');

      const result = tulisDataResetDeviceBot(currentData,chatId);
      kirimPesanBot(chatId, result);
      clearUserStateBot(chatId);
    } catch (error) {
      kirimPesanBot(chatId, '❌ Gagal menyimpan data: ' + error.message);
    }
    
    jawabCallbackQueryBot(callbackId);
    return;
  }

  // New callbacks for bot features
  switch(data) {
    case 'case_telegram':
      submitCaseDataBot(chatId, { case: 'Monitoring Telegram', keterangan: '' });
      break;

    case 'case_spam':
      submitCaseDataBot(chatId, { case: 'Blokir Spam', keterangan: '' });
      break;

    case 'case_email':
      submitCaseDataBot(chatId, { case: 'Monitoring Email', keterangan: '' });
      break;

    case 'case_manual':
      setUserStateBot(chatId, 'ASK_CASE_MANUAL', {});
      kirimPesanBot(chatId, '✏️ Masukkan case manual:');
      break;

    case 'pwd_12':
      generateAndSendPasswordBot(chatId, 12, true, true, true);
      break;

    case 'pwd_16':
      generateAndSendPasswordBot(chatId, 16, true, true, true);
      break;

    case 'pwd_strong':
      generateAndSendPasswordBot(chatId, 13, true, true, true);
      break;

    case 'pwd_custom':
      setUserStateBot(chatId, 'ASK_PASSWORD_LENGTH', {});
      kirimPesanBot(chatId, '🔢 Masukkan panjang password (4-50):');
      break;

    case 'search_nama':
      setUserStateBot(chatId, 'SEARCH_WHITELIST', { searchType: 'Nama' });
      kirimPesanBot(chatId, '🔍 Masukkan kata kunci nama:');
      break;

    case 'search_info':
      setUserStateBot(chatId, 'SEARCH_WHITELIST', { searchType: 'Informations' });
      kirimPesanBot(chatId, '🔍 Masukkan kata kunci informasi:');
      break;
  }

  jawabCallbackQueryBot(callbackId);
}

// === SUPPORT FUNCTIONS ===
function generateAndSendPasswordBot(chatId, length, useNumbers, useMixCase, useSymbols) {
  const password = generatePasswordBot(length, useNumbers, useMixCase, useSymbols);
  const message = `🔐 Password ${length} karakter:\n\n<code>${password}</code>\n\nKlik teks untuk copy.`;
  kirimPesanBot(chatId, message);
}

// Modifikasi submitCaseDataBot untuk menyertakan chatId
function submitCaseDataBot(chatId, data) {
  try {
    const userEmail = getUserEmailByChatId(chatId);
    
    if (!userEmail) {
      kirimPesanBot(chatId, "❌ Akun Anda belum terhubung. Silakan hubungkan terlebih dahulu dengan menggunakan token dari website.");
      return;
    }
    
    // Logging ke master sheet
    logToMasterSheet(userEmail, 'TELEGRAM_INPUT_CASE', `Case: ${data.case}, Keterangan: ${data.keterangan || ''}`, 'Bot Telegram');
    
    // Logging ke user sheet
    logToUserSheet(userEmail, 'TELEGRAM_INPUT_CASE', `Case: ${data.case}, Keterangan: ${data.keterangan || ''}`, 'Bot Telegram');
    
    const result = tulisNamaKeSheetBot(data.case, data.keterangan || '', chatId);
    kirimPesanBot(chatId, result);
    clearUserStateBot(chatId);
  } catch (error) {
    Logger.log('Error in submitCaseDataBot: ' + error.toString());
    kirimPesanBot(chatId, '❌ Gagal menyimpan case: ' + error.message);
  }
}

// Modifikasi submitLisensiDataBot untuk menyertakan chatId
function submitLisensiDataBot(chatId, data) {
  try {
    const result = tulisLisensiBot(data.email, data.hostname, data.password, data.expired, data.keterangan, chatId);
    kirimPesanBot(chatId, result);
    clearUserStateBot(chatId);
  } catch (error) {
    Logger.log('Error in submitLisensiDataBot: ' + error.toString());
    kirimPesanBot(chatId, '❌ Gagal menyimpan lisensi: ' + error.message);
  }
}

// Modifikasi searchAndSendResultsBot untuk menyertakan chatId
function searchAndSendResultsBot(chatId, searchType, keyword) {
  try {
    if (keyword.length < 2) {
      kirimPesanBot(chatId, '❌ Masukkan minimal 2 karakter untuk pencarian.');
      return;
    }

    const results = searchDataByColumnBot(searchType, keyword, chatId);
    
    if (results.length === 0) {
      kirimPesanBot(chatId, '❌ Tidak ada data ditemukan.');
      return;
    }

    let message = `🔍 Hasil pencarian "${keyword}" di ${searchType}:\n\n`;
    results.slice(0, 5).forEach((item, index) => {
      message += `${index + 1}. ${item.Nama || item.Informations}\n`;
      if (item.Keterangan) message += `   📝 ${item.Keterangan}\n`;
      message += '\n';
    });

    if (results.length > 5) {
      message += `📊 ... dan ${results.length - 5} data lainnya.`;
    }

    kirimPesanBot(chatId, message);
    clearUserStateBot(chatId);
  } catch (error) {
    Logger.log('Error in searchAndSendResultsBot: ' + error.toString());
    kirimPesanBot(chatId, '❌ Gagal mencari data: ' + error.message);
  }
}

// Modifikasi handleApprovalQuestionBot untuk menyertakan chatId
function handleApprovalQuestionBot(chatId, data) {
  const approveKeyboard = {
    inline_keyboard: [
      [
        { text: '✅ Ya', callback_data: 'approve_YA' },
        { text: '❌ Tidak', callback_data: 'approve_TIDAK' },
      ],
    ],
  };

  const message = `📋 Konfirmasi Data:\n\n` +
                 `👤 Nama: ${data.nama}\n` +
                 `🏢 Divisi: ${data.divisi}\n` +
                 `💻 Device: ${data.device}\n` +
                 `🔢 Serial: ${data.serial}\n\n` +
                 `✅ Apakah sudah di-approve?`;

  kirimPesanDenganKeyboardBot(chatId, message, approveKeyboard);
  setUserStateBot(chatId, 'WAIT_APPROVE_CHOICE', data);
}

// === UTILITY FUNCTIONS ===
function kirimPesanBot(chatId, text, keyboard = null) {
  const url = `${telegramUrl}/sendMessage`;
  const payload = { 
    chat_id: chatId, 
    text: text,
    parse_mode: 'HTML'
  };
  
  if (keyboard) {
    payload.reply_markup = keyboard;
  }
  
  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    });
  } catch (error) {
    Logger.log('Error kirim pesan: ' + error);
  }
}

function kirimPesanDenganKeyboardBot(chatId, text, keyboard) {
  kirimPesanBot(chatId, text, keyboard);
}

function jawabCallbackQueryBot(callbackId) {
  UrlFetchApp.fetch(`${telegramUrl}/answerCallbackQuery?callback_query_id=${callbackId}`);
}

// === SESSION MANAGEMENT WITH TIMEOUT ===
const SESSION_TIMEOUT = 30 * 60 * 1000; // 30 menit

function setUserStateBot(chatId, state, data = {}) {
  const cache = CacheService.getUserCache();
  const info = { 
    state: state, 
    data: data,
    timestamp: new Date().getTime()
  };
  cache.put(String(chatId), JSON.stringify(info), 1800); // 30 menit expiry
}

function getUserStateBot(chatId) {
  try {
    // Dapatkan email user berdasarkan Chat ID
    const userEmail = getUserEmailByChatId(chatId);
    
    if (!userEmail) {
      return "❌ Akun Anda belum terhubung. Silakan hubungkan terlebih dahulu dengan menggunakan token dari website.";
    }
  
    const cache = CacheService.getUserCache();
    const raw = cache.get(String(chatId));
    if (!raw) return null;
    
    const stateObj = JSON.parse(raw);
    
    // Cek session timeout
    const now = new Date().getTime();
    if (now - stateObj.timestamp > SESSION_TIMEOUT) {
      cache.remove(String(chatId));
      kirimPesanBot(chatId, '⏰ Sesi input telah kadaluarsa (10 menit). Silakan mulai ulang.');
      return null;
    }
    
    // Update timestamp untuk extend session
    stateObj.timestamp = now;
    cache.put(String(chatId), JSON.stringify(stateObj), 1800);
    
    return stateObj;
  } catch (error) {
    return null;
  }
}

// === CANCEL COMMAND ===
function handleCancelCommand(chatId) {
  const stateObj = getUserStateBot(chatId);
  if (stateObj) {
    clearUserStateBot(chatId);
    kirimPesanBot(chatId, '❌ Input dibatalkan. Kembali ke menu utama.');
    showMainMenuBot(chatId);
    return true;
  }
  return false;
}

function clearUserStateBot(chatId) {
  const cache = CacheService.getUserCache();
  cache.remove(String(chatId));
}

function logDebugBot(chatId, username, state, action, message, error = '') {
  try {
    // Dapatkan email user berdasarkan Chat ID
    const userEmail = getUserEmailByChatId(chatId);
    
    if (!userEmail) {
      return "❌ Akun Anda belum terhubung. Silakan hubungkan terlebih dahulu dengan menggunakan token dari website.";
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    if (!sheet) return;
    const timestamp = new Date();
    sheet.appendRow([timestamp, chatId, username || '-', state || '-', action || '-', message || '-', error || '-']);
  } catch (err) {
    Logger.log('Gagal menulis log: ' + err);
  }
}

// Tambahkan ini di TelegramBot.txt
function handleConnectCommand(chatId, token) {
  try {
    // Cari user berdasarkan token
    const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    const userSheet = masterSS.getSheetByName('Users');
    
    if (!userSheet) {
      kirimPesanBot(chatId, '❌ Terjadi kesalahan sistem. Silakan coba lagi nanti.');
      return;
    }
    
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Gunakan findColumnIndex untuk pencarian kolom yang lebih aman
    const emailCol = findColumnIndex(headers, ['Email', 'email', 'EMAIL']);
    const telegramTokenCol = findColumnIndex(headers, ['TelegramToken', 'Telegram Token', 'TelegramID', 'telegramId']);
    const chatIdCol = findColumnIndex(headers, ['ChatID', 'Chat Id', 'chatid']);
    const fullNameCol = findColumnIndex(headers, ['FullName', 'Full Name', 'fullName', 'Nama']);
    
    if (emailCol === -1 || telegramTokenCol === -1 || chatIdCol === -1) {
      Logger.log('[CONNECT] ERROR: Required columns not found in Users sheet.');
      kirimPesanBot(chatId, '❌ Terjadi kesalahan sistem. Struktur database tidak valid.');
      return;
    }
    
    // Buat perbandingan menjadi tidak case-sensitive
    const inputToken = token.toLowerCase();

    let userFound = false;
    for (let i = 1; i < data.length; i++) {
      const storedTokenValue = data[i][telegramTokenCol];
      // Pastikan nilai dari sel diubah ke string dan huruf kecil sebelum dibandingkan
      const storedToken = storedTokenValue ? storedTokenValue.toString().toLowerCase() : '';
      
      Logger.log(`[CONNECT] Comparing with row ${i+1}: Input="${inputToken}", Stored="${storedToken}"`);
      
      if (storedToken === inputToken) {
        const userEmail = data[i][emailCol];
        const fullName = data[i][fullNameCol] || userEmail.split('@')[0];
        
        // Update Chat ID di sheet
        userSheet.getRange(i + 1, chatIdCol + 1).setValue(chatId);
        
        // Simpan mapping Chat ID ke Email di cache untuk akses lebih cepat
        const cache = CacheService.getUserCache();
        cache.put(`CHAT_${chatId}`, userEmail, 86400); // Cache for 24 hours
        
        // Log aktivitas
        logActivity(userEmail, 'TELEGRAM_CONNECTED', `Chat ID: ${chatId}`);
        
        Logger.log(`[CONNECT] SUCCESS: Chat ID ${chatId} connected to ${userEmail}`);
        
        // Kirim pesan konfirmasi
        kirimPesanBot(chatId, `✅ Akun Telegram Anda berhasil terhubung dengan ${fullName} (${userEmail})\n\nSekarang Anda dapat menggunakan semua fitur bot.`);
        
        // Tampilkan menu utama
        showMainMenuBot(chatId);
        userFound = true;
        break; // Keluar dari loop setelah user ditemukan
      }
    }
    
    // Jika token tidak ditemukan setelah loop selesai
    if (!userFound) {
      Logger.log(`[CONNECT] FAILED: Token "${token}" not found in the database.`);
      kirimPesanBot(chatId, '❌ Token tidak valid. Pastikan Anda memasukkan token dengan benar.');
    }
    
  } catch (error) {
    Logger.log('[CONNECT] ERROR: ' + error.toString());
    kirimPesanBot(chatId, '❌ Terjadi kesalahan sistem. Silakan coba lagi nanti.');
  }
}

// Pastikan fungsi ini ada di file yang sama, atau salin ke sini
function findColumnIndex(headers, possibleNames) {
  for (const name of possibleNames) {
    const index = headers.indexOf(name);
    if (index !== -1) return index;
  }
  return -1;
}

function generateTelegramToken(userEmail) {
  try {
    // Generate token unik
    const token = Utilities.getUuid();
    Logger.log(`[TOKEN GEN] Generated token for ${userEmail}: ${token}`);
    
    // Simpan token ke database user
    const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    const userSheet = masterSS.getSheetByName('Users');
    
    if (!userSheet) {
      return {success: false, message: 'Sheet Users tidak ditemukan'};
    }
    
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Gunakan findColumnIndex untuk pencarian kolom yang lebih aman
    const emailCol = findColumnIndex(headers, ['Email', 'email', 'EMAIL']);
    const telegramTokenCol = findColumnIndex(headers, ['TelegramToken', 'Telegram Token', 'TelegramID', 'telegramId']);
    
    if (emailCol === -1 || telegramTokenCol === -1) {
      return {success: false, message: 'Struktur sheet tidak valid, kolom Email atau TelegramToken tidak ditemukan'};
    }
    
    // Update token untuk user
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailCol] === userEmail) {
        userSheet.getRange(i + 1, telegramTokenCol + 1).setValue(token);
        
        // Log aktivitas
        logActivity(userEmail, 'GENERATE_TELEGRAM_TOKEN', `Token: ${token}`);
        
        return {
          success: true,
          token: token,
          botUsername: CONFIG.BOT_USERNAME,
          message: 'Token berhasil dibuat'
        };
      }
    }
    
    return {success: false, message: 'User tidak ditemukan'};
    
  } catch (error) {
    logActivity(userEmail, 'GENERATE_TELEGRAM_TOKEN_ERROR', error.toString());
    return {
      success: false,
      message: 'Error: ' + error.message
    };
  }
}