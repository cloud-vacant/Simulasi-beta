// ============================================
// MAIN CONFIGURATION
// ============================================
const CONFIG = {
  MASTER_SPREADSHEET_ID: '180H7RpJnQ5-VgmDs7o7QU061EvrFQDqYXYrqSS0hTU4', // Ganti dengan ID Master Spreadsheet
  TEMPLATE_SPREADSHEET_ID: '1pIfFt2QQYia6trFAUEZXmJMqz0oxnqbUeYLAMIcHELc', // Ganti dengan ID Template Spreadsheet
  ADMIN_EMAIL: 'dbcloudvacant@gmail.com',
  ADMIN_CHAT_ID: '1979054564',
  TELEGRAM_TOKEN: '8312352125:AAFgCzZ16n9uQJiL6woitfFRA14YwPQQ324',
  BOT_USERNAME: 'Nvernan_bot',
  LOG_SHEET_NAME: 'Logs',
  USER_SHEET_NAME: 'Users'
};

/**
 * Konfigurasi Template
 * Definisikan semua template yang Anda miliki di sini.
 */
const TEMPLATE_CONFIG = {
  LAPORAN_HARIAN: {
    SPREADSHEET_ID: '1pIfFt2QQYia6trFAUEZXmJMqz0oxnqbUeYLAMIcHELc', // ID template lama
    SHEET_NAME: 'Report',
    TABLE_START_ROW: 3, // Tabel laporan harian mulai dari baris 3
    FOOTER_MARKER_TEXT: 'BATAS PENGESAHAN',
    formattingRules: {
      // Tanggal ada di kolom B (indeks 2)
      dateFormat: { column: 2, format: 'dd/mm/yyyy' }, 
      // Rata tengah untuk kolom No (A), Date (B), dan Progress (F)
      centerColumns: [1, 2, 3, 6],
      // Kolom Progress (E) diformat sebagai persentase
      // percentColumn: { column: 5, format: '0%' }
      // --- ATURAN BARU UNTUK MERGE ---
      // Gabungkan 2 kolom dimulai dari kolom D (indeks 4)
      mergeColumns: { startColumn: 4, numColumns: 2 } 
    }
  },
  RESET_DEVICE: {
    SPREADSHEET_ID: '1pIfFt2QQYia6trFAUEZXmJMqz0oxnqbUeYLAMIcHELc', // Ganti dengan ID template baru
    SHEET_NAME: 'LapReset', // Nama sheet di template baru
    TABLE_START_ROW: 3, // Sesuaikan jika tabel di template baru mulai dari baris 2
    FOOTER_MARKER_TEXT: 'BATAS PENGESAHAN',
    formattingRules: {
      // Tanggal ada di kolom B (indeks 2)
      dateFormat: { column: 3, format: 'dd/mm/yyyy' }, 
      // Rata tengah untuk kolom No (A), Date (B), dan Progress (F)
      centerColumns: [1, 2, 3, 4, 5, 6, 7]
      // Kolom Progress (E) diformat sebagai persentase
      // percentColumn: { column: 5, format: '0%' } 
    }
  }
};

// ============================================
// UTILITY FUNCTIONS
// ============================================
// Di Main.txt, tambahkan fungsi berikut:
/**
 * Get bot configuration for frontend
 * @return {object} Bot configuration
 */
function getBotConfig() {
  return {
    botUsername: CONFIG.BOT_USERNAME,
    telegramUrl: `https://t.me/${CONFIG.BOT_USERNAME}`
  };
}

// ============================================
// WEB APP ENTRY POINT
// ============================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('IT HELPDESK')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// ============================================
// INCLUDES (untuk memisahkan fungsi ke file lain)
// ============================================
// Fungsi-fungsi lain akan diletakkan di file terpisah dan dipanggil di sini.
// Untuk memudahkan, kita akan menggabungkan semua fungsi dalam satu project.
// Namun, dalam prakteknya, Anda dapat memisahkan ke file terpisah.