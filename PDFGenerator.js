/**
 * FUNSI UTAMA (Versi Fleksibel): Menghasilkan PDF dari template mana pun.
 * @param {string} userEmail Email pengguna.
 * @param {object} dataToInsert Object berisi semua data untuk placeholder.
 * @param {object} templateConfig Konfigurasi template yang akan digunakan.
 * @param {Array<Array>} laporanData Data tabel yang akan dimasukkan.
 * @returns {string} URL dari file PDF yang berhasil dibuat.
 */
function generatePdfFromTemplate(userEmail, dataToInsert, templateConfig, laporanData) {
  let tempSheet;
  let filePrefix;
  try {
    logActivity(userEmail, 'GENERATE_PDF', `Memulai generate PDF dengan template: ${templateConfig.SHEET_NAME}`);
    Logger.log('Data untuk placeholder: ' + JSON.stringify(dataToInsert));

    const templateSpreadsheet = SpreadsheetApp.openById(templateConfig.SPREADSHEET_ID);
    const templateSheet = templateSpreadsheet.getSheetByName(templateConfig.SHEET_NAME);
    if (!templateSheet) {
      throw new Error(`Sheet template "${templateConfig.SHEET_NAME}" tidak ditemukan.`);
    }

    const tempSheetName = 'Temp_' + new Date().getTime();
    tempSheet = templateSheet.copyTo(templateSpreadsheet).setName(tempSheetName);

    // --- ISI DATA DINAMIS (TABEL) ---
    const footerFinder = tempSheet.createTextFinder(templateConfig.FOOTER_MARKER_TEXT);
    const footerCell = footerFinder.findNext();
    if (!footerCell) throw new Error(`Marker "${templateConfig.FOOTER_MARKER_TEXT}" tidak ditemukan di template.`);
    
    const footerRow = footerCell.getRow();
    const numDataRows = laporanData.length;
    const numPlaceholderRows = footerRow - templateConfig.TABLE_START_ROW;
    const rowDifference = numDataRows - numPlaceholderRows;

    if (rowDifference > 0) {
      tempSheet.insertRowsBefore(templateConfig.TABLE_START_ROW, rowDifference);
    } else if (rowDifference < 0) {
      tempSheet.deleteRows(templateConfig.TABLE_START_ROW + numDataRows, Math.abs(rowDifference));
    }

    // --- SALIN FORMAT DARI BARIS TEMPLATE ---
    // PERBAIKAN: Hanya salin format jika ada lebih dari 1 baris data
    // Jika numDataRows = 1, tidak ada baris baru yang perlu diformat
    if (numDataRows > 1) {
      const templateRowRange = tempSheet.getRange(templateConfig.TABLE_START_ROW, 1, 1, tempSheet.getLastColumn());
      // Salin format ke baris kedua hingga baris terakhir
      const newRowsRange = tempSheet.getRange(templateConfig.TABLE_START_ROW + 1, 1, numDataRows - 1, tempSheet.getLastColumn());
      templateRowRange.copyTo(newRowsRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }
    
    const dataRange = tempSheet.getRange(templateConfig.TABLE_START_ROW, 1, numDataRows, laporanData[0].length);
    dataRange.setValues(laporanData);

    // --- TERAPKAN FORMAT KHUSUS (DINAMIS BERDASARKAN KONFIGURASI) ---
    if (numDataRows > 0 && templateConfig.formattingRules) {
      const rules = templateConfig.formattingRules;

      // Terapkan format tanggal jika aturannya ada
      if (rules.dateFormat) {
        const dateCol = rules.dateFormat.column;
        const dateFmt = rules.dateFormat.format;
        tempSheet.getRange(templateConfig.TABLE_START_ROW, dateCol, numDataRows).setNumberFormat(dateFmt);
        Logger.log(`Format tanggal ${dateFmt} diterapkan ke kolom ${dateCol}.`);
      }

      // Terapkan perataan tengah jika aturannya ada
      if (rules.centerColumns && Array.isArray(rules.centerColumns)) {
        rules.centerColumns.forEach(col => {
          tempSheet.getRange(templateConfig.TABLE_START_ROW, col, numDataRows, 1).setHorizontalAlignment('center');
        });
        Logger.log(`Perataan tengah diterapkan ke kolom: ${rules.centerColumns.join(', ')}.`);
      }
      
      // Terapkan format persentase jika aturannya ada
      // if (rules.percentColumn) {
      //   const percentCol = rules.percentColumn.column;
      //   const percentFmt = rules.percentColumn.format;
      //   tempSheet.getRange(templateConfig.TABLE_START_ROW, percentCol, numDataRows).setNumberFormat(percentFmt);
      //   Logger.log(`Format persentase ${percentFmt} diterapkan ke kolom ${percentCol}.`);
      // }

      // --- TERAPKAN MERGE CELL (ATURAN BARU) ---
      if (rules.mergeColumns) {
        const startCol = rules.mergeColumns.startColumn;
        const numCols = rules.mergeColumns.numColumns;
        Logger.log(`Menerapkan merge dari kolom ${startCol} selama ${numCols} kolom untuk ${numDataRows} baris.`);
        
        // Loop untuk setiap baris data dan lakukan merge
        for (let i = 0; i < numDataRows; i++) {
          const rowToMerge = templateConfig.TABLE_START_ROW + i;
          tempSheet.getRange(rowToMerge, startCol, 1, numCols).merge();
        }
      }
    }

    // --- GANTI PLACEHOLDER TEKS ---
    const fullRange = tempSheet.getDataRange();
    const fullData = fullRange.getValues();
    for (let i = 0; i < fullData.length; i++) {
      for (let j = 0; j < fullData[i].length; j++) {
        let cellValue = fullData[i][j];
        if (typeof cellValue === 'string') {
          const originalValue = cellValue;
          for (const key in dataToInsert) {
            const placeholder = '{' + key + '}';
            const regex = new RegExp(placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
            cellValue = cellValue.replace(regex, dataToInsert[key]);
          }
          if (cellValue !== originalValue) {
            Logger.log(`Placeholder diganti di [Baris ${i+1}, Kolom ${j+1}]: "${originalValue}" -> "${cellValue}"`);
          }
          fullData[i][j] = cellValue;
        }
      }
    }
    fullRange.setValues(fullData);

    // --- EKSPOR KE PDF ---
    SpreadsheetApp.flush(); 

    const sheetId = tempSheet.getSheetId();
    const url = `https://docs.google.com/spreadsheets/d/${templateSpreadsheet.getId()}/export?`;
    const exportOptions =
      'exportFormat=pdf&format=pdf&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false&gid=' + sheetId;
    
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url + exportOptions, {
      headers: { 'Authorization': 'Bearer ' + token }
    });

    if (templateConfig.SHEET_NAME === 'Report') {
      filePrefix = 'Laporan_Harian';
    } else if (templateConfig.SHEET_NAME === 'LapReset') {
      filePrefix = 'Laporan_ResetDevice';
    } else {
      filePrefix = 'Laporan'; // fallback aman
    }

    const blob = response.getBlob().setName(
      `${filePrefix}_${dataToInsert.username}_${Utilities.formatDate(
        new Date(),
        'GMT+7',
        'yyyy-MM-dd'
      )}.pdf`
    );
    
    const pdfFile = DriveApp.createFile(blob);
    pdfFile.addViewer(userEmail);

    templateSpreadsheet.deleteSheet(tempSheet);

    const pdfUrl = pdfFile.getUrl();
    logActivity(userEmail, 'GENERATE_PDF_SUCCESS', `URL: ${pdfUrl}`);
    return pdfUrl;

  } catch (error) {
    if (tempSheet) {
      try {
        templateSpreadsheet.deleteSheet(tempSheet);
      } catch (e) {
        logActivity(userEmail, 'CLEANUP_ERROR', `Gagal menghapus sheet sementara: ${e.toString()}`);
      }
    }
    logActivity(userEmail, 'GENERATE_PDF_ERROR', error.toString());
    throw new Error('Gagal generate PDF: ' + error.message);
  }
}

/**
 * FUNSI PEMANGIL (WRAPPER) untuk format Laporan Harian.
 * Ini adalah fungsi yang Anda panggil dari frontend untuk format yang sudah ada.
 */
function generatePdfLaporan(userEmail, fullName, assessor1, assessor2, customDate = null) {
  const userSS = openUserSpreadsheet(userEmail);
  const dataSheet = userSS.getSheetByName('Formulir');
  if (!dataSheet) throw new Error('Sheet "Formulir" tidak ditemukan.');
  const data = dataSheet.getDataRange().getValues();
  if (data.length <= 1) throw new Error('Tidak ada data laporan untuk dibuatkan PDF.');
  
  const laporanData = data.slice(1).map(row => [row[0], row[1], row[2], row[3], row[4], row[5], row[6]]);

  let reportDate = customDate ? new Date(customDate + 'T00:00:00') : new Date();
  const dataToInsert = {
    username: fullName, assessor1: assessor1, assessor2: assessor2,
    tanggal: Utilities.formatDate(reportDate, 'GMT+7', 'dd MMMM yyyy')
  };

  return generatePdfFromTemplate(userEmail, dataToInsert, TEMPLATE_CONFIG.LAPORAN_HARIAN, laporanData);
}

/**
 * FUNSI PEMANGIL (WRAPPER) untuk format Laporan Reset Device.
 * Ini adalah fungsi baru yang akan Anda panggil dari frontend.
 */
/**
 * FUNSI PEMANGIL (WRAPPER) untuk format Laporan Reset Device.
 * Ini adalah fungsi baru yang akan Anda panggil dari frontend.
 */
function generatePdfResetDevice(userEmail, fullName, assessor1, assessor2, customDate = null, startDate = null, endDate = null) {
  // --- 1. AMBIL DATA DARI SHEET 'RESET DEVICE' ---
  const userSS = openUserSpreadsheet(userEmail);
  const dataSheet = userSS.getSheetByName('ResetDevice');
  if (!dataSheet) throw new Error('Sheet "ResetDevice" tidak ditemukan di spreadsheet pengguna.');

  const data = dataSheet.getDataRange().getValues();
  if (data.length <= 1) throw new Error('Tidak ada data reset device untuk dibuatkan PDF.');
  
  // --- 2. FILTER DATA BERDASARKAN PERIODE TANGGAL ---
  let filteredData = data.slice(1); // Lewati header
  
  if (startDate || endDate) {
    const start = startDate ? new Date(startDate) : new Date('1900-01-01');
    const end = endDate ? new Date(endDate) : new Date('2100-12-31');
    
    // Set jam ke 0 untuk perbandingan tanggal saja
    start.setHours(0, 0, 0, 0);
    end.setHours(23, 59, 59, 999);
    
    filteredData = filteredData.filter(row => {
      const rowDate = row[2]; // Kolom C adalah Tanggal
      if (!rowDate || !(rowDate instanceof Date)) return false;
      
      return rowDate >= start && rowDate <= end;
    });
    
    if (filteredData.length === 0) {
      throw new Error('Tidak ada data reset device dalam periode yang dipilih.');
    }
  }
  
  // --- 3. BUAT ULANG DATA DENGAN NOMOR URUT YANG BARU (AUTO-INCREMENT) ---
  // Ini adalah bagian terpenting. Kita mengabaikan nomor lama dari sheet (row[0])
  // dan membuat nomor baru yang berurutan (index + 1).
  const laporanData = filteredData.map((row, index) => {
    return [
      index + 1, // Nomor urut baru, dimulai dari 1
      row[1],     // Name
      row[2],     // Date
      row[3],     // Divisi
      row[4],     // Device Model
      row[5],     // Serial Number
      row[6]      // Action By
    ];
  });

  // --- 4. SIAPKAN DATA UNTUK PLACEHOLDER (Sama seperti sebelumnya) ---
  let reportDate = customDate ? new Date(customDate + 'T00:00:00') : new Date();
  const dataToInsert = {
    username: fullName, 
    assessor1: assessor1, 
    assessor2: assessor2,
    tanggal: Utilities.formatDate(reportDate, 'GMT+7', 'dd MMMM yyyy'),
    // Tambahkan informasi periode ke placeholder
    periode: (startDate || endDate) ? 
      `${startDate ? Utilities.formatDate(new Date(startDate), 'GMT+7', 'dd MMMM yyyy') : 'Awal'} - ${endDate ? Utilities.formatDate(new Date(endDate), 'GMT+7', 'dd MMMM yyyy') : 'Sekarang'}` : 
      'Semua Data'
  };

  // --- 5. PANGGIL FUNGSI UTAMA DENGAN TEMPLATE RESET DEVICE ---
  return generatePdfFromTemplate(userEmail, dataToInsert, TEMPLATE_CONFIG.RESET_DEVICE, laporanData);
}