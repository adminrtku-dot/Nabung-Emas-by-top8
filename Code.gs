/**
 * Aplikasi Nabung Emas - Google Apps Script Backend (Database Spreadsheet)
 * File: Code.gs
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Nabung Emas - Portofolio & Target')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Mendapatkan ID Spreadsheet yang tersimpan atau membuatnya secara dinamis jika belum ada.
 */
function getOrCreateDbSpreadsheet() {
  const userProperties = PropertiesService.getUserProperties();
  let ssId = userProperties.getProperty('GOLD_DB_SS_ID');
  
  if (!ssId) {
    try {
      // Buat spreadsheet baru di Google Drive
      const ss = SpreadsheetApp.create('Nabung Emas Database - by top8');
      ssId = ss.getId();
      userProperties.setProperty('GOLD_DB_SS_ID', ssId);
      
      // Inisialisasi Sheet Settings
      const settingsSheet = ss.insertSheet('Settings');
      ss.deleteSheet(ss.getSheets()[0]); // Hapus sheet default bawaan
      settingsSheet.appendRow(['Parameter', 'Nilai']);
      settingsSheet.appendRow(['hargaBeli', '1450000']);
      settingsSheet.appendRow(['hargaBuyback', '1320000']);
      settingsSheet.appendRow(['pph22Percent', '0.45']);
      settingsSheet.appendRow(['biayaMaterai', '10000']);
      settingsSheet.appendRow(['batasMaterai', '10000000']);
      settingsSheet.appendRow(['pin', '1234']);
      settingsSheet.appendRow(['theme', 'dark']);
      settingsSheet.appendRow(['profileLabel', 'by top8']);
      
      // Inisialisasi Sheet Goals
      const goalsSheet = ss.insertSheet('Goals');
      goalsSheet.appendRow(['id', 'nama', 'targetGram', 'warna']);
      goalsSheet.appendRow(['g1', 'DP Rumah', '100', 'bg-blue-500']);
      goalsSheet.appendRow(['g2', 'Pendidikan Anak', '50', 'bg-emerald-500']);
      goalsSheet.appendRow(['g3', 'Dana Darurat', '20', 'bg-amber-500']);
      
      // Inisialisasi Sheet Transactions
      const txSheet = ss.insertSheet('Transactions');
      txSheet.appendRow(['id', 'tanggal', 'tipe', 'gram', 'hargaSatuan', 'totalRp', 'goalId', 'catatan']);
      txSheet.appendRow(['t1', '2026-05-10', 'Beli', '5.00', '1410000', '7050000', 'g3', 'Awal dana darurat']);
      txSheet.appendRow(['t2', '2026-05-15', 'Beli', '10.00', '1420000', '14200000', 'g1', 'Gaji Mei']);
      
      // Berikan style header bold
      [settingsSheet, goalsSheet, txSheet].forEach(sheet => {
        sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
      });
      
    } catch (e) {
      throw new Error("Gagal membuat database Spreadsheet: " + e.toString());
    }
  }
  return ssId;
}

/**
 * Mengambil seluruh data aplikasi dari Spreadsheet
 */
function getAppData() {
  const ssId = getOrCreateDbSpreadsheet();
  const ss = SpreadsheetApp.openById(ssId);
  
  // 1. Ambil Settings
  const settingsSheet = ss.getSheetByName('Settings');
  const settingsData = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 2).getValues();
  const settings = {};
  settingsData.forEach(row => {
    let val = row[1];
    if (row[0] !== 'theme' && row[0] !== 'pin' && row[0] !== 'profileLabel') {
      val = parseFloat(val);
    }
    settings[row[0]] = val;
  });
  
  // 2. Ambil Goals
  const goalsSheet = ss.getSheetByName('Goals');
  let goals = [];
  if (goalsSheet.getLastRow() > 1) {
    const goalsData = goalsSheet.getRange(2, 1, goalsSheet.getLastRow() - 1, 4).getValues();
    goals = goalsData.map(row => ({
      id: row[0].toString(),
      nama: row[1].toString(),
      targetGram: parseFloat(row[2]),
      warna: row[3].toString()
    }));
  }
  
  // 3. Ambil Transactions
  const txSheet = ss.getSheetByName('Transactions');
  let transactions = [];
  if (txSheet.getLastRow() > 1) {
    const txData = txSheet.getRange(2, 1, txSheet.getLastRow() - 1, 8).getValues();
    transactions = txData.map(row => {
      let tgl = row[1];
      if (tgl instanceof Date) {
        tgl = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        tgl = tgl.toString();
      }
      return {
        id: row[0].toString(),
        tanggal: tgl,
        tipe: row[2].toString(),
        gram: parseFloat(row[3]),
        hargaSatuan: parseFloat(row[4]),
        totalRp: parseFloat(row[5]),
        goalId: row[6].toString(),
        catatan: row[7].toString()
      };
    });
  }
  
  return {
    settings: settings,
    goals: goals,
    transactions: transactions,
    spreadsheetUrl: ss.getUrl()
  };
}

/**
 * Validasi kecocokan PIN Keamanan
 */
function verifyPinInternal(ss, inputPin) {
  const settingsSheet = ss.getSheetByName('Settings');
  const data = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 2).getValues();
  const pinRow = data.find(row => row[0] === 'pin');
  const actualPin = pinRow ? pinRow[1].toString() : '';
  return actualPin === inputPin.toString();
}

/**
 * Menyimpan seluruh pengaturan baru ke Spreadsheet
 */
function saveSettings(newSettings, pinOtorisasi) {
  try {
    const ssId = getOrCreateDbSpreadsheet();
    const ss = SpreadsheetApp.openById(ssId);
    
    if (!verifyPinInternal(ss, pinOtorisasi)) {
      return { success: false, message: "Otorisasi PIN Gagal! PIN tidak sesuai." };
    }
    
    const settingsSheet = ss.getSheetByName('Settings');
    settingsSheet.clearContents();
    
    settingsSheet.appendRow(['Parameter', 'Nilai']);
    settingsSheet.appendRow(['hargaBeli', newSettings.hargaBeli.toString()]);
    settingsSheet.appendRow(['hargaBuyback', newSettings.hargaBuyback.toString()]);
    settingsSheet.appendRow(['pph22Percent', newSettings.pph22Percent.toString()]);
    settingsSheet.appendRow(['biayaMaterai', newSettings.biayaMaterai.toString()]);
    settingsSheet.appendRow(['batasMaterai', newSettings.batasMaterai.toString()]);
    settingsSheet.appendRow(['pin', newSettings.pin.toString()]);
    settingsSheet.appendRow(['theme', newSettings.theme.toString()]);
    settingsSheet.appendRow(['profileLabel', newSettings.profileLabel.toString()]);
    
    settingsSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    
    return { success: true, message: "Pengaturan berhasil disinkronkan ke Spreadsheet!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Menambahkan transaksi baru ke Spreadsheet
 */
function addTransaction(tx, pinOtorisasi) {
  try {
    const ssId = getOrCreateDbSpreadsheet();
    const ss = SpreadsheetApp.openById(ssId);
    
    if (!verifyPinInternal(ss, pinOtorisasi)) {
      return { success: false, message: "Otorisasi PIN Gagal! Transaksi dibatalkan." };
    }
    
    const txSheet = ss.getSheetByName('Transactions');
    
    const newId = "tx_" + new Date().getTime();
    let gram = parseFloat(tx.gram);
    let totalRp = parseFloat(tx.totalRp);
    
    if (tx.tipe === "Jual") {
      if (gram > 0) gram = -gram;
      if (totalRp > 0) totalRp = -totalRp;
    }
    
    txSheet.appendRow([
      newId,
      tx.tanggal,
      tx.tipe,
      gram,
      parseFloat(tx.hargaSatuan),
      totalRp,
      tx.goalId || "",
      tx.catatan || ""
    ]);
    
    return { success: true, message: "Transaksi berhasil dicatat & disinkronkan!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Memperbarui transaksi yang sudah ada di Spreadsheet
 */
function updateTransaction(updatedTx, pinOtorisasi) {
  try {
    const ssId = getOrCreateDbSpreadsheet();
    const ss = SpreadsheetApp.openById(ssId);
    
    if (!verifyPinInternal(ss, pinOtorisasi)) {
      return { success: false, message: "Otorisasi PIN Gagal! Perubahan dibatalkan." };
    }
    
    const txSheet = ss.getSheetByName('Transactions');
    const rows = txSheet.getRange(2, 1, txSheet.getLastRow() - 1, 8).getValues();
    
    let rowIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0].toString() === updatedTx.id) {
        rowIndex = i + 2; 
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, message: "Catatan transaksi tidak ditemukan." };
    
    let gram = parseFloat(updatedTx.gram);
    let totalRp = parseFloat(updatedTx.totalRp);
    if (updatedTx.tipe === "Jual") {
      if (gram > 0) gram = -gram;
      if (totalRp > 0) totalRp = -totalRp;
    } else {
      if (gram < 0) gram = -gram;
      if (totalRp < 0) totalRp = -totalRp;
    }
    
    txSheet.getRange(rowIndex, 1, 1, 8).setValues([[
      updatedTx.id,
      updatedTx.tanggal,
      updatedTx.tipe,
      gram,
      parseFloat(updatedTx.hargaSatuan),
      totalRp,
      updatedTx.goalId || "",
      updatedTx.catatan || ""
    ]]);
    
    return { success: true, message: "Perubahan transaksi berhasil disinkronkan!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Menghapus transaksi dari Spreadsheet
 */
function deleteTransaction(txId, pinOtorisasi) {
  try {
    const ssId = getOrCreateDbSpreadsheet();
    const ss = SpreadsheetApp.openById(ssId);
    
    if (!verifyPinInternal(ss, pinOtorisasi)) {
      return { success: false, message: "Otorisasi PIN Gagal! Penghapusan dibatalkan." };
    }
    
    const txSheet = ss.getSheetByName('Transactions');
    const rows = txSheet.getRange(2, 1, txSheet.getLastRow() - 1, 1).getValues();
    
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0].toString() === txId) {
        txSheet.deleteRow(i + 2);
        return { success: true, message: "Transaksi berhasil dihapus dari Spreadsheet!" };
      }
    }
    return { success: false, message: "Transaksi tidak ditemukan." };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Menambahkan target baru ke Spreadsheet
 */
function addGoal(goal, pinOtorisasi) {
  try {
    const ssId = getOrCreateDbSpreadsheet();
    const ss = SpreadsheetApp.openById(ssId);
    
    if (!verifyPinInternal(ss, pinOtorisasi)) {
      return { success: false, message: "Otorisasi PIN Gagal! Target gagal dibuat." };
    }
    
    const goalsSheet = ss.getSheetByName('Goals');
    const newId = "g_" + new Date().getTime();
    
    goalsSheet.appendRow([
      newId,
      goal.nama,
      parseFloat(goal.targetGram),
      goal.warna || "bg-emerald-500"
    ]);
    
    return { success: true, message: "Target tabungan baru tersimpan di Spreadsheet!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Memperbarui target di Spreadsheet
 */
function updateGoal(updatedGoal, pinOtorisasi) {
  try {
    const ssId = getOrCreateDbSpreadsheet();
    const ss = SpreadsheetApp.openById(ssId);
    
    if (!verifyPinInternal(ss, pinOtorisasi)) {
      return { success: false, message: "Otorisasi PIN Gagal! Perubahan target dibatalkan." };
    }
    
    const goalsSheet = ss.getSheetByName('Goals');
    const rows = goalsSheet.getRange(2, 1, goalsSheet.getLastRow() - 1, 4).getValues();
    
    let rowIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0].toString() === updatedGoal.id) {
        rowIndex = i + 2;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, message: "Target tidak ditemukan." };
    
    goalsSheet.getRange(rowIndex, 1, 1, 4).setValues([[
      updatedGoal.id,
      updatedGoal.nama,
      parseFloat(updatedGoal.targetGram),
      updatedGoal.warna
    ]]);
    
    return { success: true, message: "Perubahan target sukses disinkronkan!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Menghapus target dari Spreadsheet
 */
function deleteGoal(goalId, pinOtorisasi) {
  try {
    const ssId = getOrCreateDbSpreadsheet();
    const ss = SpreadsheetApp.openById(ssId);
    
    if (!verifyPinInternal(ss, pinOtorisasi)) {
      return { success: false, message: "Otorisasi PIN Gagal!" };
    }
    
    const goalsSheet = ss.getSheetByName('Goals');
    const goalsRows = goalsSheet.getRange(2, 1, goalsSheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < goalsRows.length; i++) {
      if (goalsRows[i][0].toString() === goalId) {
        goalsSheet.deleteRow(i + 2);
        break;
      }
    }
    
    const txSheet = ss.getSheetByName('Transactions');
    if (txSheet.getLastRow() > 1) {
      const txRows = txSheet.getRange(2, 7, txSheet.getLastRow() - 1, 1).getValues(); 
      for (let j = 0; j < txRows.length; j++) {
        if (txRows[j][0].toString() === goalId) {
          txSheet.getRange(j + 2, 7).setValue("");
        }
      }
    }
    
    return { success: true, message: "Target dihapus. Seluruh transaksi terkait dialokasikan ke Umum." };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Reset data aplikasi di Google Sheet
 */
function resetAllData() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    userProperties.deleteProperty('GOLD_DB_SS_ID');
    return { success: true, message: "Aplikasi berhasil di-reset. Spreadsheet database baru akan digenerate ulang." };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
