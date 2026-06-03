/**
 * NABUNG EMAS - Google Apps Script Backend (code.gs)
 * Menghubungkan aplikasi web dengan Google Sheets secara aman menggunakan otorisasi PIN.
 * Versi Perbaikan: Otomatis mendeteksi skrip Standalone vs Bound & mencegah error null spreadsheet.
 */

const APP_TITLE = "Nabung Emas - by top8";

// 1. FUNGSI UTAMA UNTUK MENAMPILKAN APLIKASI WEB
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 2. INITIALIZE SPREADSHEET DAN SHEET SECARA AMAN (MENDUKUNG STANDALONE & BOUND)
function getSpreadsheet() {
  let ss = null;
  
  // Mencoba mengambil Spreadsheet aktif (Jika skrip tipe Bound/Melekat pada Sheet)
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    Logger.log("Bukan bound script, mencoba metode standalone.");
  }
  
  // Jika ss masih null (Skrip bertipe Standalone/Dibuat langsung dari script.google.com)
  if (!ss) {
    const props = PropertiesService.getScriptProperties();
    let ssId = props.getProperty("SPREADSHEET_ID");
    
    if (ssId) {
      try {
        ss = SpreadsheetApp.openById(ssId);
      } catch(e) {
        // Jika ID sheet sebelumnya tidak sengaja terhapus dari Google Drive, reset ID-nya
        ssId = null;
      }
    }
    
    // Jika belum ada database Sheet sama sekali, buat baru secara otomatis di Google Drive
    if (!ssId || !ss) {
      ss = SpreadsheetApp.create("Nabung Emas - Database");
      props.setProperty("SPREADSHEET_ID", ss.getId());
      Logger.log("Database spreadsheet baru berhasil dibuat secara otomatis dengan ID: " + ss.getId());
    }
  }

  if (!ss) {
    throw new Error("Gagal menginisialisasi Google Spreadsheet.");
  }

  // Pastikan Sheet "Settings" ada
  let sheetSettings = ss.getSheetByName("Settings");
  if (!sheetSettings) {
    sheetSettings = ss.insertSheet("Settings");
    // Header & Nilai Default Awal
    sheetSettings.appendRow(["Parameter", "Nilai"]);
    sheetSettings.appendRow(["hargaBeli", 2774000]);
    sheetSettings.appendRow(["hargaBuyback", 2584000]);
    sheetSettings.appendRow(["pph22Percent", 0.45]);
    sheetSettings.appendRow(["biayaMaterai", 10000]);
    sheetSettings.appendRow(["batasMaterai", 10000000]);
    sheetSettings.appendRow(["pin", "1234"]);
    sheetSettings.appendRow(["theme", "dark"]);
    sheetSettings.appendRow(["profileLabel", "by top8"]);
    sheetSettings.appendRow(["zakatNishab", 85.00]);
    sheetSettings.appendRow(["zakatRate", 2.5]);
    sheetSettings.appendRow(["zakatHargaAcuan", 2774000]);
  }

  // Pastikan Sheet "Goals" ada
  let sheetGoals = ss.getSheetByName("Goals");
  if (!sheetGoals) {
    sheetGoals = ss.insertSheet("Goals");
    sheetGoals.appendRow(["ID", "Nama Target", "Target Gram", "Warna"]);
    sheetGoals.appendRow(["g1", "Haji", 40, "bg-blue-500"]);
    sheetGoals.appendRow(["g2", "Pendidikan Anak", 40, "bg-emerald-500"]);
  }

  // Pastikan Sheet "Transactions" ada
  let sheetTransactions = ss.getSheetByName("Transactions");
  if (!sheetTransactions) {
    sheetTransactions = ss.insertSheet("Transactions");
    sheetTransactions.appendRow(["ID", "Tanggal", "Tipe", "Gram", "Harga Satuan", "Total Rp", "Goal ID", "Catatan"]);
    sheetTransactions.appendRow(["t1", "2026-05-10", "Beli", 28.00, 1758130, 49227652, "g1", "Modal Haji"]);
    sheetTransactions.appendRow(["t2", "2026-05-25", "Beli", 4.00, 2907500, 11630000, "g2", "Pendidikan Anak"]);
  }

  // Hapus sheet bawaan Google Sheets (seperti "Sheet1" atau "Sheet asli") 
  // HANYA jika sheet aplikasi kita sudah berhasil dibuat, agar tidak terjadi error "must have at least one sheet"
  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 1) {
    try {
      ss.deleteSheet(defaultSheet);
    } catch(e) {
      // Abaikan jika gagal menghapus sheet default
    }
  }

  return ss;
}

// 3. AMBIL SEMUA DATA APLIKASI (GET APP DATA DENGAN ERROR HANDLING KUAT)
function getAppData() {
  try {
    const ss = getSpreadsheet();
    
    // Ambil Settings
    const sheetSettings = ss.getSheetByName("Settings");
    const settingsRows = sheetSettings.getDataRange().getValues();
    const settings = {};
    for (let i = 1; i < settingsRows.length; i++) {
      const key = settingsRows[i][0];
      let val = settingsRows[i][1];
      
      if (key === undefined || key === "") continue; // Lewati jika baris kosong

      // Konversi angka secara aman
      if (val !== "" && !isNaN(val)) {
        val = Number(val);
      }
      settings[key] = val;
    }

    // Ambil Goals
    const sheetGoals = ss.getSheetByName("Goals");
    const goalsRows = sheetGoals.getDataRange().getValues();
    const goals = [];
    for (let i = 1; i < goalsRows.length; i++) {
      if (goalsRows[i][0] === undefined || goalsRows[i][0] === "") continue;
      goals.push({
        id: String(goalsRows[i][0]),
        nama: String(goalsRows[i][1] || ""),
        targetGram: Number(goalsRows[i][2] || 0),
        warna: String(goalsRows[i][3] || "bg-blue-500")
      });
    }

    // Ambil Transactions
    const sheetTransactions = ss.getSheetByName("Transactions");
    const txRows = sheetTransactions.getDataRange().getValues();
    const transactions = [];
    for (let i = 1; i < txRows.length; i++) {
      if (txRows[i][0] === undefined || txRows[i][0] === "") continue;
      
      let formattedDate = "";
      try {
        formattedDate = Utilities.formatDate(new Date(txRows[i][1]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      } catch(e) {
        formattedDate = String(txRows[i][1]); // Fallback jika format tanggal bermasalah
      }

      transactions.push({
        id: String(txRows[i][0]),
        tanggal: formattedDate,
        tipe: String(txRows[i][2] || "Beli"),
        gram: Number(txRows[i][3] || 0),
        hargaSatuan: Number(txRows[i][4] || 0),
        totalRp: Number(txRows[i][5] || 0),
        goalId: String(txRows[i][6] || ""),
        catatan: String(txRows[i][7] || "")
      });
    }

    return {
      settings: settings,
      goals: goals,
      transactions: transactions,
      spreadsheetUrl: ss.getUrl()
    };
  } catch (error) {
    Logger.log("Error loading data: " + error.toString());
    throw new Error("Gagal memuat data dari Spreadsheet: " + error.message);
  }
}

// 4. VALIDASI PIN KEAMANAN
function verifyPin(enteredPin) {
  const ss = getSpreadsheet();
  const sheetSettings = ss.getSheetByName("Settings");
  const rows = sheetSettings.getDataRange().getValues();
  let savedPin = "1234"; // default fallback
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === "pin") {
      savedPin = String(rows[i][1]);
      break;
    }
  }
  return enteredPin.toString() === savedPin;
}

// 5. SIMPAN KONFIGURASI SETTINGS BARU
function saveSettings(newSettings, pinVerification) {
  if (!verifyPin(pinVerification)) {
    return { success: false, message: "Otorisasi Gagal! PIN tidak valid." };
  }

  try {
    const ss = getSpreadsheet();
    const sheetSettings = ss.getSheetByName("Settings");
    
    // Bersihkan isi sheet dan tulis ulang untuk menghindari data tertinggal
    sheetSettings.clearContents();
    sheetSettings.appendRow(["Parameter", "Nilai"]);
    
    const keys = Object.keys(newSettings);
    keys.forEach(key => {
      sheetSettings.appendRow([key, newSettings[key]]);
    });

    return { success: true, message: "Parameter baru berhasil disimpan ke Google Sheet!" };
  } catch (error) {
    return { success: false, message: "Gagal menyimpan ke Sheet: " + error.toString() };
  }
}

// 6. TAMBAH CATATAN TRANSAKSI BARU
function addTransaction(tx, pinVerification) {
  if (!verifyPin(pinVerification)) {
    return { success: false, message: "Otorisasi Gagal!" };
  }

  try {
    const ss = getSpreadsheet();
    const sheetTransactions = ss.getSheetByName("Transactions");
    
    const id = "tx_" + Date.now();
    let gram = tx.gram;
    let totalRp = tx.totalRp;
    if (tx.tipe === "Jual") {
      gram = -Math.abs(tx.gram);
      totalRp = -Math.abs(tx.totalRp);
    }

    sheetTransactions.appendRow([
      id,
      tx.tanggal,
      tx.tipe,
      gram,
      tx.hargaSatuan,
      totalRp,
      tx.goalId || "",
      tx.catatan || ""
    ]);

    return { success: true, message: "Transaksi berhasil dicatat di Google Sheet!" };
  } catch (error) {
    return { success: false, message: "Gagal mencatat transaksi: " + error.toString() };
  }
}

// 7. PERBARUI TRANSAKSI
function updateTransaction(tx, pinVerification) {
  if (!verifyPin(pinVerification)) {
    return { success: false, message: "Otorisasi Gagal!" };
  }

  try {
    const ss = getSpreadsheet();
    const sheetTransactions = ss.getSheetByName("Transactions");
    const data = sheetTransactions.getDataRange().getValues();
    
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(tx.id)) {
        rowIndex = i + 1; // spreadsheet row index is 1-based
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "Data transaksi tidak ditemukan." };
    }

    let gram = tx.gram;
    let totalRp = tx.totalRp;
    if (tx.tipe === "Jual") {
      gram = -Math.abs(tx.gram);
      totalRp = -Math.abs(tx.totalRp);
    }

    sheetTransactions.getRange(rowIndex, 2, 1, 7).setValues([[
      tx.tanggal,
      tx.tipe,
      gram,
      tx.hargaSatuan,
      totalRp,
      tx.goalId || "",
      tx.catatan || ""
    ]]);

    return { success: true, message: "Catatan transaksi berhasil diperbarui!" };
  } catch (error) {
    return { success: false, message: "Gagal memperbarui transaksi: " + error.toString() };
  }
}

// 8. HAPUS TRANSAKSI
function deleteTransaction(id, pinVerification) {
  if (!verifyPin(pinVerification)) {
    return { success: false, message: "Otorisasi Gagal!" };
  }

  try {
    const ss = getSpreadsheet();
    const sheetTransactions = ss.getSheetByName("Transactions");
    const data = sheetTransactions.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheetTransactions.deleteRow(i + 1);
        return { success: true, message: "Transaksi berhasil dihapus dari Google Sheet." };
      }
    }
    return { success: false, message: "Transaksi tidak ditemukan." };
  } catch (error) {
    return { success: false, message: "Gagal menghapus transaksi: " + error.toString() };
  }
}

// 9. TAMBAH ALOKASI TARGET (GOAL) BARU
function addGoal(goal, pinVerification) {
  if (!verifyPin(pinVerification)) {
    return { success: false, message: "Otorisasi Gagal!" };
  }

  try {
    const ss = getSpreadsheet();
    const sheetGoals = ss.getSheetByName("Goals");
    const id = "g_" + Date.now();
    
    sheetGoals.appendRow([
      id,
      goal.nama,
      goal.targetGram,
      goal.warna
    ]);

    return { success: true, message: "Target tabungan baru berhasil dibuat!" };
  } catch (error) {
    return { success: false, message: "Gagal menambahkan target: " + error.toString() };
  }
}

// 10. PERBARUI ALOKASI TARGET
function updateGoal(goal, pinVerification) {
  if (!verifyPin(pinVerification)) {
    return { success: false, message: "Otorisasi Gagal!" };
  }

  try {
    const ss = getSpreadsheet();
    const sheetGoals = ss.getSheetByName("Goals");
    const data = sheetGoals.getDataRange().getValues();
    
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(goal.id)) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "Target tidak ditemukan." };
    }

    sheetGoals.getRange(rowIndex, 2, 1, 3).setValues([[
      goal.nama,
      goal.targetGram,
      goal.warna
    ]]);

    return { success: true, message: "Target tabungan berhasil diperbarui!" };
  } catch (error) {
    return { success: false, message: "Gagal memperbarui target: " + error.toString() };
  }
}

// 11. HAPUS ALOKASI TARGET DAN ALIKASIKAN KE "UMUM"
function deleteGoal(id, pinVerification) {
  if (!verifyPin(pinVerification)) {
    return { success: false, message: "Otorisasi Gagal!" };
  }

  try {
    const ss = getSpreadsheet();
    
    // 1. Hapus target di sheet Goals
    const sheetGoals = ss.getSheetByName("Goals");
    const goalsData = sheetGoals.getDataRange().getValues();
    for (let i = 1; i < goalsData.length; i++) {
      if (String(goalsData[i][0]) === String(id)) {
        sheetGoals.deleteRow(i + 1);
        break;
      }
    }

    // 2. Ubah goalId transaksi terkait ke kosong (Umum)
    const sheetTransactions = ss.getSheetByName("Transactions");
    const txData = sheetTransactions.getDataRange().getValues();
    for (let i = 1; i < txData.length; i++) {
      if (String(txData[i][6]) === String(id)) {
        sheetTransactions.getRange(i + 1, 7).setValue(""); // Kolom Goal ID dikosongkan
      }
    }

    return { success: true, message: "Target dihapus. Transaksi terkait dialihkan ke Umum." };
  } catch (error) {
    return { success: false, message: "Gagal menghapus target: " + error.toString() };
  }
}

// 12. RESET SELURUH DATA APLIKASI SECARA AMAN (TIDAK ERROR KARENA KOSONG)
function resetAllData() {
  try {
    const ss = getSpreadsheet();
    
    // Google Sheets mewajibkan minimal ada 1 sheet aktif dalam spreadsheet.
    // Kita buat sheet cadangan sementara ("Backup_Temp") sebelum menghapus yang lain.
    let tempSheet = ss.getSheetByName("Backup_Temp");
    if (!tempSheet) {
      tempSheet = ss.insertSheet("Backup_Temp");
    }
    
    // Hapus ketiga sheet utama jika eksis
    const sSettings = ss.getSheetByName("Settings");
    if (sSettings) ss.deleteSheet(sSettings);
    
    const sGoals = ss.getSheetByName("Goals");
    if (sGoals) ss.deleteSheet(sGoals);
    
    const sTransactions = ss.getSheetByName("Transactions");
    if (sTransactions) ss.deleteSheet(sTransactions);
    
    // Jalankan ulang fungsi inisialisasi untuk membuat kembali tabel bersih
    getSpreadsheet();
    
    // Setelah tabel bersih berhasil dibangun kembali, hapus sheet cadangan sementara secara aman
    if (ss.getSheetByName("Backup_Temp") && ss.getSheets().length > 1) {
      ss.deleteSheet(tempSheet);
    }
    
    return { success: true, message: "Aplikasi berhasil di-reset sepenuhnya!" };
  } catch (error) {
    return { success: false, message: "Gagal mereset data: " + error.toString() };
  }
}
