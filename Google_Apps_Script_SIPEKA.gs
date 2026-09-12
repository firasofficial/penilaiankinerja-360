/**
 * ==============================================================================
 * BACKEND GOOGLE APPS SCRIPT - SIPEKA 360° KABUPATEN ACEH TIMUR
 * ==============================================================================
 * Fitur:
 * 1. Pengiriman Kode OTP Verifikasi Email (Gratis via GmailApp / MailApp)
 * 2. Registrasi & Simpan Data Pengguna Baru (Password Hash / Plain)
 * 3. Reset & Pembaruan Password Pegawai
 * 4. Penerimaan Formulir Penilaian (Atasan, Rekan, Bawahan)
 * 
 * PANDUAN PEMASANGAN:
 * 1. Buka Google Sheets database Anda.
 * 2. Klik menu "Ekstensi" (Extensions) -> "Apps Script".
 * 3. Hapus semua kode default, lalu tempelkan (paste) seluruh isi file ini.
 * 4. Klik tombol "Simpan" (ikon disket).
 * 5. Klik tombol "Terapkan" (Deploy) -> "Penerapan Baru" (New Deployment).
 * 6. Pilih jenis: "Aplikasi Web" (Web App).
 *    - Deskripsi: SIPEKA 360 OTP & Database
 *    - Jalankan sebagai (Execute as): "Saya" (Me)
 *    - Siapa yang memiliki akses (Who has access): "Siapa saja" (Anyone)
 * 7. Klik "Terapkan" (Deploy), berikan izin akses Google (Authorize), lalu salin URL Web App yang dihasilkan.
 * ==============================================================================
 */

// Konfigurasi Nama Instansi & Pengirim Email
const APP_NAME = "SIPEKA 360°";
const INSTANSI_NAME = "Pemerintah Kabupaten Aceh Timur";
const EMAIL_SENDER_NAME = "SIPEKA 360° BKPSDM Aceh Timur";

function doGet(e) {
  try {
    const params = e.parameter || {};
    const action = params.action || "get_pegawai";

    // 1. Aksi: Ambil Master Data Pegawai dari Sheet 'Data_Pegawai'
    if (action === "get_pegawai") {
      return handleGetPegawai();
    }

    // 2. Aksi: Ambil Pengaturan Sistem (Termasuk Periode Penilaian Aktif & Daftar Periode)
    if (action === "get_config" || action === "get_periods") {
      return handleGetConfig();
    }

    // 3. Aksi: Ambil Semua Data Sheet Penilaian (Atasan, Bawahan, Rekan) untuk Unduh Admin
    if (action === "get_all_penilaian_data") {
      return handleGetAllPenilaianData();
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Layanan Google Apps Script SIPEKA 360° aktif dan siap menerima permintaan."
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const params = e.parameter || {};
    const action = params.action || "";

    // 1. Aksi: Ambil Master Data Pegawai
    if (action === "get_pegawai") {
      return handleGetPegawai();
    }

    // 2. Aksi: Ambil / Simpan Pengaturan
    if (action === "get_config" || action === "get_periods") {
      return handleGetConfig();
    }
    if (action === "save_config" || action === "set_active_period") {
      return handleSaveConfig(params);
    }
    if (action === "add_period") {
      return handleAddPeriod(params);
    }
    if (action === "edit_period" || action === "update_period") {
      return handleEditPeriod(params);
    }
    if (action === "delete_period") {
      return handleDeletePeriod(params);
    }

    // 3. Aksi: Ambil Semua Data Penilaian untuk Export/Download
    if (action === "get_all_penilaian_data") {
      return handleGetAllPenilaianData();
    }

    // 4. Aksi: Kirim Kode OTP ke Email
    if (action === "send_otp") {
      return handleSendOtp(params);
    }

    // 5. Aksi: Registrasi Pengguna Baru
    if (action === "register") {
      return handleRegisterUser(params);
    }

    // 6. Aksi: Reset Password
    if (action === "reset_password") {
      return handleResetPassword(params);
    }

    // 7. Default: Simpan Formulir Penilaian (Atasan / Rekan / Bawahan)
    return handleSavePenilaian(params);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Mengambil Seluruh Master Data Pegawai dari Sheet 'Data_Pegawai'
 */
function handleGetPegawai() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Data_Pegawai");
  
  if (!sheet) {
    sheet = ss.getSheets()[0]; // Ambil sheet pertama jika nama Data_Pegawai belum diset
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      total: 0,
      data: []
    })).setMimeType(ContentService.MimeType.JSON);
  }

  const rawHeaders = data[0];
  const headers = rawHeaders.map(h => String(h).trim());
  const result = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // Lewati baris kosong (cek apakah ada NIP atau Nama)
    const rawNip = row[1] !== undefined ? String(row[1]).trim() : "";
    const rawNama = row[2] !== undefined ? String(row[2]).trim() : "";
    if (!rawNip && !rawNama) continue;

    const item = {};
    headers.forEach((header, colIdx) => {
      let val = row[colIdx];
      if (typeof val === 'string') {
        val = val.trim().replace(/^'/, '');
      } else if (typeof val === 'number') {
        val = String(val);
      }
      item[header] = val !== undefined && val !== null ? val : "";
    });

    // Pastikan key standar 'nip' dan 'nama' selalu ada (lowercase format)
    if (!item.nip && item.NIP) item.nip = item.NIP;
    if (!item.nama && item.Nama) item.nama = item.Nama;
    if (!item.jabatan && item.Jabatan) item.jabatan = item.Jabatan;
    if (!item.golru && item.Golru) item.golru = item.Golru;
    if (!item['satuan kerja'] && item.SatuanKerja) item['satuan kerja'] = item.SatuanKerja;
    if (!item['unit kerja'] && item.UnitKerja) item['unit kerja'] = item.UnitKerja;

    result.push(item);
  }

  const config = getAppConfig();

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    total: result.length,
    config: config,
    data: result
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Mengambil Pengaturan Sistem (Termasuk Periode Penilaian)
 */
function handleGetConfig() {
  const config = getAppConfig();
  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    config: config
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Menyimpan / Memperbarui Pengaturan Sistem dari Admin
 */
function handleSaveConfig(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Pengaturan") || ss.getSheetByName("Config");
  if (!sheet) {
    sheet = ss.insertSheet("Pengaturan");
    sheet.appendRow(["Kunci_Pengaturan", "Nilai_Pengaturan", "Keterangan"]);
  }

  const periode = (params.periode_penilaian || params.periode_aktif || "Tahun 2025").trim();
  let found = false;
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toUpperCase() === "PERIODE_PENILAIAN_AKTIF") {
      sheet.getRange(i + 1, 2).setValue(periode);
      found = true;
      break;
    }
  }

  if (!found) {
    sheet.appendRow(["PERIODE_PENILAIAN_AKTIF", periode, "Periode penilaian aktif yang muncul di formulir"]);
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    message: "Pengaturan berhasil disimpan.",
    periode_aktif: periode
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Helper: Ambil Objek Config & Daftar Periode dari Sheet 'Pengaturan' / 'Periode_List'
 */
function getAppConfig() {
  let periodeAktif = "Tahun 2025";
  let periodeList = [];

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Ambil Periode Aktif dari Sheet 'Pengaturan'
    let configSheet = ss.getSheetByName("Pengaturan") || ss.getSheetByName("Config");
    if (configSheet) {
      const data = configSheet.getDataRange().getValues();
      for (let i = 0; i < data.length; i++) {
        const key = String(data[i][0] || "").trim().toUpperCase();
        const val = String(data[i][1] || "").trim();
        if (key === "PERIODE_PENILAIAN_AKTIF" || key === "PERIODE_AKTIF") {
          if (val) periodeAktif = val;
        }
      }
    }

    // 2. Ambil Daftar Periode dari Sheet 'Periode_List'
    let periodSheet = ss.getSheetByName("Periode_List");
    if (periodSheet) {
      const pData = periodSheet.getDataRange().getValues();
      for (let i = 1; i < pData.length; i++) {
        const namaPeriode = String(pData[i][0] || "").trim();
        const jenisPeriode = String(pData[i][1] || "Tahunan").trim();
        const tglDibuat = pData[i][2] || "";
        if (namaPeriode) {
          periodeList.push({
            nama: namaPeriode,
            jenis: jenisPeriode,
            created_at: tglDibuat,
            is_active: namaPeriode === periodeAktif
          });
        }
      }
    }

    // Fallback jika sheet Periode_List belum ada atau kosong
    if (periodeList.length === 0) {
      const defaultPeriods = [
        { nama: "Tahun 2025", jenis: "Tahunan" },
        { nama: "Tahun 2026", jenis: "Tahunan" },
        { nama: "Januari 2025", jenis: "Bulanan" },
        { nama: "Februari 2025", jenis: "Bulanan" },
        { nama: "Maret 2025", jenis: "Bulanan" },
        { nama: "Triwulan I 2025", jenis: "Triwulan" },
        { nama: "Triwulan II 2025", jenis: "Triwulan" },
        { nama: "Triwulan III 2025", jenis: "Triwulan" },
        { nama: "Triwulan IV 2025", jenis: "Triwulan" },
        { nama: "Semester I 2025", jenis: "Semester" },
        { nama: "Semester II 2025", jenis: "Semester" }
      ];
      
      // Buat dan inisialisasi sheet Periode_List
      if (!periodSheet) {
        periodSheet = ss.insertSheet("Periode_List");
        periodSheet.appendRow(["Nama_Periode", "Jenis_Periode", "Tanggal_Dibuat"]);
        defaultPeriods.forEach(p => {
          periodSheet.appendRow([p.nama, p.jenis, new Date()]);
        });
      }

      periodeList = defaultPeriods.map(p => ({
        ...p,
        created_at: new Date(),
        is_active: p.nama === periodeAktif
      }));
    }

  } catch (e) {
    Logger.log("Error getAppConfig: " + e.toString());
  }

  return {
    periode_aktif: periodeAktif,
    periode_list: periodeList
  };
}

/**
 * Menambahkan Periode Baru ke Sheet 'Periode_List'
 */
function handleAddPeriod(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Periode_List");
  if (!sheet) {
    sheet = ss.insertSheet("Periode_List");
    sheet.appendRow(["Nama_Periode", "Jenis_Periode", "Tanggal_Dibuat"]);
  }

  const nama = (params.nama_periode || params.nama || "").trim();
  const jenis = (params.jenis_periode || params.jenis || "Tahunan").trim();

  if (!nama) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "Nama periode wajib diisi."
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // Cek duplikasi
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || "").trim().toLowerCase() === nama.toLowerCase()) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: "Periode '" + nama + "' sudah ada dalam daftar."
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  sheet.appendRow([nama, jenis, new Date()]);

  // Jika parameter set_active true, langsung aktifkan
  if (params.set_active === "true" || params.set_active === true) {
    handleSaveConfig({ periode_penilaian: nama });
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    message: "Periode '" + nama + "' berhasil ditambahkan.",
    config: getAppConfig()
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Mengubah / Mengedit Periode di Sheet 'Periode_List' & Memperbarui Nama Periode di Sheet Penilaian
 */
function handleEditPeriod(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oldName = (params.old_nama_periode || params.old_nama || "").trim();
  const newName = (params.new_nama_periode || params.new_nama || "").trim();
  const newJenis = (params.new_jenis_periode || params.new_jenis || "Tahunan").trim();

  if (!oldName || !newName) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "Nama periode lama dan nama periode baru wajib diisi."
    })).setMimeType(ContentService.MimeType.JSON);
  }

  let periodSheet = ss.getSheetByName("Periode_List");
  if (!periodSheet) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "Daftar periode belum tersedia."
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // 1. Cek apakah nama baru sudah dipakai oleh baris lain
  const data = periodSheet.getDataRange().getValues();
  let foundRowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    const currentName = String(data[i][0] || "").trim();
    if (currentName.toLowerCase() === oldName.toLowerCase()) {
      foundRowIndex = i + 1; // 1-indexed for Sheet
    } else if (currentName.toLowerCase() === newName.toLowerCase()) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: "Nama periode '" + newName + "' sudah ada dalam daftar."
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (foundRowIndex === -1) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "Periode '" + oldName + "' tidak ditemukan."
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // 2. Perbarui baris di sheet 'Periode_List'
  periodSheet.getRange(foundRowIndex, 1).setValue(newName);
  periodSheet.getRange(foundRowIndex, 2).setValue(newJenis);

  // 3. Jika periode yang diedit adalah periode aktif, perbarui juga di 'Pengaturan'
  const currentConfig = getAppConfig();
  if (currentConfig.periode_aktif.toLowerCase() === oldName.toLowerCase()) {
    handleSaveConfig({ periode_penilaian: newName });
  }

  // 4. Perbarui nama periode di seluruh sheet penilaian yang relevan (Cascade Update)
  let updatedAssessmentCells = 0;
  const allSheets = ss.getSheets();

  allSheets.forEach(sheet => {
    const sheetName = sheet.getName().toLowerCase();
    if (sheetName !== "data_pegawai" && sheetName !== "users" && sheetName !== "pengaturan" && sheetName !== "config" && sheetName !== "periode_list" && sheetName !== "log_otp") {
      const values = sheet.getDataRange().getValues();
      if (values.length > 1) {
        for (let r = 1; r < values.length; r++) {
          for (let c = 0; c < values[r].length; c++) {
            const cellVal = String(values[r][c] || "").trim();
            if (cellVal.toLowerCase() === oldName.toLowerCase()) {
              sheet.getRange(r + 1, c + 1).setValue(newName);
              updatedAssessmentCells++;
            }
          }
        }
      }
    }
  });

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    message: "Periode '" + oldName + "' berhasil diubah menjadi '" + newName + "'.",
    updated_assessments_count: updatedAssessmentCells,
    config: getAppConfig()
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Menghapus Periode dari 'Periode_List' & Menghapus Semua Data Penilaian pada Periode Tersebut (Cascade Delete)
 */
function handleDeletePeriod(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nama = (params.nama_periode || params.nama || "").trim();

  if (!nama) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "Nama periode yang akan dihapus tidak ditentukan."
    })).setMimeType(ContentService.MimeType.JSON);
  }

  let deletedFromList = false;
  let periodSheet = ss.getSheetByName("Periode_List");
  if (periodSheet) {
    const data = periodSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0] || "").trim().toLowerCase() === nama.toLowerCase()) {
        periodSheet.deleteRow(i + 1);
        deletedFromList = true;
      }
    }
  }

  // Lakukan CASCADE DELETE di seluruh sheet penilaian (Atasan, Bawahan, Rekan, dsb)
  let deletedAssessmentRows = 0;
  const allSheets = ss.getSheets();

  allSheets.forEach(sheet => {
    const sheetName = sheet.getName().toLowerCase();
    // Cari di sheet penilaian (misalnya: Atasan, Bawahan, Rekan, Form, atau Respon)
    if (sheetName !== "data_pegawai" && sheetName !== "users" && sheetName !== "pengaturan" && sheetName !== "config" && sheetName !== "periode_list" && sheetName !== "log_otp") {
      const values = sheet.getDataRange().getValues();
      if (values.length > 1) {
        // Cari kolom yang berisi nama periode
        for (let r = values.length - 1; r >= 1; r--) {
          let rowContainsPeriod = false;
          for (let c = 0; c < values[r].length; c++) {
            const cellVal = String(values[r][c] || "").trim().toLowerCase();
            if (cellVal === nama.toLowerCase()) {
              rowContainsPeriod = true;
              break;
            }
          }
          if (rowContainsPeriod) {
            sheet.deleteRow(r + 1);
            deletedAssessmentRows++;
          }
        }
      }
    }
  });

  // Jika periode yang dihapus sedang aktif, ubah periode aktif ke periode pertama yang tersedia
  const currentConfig = getAppConfig();
  if (currentConfig.periode_aktif.toLowerCase() === nama.toLowerCase()) {
    const newActive = currentConfig.periode_list.length > 0 ? currentConfig.periode_list[0].nama : "Tahun 2025";
    handleSaveConfig({ periode_penilaian: newActive });
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    message: "Periode '" + nama + "' dan " + deletedAssessmentRows + " data penilaian terkait berhasil dihapus.",
    deleted_assessments_count: deletedAssessmentRows,
    config: getAppConfig()
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Mengambil Tiga Sheet Data Penilaian (Penilaian Atasan, Bawahan, Rekan) untuk Didownload oleh Admin
 */
function handleGetAllPenilaianData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const result = {};

  // Hanya ambil 3 sheet penilaian: Atasan, Bawahan, dan Rekan
  sheets.forEach(sheet => {
    const name = sheet.getName();
    const lowerName = name.toLowerCase();

    // Sheet yang diexclude (bukan data penilaian)
    const isSystemSheet = (
      lowerName === "data_pegawai" ||
      lowerName === "users" ||
      lowerName === "pengaturan" ||
      lowerName === "config" ||
      lowerName === "periode_list" ||
      lowerName === "log_otp"
    );

    const isAssessmentSheet = (
      lowerName.includes("atasan") ||
      lowerName.includes("bawahan") ||
      lowerName.includes("rekan") ||
      lowerName.includes("sejawat")
    );

    if (!isSystemSheet && isAssessmentSheet) {
      const dataRange = sheet.getDataRange();
      const values = dataRange.getValues();
      if (values && values.length > 0) {
        result[name] = values;
      }
    }
  });

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    spreadsheet_name: ss.getName(),
    sheets: result
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Mengirim Email OTP Verifikasi ke Email Pegawai
 */
function handleSendOtp(params) {
  const email = (params.email || "").trim();
  const nama = (params.nama || "Pegawai").trim();
  const nip = (params.nip || "-").trim();
  const otp = (params.otp || generateOtpCode()).trim();

  if (!email || email.indexOf("@") === -1) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "Alamat email tidak valid."
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // Simpan OTP sementara ke ScriptProperties (Expired dalam 10 menit)
  const scriptProperties = PropertiesService.getScriptProperties();
  const otpPayload = {
    otp: otp,
    email: email,
    nip: nip,
    createdAt: new Date().getTime()
  };
  scriptProperties.setProperty("OTP_" + nip, JSON.stringify(otpPayload));

  // Simpan riwayat OTP ke Sheet 'Log_OTP' (opsional jika sheet ada)
  logOtpToSheet(nip, nama, email, otp);

  // Template Email HTML Modern & Elegan
  const subject = `[${APP_NAME}] Kode Verifikasi Pendaftaran Akun - ${otp}`;
  const htmlBody = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="utf-8">
    <style>
      body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f1f5f9; margin: 0; padding: 20px; color: #1e293b; }
      .email-card { max-width: 540px; margin: 0 auto; background: #ffffff; border-radius: 16px; overflow: hidden; box-shadow: 0 10px 25px rgba(0,0,0,0.06); border: 1px solid #e2e8f0; }
      .header { background: linear-gradient(135deg, #1e3a8a, #3b82f6); padding: 30px 24px; text-align: center; color: #ffffff; }
      .header h1 { margin: 0; font-size: 22px; font-weight: 700; letter-spacing: 0.5px; }
      .header p { margin: 6px 0 0 0; font-size: 13px; opacity: 0.9; }
      .body-content { padding: 30px 24px; }
      .greeting { font-size: 15px; line-height: 1.6; color: #334155; margin-bottom: 20px; }
      .otp-container { background: #f8fafc; border: 2px dashed #93c5fd; border-radius: 12px; padding: 20px; text-align: center; margin: 24px 0; }
      .otp-title { font-size: 12px; font-weight: 600; text-transform: uppercase; color: #64748b; letter-spacing: 1px; margin-bottom: 8px; }
      .otp-code { font-size: 36px; font-weight: 800; color: #1e40af; letter-spacing: 8px; font-family: 'Courier New', monospace; margin: 0; }
      .otp-note { font-size: 12px; color: #64748b; margin-top: 8px; }
      .info-box { background: #eff6ff; border-left: 4px solid #3b82f6; padding: 12px 16px; border-radius: 4px; font-size: 13px; color: #1e40af; margin-bottom: 20px; line-height: 1.5; }
      .footer { background: #f8fafc; padding: 20px; text-align: center; font-size: 12px; color: #94a3b8; border-top: 1px solid #e2e8f0; }
    </style>
  </head>
  <body>
    <div class="email-card">
      <div class="header">
        <h1>${APP_NAME}</h1>
        <p>${INSTANSI_NAME}</p>
      </div>
      <div class="body-content">
        <p class="greeting">
          Yth. Bapak/Ibu <strong>${nama}</strong>,<br>
          (NIP: <strong>${nip}</strong>)
        </p>
        <p style="font-size: 14px; color: #475569; line-height: 1.6;">
          Kami menerima permintaan pembuatan / pendaftaran password akun Anda pada Sistem Informasi Penilaian Kinerja 360 Derajat.
        </p>
        
        <div class="otp-container">
          <div class="otp-title">KODE VERIFIKASI (OTP) ANDA</div>
          <div class="otp-code">${otp}</div>
          <div class="otp-note">⏱️ Berlaku selama <strong>10 menit</strong></div>
        </div>

        <div class="info-box">
          ⚠️ <strong>PENTING:</strong> Jangan berikan kode ini kepada siapa pun termasuk rekan kerja atau atasan. Petugas tidak akan pernah meminta kode verifikasi Anda.
        </div>

        <p style="font-size: 13px; color: #64748b; line-height: 1.5;">
          Jika Anda tidak merasa melakukan pendaftaran ini, Anda dapat mengabaikan email ini atau segera menghubungi Administrator Kepegawaian BKPSDM.
        </p>
      </div>
      <div class="footer">
        &copy; ${new Date().getFullYear()} ${INSTANSI_NAME}<br>
        Sistem Informasi Perilaku Kerja Aparatur 360 Derajat
      </div>
    </div>
  </body>
  </html>
  `;

  // Kirim email menggunakan MailApp atau GmailApp
  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      name: EMAIL_SENDER_NAME
    });
  } catch (errMail) {
    GmailApp.sendEmail(email, subject, `Kode Verifikasi SIPEKA 360 Anda adalah: ${otp}`, {
      name: EMAIL_SENDER_NAME,
      htmlBody: htmlBody
    });
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    message: "Kode OTP berhasil dikirimkan ke email.",
    emailMasked: maskEmail(email)
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Menyimpan Akun Pengguna Baru ke Google Sheets
 */
function handleRegisterUser(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Users");
  
  if (!sheet) {
    sheet = ss.insertSheet("Users");
    sheet.appendRow(["Timestamp", "NIP", "Nama Pegawai", "Email", "Nomor HP", "Password", "Pertanyaan Pengaman", "Jawaban Pengaman"]);
    sheet.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#dbeafe");
  }

  const timestamp = new Date();
  const nip = "'" + (params.nip || "").trim(); // kutip satu agar NIP tidak terpotong leading zero
  const rawNip = (params.nip || "").trim();
  const nama = params.nama || "-";
  const email = params.email || params.no_hp || "-";
  const noHp = params.no_hp || "-";
  const password = params.password || "-";
  const pertanyaan = params.pertanyaan || "-";
  const jawaban = params.jawaban || "-";

  // Cek apakah NIP sudah ada di sheet Users, jika ada maka update
  const data = sheet.getDataRange().getValues();
  let rowIndexToUpdate = -1;
  for (let i = 1; i < data.length; i++) {
    const rowNip = String(data[i][1]).replace(/['\s]/g, "");
    if (rowNip === rawNip) {
      rowIndexToUpdate = i + 1; // 1-indexed
      break;
    }
  }

  if (rowIndexToUpdate > 0) {
    sheet.getRange(rowIndexToUpdate, 1, 1, 8).setValues([[
      timestamp, nip, nama, email, noHp, password, pertanyaan, jawaban
    ]]);
  } else {
    sheet.appendRow([timestamp, nip, nama, email, noHp, password, pertanyaan, jawaban]);
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    message: "Registrasi pengguna berhasil disimpan."
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Reset / Update Password Pengguna
 */
function handleResetPassword(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Users");
  const rawNip = (params.nip || "").trim();
  const newPassword = params.password || "";
  const timestamp = new Date();

  if (!sheet) {
    sheet = ss.insertSheet("Users");
    sheet.appendRow(["Timestamp", "NIP", "Nama Pegawai", "Email", "Nomor HP", "Password", "Pertanyaan Pengaman", "Jawaban Pengaman"]);
  }

  const data = sheet.getDataRange().getValues();
  let found = false;

  for (let i = 1; i < data.length; i++) {
    const rowNip = String(data[i][1]).replace(/['\s]/g, "");
    if (rowNip === rawNip) {
      sheet.getRange(i + 1, 1).setValue(timestamp);
      sheet.getRange(i + 1, 6).setValue(newPassword);
      found = true;
      break;
    }
  }

  if (!found) {
    sheet.appendRow([timestamp, "'" + rawNip, "-", "-", "-", newPassword, "-", "-"]);
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    message: "Password berhasil diperbarui."
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Menyimpan data formulir penilaian umum (Atasan / Bawahan / Rekan)
 */
function handleSavePenilaian(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getActiveSheet();

  const timestamp = new Date();
  const rowData = [timestamp];

  for (let key in params) {
    if (key !== "action") {
      let val = params[key];
      if (key.toLowerCase().includes("nip")) {
        val = "'" + val;
      }
      rowData.push(val);
    }
  }

  sheet.appendRow(rowData);

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    message: "Data penilaian berhasil disimpan."
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Helper: Generate 6-Digit Angka Acak
 */
function generateOtpCode() {
  return Math.floor(100000 + Math.random() * 900000).toString();
}

/**
 * Helper: Sensor sebagian alamat email (contoh: f***@gmail.com)
 */
function maskEmail(email) {
  if (!email || email.indexOf("@") === -1) return email;
  const parts = email.split("@");
  const name = parts[0];
  const domain = parts[1];
  if (name.length <= 2) {
    return name[0] + "***@" + domain;
  }
  return name.substring(0, 2) + "***" + name.substring(name.length - 1) + "@" + domain;
}

/**
 * Helper: Log OTP ke sheet terpisah untuk audit log (opsional)
 */
function logOtpToSheet(nip, nama, email, otp) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Log_OTP");
    if (!sheet) {
      sheet = ss.insertSheet("Log_OTP");
      sheet.appendRow(["Timestamp", "NIP", "Nama", "Email", "Kode OTP"]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
    }
    sheet.appendRow([new Date(), "'" + nip, nama, email, otp]);
  } catch (e) {
    // Abaikan error logging jika ada kendala hak akses sheet
  }
}

/**
 * FUNGSI TEST PENGIRIMAN EMAIL & OTORISASI GOOGLE
 * Jalankan fungsi ini dari editor Apps Script untuk mengizinkan (Authorize) pengiriman email.
 */
function testKirimEmailManual() {
  const userEmail = Session.getActiveUser().getEmail() || "firasofficial@gmail.com";
  const dummyOtp = "123456";
  Logger.log("Mengirim email percobaan ke: " + userEmail);
  
  handleSendOtp({
    email: userEmail,
    nama: "Administrator Testing",
    nip: "199001012015011001",
    otp: dummyOtp
  });
  
  Logger.log("Email berhasil dikirim! Silakan periksa inbox / spam email Anda.");
}

