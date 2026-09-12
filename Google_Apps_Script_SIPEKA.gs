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

    // Aksi: Ambil Master Data Pegawai dari Sheet 'Data_Pegawai'
    if (action === "get_pegawai") {
      return handleGetPegawai();
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

    // 2. Aksi: Kirim Kode OTP ke Email
    if (action === "send_otp") {
      return handleSendOtp(params);
    }

    // 3. Aksi: Registrasi Pengguna Baru
    if (action === "register") {
      return handleRegisterUser(params);
    }

    // 4. Aksi: Reset Password
    if (action === "reset_password") {
      return handleResetPassword(params);
    }

    // 5. Default: Simpan Formulir Penilaian (Atasan / Rekan / Bawahan)
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

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    total: result.length,
    data: result
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

