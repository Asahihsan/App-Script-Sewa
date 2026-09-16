// ==========================================
// 1. ROUTING SYSTEM & LOGIN MULTI-ROLE
// ==========================================
function doGet(e) {
  var emailUser = Session.getActiveUser().getEmail();
  var userData = getUserDetails(emailUser);
  var page = (e && e.parameter) ? e.parameter.page : null;
  var requestedRole = (e && e.parameter) ? e.parameter.role : null;

  if (userData.status !== "Aktif" || !userData.roles || userData.roles.length === 0) {
    var tmplDenied = HtmlService.createTemplateFromFile('AccessDenied');
    tmplDenied.emailUser = emailUser;
    return tmplDenied.evaluate()
      .setTitle('SIMPEWA - Akses Ditolak')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (!page) {
    if (userData.roles.length === 1) {
      return renderDashboardByRole(userData.roles[0], emailUser, userData.namaUser);
    }

    var template = HtmlService.createTemplateFromFile('PortalLogin');
    template.emailAwal = emailUser;
    template.namaUser = userData.namaUser;
    template.userRoles = userData.roles;
    template.appUrl = ScriptApp.getService().getUrl();
    return template.evaluate()
      .setTitle('SIMPEWA - Portal Akses')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === 'dashboard') {
    if (!requestedRole) requestedRole = userData.roles[0];

    if (!userData.roles.includes(requestedRole)) {
      var tmplDenied = HtmlService.createTemplateFromFile('AccessDenied');
      tmplDenied.emailUser = emailUser;
      return tmplDenied.evaluate()
        .setTitle('SIMPEWA - Akses Ditolak')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    return renderDashboardByRole(requestedRole, emailUser, userData.namaUser);
  }

  var templateFallback = HtmlService.createTemplateFromFile('PortalLogin');
  templateFallback.emailAwal = emailUser;
  templateFallback.namaUser = userData.namaUser;
  templateFallback.userRoles = userData.roles;
  templateFallback.appUrl = ScriptApp.getService().getUrl();
  return templateFallback.evaluate()
    .setTitle('SIMPEWA - Portal Akses')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function renderDashboardByRole(role, emailUser, namaUser) {
  var fileName = "";
  if (role === "Negosiator") fileName = "DashboardNegosiator";
  else if (role === "Sekretaris") fileName = "DashboardSekretaris";
  else if (role === "Direktur") fileName = "DashboardDirektur";
  else if (role === "Keuangan") fileName = "DashboardKeuangan";
  else if (role === "Admin") fileName = "DashboardAdmin";
  else if (role === "MCX") fileName = "DashboardMCX";
  else if (role === "Informasi") fileName = "DashboardInformasi";
  else fileName = "AccessDenied";

  var tmpl = HtmlService.createTemplateFromFile(fileName);
  tmpl.emailUser = emailUser;
  tmpl.namaUser = namaUser;
  tmpl.currentRole = role;
  tmpl.appUrl = ScriptApp.getService().getUrl();
  return tmpl.evaluate()
    .setTitle('SIMPEWA - Dashboard ' + role)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getUserDetails(email) {
  var defaultResult = { namaUser: "Pengguna", roles: [], status: "Nonaktif" };
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DATA_USER");
    if (!sheet) return defaultResult;
    var data = sheet.getDataRange().getValues();

    var userRoles = [];
    var namaUser = "Pengguna";
    var statusUser = "Nonaktif";
    var targetEmail = String(email).trim().toLowerCase();

    for (var i = 1; i < data.length; i++) {
      var rowEmail = String(data[i][0]).trim().toLowerCase();

      if (rowEmail === targetEmail) {
        namaUser = String(data[i][1]).trim() || namaUser;
        var roleStr = String(data[i][2]).trim();
        var statusStr = String(data[i][3]).trim() || "Aktif";

        if (statusStr === "Aktif" && roleStr !== "" && !userRoles.includes(roleStr)) {
          userRoles.push(roleStr);
          statusUser = "Aktif";
        }
      }
    }

    return { namaUser: namaUser, roles: userRoles, status: statusUser };
  } catch (e) {
    return defaultResult;
  }
}

// ==========================================
// 2. BACKEND NEGOSIATOR (DRAFT, RELOKASI & PRIVASI)
// ==========================================
function simpanPengajuanBaru(data, statusSubmit) {
  var roles = getUserDetails(Session.getActiveUser().getEmail()).roles;
  if (!roles.includes("Negosiator") && !roles.includes("Admin")) throw new Error("Akses Ditolak: Hak akses Negosiator dibutuhkan.");

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("Sistem sibuk."); }

  try {
    var emailNegosiator = Session.getActiveUser().getEmail();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("NEW_INPUT");

    if (!statusSubmit) statusSubmit = "FINAL";

    if (statusSubmit === "FINAL") {
      if (!data.lokasi) throw new Error("Lokasi wajib dipilih!");
      if (!data.hargaBaru || parseFloat(String(data.hargaBaru).replace(/\./g, '')) <= 0) throw new Error("Harga sewa baru wajib diisi!");
      if (!data.spkData && !data.existingSpk) throw new Error("Draft SPK Wajib diunggah!");
      if (!data.bankNama || !data.bankNoRek || !data.bankAn) throw new Error("Informasi Rekening Bank Pemilik wajib dilengkapi!");
      if (!data.namaPemilik) throw new Error("Nama Pemilik Lokasi wajib diisi!");
      if (data.statusSewa === "Relokasi" && !data.lokasiLama) throw new Error("Untuk status Relokasi, Lokasi Lama wajib diisi!");
    }

    var primaryKey = data.kodeId;
    var isEdit = false;

    if (primaryKey && primaryKey !== "") {
      isEdit = true;
    } else {
      var maxSeq = 0;
      var findData = sheet.getDataRange().getValues();
      var kodeLokasi = data.lokasi ? data.lokasi.substring(0, 3).toUpperCase() : "REL";
      var tahun = new Date().getFullYear();

      for (var r = 4; r < findData.length; r++) {
        var existingId = String(findData[r][0] || "");
        if (existingId.includes("-")) {
          var parts = existingId.split("-");
          if (parts.length === 3) {
            var seq = parseInt(parts[2], 10);
            if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
          }
        }
      }
      primaryKey = kodeLokasi + "-" + tahun + "-" + ("0000" + (maxSeq + 1)).slice(-4);
    }

    var spkCellContent = data.existingSpk || "";
    if (data.spkData) {
      var spkFolder = DriveApp.getFolderById('1KF1yPXGPflR1T9U6Rt3e1PRV8lMzLypS');
      var fileSPK = spkFolder.createFile(Utilities.newBlob(Utilities.base64Decode(data.spkData), data.spkMimeType, "SPK_" + primaryKey + "_" + data.spkName));
      spkCellContent = fileSPK.getUrl() + "|ID:" + fileSPK.getId();
    }

    var memoCellContent = data.existingMemo || "";
    if (statusSubmit === "FINAL") {
      var memoResult = buatMemoPDF(data, primaryKey);
      memoCellContent = memoResult.url + "|ID:" + memoResult.id;
    }

    var hargaBaru = parseFloat(String(data.hargaBaru || "0").replace(/\./g, '')) || 0;
    var nominalPajak = Math.round(data.dipotongPajak && data.dipotongPajak.includes("Ya") ? (hargaBaru * 0.10) : ((hargaBaru / 0.9) * 0.10));
    var hargaSetelahPajak = Math.round(data.dipotongPajak && data.dipotongPajak.includes("Ya") ? (hargaBaru - nominalPajak) : hargaBaru);
    var nilaiGrossUp = Math.round(data.dipotongPajak && data.dipotongPajak.includes("Ya") ? hargaBaru : (hargaBaru / 0.9));
    var infoRekeningBank = data.bankNama ? (data.bankNama.toUpperCase() + " - " + data.bankNoRek + " a.n " + data.bankAn.toUpperCase()) : "";
    var tahunSewaLama = data.tanggalHabisLama ? String(data.tanggalHabisLama).substring(0, 4) : (data.tahunSewaLama || "");

    var barisData = [
      primaryKey,
      new Date().getFullYear(),
      data.lokasi || "",
      data.statusSewa || "Kontrak Baru",
      data.pembayarPBB || "Pemilik",
      hargaBaru,
      nominalPajak,
      hargaSetelahPajak,
      parseInt(data.masaSewa) || 1,
      data.tglAkhirBaru || "",
      nilaiGrossUp,
      data.tglCicilan1 || "", parseFloat(String(data.jmlCicilan1 || "0").replace(/\./g, '')) || 0,
      data.tglCicilan2 || "", parseFloat(String(data.jmlCicilan2 || "0").replace(/\./g, '')) || 0,
      data.tglCicilan3 || "", parseFloat(String(data.jmlCicilan3 || "0").replace(/\./g, '')) || 0,
      parseFloat(String(data.depositBaru || "0").replace(/\./g, '')) || 0,
      spkCellContent,
      memoCellContent,
      tahunSewaLama,
      data.tanggalHabisLama || "",
      parseInt(data.masaSewaLama) || 0,
      parseFloat(String(data.hargaLama || "0").replace(/\./g, '')) || 0,
      parseFloat(String(data.depositLama || "0").replace(/\./g, '')) || 0,
      statusSubmit === "FINAL" ? "PENDING" : "",
      "", "", "", "", "",
      infoRekeningBank,
      "", "", "", "", "",
      statusSubmit,
      data.namaPemilik || "",
      data.lokasiLama || "",
      emailNegosiator
    ];

    if (isEdit) {
      var findData = sheet.getDataRange().getValues();
      for (var row = 4; row < findData.length; row++) {
        if (findData[row][0] === primaryKey) {
          sheet.getRange(row + 1, 1, 1, barisData.length).setValues([barisData]);
          break;
        }
      }
    } else {
      sheet.appendRow(barisData);
    }

    return primaryKey;
  } finally { lock.releaseLock(); }
}

function getRiwayatPengajuan() {
  var emailLogin = Session.getActiveUser().getEmail().trim().toLowerCase();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 4; i < data.length; i++) {
    var emailPembuat = String(data[i][40] || "").trim().toLowerCase();

    if (data[i][0] !== "" && emailPembuat === emailLogin) {
      result.push({
        kodeId: data[i][0],
        lokasi: data[i][2],
        statusSewa: data[i][3],
        hargaBaru: data[i][5],
        statusDraft: data[i][37] || "FINAL",
        statusMcx: data[i][25] || "-",
        catatanMcx: data[i][26] || "-",
        statusSekretaris: data[i][27],
        catatanSekretaris: data[i][28] || "-",
        statusDirektur: data[i][29],
        ntpn: data[i][34]
      });
    }
  }
  return result;
}

function getDetailPengajuan(kodeId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();

  for (var i = 4; i < data.length; i++) {
    if (data[i][0] === kodeId) {
      var hrgBaru = parseFloat(data[i][5]) || 0;
      var grsUp = parseFloat(data[i][10]) || 0;
      var fmtTgl = function (v) { return (v instanceof Date) ? v.toISOString().split('T')[0] : (v || ""); };

      var infoRek = String(data[i][31] || "");
      var bankNama = "", bankNoRek = "", bankAn = "";
      if (infoRek.includes(" - ") && infoRek.includes(" a.n ")) {
        var parts1 = infoRek.split(" - ");
        bankNama = parts1[0];
        var parts2 = parts1[1].split(" a.n ");
        bankNoRek = parts2[0];
        bankAn = parts2[1];
      }

      var tglAkhirFix = fmtTgl(data[i][9]);
      var masaSewaFix = parseInt(data[i][8]) || 1;
      var tglMulaiFix = "";
      if (tglAkhirFix && masaSewaFix) {
        var d = new Date(tglAkhirFix);
        d.setFullYear(d.getFullYear() - masaSewaFix);
        d.setDate(d.getDate() + 1);
        tglMulaiFix = d.toISOString().split('T')[0];
      }

      return {
        kodeId: data[i][0],
        lokasi: data[i][2],
        statusSewa: data[i][3],
        pembayarPBB: data[i][4] || "Pemilik",
        hargaBaru: hrgBaru,
        dipotongPajak: (Math.round(grsUp) === Math.round(hrgBaru)) ? "Ya, dipotong" : "Tidak (Gross Up)",
        masaSewaBaru: masaSewaFix,
        tglMulaiBaru: tglMulaiFix,
        tglAkhirBaru: tglAkhirFix,
        tglCicilan1: fmtTgl(data[i][11]),
        jmlCicilan1: data[i][12] || 0,
        tglCicilan2: fmtTgl(data[i][13]),
        jmlCicilan2: data[i][14] || 0,
        tglCicilan3: fmtTgl(data[i][15]),
        jmlCicilan3: data[i][16] || 0,
        depositBaru: data[i][17] || 0,
        spkDokumen: data[i][18] || "",
        memoDokumen: data[i][19] || "",
        tahunSewaLama: data[i][20] || "",
        tanggalHabisLama: fmtTgl(data[i][21]),
        masaSewaLama: data[i][22] || "",
        hargaLama: data[i][23] || 0,
        depositLama: data[i][24] || 0,
        catatanPenolakan: data[i][30] ? data[i][30] : (data[i][28] ? data[i][28] : (data[i][26] || "Tidak ada catatan.")),
        bankNama: bankNama, bankNoRek: bankNoRek, bankAn: bankAn,
        statusDraft: data[i][37] || "FINAL",
        namaPemilik: data[i][38] || "",
        lokasiLama: data[i][39] || ""
      };
    }
  }
  return null;
}

function batalkanPengajuan(kodeId) {
  var lock = LockService.getScriptLock(); try { lock.waitLock(10000); } catch (e) { throw new Error("Sistem sibuk."); }
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT"); var data = sheet.getDataRange().getValues();
    for (var i = 4; i < data.length; i++) {
      if (data[i][0] === kodeId) {
        var spk = data[i][18], memo = data[i][19];
        if (spk && spk.indexOf("|ID:") !== -1) { try { DriveApp.getFileById(spk.split("|ID:")[1]).setTrashed(true); } catch (e) { } }
        if (memo && memo.indexOf("|ID:") !== -1) { try { DriveApp.getFileById(memo.split("|ID:")[1]).setTrashed(true); } catch (e) { } }
        sheet.deleteRow(i + 1); return "Pengajuan dibatalkan.";
      }
    } throw new Error("ID tidak ditemukan!");
  } finally { lock.releaseLock(); }
}

// ==========================================
// 3. BACKEND ROLE MCX (MASTER CONTROL)
// ==========================================
function getPengajuanMCX() {
  var roles = getUserDetails(Session.getActiveUser().getEmail()).roles;
  if (!roles.includes("MCX") && !roles.includes("Admin")) throw new Error("Akses Ditolak.");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 4; i < data.length; i++) {
    var statusDraft = String(data[i][37] || "").trim();
    var statusMcx = String(data[i][25] || "").trim();

    if (data[i][0] !== "" && statusDraft === "FINAL" && (statusMcx === "PENDING" || statusMcx === "")) {
      result.push({
        kodeId: data[i][0], lokasi: data[i][2], statusSewa: data[i][3], pembayarPBB: data[i][4],
        hargaPokok: data[i][5], pajak: data[i][6], hargaBersih: data[i][7], masaSewa: data[i][8],
        tglAkhirBaru: data[i][9] instanceof Date ? data[i][9].toISOString().split('T')[0] : data[i][9],
        nilaiGrossUp: data[i][10], spkDokumen: data[i][18], memoDokumen: data[i][19],
        namaPemilik: data[i][38], lokasiLama: data[i][39] || "-"
      });
    }
  }
  return result;
}

function verifikasiDanAdjustMCX(kodeId, keputusan, nominalBaruFix, catatanMcx) {
  var roles = getUserDetails(Session.getActiveUser().getEmail()).roles;
  if (!roles.includes("MCX") && !roles.includes("Admin")) throw new Error("Akses Ditolak: Anda bukan tim MCX.");

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("Sistem sibuk."); }

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
    var data = sheet.getDataRange().getValues();

    for (var i = 4; i < data.length; i++) {
      if (data[i][0] === kodeId) {

        if (keputusan === "APPROVED") {
          var hrgFix = parseFloat(String(nominalBaruFix).replace(/\./g, '')) || data[i][5];
          var dipotongPajak = (Math.round(data[i][10]) === Math.round(data[i][5])) ? "Ya, dipotong" : "Tidak (Gross Up)";

          var nominalPajak = Math.round(dipotongPajak.includes("Ya") ? (hrgFix * 0.10) : ((hrgFix / 0.9) * 0.10));
          var hargaSetelahPajak = Math.round(dipotongPajak.includes("Ya") ? (hrgFix - nominalPajak) : hrgFix);
          var nilaiGrossUp = Math.round(dipotongPajak.includes("Ya") ? hrgFix : (hrgFix / 0.9));

          sheet.getRange(i + 1, 6).setValue(hrgFix);
          sheet.getRange(i + 1, 7).setValue(nominalPajak);
          sheet.getRange(i + 1, 8).setValue(hargaSetelahPajak);
          sheet.getRange(i + 1, 11).setValue(nilaiGrossUp);

          sheet.getRange(i + 1, 26).setValue("APPROVED");
        } else {
          sheet.getRange(i + 1, 26).setValue("REVISED");
        }

        sheet.getRange(i + 1, 27).setValue(catatanMcx || "-");
        return "Verifikasi MCX untuk ID " + kodeId + " berhasil disimpan dengan status: " + keputusan;
      }
    }
    throw new Error("ID tidak ditemukan!");
  } finally { lock.releaseLock(); }
}

// ==========================================
// 4. BACKEND SEKRETARIS & DIREKTUR
// ==========================================
function getPengajuanSekretaris() {
  var roles = getUserDetails(Session.getActiveUser().getEmail()).roles;
  if (!roles.includes("Sekretaris") && !roles.includes("Admin")) throw new Error("Akses Ditolak.");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
  if (!sheet) return { pendingList: [], countPending: 0, countApproved: 0, countRejected: 0 };
  var data = sheet.getDataRange().getValues();
  var pendingList = [];
  var countApproved = 0, countRejected = 0;

  for (var i = 4; i < data.length; i++) {
    if (data[i][0] !== "") {
      var statusMcx = String(data[i][25] || "").trim();
      var statusCek = data[i][27];

      if (statusMcx === "APPROVED") {
        if (statusCek === true) {
          countApproved++;
        } else if (statusCek === false) {
          countRejected++;
        } else {
          pendingList.push({
            kodeId: data[i][0], lokasi: data[i][2], statusSewa: data[i][3], pembayarPBB: data[i][4],
            hargaPokok: data[i][5], pajak: data[i][6], hargaBersih: data[i][7], masaSewa: data[i][8],
            tglHabisBaru: data[i][9] instanceof Date ? data[i][9].toISOString().split('T')[0] : data[i][9],
            spkDokumen: data[i][18], memoDokumen: data[i][19],
            statusCek: statusCek, catatan: data[i][28] || "-",
            namaPemilik: data[i][38], lokasiLama: data[i][39] || "-"
          });
        }
      }
    }
  }
  return { pendingList: pendingList, countPending: pendingList.length, countApproved: countApproved, countRejected: countRejected };
}

function getRiwayatSekretaris() {
  var roles = getUserDetails(Session.getActiveUser().getEmail()).roles;
  if (!roles.includes("Sekretaris") && !roles.includes("Admin")) throw new Error("Akses Ditolak.");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 4; i < data.length; i++) {
    var statusMcx = String(data[i][25] || "").trim();
    var statusCek = data[i][27];

    if (data[i][0] !== "" && statusMcx === "APPROVED" && (statusCek === true || statusCek === false)) {
      result.push({
        kodeId: data[i][0], lokasi: data[i][2], statusSewa: data[i][3], pembayarPBB: data[i][4],
        hargaPokok: data[i][5], pajak: data[i][6], hargaBersih: data[i][7], masaSewa: data[i][8],
        tglHabisBaru: data[i][9] instanceof Date ? data[i][9].toISOString().split('T')[0] : data[i][9],
        spkDokumen: data[i][18], memoDokumen: data[i][19],
        statusCek: statusCek, catatan: data[i][28] || "-"
      });
    }
  }
  return result;
}

function getPengajuanSiapCetak() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 4; i < data.length; i++) {
    if (data[i][0] !== "" && data[i][29] === "Disetujui") {
      result.push({
        kodeId: data[i][0],
        lokasi: data[i][2],
        statusSewa: data[i][3],
        nilaiGrossUp: data[i][10] || data[i][5],
        hargaBaru: data[i][5],
        catatanDirektur: data[i][30] || "-",
        memoUrl: data[i][19] ? data[i][19].split("|ID:")[0] : "-",
        spkUrl: data[i][18] ? data[i][18].split("|ID:")[0] : "-"
      });
    }
  }
  return result;
}

function verifikasiPengajuan(kodeId, statusKeputusan, catatanSekretaris) {
  var role = getUserDetails(Session.getActiveUser().getEmail()).roles;
  if (!role.includes("Sekretaris") && !role.includes("Admin")) throw new Error("Akses Ditolak.");
  var lock = LockService.getScriptLock(); try { lock.waitLock(10000); } catch (e) { throw new Error("Sistem sibuk."); }
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT"); var data = sheet.getDataRange().getValues();
    for (var i = 4; i < data.length; i++) {
      if (data[i][0] === kodeId) {
        sheet.getRange(i + 1, 28).setValue(statusKeputusan === "Setuju" ? true : false);
        sheet.getRange(i + 1, 29).setValue(catatanSekretaris || "-");
        return "Berhasil memperbarui status pengajuan.";
      }
    } throw new Error("ID tidak ditemukan!");
  } finally { lock.releaseLock(); }
}

function getPengajuanDirektur() {
  var roles = getUserDetails(Session.getActiveUser().getEmail()).roles;
  if (!roles.includes("Direktur") && !roles.includes("Admin")) throw new Error("Akses Ditolak.");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
  if (!sheet) return []; var data = sheet.getDataRange().getValues(); var result = [];
  for (var i = 4; i < data.length; i++) {
    if (data[i][0] !== "" && data[i][27] === true && (!data[i][29] || data[i][29] === "")) {
      result.push({
        kodeId: data[i][0], lokasi: data[i][2], statusSewa: data[i][3], hargaPokok: data[i][5],
        nilaiGrossUp: data[i][10], masaSewa: data[i][8], tglHabisBaru: data[i][9] instanceof Date ? data[i][9].toISOString().split('T')[0] : data[i][9],
        spkDokumen: data[i][18], memoDokumen: data[i][19], catatanSekretaris: data[i][28] || "-",
        namaPemilik: data[i][38], lokasiLama: data[i][39] || "-"
      });
    }
  } return result;
}

function getExecutiveAnalytics() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
  if (!sheet) return { totalPengajuan: 0, countPending: 0, totalNominalApproved: 0, totalNominalPending: 0 };

  var data = sheet.getDataRange().getValues();
  var totalPengajuan = 0, countPending = 0;
  var totalApproved = 0, totalPending = 0;

  for (var i = 4; i < data.length; i++) {
    if (data[i][0] !== "" && data[i][37] === "FINAL") {
      totalPengajuan++;
      var gross = parseFloat(data[i][10]) || parseFloat(data[i][5]) || 0;
      var statusDir = data[i][29];

      if (statusDir === "Disetujui") {
        totalApproved += gross;
      } else {
        countPending++;
        totalPending += gross;
      }
    }
  }
  return {
    totalPengajuan: totalPengajuan,
    countPending: countPending,
    totalNominalApproved: totalApproved,
    totalNominalPending: totalPending
  };
}

function getRiwayatDirektur() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 4; i < data.length; i++) {
    if (data[i][0] !== "" && data[i][29] && data[i][29] !== "") {
      result.push({
        kodeId: data[i][0],
        lokasi: data[i][2],
        nilaiGrossUp: data[i][10] || data[i][5],
        statusDirektur: data[i][29],
        catatanDirektur: data[i][30] || "-"
      });
    }
  }
  return result;
}

function verifikasiDirektur(kodeId, statusKeputusan, catatanDirektur) {
  var roles = getUserDetails(Session.getActiveUser().getEmail()).roles;
  if (!roles.includes("Direktur") && !roles.includes("Admin")) throw new Error("Akses Ditolak.");

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("Sistem sibuk."); }

  try {
    checkAndPrepareSheets();

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetInput = ss.getSheetByName("NEW_INPUT");
    var dataInput = sheetInput.getDataRange().getValues();

    for (var i = 4; i < dataInput.length; i++) {
      if (dataInput[i][0] === kodeId) {
        sheetInput.getRange(i + 1, 30).setValue(statusKeputusan);
        sheetInput.getRange(i + 1, 31).setValue(catatanDirektur || "-");

        if (statusKeputusan === "Disetujui") {
          var sheetKeu = ss.getSheetByName("NEW_KEU");
          var dataKeu = sheetKeu.getDataRange().getValues();

          var exists = false;
          for (var j = 3; j < dataKeu.length; j++) {
            if (dataKeu[j][0] === kodeId) { exists = true; break; }
          }

          if (!exists) {
            var barisKeu = [
              dataInput[i][0], dataInput[i][1], "", dataInput[i][2], dataInput[i][3],
              dataInput[i][21], dataInput[i][23], dataInput[i][24], dataInput[i][5],
              dataInput[i][6], dataInput[i][7], dataInput[i][9], dataInput[i][10],
              dataInput[i][7], dataInput[i][6], "", "", "", dataInput[i][7],
              "ADA", dataInput[i][18], "BELUM", "", "", "", "", ""
            ];
            sheetKeu.appendRow(barisKeu);
          }
        }
        return "Pengajuan berhasil " + statusKeputusan + " dan disalurkan ke Divisi Keuangan.";
      }
    }
    throw new Error("ID Pengajuan tidak ditemukan!");
  } finally { lock.releaseLock(); }
}

// ==========================================
// 5. BACKEND DIVISI KEUANGAN
// ==========================================
function getPengajuanKeuangan() {
  var roles = getUserDetails(Session.getActiveUser().getEmail()).roles;
  if (!roles.includes("Keuangan") && !roles.includes("Admin")) throw new Error("Akses Ditolak.");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 4; i < data.length; i++) {
    if (data[i][0] !== "" && data[i][29] === "Disetujui" && (!data[i][34] || data[i][34] === "")) {
      result.push({
        kodeId: data[i][0], lokasi: data[i][2], statusSewa: data[i][3],
        hargaPokok: data[i][5], nominalPajak: data[i][6], hargaBersih: data[i][7],
        masaSewa: data[i][8], tglHabisBaru: data[i][9] instanceof Date ? data[i][9].toISOString().split('T')[0] : data[i][9],
        nilaiGrossUp: data[i][10], spkDokumen: data[i][18], memoDokumen: data[i][19],
        catatanDirektur: data[i][30] || "-"
      });
    }
  }
  return result;
}

function getRiwayatKeuangan() {
  var roles = getUserDetails(Session.getActiveUser().getEmail()).roles;
  if (!roles.includes("Keuangan") && !roles.includes("Admin")) throw new Error("Akses Ditolak.");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 4; i < data.length; i++) {
    if (data[i][0] !== "" && data[i][34] && String(data[i][34]).trim() !== "") {
      result.push({
        kodeId: data[i][0],
        lokasi: data[i][2],
        nilaiGrossUp: data[i][10] || data[i][5],
        tglBayarPajak: data[i][33] instanceof Date ? data[i][33].toISOString().split('T')[0] : (data[i][33] || "-"),
        ntpn: data[i][34],
        bukpotLink: data[i][36] || ""
      });
    }
  }
  return result;
}

// ==========================================
// BACKEND DIVISI KEUANGAN (NEW_KEU UPDATE)
// ==========================================
function simpanDataKeuangan(data) {
  var roles = getUserDetails(Session.getActiveUser().getEmail()).roles;
  if (!roles.includes("Keuangan") && !roles.includes("Admin")) throw new Error("Akses Ditolak: Anda bukan tim Keuangan.");

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("Sistem sibuk."); }

  try {
    checkAndPrepareSheets();

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetInput = ss.getSheetByName("NEW_INPUT");
    var sheetKeu = ss.getSheetByName("NEW_KEU");

    var dataInput = sheetInput.getDataRange().getValues();
    var dataKeu = sheetKeu.getDataRange().getValues();

    for (var i = 4; i < dataInput.length; i++) {
      if (dataInput[i][0] === data.kodeId) {

        // 1. Upload Bukti Potong jika ada
        var bukpotCellContent = "";
        if (data.bukpotData && data.bukpotName) {
          var bukpotFolder = DriveApp.getFolderById('1tEW_fyH69zpOWqq3oAI0u7QkT_yZsqS_');
          var blob = Utilities.newBlob(Utilities.base64Decode(data.bukpotData), data.bukpotMimeType, "BUKPOT_" + data.kodeId + "_" + data.bukpotName);
          var fileBukpot = bukpotFolder.createFile(blob);
          bukpotCellContent = fileBukpot.getUrl() + "|ID:" + fileBukpot.getId();
        }

        var nilaiPajakRealNum = parseFloat(String(data.nilaiPajakReal || "0").replace(/\./g, '')) || 0;
        var hardcopyStatus = data.hardcopySPK ? "ADA" : "BELUM";

        // 2. Update status & pajak di Master NEW_INPUT (Kolom AG s/d AK)
        sheetInput.getRange(i + 1, 33).setValue(hardcopyStatus);       // AG (32)
        sheetInput.getRange(i + 1, 34).setValue(data.tglBayarPajak);  // AH (33)
        sheetInput.getRange(i + 1, 35).setValue(data.ntpnPajak);     // AI (34)
        sheetInput.getRange(i + 1, 36).setValue(nilaiPajakRealNum);  // AJ (35)
        if (bukpotCellContent !== "") sheetInput.getRange(i + 1, 37).setValue(bukpotCellContent); // AK (36)

        // 3. Update / Isi Lengkap Baris NEW_KEU (26 Kolom A-Z)
        var targetRowKeu = -1;
        for (var k = 3; k < dataKeu.length; k++) {
          if (dataKeu[k][0] === data.kodeId) {
            targetRowKeu = k + 1;
            break;
          }
        }

        var barisKeuLengkap = [
          dataInput[i][0],                             // A (1): NOMOR SPK
          dataInput[i][1],                             // B (2): PERIODE / TAHUN INPUT
          dataInput[i][2],                             // C (3): LOKASI
          dataInput[i][3],                             // D (4): STATUS SEWA
          dataInput[i][21] || "",                      // E (5): Tanggal Habis Kontrak Lama
          dataInput[i][23] || 0,                       // F (6): Harga Sewa Lama Net
          dataInput[i][24] || 0,                       // G (7): Deposit Lama
          dataInput[i][5] || 0,                        // H (8): Harga Pokok Sewa Baru
          dataInput[i][6] || 0,                        // I (9): Pajak Sewa (Dipotong)
          dataInput[i][7] || 0,                        // J (10): Harga Net Ditransfer
          dataInput[i][9] || "",                       // K (11): Tanggal Habis Kontrak Baru
          dataInput[i][10] || 0,                       // L (12): Gross Up
          dataInput[i][7] || 0,                        // M (13): Fix Sewa Bersih
          dataInput[i][6] || 0,                        // N (14): Fix Pajak
          dataInput[i][38] || "",                      // O (15): Nama Pemilik
          "",                                          // P (16): Nomor NIK
          "",                                          // Q (17): NPWP
          dataInput[i][7] || 0,                        // R (18): Sewa Ditransfer
          "ADA",                                       // S (19): Soft File SPK
          dataInput[i][18] || "",                      // T (20): Link Soft SPK
          hardcopyStatus,                              // U (21): Hardcopy SPK
          "",                                          // V (22): Masa Pajak
          data.tglBayarPajak,                          // W (23): Tanggal Bayar
          data.ntpnPajak,                              // X (24): Nomor NTPN
          nilaiPajakRealNum,                           // Y (25): Nilai Pajak Real
          bukpotCellContent                            // Z (26): ARSIP BUKPOT
        ];

        if (targetRowKeu !== -1) {
          sheetKeu.getRange(targetRowKeu, 1, 1, barisKeuLengkap.length).setValues([barisKeuLengkap]);
        } else {
          sheetKeu.appendRow(barisKeuLengkap);
        }

        // 4. SINKRONISASI KE SHEET TERPISAH (NEW_REKAP)
        syncKeRekap(data.kodeId);

        // 5. Notifikasi Email Feedback
        kirimEmailFeedbackPencairan(data.kodeId, dataInput[i][2], data.ntpnPajak, dataInput[i][7]);

        return "Pencairan dana berhasil dicatat di NEW_KEU dan disinkronkan ke NEW_REKAP.";
      }
    }
    throw new Error("ID Pengajuan tidak ditemukan!");
  } finally { lock.releaseLock(); }
}

// ==========================================
// ENGINE SINKRONISASI MANIFEST NEW_REKAP
// ==========================================
function syncKeRekap(kodeId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInput = ss.getSheetByName("NEW_INPUT");
  var sheetRekap = ss.getSheetByName("NEW_REKAP");

  if (!sheetInput || !sheetRekap) return;

  var dataInput = sheetInput.getDataRange().getValues();
  var dataRekap = sheetRekap.getDataRange().getValues();

  for (var i = 4; i < dataInput.length; i++) {
    if (dataInput[i][0] === kodeId) {

      var hardcopyFix = dataInput[i][32] || "BELUM";
      var ntpnFix = dataInput[i][34] || "-";
      var nilaiPajakRealFix = dataInput[i][35] || 0;
      var infoBankFix = dataInput[i][31] || "-";

      // Memetakan 17 Kolom NEW_REKAP (A-Q)
      var barisRekap = [
        dataInput[i][1],            // A: PERIODE
        dataInput[i][2],            // B: WILAYAH / LOKASI
        dataInput[i][3],            // C: STATUS SEWA
        dataInput[i][21] || "-",    // D: Tanggal Habis Kontrak Lama
        dataInput[i][22] || 0,      // E: Masa Sewa Lama
        dataInput[i][23] || 0,      // F: Harga Lama Net
        dataInput[i][5] || 0,       // G: Harga Pokok Sewa Baru
        dataInput[i][6] || 0,       // H: Pajak Sewa
        dataInput[i][7] || 0,       // I: Harga Ditransfer Net
        parseInt(dataInput[i][8]) || 1, // J: Masa Sewa Baru
        dataInput[i][9],            // K: Tanggal Habis Baru
        dataInput[i][4],            // L: Pihak Pembayar Pajak PBB
        infoBankFix,                // M: INFO REKENING BANK
        "ADA",                      // N: Soft File SPK
        hardcopyFix,                // O: Hard File SPK
        ntpnFix,                    // P: Nomor NTPN
        nilaiPajakRealFix           // Q: Nilai Pajak Real
      ];

      // Cek apakah data sudah ada di NEW_REKAP (Cari berdasarkan Wilayah/Lokasi & Periode)
      var targetRowRekap = -1;
      for (var r = 5; r < dataRekap.length; r++) {
        if (dataRekap[r][1] === dataInput[i][2] && dataRekap[r][0] === dataInput[i][1]) {
          targetRowRekap = r + 1;
          break;
        }
      }

      if (targetRowRekap !== -1) {
        sheetRekap.getRange(targetRowRekap, 1, 1, barisRekap.length).setValues([barisRekap]);
      } else {
        sheetRekap.appendRow(barisRekap);
      }

      sheetRekap.getRange(sheetRekap.getLastRow(), 10).setNumberFormat("0");
      break;
    }
  }
}

// ==========================================
// DASHBOARD EXECUTIVE (BACA DARI NEW_REKAP)
// ==========================================
function getExecutiveDashboardData() {
  var roles = getUserDetails(Session.getActiveUser().getEmail()).roles;
  if (!roles.includes("Informasi") && !roles.includes("Direktur") && !roles.includes("Admin")) throw new Error("Akses Ditolak.");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("NEW_REKAP");
  if (!sheet) return { stats: {}, listData: [] };

  var data = sheet.getDataRange().getValues();
  var listData = [];
  var totalBiayaSewa = 0;
  var totalPajak = 0;
  var countPerpanjang = 0;
  var countBaru = 0;
  var countRelokasi = 0;

  // Membaca mulai dari baris ke-16 atau indeks 15 (di bawah header NEW_REKAP)
  for (var i = 5; i < data.length; i++) {
    if (data[i][1] && String(data[i][1]).trim() !== "") {
      var hrgBersih = parseFloat(data[i][8]) || 0; // Kolom I
      var pjk = parseFloat(data[i][7]) || 0;       // Kolom H
      var statusSewa = String(data[i][2] || "").trim(); // Kolom C

      totalBiayaSewa += hrgBersih;
      totalPajak += pjk;

      if (statusSewa === "Perpanjang") countPerpanjang++;
      else if (statusSewa === "Kontrak Baru") countBaru++;
      else if (statusSewa === "Relokasi") countRelokasi++;

      listData.push({
        periode: data[i][0] || "-",
        lokasi: data[i][1] || "-",
        statusSewa: statusSewa || "-",
        hargaPokok: data[i][6] || 0,
        pajak: pjk,
        hargaBersih: hrgBersih,
        masaSewa: data[i][9] || 1,
        tglHabisBaru: data[i][10] instanceof Date ? data[i][10].toISOString().split('T')[0] : (data[i][10] || "-"),
        ntpn: data[i][15] || "-"
      });
    }
  }

  var stats = {
    totalLokasi: listData.length,
    totalBiayaSewa: totalBiayaSewa,
    totalPajak: totalPajak,
    countPerpanjang: countPerpanjang,
    countBaru: countBaru,
    countRelokasi: countRelokasi
  };

  return { stats: stats, listData: listData };
}

// ==========================================
// 6. HELPER DOKUMEN & ADMIN MASTER DATA
// ==========================================
function getDaftarLokasi() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LOKASI");
  if (!sheet) return []; var data = sheet.getDataRange().getValues(); var list = [];
  for (var i = 3; i < data.length; i++) { if (data[i][4]) list.push(data[i][4]); } return list;
}

function buatMemoPDF(data, nomorPengajuan) {
  var copy = DriveApp.getFileById('1DF7P2-nImGOSjHrzd4UEQU0nuXbCiF2hqnFEtx7CUYA').makeCopy('MEMO_' + nomorPengajuan, DriveApp.getFolderById('1i1iYPvDayVa_WXre_mMPsmlP_UEGdeMd'));
  var doc = DocumentApp.openById(copy.getId()); var body = doc.getBody();
  var hargaBaru = parseFloat(String(data.hargaBaru || "0").replace(/\./g, '')) || 0;
  var pajak = data.dipotongPajak.includes("Ya") ? (hargaBaru * 0.10) : ((hargaBaru / 0.9) * 0.10);
  var alamat = "-";
  try {
    var v = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LOKASI").getDataRange().getValues();
    var idxLok = v[2].indexOf("NAMA LOKASI") !== -1 ? v[2].indexOf("NAMA LOKASI") : 4; var idxAlm = v[2].indexOf("ALAMAT") !== -1 ? v[2].indexOf("ALAMAT") : 6;
    for (var i = 3; i < v.length; i++) { if (v[i][idxLok] === data.lokasi) { alamat = v[i][idxAlm]; break; } }
  } catch (e) { }
  body.replaceText('<<NOMOR_PENGAJUAN>>', nomorPengajuan); body.replaceText('<<PERIHAL>>', data.statusSewa + " - Sewa");
  body.replaceText('<<TANGGAL>>', new Date().toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' }));
  body.replaceText('<<LOKASI>>', data.lokasi); body.replaceText('<<ALAMAT>>', alamat);
  body.replaceText('<<TGL_HABIS>>', data.tglAkhirBaru ? data.tglAkhirBaru.split("-").reverse().join("/") : "-");
  var finalValueStr = data.nilaiFinal ? data.nilaiFinal.replace("Rp ", "") : Math.round(data.dipotongPajak.includes("Ya") ? hargaBaru : (hargaBaru / 0.9)).toLocaleString('id-ID');
  body.replaceText('<<NILAI_FINAL>>', finalValueStr);
  body.replaceText('<<PAJAK>>', Math.round(pajak).toLocaleString('id-ID'));
  body.replaceText('<<MASA_SEWA>>', (data.masaSewa || 1) + " Tahun"); body.replaceText('<<CATATAN>>', data.catatanMemo || "-");
  doc.saveAndClose(); var pdf = DriveApp.getFolderById('1i1iYPvDayVa_WXre_mMPsmlP_UEGdeMd').createFile(copy.getAs('application/pdf'));
  copy.setTrashed(true); return { url: pdf.getUrl(), id: pdf.getId() };
}

function checkAndPrepareSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheetKeu = ss.getSheetByName("NEW_KEU");
  if (!sheetKeu) {
    sheetKeu = ss.insertSheet("NEW_KEU");
    var headerRow1 = ["PERIODE", "NOMOR SPK / TAHUN INPUT", "", "LOKASI", "STATUS SEWA", "INFORMASI SEWA SEBELUMNYA", "", "", "INFORMASI SEWA MENDATANG", "", "", "", "Fix Nilai (Arsip)", "", "NIK", "", "NPWP", "Sewa Ditransfer", "Arsip SPK", "", "", "", "INFO PAJAK", "", "", "", "ARSIP BUKPOT"];
    var headerRow2 = ["", "", "", "", "", "Tanggal Habis Kontrak", "Harga Setelah Pajak (Ditransfer Bersih)", "Deposit", "Harga Sebelum Pajak (Pokok) Sewa", "Pajak Sewa (Yang dipotong)", "Harga Setelah Pajak (Ditransfer Bersih)", "Tanggal Habis Kontrak", "Penilaian Harga Sewa (Gross Up)", "Sewa Bersih", "Pajak", "Nama", "Nomor", "", "Soft", "Link", "Hard", "Masa Pajak", "Tanggal Bayar", "Nomor NTPN", "Nilai Pajak", ""];
    sheetKeu.getRange(2, 1, 1, headerRow1.length).setValues([headerRow1]).setBackground("#dcfce7").setFontWeight("bold");
    sheetKeu.getRange(3, 1, 1, headerRow2.length).setValues([headerRow2]).setBackground("#f1f5f9").setFontWeight("bold");
  }

  var sheetRekap = ss.getSheetByName("NEW_REKAP");
  if (!sheetRekap) {
    sheetRekap = ss.insertSheet("NEW_REKAP");
    var headerRekap = ["PERIODE", "WILAYAH / LOKASI", "STATUS SEWA", "INFORMASI SEWA SEBELUMNYA", "", "", "INFORMASI SEWA MENDATANG", "", "", "", "", "", "Pihak Pembayar Pajak PBB", "INFO REKENING BANK (Pemilik)", "Arsip SPK", "", "INFO PAJAK", ""];
    var headerRekapSub = ["", "", "", "Tanggal Habis Kontrak", "Masa Sewa (tahun)", "Harga Setelah Pajak (Ditransfer Bersih)", "Harga Sebelum Pajak (Pokok) Sewa", "Pajak Sewa (Yang dipotong)", "Harga Setelah Pajak (Ditransfer Bersih)", "Masa Sewa (tahun)", "Tanggal Habis Kontrak", "", "", "", "Soft", "Hard", "Nomor NTPN", "Nilai Pajak"];
    sheetRekap.getRange(4, 1, 1, headerRekap.length).setValues([headerRekap]).setBackground("#2e1065").setFontColor("#ffffff").setFontWeight("bold");
    sheetRekap.getRange(5, 1, 1, headerRekapSub.length).setValues([headerRekapSub]).setBackground("#f1f5f9").setFontWeight("bold");
  }
}

function getAdminMasterData() {
  var sessionEmail = Session.getActiveUser().getEmail();
  var roles = getUserDetails(sessionEmail).roles;
  if (!roles.includes("Admin")) throw new Error("Akses Ditolak: Anda bukan Admin.");

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheetUser = ss.getSheetByName("DATA_USER");
  var dataUser = sheetUser ? sheetUser.getDataRange().getValues() : [];
  var listUser = [];
  for (var i = 1; i < dataUser.length; i++) {
    if (dataUser[i][0] !== "") {
      listUser.push({
        rowIdx: i + 1,
        email: String(dataUser[i][0]).trim(),
        nama: String(dataUser[i][1]).trim(),
        role: String(dataUser[i][2]).trim(),
        status: String(dataUser[i][3]).trim() || "Aktif"
      });
    }
  }

  var sheetLokasi = ss.getSheetByName("LOKASI");
  var dataLokasi = sheetLokasi ? sheetLokasi.getDataRange().getValues() : [];
  var listLokasi = [];
  for (var j = 3; j < dataLokasi.length; j++) {
    if (dataLokasi[j][4] !== "") {
      listLokasi.push({
        rowIdx: j + 1,
        namaLokasi: String(dataLokasi[j][4]).trim(),
        alamat: String(dataLokasi[j][6] || "-").trim()
      });
    }
  }

  return { listUser: listUser, listLokasi: listLokasi };
}

function simpanUserMaster(data) {
  var sessionEmail = Session.getActiveUser().getEmail();
  var roles = getUserDetails(sessionEmail).roles;
  if (!roles.includes("Admin")) throw new Error("Akses Ditolak.");

  if (data.email.trim().toLowerCase() === sessionEmail.trim().toLowerCase() && data.role === "Admin" && data.status === "Nonaktif") {
    throw new Error("Ditolak: Anda tidak dapat menonaktifkan akses Admin milik diri sendiri!");
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DATA_USER");
  if (data.rowIdx) {
    sheet.getRange(data.rowIdx, 1, 1, 4).setValues([[data.email.trim(), data.nama.trim(), data.role.trim(), data.status]]);
  } else {
    sheet.appendRow([data.email.trim(), data.nama.trim(), data.role.trim(), data.status]);
  }
  return "Data user berhasil disimpan.";
}

function hapusUserMaster(rowIdx, targetEmail, targetRole) {
  var sessionEmail = Session.getActiveUser().getEmail();
  var roles = getUserDetails(sessionEmail).roles;
  if (!roles.includes("Admin")) throw new Error("Akses Ditolak.");

  if (targetEmail.trim().toLowerCase() === sessionEmail.trim().toLowerCase() && targetRole === "Admin") {
    throw new Error("Ditolak: Anda tidak dapat menghapus akses Admin milik diri sendiri!");
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DATA_USER");
  sheet.deleteRow(rowIdx);
  return "Data user berhasil dihapus.";
}

// =========================================================================
// FASE 8 & 9: DB SYNC, AUTOMATED NOTIFICATION & EXECUTIVE DASHBOARD
// =========================================================================

function getDaftarLokasiEksternal(spreadsheetIdEksternal) {
  if (spreadsheetIdEksternal && spreadsheetIdEksternal.trim() !== "") {
    try {
      var ssEksternal = SpreadsheetApp.openById(spreadsheetIdEksternal.trim());
      var sheetLokasi = ssEksternal.getSheetByName("LOKASI");
      if (sheetLokasi) {
        var data = sheetLokasi.getDataRange().getValues();
        var listLokasi = [];
        for (var i = 3; i < data.length; i++) {
          if (data[i][4] && String(data[i][4]).trim() !== "") {
            listLokasi.push(String(data[i][4]).trim());
          }
        }
        if (listLokasi.length > 0) return listLokasi;
      }
    } catch (e) {
      Logger.log("Gagal mengambil DB Eksternal, menggunakan DB Lokal: " + e.message);
    }
  }
  return getDaftarLokasi();
}

function cekJatuhTempoKontrak() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("NEW_INPUT");
  if (!sheet) return "Sheet NEW_INPUT tidak ditemukan.";

  var data = sheet.getDataRange().getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var terkirimCount = 0;

  for (var i = 4; i < data.length; i++) {
    var kodeId = data[i][0];
    var lokasi = data[i][2];
    var statusDraft = String(data[i][37] || "").trim();
    var tglHabisRaw = data[i][9];
    var emailNegosiator = String(data[i][40] || "").trim();

    if (kodeId !== "" && statusDraft === "FINAL" && tglHabisRaw) {
      try {
        var tglHabis = new Date(tglHabisRaw);
        tglHabis.setHours(0, 0, 0, 0);

        var diffTime = tglHabis.getTime() - today.getTime();
        var diffDays = Math.round(diffTime / (1000 * 3600 * 24));

        if (diffDays === 60 || diffDays === 30 || diffDays === 14) {
          if (emailNegosiator && emailNegosiator.includes("@")) {
            var subject = "⚠️ PERINGATAN SEWA (H-" + diffDays + "): " + lokasi + " [" + kodeId + "]";
            var tglHabisStr = tglHabis.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
            var bodyHtml =
              "<div style='font-family: Arial, sans-serif; padding: 20px; border: 1px solid #e2e8f0; border-radius: 8px; max-width: 600px;'>" +
              "<h3 style='color: #d97706; margin-top: 0;'>⚠️ Peringatan Masa Sewa Berakhir</h3>" +
              "<p>Halo Negosiator,</p>" +
              "<p>Kontrak sewa untuk lokasi berikut akan berakhir dalam <b>" + diffDays + " hari</b>:</p>" +
              "<table style='width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 13px;'>" +
              "<tr><td style='padding: 6px; border: 1px solid #cbd5e1; background: #f8fafc;'><b>Kode ID Transaksi</b></td><td style='padding: 6px; border: 1px solid #cbd5e1;'>" + kodeId + "</td></tr>" +
              "<tr><td style='padding: 6px; border: 1px solid #cbd5e1; background: #f8fafc;'><b>Lokasi</b></td><td style='padding: 6px; border: 1px solid #cbd5e1;'>" + lokasi + "</td></tr>" +
              "<tr><td style='padding: 6px; border: 1px solid #cbd5e1; background: #f8fafc;'><b>Tanggal Habis Kontrak</b></td><td style='padding: 6px; border: 1px solid #cbd5e1; color: #dc2626;'><b>" + tglHabisStr + "</b></td></tr>" +
              "</table>" +
              "<p>Mohon untuk segera berkoordinasi dengan pihak pemilik lokasi untuk menentukan proses <b>Perpanjang</b> atau <b>Relokasi</b>.</p>" +
              "<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 20px 0;'>" +
              "<p style='font-size: 11px; color: #64748b;'>Pesan ini dikirim otomatis oleh Engine SIMPEWA System.</p>" +
              "</div>";

            MailApp.sendEmail({
              to: emailNegosiator,
              subject: subject,
              htmlBody: bodyHtml
            });
            terkirimCount++;
          }
        }
      } catch (errEmail) {
        Logger.log("Gagal mengirim email notifikasi ID " + kodeId + ": " + errEmail.message);
      }
    }
  }

  return "Pengecekan selesai. " + terkirimCount + " email notifikasi peringatan berhasil dikirim.";
}

function updateStatusHardcopySPK(kodeId, statusHardcopy) {
  var roles = getUserDetails(Session.getActiveUser().getEmail()).roles;
  if (!roles.includes("Keuangan") && !roles.includes("Admin")) throw new Error("Akses Ditolak.");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("Sistem sibuk."); }

  try {
    var sheetInput = ss.getSheetByName("NEW_INPUT");
    if (sheetInput) {
      var dataInput = sheetInput.getDataRange().getValues();
      for (var i = 4; i < dataInput.length; i++) {
        if (dataInput[i][0] === kodeId) {
          sheetInput.getRange(i + 1, 33).setValue(statusHardcopy);
          break;
        }
      }
    }

    var sheetKeu = ss.getSheetByName("NEW_KEU");
    if (sheetKeu) {
      var dataKeu = sheetKeu.getDataRange().getValues();
      for (var j = 3; j < dataKeu.length; j++) {
        if (dataKeu[j][0] === kodeId) {
          sheetKeu.getRange(j + 1, 21).setValue(statusHardcopy);
          break;
        }
      }
    }

    return "Status Hardcopy SPK Fisik untuk " + kodeId + " berhasil diperbarui secara konsisten di seluruh sheet menjadi: " + statusHardcopy;
  } finally { lock.releaseLock(); }
}

function setupTriggerJatuhTempoOtomatis() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (trigger) {
    if (trigger.getHandlerFunction() === "cekJatuhTempoKontrak") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("cekJatuhTempoKontrak")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  return "Trigger Otomatis Peringatan Jatuh Tempo berhasil diaktifkan (Setiap Hari Pukul 08:00 WIB).";
}


function kirimEmailFeedbackPencairan(kodeId, lokasi, ntpn, nominalNet) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetInput = ss.getSheetByName("NEW_INPUT");
    if (!sheetInput) return;

    var dataInput = sheetInput.getDataRange().getValues();
    var emailNegosiator = "";

    for (var i = 4; i < dataInput.length; i++) {
      if (dataInput[i][0] === kodeId) {
        emailNegosiator = String(dataInput[i][40] || "").trim();
        break;
      }
    }

    if (emailNegosiator && emailNegosiator.includes("@")) {
      var subject = "💰 SEWA LUNAS & DICAIRKAN: " + lokasi + " [" + kodeId + "]";
      var bodyHtml =
        "<div style='font-family: Arial, sans-serif; padding: 20px; border: 1px solid #10b981; border-radius: 8px; max-width: 600px;'>" +
        "<h3 style='color: #059669; margin-top: 0;'>🎉 Dana Sewa Telah Dicairkan!</h3>" +
        "<p>Halo Negosiator,</p>" +
        "<p>Pengajuan sewa lokasi Anda telah selesai diproses oleh Divisi Keuangan. Pembayaran dan NTPN Pajak telah terbit.</p>" +
        "<table style='width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 13px;'>" +
        "<tr><td style='padding: 6px; border: 1px solid #cbd5e1; background: #f8fafc;'><b>Kode ID</b></td><td style='padding: 6px; border: 1px solid #cbd5e1;'>" + kodeId + "</td></tr>" +
        "<tr><td style='padding: 6px; border: 1px solid #cbd5e1; background: #f8fafc;'><b>Lokasi</b></td><td style='padding: 6px; border: 1px solid #cbd5e1;'>" + lokasi + "</td></tr>" +
        "<tr><td style='padding: 6px; border: 1px solid #cbd5e1; background: #f8fafc;'><b>Nominal Transfer (Net)</b></td><td style='padding: 6px; border: 1px solid #cbd5e1; color: #059669;'><b>Rp " + Math.round(nominalNet).toLocaleString('id-ID') + "</b></td></tr>" +
        "<tr><td style='padding: 6px; border: 1px solid #cbd5e1; background: #f8fafc;'><b>Nomor NTPN Pajak</b></td><td style='padding: 6px; border: 1px solid #cbd5e1;'><b>" + ntpn + "</b></td></tr>" +
        "</table>" +
        "<p>Status transaksi saat ini telah diperbarui menjadi <b>LUNAS (SELESAI)</b> pada Ledger Rekapitalisasi.</p>" +
        "<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 20px 0;'>" +
        "<p style='font-size: 11px; color: #64748b;'>Pesan otomatis dari SIMPEWA System Feedback Loop.</p>" +
        "</div>";

      MailApp.sendEmail({
        to: emailNegosiator,
        subject: subject,
        htmlBody: bodyHtml
      });
    }
  } catch (e) {
    Logger.log("Gagal mengirim email feedback pencairan: " + e.message);
  }
}