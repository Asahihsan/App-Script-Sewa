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
// Baris pertama data transaksi di NEW_INPUT (0-based index array getValues()).
// Sheet punya 3 baris header (grup/sub-grup/leaf) + 1 baris metadata nomor
// urut kolom, jadi data mulai di baris ke-5 (index 4).
var NEW_INPUT_DATA_START_ROW = 4;
var NEW_INPUT_HEADER_ROWS = 3;

function simpanPengajuanBaru(data, statusSubmit) {
  var roles = getUserDetails(Session.getActiveUser().getEmail()).roles;
  if (!roles.includes("Negosiator") && !roles.includes("Admin")) throw new Error("Akses Ditolak: Hak akses Negosiator dibutuhkan.");

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("Sistem sibuk."); }

  try {
    var emailNegosiator = Session.getActiveUser().getEmail();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("NEW_INPUT");
    var colMap = getColIndexMapAuto(sheet, NEW_INPUT_HEADER_ROWS);
    var lastCol = sheet.getLastColumn();
    var idx = function (headerName) { return getColIndex(colMap, headerName); };

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
    var findData = sheet.getDataRange().getValues();

    if (primaryKey && primaryKey !== "") {
      isEdit = true;
    } else {
      var maxSeq = 0;
      var kodeLokasi = data.lokasi ? data.lokasi.substring(0, 3).toUpperCase() : "REL";
      var tahun = new Date().getFullYear();

      for (var r = NEW_INPUT_DATA_START_ROW; r < findData.length; r++) {
        var existingId = String(findData[r][idx('KODE ID')] || "");
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

    // Baris ditulis lewat peta nama header, bukan posisi angka -- kalau
    // suatu saat ada kolom baru disisipkan lagi di sheet, baris ini TIDAK
    // perlu diubah, karena getColIndex akan cari ulang posisinya otomatis.
    var rowArr = new Array(lastCol).fill("");
    var set = function (headerName, value) { rowArr[idx(headerName)] = value; };

    set('KODE ID', primaryKey);
    set('PERIODE / TAHUN INPUT', new Date().getFullYear());
    set('LOKASI', data.lokasi || "");
    set('STATUS SEWA', data.statusSewa || "Kontrak Baru");
    set('PIHAK PEMBAYAR PAJAK PBB', data.pembayarPBB || "Pemilik");
    set('NIK', data.nik || "");
    set('NPWP', data.npwp || "");
    set('HARGA SEBELUM PAJAK (POKOK) SEWA', hargaBaru);
    set('PAJAK SEWA (YANG DIPOTONG)', nominalPajak);
    set('HARGA SETELAH PAJAK (DITRANSFER BERSIH)', hargaSetelahPajak);
    set('MASA SEWA (TAHUN)', parseInt(data.masaSewa) || 1);
    set('TANGGAL HABIS KONTRAK', data.tglAkhirBaru || "");
    set('PENILAIAN HARGA SEWA (GROSS UP)', nilaiGrossUp);
    set('LUNAS / CICILAN 1 - TANGGAL', data.tglCicilan1 || "");
    set('LUNAS / CICILAN 1 - JUMLAH', parseFloat(String(data.jmlCicilan1 || "0").replace(/\./g, '')) || 0);
    set('LUNAS / CICILAN 2 - TANGGAL', data.tglCicilan2 || "");
    set('LUNAS / CICILAN 2 - JUMLAH', parseFloat(String(data.jmlCicilan2 || "0").replace(/\./g, '')) || 0);
    set('LUNAS / CICILAN 3 - TANGGAL', data.tglCicilan3 || "");
    set('LUNAS / CICILAN 3 - JUMLAH', parseFloat(String(data.jmlCicilan3 || "0").replace(/\./g, '')) || 0);
    set('JUMLAH PEMBAYARAN DEPOSIT (JIKA ADA)', parseFloat(String(data.depositBaru || "0").replace(/\./g, '')) || 0);
    set('SPK (LINK DRIVE)', spkCellContent);
    set('MEMO (LINK DRIVE)', memoCellContent);
    set('SEWA LAMA - TAHUN / PERIODE', tahunSewaLama);
    set('SEWA LAMA - TANGGAL HABIS KONTRAK', data.tanggalHabisLama || "");
    set('SEWA LAMA - MASA SEWA (TAHUN)', parseInt(data.masaSewaLama) || 0);
    set('SEWA LAMA - HARGA SETELAH PAJAK (NET)', parseFloat(String(data.hargaLama || "0").replace(/\./g, '')) || 0);
    set('SEWA LAMA - DEPOSIT', parseFloat(String(data.depositLama || "0").replace(/\./g, '')) || 0);
    set('STATUS MCX', statusSubmit === "FINAL" ? "PENDING" : "");
    set('INFO REKENING BANK', infoRekeningBank);
    set('STATUS DRAF', statusSubmit);
    set('NAMA PEMILIK', data.namaPemilik || "");
    set('LOKASI LAMA', data.lokasiLama || "");
    set('EMAIL NEGOSIATOR', emailNegosiator);
    // Kolom lain (CATATAN MCX, STATUS/CATATAN SEKRETARIS & DIREKTUR,
    // HARDCOPY SPK, info pencairan pajak, dst) sengaja dibiarkan kosong --
    // itu wilayah role selanjutnya di alur approval, bukan Negosiator.

    if (isEdit) {
      for (var row = NEW_INPUT_DATA_START_ROW; row < findData.length; row++) {
        if (findData[row][idx('KODE ID')] === primaryKey) {
          sheet.getRange(row + 1, 1, 1, rowArr.length).setValues([rowArr]);
          break;
        }
      }
    } else {
      sheet.appendRow(rowArr);
    }

    return primaryKey;
  } finally { lock.releaseLock(); }
}

function getRiwayatPengajuan() {
  var emailLogin = Session.getActiveUser().getEmail().trim().toLowerCase();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
  if (!sheet) return [];

  var colMap = getColIndexMapAuto(sheet, NEW_INPUT_HEADER_ROWS);
  var idx = function (headerName) { return getColIndex(colMap, headerName); };
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = NEW_INPUT_DATA_START_ROW; i < data.length; i++) {
    var emailPembuat = String(data[i][idx('EMAIL NEGOSIATOR')] || "").trim().toLowerCase();

    if (data[i][idx('KODE ID')] !== "" && emailPembuat === emailLogin) {
      result.push({
        kodeId: data[i][idx('KODE ID')],
        lokasi: data[i][idx('LOKASI')],
        statusSewa: data[i][idx('STATUS SEWA')],
        hargaBaru: data[i][idx('HARGA SEBELUM PAJAK (POKOK) SEWA')],
        statusDraft: data[i][idx('STATUS DRAF')] || "FINAL",
        statusMcx: data[i][idx('STATUS MCX')] || "-",
        catatanMcx: data[i][idx('CATATAN MCX')] || "-",
        statusSekretaris: data[i][idx('STATUS SEKRETARIS')],
        catatanSekretaris: data[i][idx('CATATAN SEKRETARIS')] || "-",
        statusDirektur: data[i][idx('STATUS DIREKTUR')],
        ntpn: data[i][idx('NOMOR NTPN RESMI')]
      });
    }
  }
  return result;
}

function getDetailPengajuan(kodeId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
  if (!sheet) return null;

  var colMap = getColIndexMapAuto(sheet, NEW_INPUT_HEADER_ROWS);
  var idx = function (headerName) { return getColIndex(colMap, headerName); };
  var data = sheet.getDataRange().getValues();

  for (var i = NEW_INPUT_DATA_START_ROW; i < data.length; i++) {
    if (data[i][idx('KODE ID')] === kodeId) {
      var hrgBaru = parseFloat(data[i][idx('HARGA SEBELUM PAJAK (POKOK) SEWA')]) || 0;
      var grsUp = parseFloat(data[i][idx('PENILAIAN HARGA SEWA (GROSS UP)')]) || 0;
      var fmtTgl = function (v) { return (v instanceof Date) ? v.toISOString().split('T')[0] : (v || ""); };

      var infoRek = String(data[i][idx('INFO REKENING BANK')] || "");
      var bankNama = "", bankNoRek = "", bankAn = "";
      if (infoRek.includes(" - ") && infoRek.includes(" a.n ")) {
        var parts1 = infoRek.split(" - ");
        bankNama = parts1[0];
        var parts2 = parts1[1].split(" a.n ");
        bankNoRek = parts2[0];
        bankAn = parts2[1];
      }

      var tglAkhirFix = fmtTgl(data[i][idx('TANGGAL HABIS KONTRAK')]);
      var masaSewaFix = parseInt(data[i][idx('MASA SEWA (TAHUN)')]) || 1;
      var tglMulaiFix = "";
      if (tglAkhirFix && masaSewaFix) {
        var d = new Date(tglAkhirFix);
        d.setFullYear(d.getFullYear() - masaSewaFix);
        d.setDate(d.getDate() + 1);
        tglMulaiFix = d.toISOString().split('T')[0];
      }

      return {
        kodeId: data[i][idx('KODE ID')],
        lokasi: data[i][idx('LOKASI')],
        statusSewa: data[i][idx('STATUS SEWA')],
        pembayarPBB: data[i][idx('PIHAK PEMBAYAR PAJAK PBB')] || "Pemilik",
        nik: data[i][idx('NIK')] || "",
        npwp: data[i][idx('NPWP')] || "",
        hargaBaru: hrgBaru,
        dipotongPajak: (Math.round(grsUp) === Math.round(hrgBaru)) ? "Ya, dipotong" : "Tidak (Gross Up)",
        masaSewaBaru: masaSewaFix,
        tglMulaiBaru: tglMulaiFix,
        tglAkhirBaru: tglAkhirFix,
        tglCicilan1: fmtTgl(data[i][idx('LUNAS / CICILAN 1 - TANGGAL')]),
        jmlCicilan1: data[i][idx('LUNAS / CICILAN 1 - JUMLAH')] || 0,
        tglCicilan2: fmtTgl(data[i][idx('LUNAS / CICILAN 2 - TANGGAL')]),
        jmlCicilan2: data[i][idx('LUNAS / CICILAN 2 - JUMLAH')] || 0,
        tglCicilan3: fmtTgl(data[i][idx('LUNAS / CICILAN 3 - TANGGAL')]),
        jmlCicilan3: data[i][idx('LUNAS / CICILAN 3 - JUMLAH')] || 0,
        depositBaru: data[i][idx('JUMLAH PEMBAYARAN DEPOSIT (JIKA ADA)')] || 0,
        spkDokumen: data[i][idx('SPK (LINK DRIVE)')] || "",
        memoDokumen: data[i][idx('MEMO (LINK DRIVE)')] || "",
        tahunSewaLama: data[i][idx('SEWA LAMA - TAHUN / PERIODE')] || "",
        tanggalHabisLama: fmtTgl(data[i][idx('SEWA LAMA - TANGGAL HABIS KONTRAK')]),
        masaSewaLama: data[i][idx('SEWA LAMA - MASA SEWA (TAHUN)')] || "",
        hargaLama: data[i][idx('SEWA LAMA - HARGA SETELAH PAJAK (NET)')] || 0,
        depositLama: data[i][idx('SEWA LAMA - DEPOSIT')] || 0,
        catatanPenolakan: data[i][idx('CATATAN DIREKTUR')] ? data[i][idx('CATATAN DIREKTUR')] :
          (data[i][idx('CATATAN SEKRETARIS')] ? data[i][idx('CATATAN SEKRETARIS')] :
          (data[i][idx('CATATAN MCX')] || "Tidak ada catatan.")),
        bankNama: bankNama, bankNoRek: bankNoRek, bankAn: bankAn,
        statusDraft: data[i][idx('STATUS DRAF')] || "FINAL",
        namaPemilik: data[i][idx('NAMA PEMILIK')] || "",
        lokasiLama: data[i][idx('LOKASI LAMA')] || ""
      };
    }
  }
  return null;
}

function batalkanPengajuan(kodeId) {
  var lock = LockService.getScriptLock(); try { lock.waitLock(10000); } catch (e) { throw new Error("Sistem sibuk."); }
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
    var colMap = getColIndexMapAuto(sheet, NEW_INPUT_HEADER_ROWS);
    var idx = function (headerName) { return getColIndex(colMap, headerName); };
    var data = sheet.getDataRange().getValues();

    for (var i = NEW_INPUT_DATA_START_ROW; i < data.length; i++) {
      if (data[i][idx('KODE ID')] === kodeId) {
        var spk = data[i][idx('SPK (LINK DRIVE)')], memo = data[i][idx('MEMO (LINK DRIVE)')];
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

  var colMap = getColIndexMapAuto(sheet, NEW_INPUT_HEADER_ROWS);
  var idx = function (headerName) { return getColIndex(colMap, headerName); };
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = NEW_INPUT_DATA_START_ROW; i < data.length; i++) {
    var statusDraft = String(data[i][idx('STATUS DRAF')] || "").trim();
    var statusMcx = String(data[i][idx('STATUS MCX')] || "").trim();

    if (data[i][idx('KODE ID')] !== "" && statusDraft === "FINAL" && (statusMcx === "PENDING" || statusMcx === "")) {
      var tglAkhir = data[i][idx('TANGGAL HABIS KONTRAK')];
      result.push({
        kodeId: data[i][idx('KODE ID')],
        lokasi: data[i][idx('LOKASI')],
        statusSewa: data[i][idx('STATUS SEWA')],
        pembayarPBB: data[i][idx('PIHAK PEMBAYAR PAJAK PBB')],
        hargaPokok: data[i][idx('HARGA SEBELUM PAJAK (POKOK) SEWA')],
        pajak: data[i][idx('PAJAK SEWA (YANG DIPOTONG)')],
        hargaBersih: data[i][idx('HARGA SETELAH PAJAK (DITRANSFER BERSIH)')],
        masaSewa: data[i][idx('MASA SEWA (TAHUN)')],
        tglAkhirBaru: tglAkhir instanceof Date ? tglAkhir.toISOString().split('T')[0] : tglAkhir,
        nilaiGrossUp: data[i][idx('PENILAIAN HARGA SEWA (GROSS UP)')],
        spkDokumen: data[i][idx('SPK (LINK DRIVE)')],
        memoDokumen: data[i][idx('MEMO (LINK DRIVE)')],
        namaPemilik: data[i][idx('NAMA PEMILIK')],
        lokasiLama: data[i][idx('LOKASI LAMA')] || "-"
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
    var colMap = getColIndexMapAuto(sheet, NEW_INPUT_HEADER_ROWS);
    var idx = function (headerName) { return getColIndex(colMap, headerName); };
    var data = sheet.getDataRange().getValues();

    for (var i = NEW_INPUT_DATA_START_ROW; i < data.length; i++) {
      if (data[i][idx('KODE ID')] === kodeId) {
        var rowNum = i + 1;
        // setValue pakai kolom 1-based, sedangkan idx() 0-based -- jadi +1.
        var setCell = function (headerName, value) {
          sheet.getRange(rowNum, idx(headerName) + 1).setValue(value);
        };

        if (keputusan === "APPROVED") {
          var hargaPokokLama = data[i][idx('HARGA SEBELUM PAJAK (POKOK) SEWA')];
          var grossUpLama = data[i][idx('PENILAIAN HARGA SEWA (GROSS UP)')];
          var hrgFix = parseFloat(String(nominalBaruFix).replace(/\./g, '')) || hargaPokokLama;
          var dipotongPajak = (Math.round(grossUpLama) === Math.round(hargaPokokLama)) ? "Ya, dipotong" : "Tidak (Gross Up)";

          var nominalPajak = Math.round(dipotongPajak.includes("Ya") ? (hrgFix * 0.10) : ((hrgFix / 0.9) * 0.10));
          var hargaSetelahPajak = Math.round(dipotongPajak.includes("Ya") ? (hrgFix - nominalPajak) : hrgFix);
          var nilaiGrossUp = Math.round(dipotongPajak.includes("Ya") ? hrgFix : (hrgFix / 0.9));

          setCell('HARGA SEBELUM PAJAK (POKOK) SEWA', hrgFix);
          setCell('PAJAK SEWA (YANG DIPOTONG)', nominalPajak);
          setCell('HARGA SETELAH PAJAK (DITRANSFER BERSIH)', hargaSetelahPajak);
          setCell('PENILAIAN HARGA SEWA (GROSS UP)', nilaiGrossUp);

          setCell('STATUS MCX', "APPROVED");
          // Jejak waktu untuk NEW_KEU nanti. Kolom "TANGGAL ACC MCX" perlu
          // ditambahkan manual di sheet NEW_INPUT kalau belum ada.
          setCell('TANGGAL ACC MCX', new Date());
        } else {
          setCell('STATUS MCX', "REVISED");
        }

        setCell('CATATAN MCX', catatanMcx || "-");
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

  var colMap = getColIndexMapAuto(sheet, NEW_INPUT_HEADER_ROWS);
  var idx = function (headerName) { return getColIndex(colMap, headerName); };
  var data = sheet.getDataRange().getValues();
  var pendingList = [];
  var countApproved = 0, countRejected = 0;

  for (var i = NEW_INPUT_DATA_START_ROW; i < data.length; i++) {
    if (data[i][idx('KODE ID')] !== "") {
      var statusMcx = String(data[i][idx('STATUS MCX')] || "").trim();
      var statusCek = data[i][idx('STATUS SEKRETARIS')];

      if (statusMcx === "APPROVED") {
        if (statusCek === true) {
          countApproved++;
        } else if (statusCek === false) {
          countRejected++;
        } else {
          var tglHabis = data[i][idx('TANGGAL HABIS KONTRAK')];
          pendingList.push({
            kodeId: data[i][idx('KODE ID')],
            lokasi: data[i][idx('LOKASI')],
            statusSewa: data[i][idx('STATUS SEWA')],
            pembayarPBB: data[i][idx('PIHAK PEMBAYAR PAJAK PBB')],
            hargaPokok: data[i][idx('HARGA SEBELUM PAJAK (POKOK) SEWA')],
            pajak: data[i][idx('PAJAK SEWA (YANG DIPOTONG)')],
            hargaBersih: data[i][idx('HARGA SETELAH PAJAK (DITRANSFER BERSIH)')],
            masaSewa: data[i][idx('MASA SEWA (TAHUN)')],
            tglHabisBaru: tglHabis instanceof Date ? tglHabis.toISOString().split('T')[0] : tglHabis,
            spkDokumen: data[i][idx('SPK (LINK DRIVE)')],
            memoDokumen: data[i][idx('MEMO (LINK DRIVE)')],
            statusCek: statusCek,
            catatan: data[i][idx('CATATAN SEKRETARIS')] || "-",
            namaPemilik: data[i][idx('NAMA PEMILIK')],
            lokasiLama: data[i][idx('LOKASI LAMA')] || "-"
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

  var colMap = getColIndexMapAuto(sheet, NEW_INPUT_HEADER_ROWS);
  var idx = function (headerName) { return getColIndex(colMap, headerName); };
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = NEW_INPUT_DATA_START_ROW; i < data.length; i++) {
    var statusMcx = String(data[i][idx('STATUS MCX')] || "").trim();
    var statusCek = data[i][idx('STATUS SEKRETARIS')];

    if (data[i][idx('KODE ID')] !== "" && statusMcx === "APPROVED" && (statusCek === true || statusCek === false)) {
      var tglHabis = data[i][idx('TANGGAL HABIS KONTRAK')];
      result.push({
        kodeId: data[i][idx('KODE ID')],
        lokasi: data[i][idx('LOKASI')],
        statusSewa: data[i][idx('STATUS SEWA')],
        pembayarPBB: data[i][idx('PIHAK PEMBAYAR PAJAK PBB')],
        hargaPokok: data[i][idx('HARGA SEBELUM PAJAK (POKOK) SEWA')],
        pajak: data[i][idx('PAJAK SEWA (YANG DIPOTONG)')],
        hargaBersih: data[i][idx('HARGA SETELAH PAJAK (DITRANSFER BERSIH)')],
        masaSewa: data[i][idx('MASA SEWA (TAHUN)')],
        tglHabisBaru: tglHabis instanceof Date ? tglHabis.toISOString().split('T')[0] : tglHabis,
        spkDokumen: data[i][idx('SPK (LINK DRIVE)')],
        memoDokumen: data[i][idx('MEMO (LINK DRIVE)')],
        statusCek: statusCek,
        catatan: data[i][idx('CATATAN SEKRETARIS')] || "-"
      });
    }
  }
  return result;
}

// FIX KEAMANAN: sebelumnya fungsi ini bisa dipanggil siapapun tanpa role-check.
// Dicek pemakaiannya: hanya dipanggil dari DashboardSekretaris.html, jadi
// role-check disamakan dengan fungsi Sekretaris lainnya.
function getPengajuanSiapCetak() {
  var roles = getUserDetails(Session.getActiveUser().getEmail()).roles;
  if (!roles.includes("Sekretaris") && !roles.includes("Admin")) throw new Error("Akses Ditolak.");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
  if (!sheet) return [];

  var colMap = getColIndexMapAuto(sheet, NEW_INPUT_HEADER_ROWS);
  var idx = function (headerName) { return getColIndex(colMap, headerName); };
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = NEW_INPUT_DATA_START_ROW; i < data.length; i++) {
    if (data[i][idx('KODE ID')] !== "" && data[i][idx('STATUS DIREKTUR')] === "Disetujui") {
      var spkVal = data[i][idx('SPK (LINK DRIVE)')];
      var memoVal = data[i][idx('MEMO (LINK DRIVE)')];
      result.push({
        kodeId: data[i][idx('KODE ID')],
        lokasi: data[i][idx('LOKASI')],
        statusSewa: data[i][idx('STATUS SEWA')],
        nilaiGrossUp: data[i][idx('PENILAIAN HARGA SEWA (GROSS UP)')] || data[i][idx('HARGA SEBELUM PAJAK (POKOK) SEWA')],
        hargaBaru: data[i][idx('HARGA SEBELUM PAJAK (POKOK) SEWA')],
        catatanDirektur: data[i][idx('CATATAN DIREKTUR')] || "-",
        memoUrl: memoVal ? memoVal.split("|ID:")[0] : "-",
        spkUrl: spkVal ? spkVal.split("|ID:")[0] : "-"
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
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
    var colMap = getColIndexMapAuto(sheet, NEW_INPUT_HEADER_ROWS);
    var idx = function (headerName) { return getColIndex(colMap, headerName); };
    var data = sheet.getDataRange().getValues();

    for (var i = NEW_INPUT_DATA_START_ROW; i < data.length; i++) {
      if (data[i][idx('KODE ID')] === kodeId) {
        sheet.getRange(i + 1, idx('STATUS SEKRETARIS') + 1).setValue(statusKeputusan === "Setuju" ? true : false);
        sheet.getRange(i + 1, idx('CATATAN SEKRETARIS') + 1).setValue(catatanSekretaris || "-");
        if (statusKeputusan === "Setuju") {
          // Jejak waktu untuk NEW_KEU nanti. Kolom "TANGGAL ACC SEKRETARIS"
          // perlu ditambahkan manual di sheet NEW_INPUT kalau belum ada.
          sheet.getRange(i + 1, idx('TANGGAL ACC SEKRETARIS') + 1).setValue(new Date());
        }
        return "Berhasil memperbarui status pengajuan.";
      }
    } throw new Error("ID tidak ditemukan!");
  } finally { lock.releaseLock(); }
}

function getPengajuanDirektur() {
  var roles = getUserDetails(Session.getActiveUser().getEmail()).roles;
  if (!roles.includes("Direktur") && !roles.includes("Admin")) throw new Error("Akses Ditolak.");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
  if (!sheet) return [];

  var colMap = getColIndexMapAuto(sheet, NEW_INPUT_HEADER_ROWS);
  var idx = function (headerName) { return getColIndex(colMap, headerName); };
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = NEW_INPUT_DATA_START_ROW; i < data.length; i++) {
    var statusCek = data[i][idx('STATUS SEKRETARIS')];
    var statusDirektur = data[i][idx('STATUS DIREKTUR')];
    if (data[i][idx('KODE ID')] !== "" && statusCek === true && (!statusDirektur || statusDirektur === "")) {
      var tglHabis = data[i][idx('TANGGAL HABIS KONTRAK')];
      result.push({
        kodeId: data[i][idx('KODE ID')],
        lokasi: data[i][idx('LOKASI')],
        statusSewa: data[i][idx('STATUS SEWA')],
        hargaPokok: data[i][idx('HARGA SEBELUM PAJAK (POKOK) SEWA')],
        nilaiGrossUp: data[i][idx('PENILAIAN HARGA SEWA (GROSS UP)')],
        masaSewa: data[i][idx('MASA SEWA (TAHUN)')],
        tglHabisBaru: tglHabis instanceof Date ? tglHabis.toISOString().split('T')[0] : tglHabis,
        spkDokumen: data[i][idx('SPK (LINK DRIVE)')],
        memoDokumen: data[i][idx('MEMO (LINK DRIVE)')],
        catatanSekretaris: data[i][idx('CATATAN SEKRETARIS')] || "-",
        namaPemilik: data[i][idx('NAMA PEMILIK')],
        lokasiLama: data[i][idx('LOKASI LAMA')] || "-"
      });
    }
  }
  return result;
}

function getExecutiveAnalytics() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
  if (!sheet) return { totalPengajuan: 0, countPending: 0, totalNominalApproved: 0, totalNominalPending: 0 };

  var colMap = getColIndexMapAuto(sheet, NEW_INPUT_HEADER_ROWS);
  var idx = function (headerName) { return getColIndex(colMap, headerName); };
  var data = sheet.getDataRange().getValues();
  var totalPengajuan = 0, countPending = 0;
  var totalApproved = 0, totalPending = 0;

  for (var i = NEW_INPUT_DATA_START_ROW; i < data.length; i++) {
    if (data[i][idx('KODE ID')] !== "" && data[i][idx('STATUS DRAF')] === "FINAL") {
      totalPengajuan++;
      var gross = parseFloat(data[i][idx('PENILAIAN HARGA SEWA (GROSS UP)')]) || parseFloat(data[i][idx('HARGA SEBELUM PAJAK (POKOK) SEWA')]) || 0;
      var statusDir = data[i][idx('STATUS DIREKTUR')];

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

  var colMap = getColIndexMapAuto(sheet, NEW_INPUT_HEADER_ROWS);
  var idx = function (headerName) { return getColIndex(colMap, headerName); };
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = NEW_INPUT_DATA_START_ROW; i < data.length; i++) {
    var statusDir = data[i][idx('STATUS DIREKTUR')];
    if (data[i][idx('KODE ID')] !== "" && statusDir && statusDir !== "") {
      result.push({
        kodeId: data[i][idx('KODE ID')],
        lokasi: data[i][idx('LOKASI')],
        nilaiGrossUp: data[i][idx('PENILAIAN HARGA SEWA (GROSS UP)')] || data[i][idx('HARGA SEBELUM PAJAK (POKOK) SEWA')],
        statusDirektur: statusDir,
        catatanDirektur: data[i][idx('CATATAN DIREKTUR')] || "-"
      });
    }
  }
  return result;
}

// PERUBAHAN DESAIN (Opsi B, sesuai kesepakatan): fungsi ini SEBELUMNYA
// langsung membuat baris parsial di NEW_KEU begitu Direktur approve.
// Sekarang TIDAK LAGI -- baris NEW_KEU hanya dibuat lengkap sekali oleh
// simpanDataKeuangan() saat Keuangan submit. Direktur di sini hanya
// mengubah status di NEW_INPUT.
function verifikasiDirektur(kodeId, statusKeputusan, catatanDirektur) {
  var roles = getUserDetails(Session.getActiveUser().getEmail()).roles;
  if (!roles.includes("Direktur") && !roles.includes("Admin")) throw new Error("Akses Ditolak.");

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("Sistem sibuk."); }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetInput = ss.getSheetByName("NEW_INPUT");
    var colMap = getColIndexMapAuto(sheetInput, NEW_INPUT_HEADER_ROWS);
    var idx = function (headerName) { return getColIndex(colMap, headerName); };
    var dataInput = sheetInput.getDataRange().getValues();

    for (var i = NEW_INPUT_DATA_START_ROW; i < dataInput.length; i++) {
      if (dataInput[i][idx('KODE ID')] === kodeId) {
        sheetInput.getRange(i + 1, idx('STATUS DIREKTUR') + 1).setValue(statusKeputusan);
        sheetInput.getRange(i + 1, idx('CATATAN DIREKTUR') + 1).setValue(catatanDirektur || "-");

        if (statusKeputusan === "Disetujui") {
          // Catat jejak waktu approval Direktur -- dipakai nanti oleh
          // simpanDataKeuangan() untuk mengisi NEW_KEU secara lengkap.
          // CATATAN: kolom "TANGGAL ACC DIREKTUR" perlu ditambahkan manual
          // di sheet NEW_INPUT (paling gampang: taruh di kolom paling akhir,
          // setelah EMAIL NEGOSIATOR). Kalau kolom ini belum ada, baris di
          // bawah akan gagal dengan pesan error yang jelas -- bukan silent bug.
          sheetInput.getRange(i + 1, idx('TANGGAL ACC DIREKTUR') + 1).setValue(new Date());
        }

        return "Pengajuan berhasil " + statusKeputusan + ". Siap diproses oleh Divisi Keuangan.";
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

  var colMap = getColIndexMapAuto(sheet, NEW_INPUT_HEADER_ROWS);
  var idx = function (headerName) { return getColIndex(colMap, headerName); };
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = NEW_INPUT_DATA_START_ROW; i < data.length; i++) {
    var statusDir = data[i][idx('STATUS DIREKTUR')];
    var ntpn = data[i][idx('NOMOR NTPN RESMI')];
    if (data[i][idx('KODE ID')] !== "" && statusDir === "Disetujui" && (!ntpn || ntpn === "")) {
      var tglHabis = data[i][idx('TANGGAL HABIS KONTRAK')];
      result.push({
        kodeId: data[i][idx('KODE ID')],
        lokasi: data[i][idx('LOKASI')],
        statusSewa: data[i][idx('STATUS SEWA')],
        hargaPokok: data[i][idx('HARGA SEBELUM PAJAK (POKOK) SEWA')],
        nominalPajak: data[i][idx('PAJAK SEWA (YANG DIPOTONG)')],
        hargaBersih: data[i][idx('HARGA SETELAH PAJAK (DITRANSFER BERSIH)')],
        masaSewa: data[i][idx('MASA SEWA (TAHUN)')],
        tglHabisBaru: tglHabis instanceof Date ? tglHabis.toISOString().split('T')[0] : tglHabis,
        nilaiGrossUp: data[i][idx('PENILAIAN HARGA SEWA (GROSS UP)')],
        spkDokumen: data[i][idx('SPK (LINK DRIVE)')],
        memoDokumen: data[i][idx('MEMO (LINK DRIVE)')],
        catatanDirektur: data[i][idx('CATATAN DIREKTUR')] || "-"
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

  var colMap = getColIndexMapAuto(sheet, NEW_INPUT_HEADER_ROWS);
  var idx = function (headerName) { return getColIndex(colMap, headerName); };
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = NEW_INPUT_DATA_START_ROW; i < data.length; i++) {
    var ntpn = data[i][idx('NOMOR NTPN RESMI')];
    if (data[i][idx('KODE ID')] !== "" && ntpn && String(ntpn).trim() !== "") {
      var tglBayar = data[i][idx('TANGGAL BAYAR PAJAK')];
      result.push({
        kodeId: data[i][idx('KODE ID')],
        lokasi: data[i][idx('LOKASI')],
        nilaiGrossUp: data[i][idx('PENILAIAN HARGA SEWA (GROSS UP)')] || data[i][idx('HARGA SEBELUM PAJAK (POKOK) SEWA')],
        tglBayarPajak: tglBayar instanceof Date ? tglBayar.toISOString().split('T')[0] : (tglBayar || "-"),
        ntpn: ntpn,
        bukpotLink: data[i][idx('BUKTI TRANSFER / BUKPOT')] || ""
      });
    }
  }
  return result;
}

// ==========================================
// BACKEND DIVISI KEUANGAN (NEW_KEU UPDATE)
// ==========================================
// Konstanta struktur sheet NEW_KEU: sama seperti NEW_INPUT, ada kolom
// "Fix Nilai (Arsip) > Sewa Bersih / Pajak" yang berjenjang 3 baris,
// plus 1 baris metadata nomor kolom sebelum data mulai.
var NEW_KEU_HEADER_ROWS = 3;
var NEW_KEU_DATA_START_ROW = 4;

// PERUBAHAN DESAIN (Opsi B, sesuai kesepakatan): NEW_KEU sekarang HANYA
// diisi lengkap sekali di sini, tidak lagi ada baris parsial dari Direktur.
// Kalau untuk suatu alasan baris untuk kodeId ini SUDAH ada (misal koreksi
// ulang oleh Keuangan), baris itu ditimpa lengkap -- bukan ditambah baris baru.
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

    var colMapInput = getColIndexMapAuto(sheetInput, NEW_INPUT_HEADER_ROWS);
    var idxIn = function (headerName) { return getColIndex(colMapInput, headerName); };
    var colMapKeu = getColIndexMapAuto(sheetKeu, NEW_KEU_HEADER_ROWS);
    var idxKeu = function (headerName) { return getColIndex(colMapKeu, headerName); };

    var dataInput = sheetInput.getDataRange().getValues();
    var dataKeu = sheetKeu.getDataRange().getValues();

    for (var i = NEW_INPUT_DATA_START_ROW; i < dataInput.length; i++) {
      if (dataInput[i][idxIn('KODE ID')] === data.kodeId) {

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
        var tglCair = new Date();

        // 2. Update status & pajak di Master NEW_INPUT
        sheetInput.getRange(i + 1, idxIn('HARDCOPY SPK (FISIK)') + 1).setValue(hardcopyStatus);
        sheetInput.getRange(i + 1, idxIn('TANGGAL BAYAR PAJAK') + 1).setValue(data.tglBayarPajak);
        sheetInput.getRange(i + 1, idxIn('NOMOR NTPN RESMI') + 1).setValue(data.ntpnPajak);
        sheetInput.getRange(i + 1, idxIn('NILAI REAL PAJAK') + 1).setValue(nilaiPajakRealNum);
        if (bukpotCellContent !== "") sheetInput.getRange(i + 1, idxIn('BUKTI TRANSFER / BUKPOT') + 1).setValue(bukpotCellContent);

        // 3. Bangun baris NEW_KEU LENGKAP (Opsi B: satu-satunya penulis baris NEW_KEU)
        var lastColKeu = sheetKeu.getLastColumn();
        var barisKeu = new Array(lastColKeu).fill("");
        var setKeu = function (headerName, value) { barisKeu[idxKeu(headerName)] = value; };

        setKeu('KODE ID', dataInput[i][idxIn('KODE ID')]);
        setKeu('PERIODE / TAHUN INPUT', dataInput[i][idxIn('PERIODE / TAHUN INPUT')]);
        setKeu('LOKASI', dataInput[i][idxIn('LOKASI')]);
        setKeu('STATUS SEWA', dataInput[i][idxIn('STATUS SEWA')]);
        setKeu('INFORMASI SEWA SEBELUMNYA - TANGGAL HABIS KONTRAK', dataInput[i][idxIn('SEWA LAMA - TANGGAL HABIS KONTRAK')] || "");
        setKeu('INFORMASI SEWA SEBELUMNYA - HARGA SETELAH PAJAK (DITRANSFER BERSIH)', dataInput[i][idxIn('SEWA LAMA - HARGA SETELAH PAJAK (NET)')] || 0);
        setKeu('DEPOSIT', dataInput[i][idxIn('SEWA LAMA - DEPOSIT')] || 0);
        setKeu('HARGA SEBELUM PAJAK (POKOK) SEWA', dataInput[i][idxIn('HARGA SEBELUM PAJAK (POKOK) SEWA')] || 0);
        setKeu('PAJAK SEWA (YANG DIPOTONG)', dataInput[i][idxIn('PAJAK SEWA (YANG DIPOTONG)')] || 0);
        setKeu('INFORMASI SEWA MENDATANG - HARGA SETELAH PAJAK (DITRANSFER BERSIH)', dataInput[i][idxIn('HARGA SETELAH PAJAK (DITRANSFER BERSIH)')] || 0);
        setKeu('INFORMASI SEWA MENDATANG - TANGGAL HABIS KONTRAK', dataInput[i][idxIn('TANGGAL HABIS KONTRAK')] || "");
        setKeu('PENILAIAN HARGA SEWA (GROSS UP)', dataInput[i][idxIn('PENILAIAN HARGA SEWA (GROSS UP)')] || 0);
        setKeu('SEWA BERSIH', dataInput[i][idxIn('HARGA SETELAH PAJAK (DITRANSFER BERSIH)')] || 0);
        setKeu('PAJAK', dataInput[i][idxIn('PAJAK SEWA (YANG DIPOTONG)')] || 0);
        setKeu('NAMA', dataInput[i][idxIn('NAMA PEMILIK')] || "");
        setKeu('NOMOR', dataInput[i][idxIn('NIK')] || "");
        setKeu('NPWP', dataInput[i][idxIn('NPWP')] || "");
        setKeu('SEWA DITRANSFER', dataInput[i][idxIn('HARGA SETELAH PAJAK (DITRANSFER BERSIH)')] || 0);
        setKeu('SOFT', "ADA");
        setKeu('LINK', dataInput[i][idxIn('SPK (LINK DRIVE)')] || "");
        setKeu('HARD', hardcopyStatus);
        setKeu('TANGGAL BAYAR', data.tglBayarPajak);
        setKeu('NOMOR NTPN', data.ntpnPajak);
        setKeu('NILAI PAJAK', nilaiPajakRealNum);
        setKeu('ARSIP BUKPOT', bukpotCellContent);
        setKeu('EMAIL NEGOSIATOR', dataInput[i][idxIn('EMAIL NEGOSIATOR')] || "");
        // Kolom jejak waktu ini butuh kolom baru di NEW_INPUT (lihat catatan
        // di verifikasiDanAdjustMCX / verifikasiPengajuan / verifikasiDirektur).
        setKeu('TANGGAL ACC MCX', dataInput[i][idxIn('TANGGAL ACC MCX')] || "");
        setKeu('TANGGAL ACC SEKRETARIS', dataInput[i][idxIn('TANGGAL ACC SEKRETARIS')] || "");
        setKeu('TANGGAL ACC DIREKTUR', dataInput[i][idxIn('TANGGAL ACC DIREKTUR')] || "");
        setKeu('TANGGAL CAIR', tglCair);

        var targetRowKeu = -1;
        for (var k = NEW_KEU_DATA_START_ROW; k < dataKeu.length; k++) {
          if (dataKeu[k][idxKeu('KODE ID')] === data.kodeId) { targetRowKeu = k + 1; break; }
        }

        if (targetRowKeu !== -1) {
          sheetKeu.getRange(targetRowKeu, 1, 1, barisKeu.length).setValues([barisKeu]);
        } else {
          sheetKeu.appendRow(barisKeu);
        }

        // 4. Notifikasi Email Feedback
        kirimEmailFeedbackPencairan(data.kodeId, dataInput[i][idxIn('LOKASI')], data.ntpnPajak, dataInput[i][idxIn('HARGA SETELAH PAJAK (DITRANSFER BERSIH)')]);

        return "Pencairan dana berhasil dicatat lengkap di NEW_KEU.";
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
    var headerRow1 = ["KODE ID", "PERIODE / TAHUN INPUT", "LOKASI", "STATUS SEWA", "INFORMASI SEWA SEBELUMNYA", "", "", "INFORMASI SEWA MENDATANG", "", "", "", "", "Fix Nilai (Arsip)", "", "N I K", "", "NPWP", "Sewa Ditransfer", "Arsip SPK", "", "", "INFO PAJAK", "", "", "", "ARSIP BUKPOT", "RENTANG WAKTU", "", "", "", ""];
    var headerRow2 = ["", "", "", "", "Tanggal Habis Kontrak", "Harga Setelah Pajak (Ditransfer Bersih)", "Deposit", "Harga Sebelum Pajak (Pokok) Sewa", "Pajak Sewa (Yang dipotong)", "Harga Setelah Pajak (Ditransfer Bersih)", "Tanggal Habis Kontrak", "Penilaian Harga Sewa (Gross Up)", "Sewa Bersih", "Pajak", "Nama", "Nomor", "", "", "Soft", "Link", "Hard", "Masa Pajak", "Tanggal Bayar", "Nomor NTPN", "Nilai Pajak", "", "Email Negosiator", "Tanggal ACC MCX", "Tanggal ACC Sekretaris", "Tanggal ACC Direktur", "Tanggal Cair"];
    sheetKeu.getRange(1, 1, 1, headerRow1.length).setValues([headerRow1]).setBackground("#dcfce7").setFontWeight("bold");
    sheetKeu.getRange(2, 1, 1, headerRow2.length).setValues([headerRow2]).setBackground("#f1f5f9").setFontWeight("bold");
    // Baris 3 dibiarkan kosong sebagai pemisah visual, data mulai baris 4
    // (konsisten dengan NEW_KEU_HEADER_ROWS=3 / NEW_KEU_DATA_START_ROW=4).
  }

  var sheetRekap = ss.getSheetByName("NEW_REKAP");
  if (!sheetRekap) {
    sheetRekap = ss.insertSheet("NEW_REKAP");
    // FIX: skema lama tidak punya KODE ID sama sekali (tidak bisa trace
    // balik ke transaksi asli). Ditambahkan sebagai kolom pertama.
    var headerRekap = ["KODE ID", "PERIODE", "WILAYAH / LOKASI", "STATUS SEWA", "INFORMASI SEWA SEBELUMNYA", "", "INFORMASI SEWA MENDATANG", "", "", "", "", "Pihak Pembayar Pajak PBB", "INFO REKENING BANK (Pemilik)", "Arsip SPK", "", "INFO PAJAK", ""];
    var headerRekapSub = ["", "", "", "", "Tanggal Habis Kontrak", "Masa Sewa (tahun)", "Harga Sebelum Pajak (Pokok) Sewa", "Pajak Sewa (Yang dipotong)", "Harga Setelah Pajak (Ditransfer Bersih)", "Masa Sewa (tahun)", "Tanggal Habis Kontrak", "", "", "", "Soft", "Hard", "Nomor NTPN", "Nilai Pajak"];
    sheetRekap.getRange(1, 1, 1, headerRekap.length).setValues([headerRekap]).setBackground("#2e1065").setFontColor("#ffffff").setFontWeight("bold");
    sheetRekap.getRange(2, 1, 1, headerRekapSub.length).setValues([headerRekapSub]).setBackground("#f1f5f9").setFontWeight("bold");
    // CATATAN: baris header di sini sengaja dipindah ke baris 1-2 (bukan 4-5
    // seperti versi lama) supaya konsisten dengan rebuildNewRekap() yang akan
    // ditulis di sesi berikutnya. Sheet NEW_REKAP produksi kamu yang sudah
    // ada TIDAK terpengaruh oleh perubahan ini -- fungsi ini hanya jalan
    // kalau sheet belum ada sama sekali.
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

  var colMap = getColIndexMapAuto(sheet, NEW_INPUT_HEADER_ROWS);
  var idx = function (headerName) { return getColIndex(colMap, headerName); };
  var data = sheet.getDataRange().getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var terkirimCount = 0;

  for (var i = NEW_INPUT_DATA_START_ROW; i < data.length; i++) {
    var kodeId = data[i][idx('KODE ID')];
    var lokasi = data[i][idx('LOKASI')];
    var statusDraft = String(data[i][idx('STATUS DRAF')] || "").trim();
    var tglHabisRaw = data[i][idx('TANGGAL HABIS KONTRAK')];
    var emailNegosiator = String(data[i][idx('EMAIL NEGOSIATOR')] || "").trim();

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
      var colMapIn = getColIndexMapAuto(sheetInput, NEW_INPUT_HEADER_ROWS);
      var dataInput = sheetInput.getDataRange().getValues();
      var kodeIdColIn = getColIndex(colMapIn, 'KODE ID');
      for (var i = NEW_INPUT_DATA_START_ROW; i < dataInput.length; i++) {
        if (dataInput[i][kodeIdColIn] === kodeId) {
          sheetInput.getRange(i + 1, getColIndex(colMapIn, 'HARDCOPY SPK (FISIK)') + 1).setValue(statusHardcopy);
          break;
        }
      }
    }

    var sheetKeu = ss.getSheetByName("NEW_KEU");
    if (sheetKeu) {
      var colMapKeu = getColIndexMapAuto(sheetKeu, NEW_KEU_HEADER_ROWS);
      var dataKeu = sheetKeu.getDataRange().getValues();
      var kodeIdColKeu = getColIndex(colMapKeu, 'KODE ID');
      for (var j = NEW_KEU_DATA_START_ROW; j < dataKeu.length; j++) {
        if (dataKeu[j][kodeIdColKeu] === kodeId) {
          sheetKeu.getRange(j + 1, getColIndex(colMapKeu, 'HARD') + 1).setValue(statusHardcopy);
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

    var colMap = getColIndexMapAuto(sheetInput, NEW_INPUT_HEADER_ROWS);
    var dataInput = sheetInput.getDataRange().getValues();
    var emailNegosiator = "";
    var kodeIdCol = getColIndex(colMap, 'KODE ID');
    var emailCol = getColIndex(colMap, 'EMAIL NEGOSIATOR');

    for (var i = NEW_INPUT_DATA_START_ROW; i < dataInput.length; i++) {
      if (dataInput[i][kodeIdCol] === kodeId) {
        emailNegosiator = String(dataInput[i][emailCol] || "").trim();
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