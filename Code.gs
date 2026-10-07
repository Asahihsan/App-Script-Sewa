// ==========================================
// HELPER PERHITUNGAN PAJAK -- SATU SUMBER KEBENARAN
// ==========================================
// Sebelumnya rumus ini di-copy-paste di 3 tempat (simpanPengajuanBaru,
// verifikasiDanAdjustMCX, buatMemoPDF) dengan logic identik. Kalau rumus
// pajak berubah (misal tarif PPh berubah dari 10%), sekarang cukup ubah
// di SATU tempat ini saja.
//
// @param {number} harga - nilai harga pokok sewa
// @param {boolean} dipotongDiSumber - true kalau pajak dipotong dari harga
//   (harga yang diinput adalah harga kotor sebelum potong pajak), false
//   kalau skema gross-up (harga yang diinput sudah bersih, pajak ditanggung
//   terpisah di atas harga tsb).
// @return {{ nominalPajak: number, hargaBersih: number, nilaiGrossUp: number }}
function hitungPajak(harga, dipotongDiSumber) {
  var nominalPajak, hargaBersih, nilaiGrossUp;
  if (dipotongDiSumber) {
    nominalPajak = Math.round(harga * 0.10);
    hargaBersih = Math.round(harga - nominalPajak);
    nilaiGrossUp = Math.round(harga);
  } else {
    nominalPajak = Math.round((harga / 0.9) * 0.10);
    hargaBersih = Math.round(harga);
    nilaiGrossUp = Math.round(harga / 0.9);
  }
  return { nominalPajak: nominalPajak, hargaBersih: hargaBersih, nilaiGrossUp: nilaiGrossUp };
}

// Helper kecil: field "dipotongPajak" di berbagai tempat kadang berupa
// string ("Ya, dipotong" / "Tidak (Gross Up)") -- ini menyeragamkan jadi
// boolean murni untuk dipakai hitungPajak().
function isDipotongDiSumber(dipotongPajakField) {
  return !!(dipotongPajakField && String(dipotongPajakField).includes("Ya"));
}

// ==========================================
// HELPER SKEMA CICILAN -- SATU SUMBER KEBENARAN
// ==========================================
// Desain cicilan (Opsi A, disepakati): negosiator mengisi SELURUH jadwal
// cicilan (maks 3x) di awal. Pass pertama (approval penuh MCX->Sekretaris->
// Direktur->Keuangan) otomatis melunasi Cicilan 1. Cicilan 2 & 3 masing-
// masing cukup lewat jalur ringkas Direktur (transfer) -> Keuangan (pajak +
// catat), tanpa perlu verifikasi ulang MCX/Sekretaris karena deal-nya sudah
// diverifikasi sekali di pass pertama.
//
// Kolom BARU yang perlu ditambahkan manual oleh user di sheet NEW_INPUT
// (pola sama seperti penambahan NIK/NPWP sebelumnya -- getColIndexMapAuto
// akan menemukan posisinya otomatis, tidak peduli taruh di kolom mana):
//   - "CICILAN AKTIF"              -> angka 1/2/3, kosong = tidak ada siklus cicilan berjalan
//   - "STATUS TRANSFER CICILAN"    -> kosong / "DITRANSFER"
//   - "TANGGAL TRANSFER CICILAN"   -> diisi otomatis (trySetCell, opsional)
//   - "BUKTI TRANSFER SEWA (DIREKTUR)" -> link Drive hasil upload Direktur
//   - "CICILAN TERAKHIR LUNAS"     -> angka 0/1/2/3, 0 = belum ada yang lunas
// Kolom BARU di sheet NEW_KEU:
//   - "CICILAN KE" -> "LUNAS" (sewa non-cicilan) atau "1"/"2"/"3"
//
// @return {{ jml: number[], tgl: any[], lastSlot: number }} lastSlot = nomor
//   cicilan terakhir yang diisi negosiator (0 kalau bukan sewa cicilan sama
//   sekali / dibayar lunas sekaligus).
function getCicilanSchedule(rowArr, colMap) {
  var jml = [
    parseFloat(tryGetCell(rowArr, colMap, 'LUNAS / CICILAN 1 - JUMLAH', 0)) || 0,
    parseFloat(tryGetCell(rowArr, colMap, 'LUNAS / CICILAN 2 - JUMLAH', 0)) || 0,
    parseFloat(tryGetCell(rowArr, colMap, 'LUNAS / CICILAN 3 - JUMLAH', 0)) || 0
  ];
  var tgl = [
    tryGetCell(rowArr, colMap, 'LUNAS / CICILAN 1 - TANGGAL', ""),
    tryGetCell(rowArr, colMap, 'LUNAS / CICILAN 2 - TANGGAL', ""),
    tryGetCell(rowArr, colMap, 'LUNAS / CICILAN 3 - TANGGAL', "")
  ];
  var lastSlot = 0;
  for (var n = 0; n < 3; n++) { if (jml[n] > 0) lastSlot = n + 1; }
  return { jml: jml, tgl: tgl, lastSlot: lastSlot };
}

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
    var pajakCalc = hitungPajak(hargaBaru, isDipotongDiSumber(data.dipotongPajak));
    var nominalPajak = pajakCalc.nominalPajak;
    var hargaSetelahPajak = pajakCalc.hargaBersih;
    var nilaiGrossUp = pajakCalc.nilaiGrossUp;
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
      var jadwalRiw = getCicilanSchedule(data[i], colMap);
      var isCicilanRiw = jadwalRiw.lastSlot >= 2;
      var cicilanAktifRiw = parseInt(tryGetCell(data[i], colMap, 'CICILAN AKTIF', 0)) || 0;
      var cicilanLunasRiw = parseInt(tryGetCell(data[i], colMap, 'CICILAN TERAKHIR LUNAS', 0)) || 0;

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
        ntpn: data[i][idx('NOMOR NTPN RESMI')],
        // Info progres cicilan -- dipakai dashboard Negosiator utk nampilin
        // tombol "Ajukan Pencairan Cicilan ke-N" begitu tahap sebelumnya lunas.
        isCicilan: isCicilanRiw,
        cicilanTotalTahap: jadwalRiw.lastSlot,
        cicilanAktif: cicilanAktifRiw,
        cicilanTerakhirLunas: cicilanLunasRiw,
        bisaAjukanCicilanBerikutnya: isCicilanRiw && cicilanAktifRiw === 0 && cicilanLunasRiw < jadwalRiw.lastSlot && data[i][idx('STATUS DIREKTUR')] === "Disetujui"
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
          var dipotongDiSumber = (Math.round(grossUpLama) === Math.round(hargaPokokLama));

          var pajakCalc = hitungPajak(hrgFix, dipotongDiSumber);
          var nominalPajak = pajakCalc.nominalPajak;
          var hargaSetelahPajak = pajakCalc.hargaBersih;
          var nilaiGrossUp = pajakCalc.nilaiGrossUp;

          setCell('HARGA SEBELUM PAJAK (POKOK) SEWA', hrgFix);
          setCell('PAJAK SEWA (YANG DIPOTONG)', nominalPajak);
          setCell('HARGA SETELAH PAJAK (DITRANSFER BERSIH)', hargaSetelahPajak);
          setCell('PENILAIAN HARGA SEWA (GROSS UP)', nilaiGrossUp);

          setCell('STATUS MCX', "APPROVED");
          // Jejak waktu opsional untuk NEW_KEU nanti -- pakai trySetCell
          // supaya kalau kolom "TANGGAL ACC MCX" belum ada di sheet, proses
          // approval INTI (status, harga) tetap jalan, tidak ikut gagal.
          trySetCell(sheet, rowNum, colMap, 'TANGGAL ACC MCX', new Date());
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
          // Jejak waktu opsional -- gagal aman kalau kolom belum ada.
          trySetCell(sheet, i + 1, colMap, 'TANGGAL ACC SEKRETARIS', new Date());
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

    // Antrean "transfer cicilan lanjutan" (Cicilan 2/3): deal-nya sudah
    // "Disetujui" sejak pass pertama, jadi tidak lewat gate STATUS SEKRETARIS
    // lagi -- cukup dipicu negosiator via ajukanPencairanCicilan().
    var cicilanAktif = parseInt(tryGetCell(data[i], colMap, 'CICILAN AKTIF', 0)) || 0;
    var statusTransferCicilan = tryGetCell(data[i], colMap, 'STATUS TRANSFER CICILAN', '');
    if (data[i][idx('KODE ID')] !== "" && cicilanAktif >= 2 && statusTransferCicilan !== 'DITRANSFER') {
      var jadwalC = getCicilanSchedule(data[i], colMap);
      result.push({
        kodeId: data[i][idx('KODE ID')],
        lokasi: data[i][idx('LOKASI')],
        statusSewa: data[i][idx('STATUS SEWA')],
        namaPemilik: data[i][idx('NAMA PEMILIK')],
        tipePengajuan: 'CICILAN_LANJUTAN',
        cicilanKe: cicilanAktif,
        jumlahCicilan: jadwalC.jml[cicilanAktif - 1],
        tglJatuhTempoCicilan: jadwalC.tgl[cicilanAktif - 1]
      });
      continue;
    }

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
        lokasiLama: data[i][idx('LOKASI LAMA')] || "-",
        tipePengajuan: 'DEAL_BARU'
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
// PERUBAHAN: Direktur sekarang sekaligus mentransfer dana sewa, jadi di sini
// juga menerima upload bukti kirim (bukti transfer) -- opsional di parameter
// supaya tidak memaksa dashboard lama (yang belum update UI) jadi error saat
// menolak pengajuan (di mana bukti transfer tidak relevan).
// Kalau deal ini sewa cicilan (lastSlot >= 2, lihat getCicilanSchedule),
// approval ini otomatis dianggap juga sebagai transfer utk Cicilan 1 --
// CICILAN AKTIF & STATUS TRANSFER CICILAN diset supaya Keuangan tahu harus
// memproses pajak atas nilai Cicilan 1 saja (bukan nilai total deal).
function verifikasiDirektur(kodeId, statusKeputusan, catatanDirektur, buktiTransferData, buktiTransferMimeType, buktiTransferName) {
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
          // trySetCell = gagal aman, tidak menghentikan proses approval inti
          // kalau kolom "TANGGAL ACC DIREKTUR" belum ada di sheet.
          trySetCell(sheetInput, i + 1, colMap, 'TANGGAL ACC DIREKTUR', new Date());

          // Upload bukti transfer (gagal aman: kalau kolom belum ditambah
          // user atau Direktur tidak melampirkan apa pun, cukup dilewati).
          if (buktiTransferData && buktiTransferName) {
            try {
              var folderBukti = DriveApp.getFolderById('1tEW_fyH69zpOWqq3oAI0u7QkT_yZsqS_');
              var blobBukti = Utilities.newBlob(Utilities.base64Decode(buktiTransferData), buktiTransferMimeType, "BUKTI_TRANSFER_" + kodeId + "_" + buktiTransferName);
              var fileBukti = folderBukti.createFile(blobBukti);
              trySetCell(sheetInput, i + 1, colMap, 'BUKTI TRANSFER SEWA (DIREKTUR)', fileBukti.getUrl() + "|ID:" + fileBukti.getId());
            } catch (eUp) { Logger.log("Gagal upload bukti transfer Direktur (" + kodeId + "): " + eUp.message); }
          }
          trySetCell(sheetInput, i + 1, colMap, 'TANGGAL TRANSFER CICILAN', new Date());

          var jadwal = getCicilanSchedule(dataInput[i], colMap);
          if (jadwal.lastSlot >= 2) {
            // Sewa cicilan sungguhan -- approval+transfer ini melunasi Cicilan 1.
            trySetCell(sheetInput, i + 1, colMap, 'CICILAN AKTIF', 1);
            trySetCell(sheetInput, i + 1, colMap, 'STATUS TRANSFER CICILAN', 'DITRANSFER');
          }
        }

        return "Pengajuan berhasil " + statusKeputusan + ". Siap diproses oleh Divisi Keuangan.";
      }
    }
    throw new Error("ID Pengajuan tidak ditemukan!");
  } finally { lock.releaseLock(); }
}

// Dipanggil Negosiator untuk membuka siklus pencairan Cicilan berikutnya
// (ke-2 atau ke-3) setelah cicilan sebelumnya lunas. Tidak perlu lewat
// MCX/Sekretaris lagi -- deal-nya sudah diverifikasi sekali di pass pertama.
// Begitu dipanggil, pengajuan ini langsung muncul di antrean Direktur
// (lihat getPengajuanDirektur) sebagai permintaan transfer cicilan.
function ajukanPencairanCicilan(kodeId) {
  var roles = getUserDetails(Session.getActiveUser().getEmail()).roles;
  if (!roles.includes("Negosiator") && !roles.includes("Admin")) throw new Error("Akses Ditolak.");

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("Sistem sibuk."); }

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
    var colMap = getColIndexMapAuto(sheet, NEW_INPUT_HEADER_ROWS);
    var idx = function (headerName) { return getColIndex(colMap, headerName); };
    var data = sheet.getDataRange().getValues();

    for (var i = NEW_INPUT_DATA_START_ROW; i < data.length; i++) {
      if (data[i][idx('KODE ID')] === kodeId) {
        var jadwal = getCicilanSchedule(data[i], colMap);
        if (jadwal.lastSlot < 2) throw new Error("Pengajuan ini bukan sewa cicilan (atau hanya 1x bayar).");

        var cicilanAktifSkrg = parseInt(tryGetCell(data[i], colMap, 'CICILAN AKTIF', 0)) || 0;
        if (cicilanAktifSkrg > 0) throw new Error("Masih ada siklus pencairan cicilan yang berjalan untuk ID ini.");

        var lunasTerakhir = parseInt(tryGetCell(data[i], colMap, 'CICILAN TERAKHIR LUNAS', 0)) || 0;
        var berikutnya = lunasTerakhir + 1;
        if (berikutnya > jadwal.lastSlot) throw new Error("Seluruh jadwal cicilan untuk ID ini sudah lunas.");

        trySetCell(sheet, i + 1, colMap, 'CICILAN AKTIF', berikutnya);
        trySetCell(sheet, i + 1, colMap, 'STATUS TRANSFER CICILAN', '');
        return "Pengajuan pencairan Cicilan ke-" + berikutnya + " berhasil diajukan ke Direktur.";
      }
    }
    throw new Error("ID tidak ditemukan!");
  } finally { lock.releaseLock(); }
}

// Dipanggil Direktur khusus untuk siklus transfer Cicilan ke-2/3 (bukan
// approval deal baru -- deal-nya sudah "Disetujui" sejak pass pertama).
// Hanya menandai transfer dilakukan + simpan bukti; Keuangan yang akan
// memproses pajak & mencatatnya sebagai baris NEW_KEU baru.
function verifikasiTransferCicilan(kodeId, buktiTransferData, buktiTransferMimeType, buktiTransferName) {
  var roles = getUserDetails(Session.getActiveUser().getEmail()).roles;
  if (!roles.includes("Direktur") && !roles.includes("Admin")) throw new Error("Akses Ditolak.");

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("Sistem sibuk."); }

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NEW_INPUT");
    var colMap = getColIndexMapAuto(sheet, NEW_INPUT_HEADER_ROWS);
    var idx = function (headerName) { return getColIndex(colMap, headerName); };
    var data = sheet.getDataRange().getValues();

    for (var i = NEW_INPUT_DATA_START_ROW; i < data.length; i++) {
      if (data[i][idx('KODE ID')] === kodeId) {
        var cicilanAktif = parseInt(tryGetCell(data[i], colMap, 'CICILAN AKTIF', 0)) || 0;
        if (cicilanAktif < 1) throw new Error("Tidak ada siklus pencairan cicilan yang berjalan untuk ID ini.");

        if (buktiTransferData && buktiTransferName) {
          try {
            var folderBukti = DriveApp.getFolderById('1tEW_fyH69zpOWqq3oAI0u7QkT_yZsqS_');
            var blobBukti = Utilities.newBlob(Utilities.base64Decode(buktiTransferData), buktiTransferMimeType, "BUKTI_TRANSFER_" + kodeId + "_CICILAN" + cicilanAktif + "_" + buktiTransferName);
            var fileBukti = folderBukti.createFile(blobBukti);
            trySetCell(sheet, i + 1, colMap, 'BUKTI TRANSFER SEWA (DIREKTUR)', fileBukti.getUrl() + "|ID:" + fileBukti.getId());
          } catch (eUp) { Logger.log("Gagal upload bukti transfer cicilan (" + kodeId + "): " + eUp.message); }
        }
        trySetCell(sheet, i + 1, colMap, 'STATUS TRANSFER CICILAN', 'DITRANSFER');
        trySetCell(sheet, i + 1, colMap, 'TANGGAL TRANSFER CICILAN', new Date());

        return "Transfer Cicilan ke-" + cicilanAktif + " berhasil dicatat. Siap diproses oleh Divisi Keuangan.";
      }
    }
    throw new Error("ID tidak ditemukan!");
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

    // Antrean "pajak cicilan lanjutan" (Cicilan 2/3): muncul begitu Direktur
    // sudah mencatat transfer-nya (verifikasiTransferCicilan), terpisah dari
    // gate NTPN di bawah karena NTPN RESMI di NEW_INPUT cuma menampung 1
    // nilai sedangkan cicilan butuh baris NEW_KEU sendiri per tahap.
    var cicilanAktifQ = parseInt(tryGetCell(data[i], colMap, 'CICILAN AKTIF', 0)) || 0;
    var statusTransferQ = tryGetCell(data[i], colMap, 'STATUS TRANSFER CICILAN', '');
    if (data[i][idx('KODE ID')] !== "" && cicilanAktifQ >= 2 && statusTransferQ === 'DITRANSFER') {
      var jadwalQ = getCicilanSchedule(data[i], colMap);
      result.push({
        kodeId: data[i][idx('KODE ID')],
        lokasi: data[i][idx('LOKASI')],
        statusSewa: data[i][idx('STATUS SEWA')],
        tipePengajuan: 'CICILAN_LANJUTAN',
        cicilanKe: cicilanAktifQ,
        hargaPokok: jadwalQ.jml[cicilanAktifQ - 1],
        spkDokumen: data[i][idx('SPK (LINK DRIVE)')],
        memoDokumen: data[i][idx('MEMO (LINK DRIVE)')],
        buktiTransferDirektur: tryGetCell(data[i], colMap, 'BUKTI TRANSFER SEWA (DIREKTUR)', '')
      });
      continue;
    }

    if (data[i][idx('KODE ID')] !== "" && statusDir === "Disetujui" && (!ntpn || ntpn === "")) {
      var tglHabis = data[i][idx('TANGGAL HABIS KONTRAK')];
      var jadwalN = getCicilanSchedule(data[i], colMap);
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
        catatanDirektur: data[i][idx('CATATAN DIREKTUR')] || "-",
        // Kalau ini sewa cicilan, ini otomatis mewakili pajak Cicilan 1 saja
        // (bukan nilai total deal) -- lihat simpanDataKeuangan().
        tipePengajuan: jadwalN.lastSlot >= 2 ? 'CICILAN_1' : 'LUNAS',
        cicilanKe: jadwalN.lastSlot >= 2 ? 1 : 0,
        jumlahCicilanTahap: jadwalN.lastSlot >= 2 ? jadwalN.jml[0] : null
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

        // Cicilan: tentukan apakah deal ini sewa cicilan sungguhan (lastSlot
        // >= 2) dan tahap keberapa yang sedang diproses. CICILAN AKTIF di
        // NEW_INPUT adalah sumber kebenaran -- diset otomatis ke 1 oleh
        // verifikasiDirektur() saat approval pertama, lalu di-set ulang ke
        // 2/3 oleh ajukanPencairanCicilan()+verifikasiTransferCicilan().
        // Server TIDAK mempercayai cicilanKe dari client demi keamanan data.
        var jadwalKeu = getCicilanSchedule(dataInput[i], colMapInput);
        var isCicilanDeal = jadwalKeu.lastSlot >= 2;
        var cicilanKeFinal = isCicilanDeal ? (parseInt(tryGetCell(dataInput[i], colMapInput, 'CICILAN AKTIF', 0)) || 1) : 0;

        // Basis pajak tahap ini: untuk sewa cicilan, HANYA nominal tahap
        // tsb (bukan nilai total deal). Untuk sewa biasa, tetap nilai penuh
        // seperti sebelumnya (tidak ada perubahan perilaku).
        var hargaPokokFull = parseFloat(dataInput[i][idxIn('HARGA SEBELUM PAJAK (POKOK) SEWA')]) || 0;
        var grossUpFull = parseFloat(dataInput[i][idxIn('PENILAIAN HARGA SEWA (GROSS UP)')]) || 0;
        var dipotongDiSumberFlag = Math.round(grossUpFull) === Math.round(hargaPokokFull);
        var baseHargaTahap = isCicilanDeal ? jadwalKeu.jml[cicilanKeFinal - 1] : hargaPokokFull;
        var pajakTahap = hitungPajak(baseHargaTahap, dipotongDiSumberFlag);

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
        // Variant gagal-aman khusus kolom cicilan yang mungkin belum ada di
        // NEW_KEU (user belum sempat tambahkan manual) -- dilewati + dicatat
        // ke log, tidak menghentikan pencatatan pencairan dana inti.
        var setKeuSafe = function (headerName, value) {
          try { barisKeu[idxKeu(headerName)] = value; }
          catch (eCol) { Logger.log('Kolom opsional "' + headerName + '" belum ada di NEW_KEU, dilewati: ' + eCol.message); }
        };

        setKeu('KODE ID', dataInput[i][idxIn('KODE ID')]);
        setKeu('PERIODE / TAHUN INPUT', dataInput[i][idxIn('PERIODE / TAHUN INPUT')]);
        setKeu('LOKASI', dataInput[i][idxIn('LOKASI')]);
        setKeu('STATUS SEWA', dataInput[i][idxIn('STATUS SEWA')]);
        setKeu('INFORMASI SEWA SEBELUMNYA - TANGGAL HABIS KONTRAK', dataInput[i][idxIn('SEWA LAMA - TANGGAL HABIS KONTRAK')] || "");
        setKeu('INFORMASI SEWA SEBELUMNYA - HARGA SETELAH PAJAK (DITRANSFER BERSIH)', dataInput[i][idxIn('SEWA LAMA - HARGA SETELAH PAJAK (NET)')] || 0);
        setKeu('DEPOSIT', dataInput[i][idxIn('SEWA LAMA - DEPOSIT')] || 0);
        setKeu('HARGA SEBELUM PAJAK (POKOK) SEWA', baseHargaTahap);
        setKeu('PAJAK SEWA (YANG DIPOTONG)', pajakTahap.nominalPajak);
        setKeu('INFORMASI SEWA MENDATANG - HARGA SETELAH PAJAK (DITRANSFER BERSIH)', pajakTahap.hargaBersih);
        setKeu('INFORMASI SEWA MENDATANG - TANGGAL HABIS KONTRAK', dataInput[i][idxIn('TANGGAL HABIS KONTRAK')] || "");
        setKeu('PENILAIAN HARGA SEWA (GROSS UP)', pajakTahap.nilaiGrossUp);
        setKeu('SEWA BERSIH', pajakTahap.hargaBersih);
        setKeu('PAJAK', pajakTahap.nominalPajak);
        setKeu('NAMA', dataInput[i][idxIn('NAMA PEMILIK')] || "");
        setKeu('NOMOR', dataInput[i][idxIn('NIK')] || "");
        setKeu('NPWP', dataInput[i][idxIn('NPWP')] || "");
        setKeu('SEWA DITRANSFER', pajakTahap.hargaBersih);
        setKeu('SOFT', "ADA");
        setKeu('LINK', dataInput[i][idxIn('SPK (LINK DRIVE)')] || "");
        setKeu('HARD', hardcopyStatus);
        setKeu('TANGGAL BAYAR', data.tglBayarPajak);
        setKeu('NOMOR NTPN', data.ntpnPajak);
        setKeu('NILAI PAJAK', nilaiPajakRealNum);
        setKeu('ARSIP BUKPOT', bukpotCellContent);
        setKeu('EMAIL NEGOSIATOR', dataInput[i][idxIn('EMAIL NEGOSIATOR')] || "");
        // Gagal aman: kalau kolom jejak waktu belum ada di NEW_INPUT, isi
        // kosong saja -- tidak menghentikan pencatatan pencairan dana.
        setKeu('TANGGAL ACC MCX', tryGetCell(dataInput[i], colMapInput, 'TANGGAL ACC MCX', ""));
        setKeu('TANGGAL ACC SEKRETARIS', tryGetCell(dataInput[i], colMapInput, 'TANGGAL ACC SEKRETARIS', ""));
        setKeu('TANGGAL ACC DIREKTUR', tryGetCell(dataInput[i], colMapInput, 'TANGGAL ACC DIREKTUR', ""));
        setKeu('TANGGAL CAIR', tglCair);
        // Cicilan: "LUNAS" utk sewa biasa/1x bayar, "1"/"2"/"3" utk tiap
        // tahap sewa cicilan -- ini yg bikin tiap tahap jadi baris sendiri
        // di NEW_KEU, bukan saling menimpa.
        setKeuSafe('CICILAN KE', isCicilanDeal ? String(cicilanKeFinal) : 'LUNAS');
        // Bukti transfer Direktur utk tahap ini (gagal aman: kolom opsional).
        setKeuSafe('BUKTI TRANSFER SEWA (DIREKTUR)', tryGetCell(dataInput[i], colMapInput, 'BUKTI TRANSFER SEWA (DIREKTUR)', ''));

        // Baris target: untuk sewa cicilan, kunci pencarian adalah KODE ID +
        // CICILAN KE (supaya tiap tahap jadi baris sendiri, bukan saling
        // menimpa). Untuk sewa biasa, tetap KODE ID saja seperti semula.
        var cicilanKeyExpected = isCicilanDeal ? String(cicilanKeFinal) : 'LUNAS';
        var targetRowKeu = -1;
        for (var k = NEW_KEU_DATA_START_ROW; k < dataKeu.length; k++) {
          if (dataKeu[k][idxKeu('KODE ID')] !== data.kodeId) continue;
          var cicilanKeExisting = String(tryGetCell(dataKeu[k], colMapKeu, 'CICILAN KE', 'LUNAS') || 'LUNAS');
          if (cicilanKeExisting === cicilanKeyExpected) { targetRowKeu = k + 1; break; }
        }

        if (targetRowKeu !== -1) {
          sheetKeu.getRange(targetRowKeu, 1, 1, barisKeu.length).setValues([barisKeu]);
        } else {
          sheetKeu.appendRow(barisKeu);
        }

        // 3b. Tutup siklus cicilan tahap ini di NEW_INPUT (gagal aman):
        // tandai tahap ini lunas, bersihkan slot "aktif" supaya negosiator
        // bisa ajukan tahap berikutnya (ajukanPencairanCicilan), dan
        // reset status transfer supaya tidak dikira sudah ditransfer lagi
        // utk tahap selanjutnya.
        if (isCicilanDeal) {
          trySetCell(sheetInput, i + 1, colMapInput, 'CICILAN TERAKHIR LUNAS', cicilanKeFinal);
          trySetCell(sheetInput, i + 1, colMapInput, 'CICILAN AKTIF', '');
          trySetCell(sheetInput, i + 1, colMapInput, 'STATUS TRANSFER CICILAN', '');
        }

        // 4. Notifikasi Email Feedback
        kirimEmailFeedbackPencairan(data.kodeId, dataInput[i][idxIn('LOKASI')], data.ntpnPajak, pajakTahap.hargaBersih);

        return isCicilanDeal
          ? "Pencairan Cicilan ke-" + cicilanKeFinal + " berhasil dicatat lengkap di NEW_KEU."
          : "Pencairan dana berhasil dicatat lengkap di NEW_KEU.";
      }
    }
    throw new Error("ID Pengajuan tidak ditemukan!");
  } finally { lock.releaseLock(); }
}

// ==========================================
// ENGINE SINKRONISASI MANIFEST NEW_REKAP
// ==========================================
// Konstanta struktur sheet NEW_REKAP: header berjenjang di baris 3-5
// (banner filter periode di baris 1, baris 2 sengaja dikosongkan), lalu
// 1 baris metadata nomor kolom (pola sama seperti NEW_INPUT & NEW_KEU)
// sebelum data mulai.
var NEW_REKAP_HEADER_ROWS = 3;
var NEW_REKAP_HEADER_START_ROW = 3;
var NEW_REKAP_DATA_START_ROW = 6; // 0-based -> baris sheet ke-7
// CATATAN: kalau ternyata baris data pertama BUKAN baris 7 di sheet kamu,
// nilai ini perlu disesuaikan -- cek langsung baris berapa row pertama yang
// kosong siap diisi data transaksi di NEW_REKAP.

/**
 * PERUBAHAN DESAIN BESAR: NEW_REKAP sekarang adalah "materialized view" --
 * bukan lagi ditulis inkremental saat Keuangan submit (fungsi syncKeRekap
 * yang lama sudah dihapus total). Fungsi ini MENGOSONGKAN seluruh data lama
 * di NEW_REKAP lalu menulis ulang dari nol, hasil gabungan NEW_KEU (sumber
 * data transaksi yang sudah lunas) + NEW_INPUT (untuk field yang tidak
 * disimpan di NEW_KEU, seperti Masa Sewa, Pihak Pembayar Pajak PBB, dan
 * Info Rekening Bank). Ini menghilangkan race condition/tabrakan data yang
 * dulu terjadi karena banyak role menulis ke sheet yang sama.
 *
 * Dipanggil otomatis oleh getExecutiveDashboardData() setiap kali dashboard
 * dibuka/di-refresh, jadi datanya selalu up-to-date tanpa perlu trigger
 * terpisah. Kalau nanti datanya sudah sangat banyak dan proses ini mulai
 * terasa lambat, PERTIMBANGKAN pindah ke trigger time-based (jalan sekali
 * semalam) -- ini sudah pernah dibahas di roadmap performa.
 */
function rebuildNewRekap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetKeu = ss.getSheetByName("NEW_KEU");
  var sheetInput = ss.getSheetByName("NEW_INPUT");
  var sheetRekap = ss.getSheetByName("NEW_REKAP");
  if (!sheetKeu || !sheetInput || !sheetRekap) return;

  var colMapKeu = getColIndexMapAuto(sheetKeu, NEW_KEU_HEADER_ROWS);
  var colMapInput = getColIndexMapAuto(sheetInput, NEW_INPUT_HEADER_ROWS);
  var colMapRekap = getColIndexMapAuto(sheetRekap, NEW_REKAP_HEADER_ROWS, NEW_REKAP_HEADER_START_ROW);

  var cell = function (rowArr, colMap, headerName) { return rowArr[getColIndex(colMap, headerName)]; };

  var dataKeu = sheetKeu.getDataRange().getValues();
  var dataInput = sheetInput.getDataRange().getValues();

  // Index NEW_INPUT by kodeId sekali di awal, biar tidak scan ulang per baris.
  var inputByKode = {};
  var kodeIdColIn = getColIndex(colMapInput, 'KODE ID');
  for (var i = NEW_INPUT_DATA_START_ROW; i < dataInput.length; i++) {
    var kid = dataInput[i][kodeIdColIn];
    if (kid) inputByKode[kid] = dataInput[i];
  }

  var kodeIdColKeu = getColIndex(colMapKeu, 'KODE ID');
  var lastColRekap = sheetRekap.getLastColumn();
  var outputRows = [];

  for (var k = NEW_KEU_DATA_START_ROW; k < dataKeu.length; k++) {
    var keuRow = dataKeu[k];
    var kid2 = keuRow[kodeIdColKeu];
    if (!kid2) continue;
    var inRow = inputByKode[kid2] || null;

    var rekapRow = new Array(lastColRekap).fill("");
    var setR = function (headerName, value) { rekapRow[getColIndex(colMapRekap, headerName)] = value; };

    setR('KODE ID', kid2);
    setR('WILAYAH / LOKASI', cell(keuRow, colMapKeu, 'LOKASI'));
    setR('STATUS SEWA', cell(keuRow, colMapKeu, 'STATUS SEWA'));
    setR('SEWA LAMA - TANGGAL HABIS KONTRAK', cell(keuRow, colMapKeu, 'INFORMASI SEWA SEBELUMNYA - TANGGAL HABIS KONTRAK'));
    setR('SEWA LAMA - MASA SEWA (TAHUN)', inRow ? cell(inRow, colMapInput, 'SEWA LAMA - MASA SEWA (TAHUN)') : "");
    setR('SEWA LAMA - HARGA SETELAH PAJAK (NET)', cell(keuRow, colMapKeu, 'INFORMASI SEWA SEBELUMNYA - HARGA SETELAH PAJAK (DITRANSFER BERSIH)'));
    setR('SEWA BARU - HARGA SEBELUM PAJAK (NET)', cell(keuRow, colMapKeu, 'HARGA SEBELUM PAJAK (POKOK) SEWA'));
    setR('PAJAK SEWA (YANG DIPOTONG)', cell(keuRow, colMapKeu, 'PAJAK SEWA (YANG DIPOTONG)'));
    setR('SEWA BARU - HARGA SETELAH PAJAK (NET)', cell(keuRow, colMapKeu, 'INFORMASI SEWA MENDATANG - HARGA SETELAH PAJAK (DITRANSFER BERSIH)'));
    setR('SEWA BARU - MASA SEWA (TAHUN)', inRow ? cell(inRow, colMapInput, 'MASA SEWA (TAHUN)') : "");
    setR('SEWA BARU - TANGGAL HABIS KONTRAK', cell(keuRow, colMapKeu, 'INFORMASI SEWA MENDATANG - TANGGAL HABIS KONTRAK'));
    setR('PIHAK PEMBAYAR PAJAK PBB', inRow ? cell(inRow, colMapInput, 'PIHAK PEMBAYAR PAJAK PBB') : "");
    setR('INFO REKENING BANK (PEMILIK)', inRow ? cell(inRow, colMapInput, 'INFO REKENING BANK') : "");
    setR('SOFT', cell(keuRow, colMapKeu, 'SOFT'));
    setR('HARD', cell(keuRow, colMapKeu, 'HARD'));
    setR('NOMOR NTPN', cell(keuRow, colMapKeu, 'NOMOR NTPN'));
    setR('NILAI PAJAK', cell(keuRow, colMapKeu, 'NILAI PAJAK'));

    outputRows.push(rekapRow);
  }

  // Kosongkan data lama (dari NEW_REKAP_DATA_START_ROW sampai baris terakhir
  // yang pernah dipakai), lalu tulis ulang dari nol -- ini "materialized view".
  var lastRowRekap = sheetRekap.getLastRow();
  if (lastRowRekap >= NEW_REKAP_DATA_START_ROW + 1) {
    sheetRekap.getRange(NEW_REKAP_DATA_START_ROW + 1, 1, lastRowRekap - NEW_REKAP_DATA_START_ROW, lastColRekap).clearContent();
  }
  if (outputRows.length > 0) {
    sheetRekap.getRange(NEW_REKAP_DATA_START_ROW + 1, 1, outputRows.length, lastColRekap).setValues(outputRows);
  }
}

// ==========================================
// DASHBOARD EXECUTIVE (BACA DARI NEW_REKAP)
// ==========================================
function getExecutiveDashboardData() {
  var roles = getUserDetails(Session.getActiveUser().getEmail()).roles;
  if (!roles.includes("Informasi") && !roles.includes("Direktur") && !roles.includes("Admin")) throw new Error("Akses Ditolak.");

  // NEW_REKAP adalah materialized view -- rebuild dulu supaya datanya
  // selalu segar setiap kali dashboard ini dibuka/di-refresh.
  rebuildNewRekap();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("NEW_REKAP");
  if (!sheet) return { stats: {}, listData: [] };

  var colMap = getColIndexMapAuto(sheet, NEW_REKAP_HEADER_ROWS, NEW_REKAP_HEADER_START_ROW);
  var idx = function (headerName) { return getColIndex(colMap, headerName); };
  var data = sheet.getDataRange().getValues();
  var listData = [];
  var totalBiayaSewa = 0;
  var totalPajak = 0;
  var countPerpanjang = 0;
  var countBaru = 0;
  var countRelokasi = 0;

  for (var i = NEW_REKAP_DATA_START_ROW; i < data.length; i++) {
    var lokasi = data[i][idx('WILAYAH / LOKASI')];
    if (lokasi && String(lokasi).trim() !== "") {
      var hrgBersih = parseFloat(data[i][idx('SEWA BARU - HARGA SETELAH PAJAK (NET)')]) || 0;
      var pjk = parseFloat(data[i][idx('PAJAK SEWA (YANG DIPOTONG)')]) || 0;
      var statusSewa = String(data[i][idx('STATUS SEWA')] || "").trim();
      var tglHabis = data[i][idx('SEWA BARU - TANGGAL HABIS KONTRAK')];

      totalBiayaSewa += hrgBersih;
      totalPajak += pjk;

      if (statusSewa === "Perpanjang") countPerpanjang++;
      else if (statusSewa === "Kontrak Baru") countBaru++;
      else if (statusSewa === "Relokasi") countRelokasi++;

      listData.push({
        kodeId: data[i][idx('KODE ID')] || "-",
        lokasi: lokasi || "-",
        statusSewa: statusSewa || "-",
        hargaPokok: data[i][idx('SEWA BARU - HARGA SEBELUM PAJAK (NET)')] || 0,
        pajak: pjk,
        hargaBersih: hrgBersih,
        masaSewa: data[i][idx('SEWA BARU - MASA SEWA (TAHUN)')] || 1,
        tglHabisBaru: tglHabis instanceof Date ? tglHabis.toISOString().split('T')[0] : (tglHabis || "-"),
        ntpn: data[i][idx('NOMOR NTPN')] || "-"
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
  var pajakCalc = hitungPajak(hargaBaru, isDipotongDiSumber(data.dipotongPajak));
  var pajak = pajakCalc.nominalPajak;
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
  var finalValueStr = data.nilaiFinal ? data.nilaiFinal.replace("Rp ", "") : pajakCalc.nilaiGrossUp.toLocaleString('id-ID');
  body.replaceText('<<NILAI_FINAL>>', finalValueStr);
  body.replaceText('<<PAJAK>>', Math.round(pajak).toLocaleString('id-ID'));
  body.replaceText('<<MASA_SEWA>>', (data.masaSewa || 1) + " Tahun"); body.replaceText('<<CATATAN>>', data.catatanMemo || "-");

  // Skema Pembayaran: hanya menampilkan baris cicilan yang benar-benar terisi
  // (jumlah > 0). Kalau negosiator tidak mengisi cicilan sama sekali, tampil
  // sebagai pembayaran lunas sekali bayar. Sengaja ringkas (bukan paragraf
  // panjang) sesuai permintaan -- cukup label, nominal, dan tanggal.
  var skemaLines = [];
  var cicilanDefs = [
    { label: 'Cicilan 1', tgl: data.tglCicilan1, jml: data.jmlCicilan1 },
    { label: 'Cicilan 2', tgl: data.tglCicilan2, jml: data.jmlCicilan2 },
    { label: 'Cicilan 3', tgl: data.tglCicilan3, jml: data.jmlCicilan3 }
  ];
  cicilanDefs.forEach(function (c) {
    var jmlNum = parseFloat(String(c.jml || "0").replace(/\./g, '')) || 0;
    if (jmlNum > 0) {
      var tglFmt = c.tgl ? String(c.tgl).split("-").reverse().join("/") : "-";
      skemaLines.push(c.label + "  Rp " + Math.round(jmlNum).toLocaleString('id-ID') + "   " + tglFmt);
    }
  });
  var skemaPembayaran = skemaLines.length > 0 ? skemaLines.join("\n") : "Lunas (1x Pembayaran)";
  body.replaceText('<<SKEMA_PEMBAYARAN>>', skemaPembayaran);

  doc.saveAndClose(); var pdf = DriveApp.getFolderById('1i1iYPvDayVa_WXre_mMPsmlP_UEGdeMd').createFile(copy.getAs('application/pdf'));
  copy.setTrashed(true); return { url: pdf.getUrl(), id: pdf.getId() };
}

function checkAndPrepareSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheetKeu = ss.getSheetByName("NEW_KEU");
  if (!sheetKeu) {
    sheetKeu = ss.insertSheet("NEW_KEU");
    // Kolom "Cicilan Ke" & "Bukti Transfer Sewa (Direktur)" ditambahkan di
    // ujung -- dipakai fitur pencairan cicilan (lihat getCicilanSchedule).
    var headerRow1 = ["KODE ID", "PERIODE / TAHUN INPUT", "LOKASI", "STATUS SEWA", "INFORMASI SEWA SEBELUMNYA", "", "", "INFORMASI SEWA MENDATANG", "", "", "", "", "Fix Nilai (Arsip)", "", "N I K", "", "NPWP", "Sewa Ditransfer", "Arsip SPK", "", "", "INFO PAJAK", "", "", "", "ARSIP BUKPOT", "RENTANG WAKTU", "", "", "", "", "CICILAN", ""];
    var headerRow2 = ["", "", "", "", "Tanggal Habis Kontrak", "Harga Setelah Pajak (Ditransfer Bersih)", "Deposit", "Harga Sebelum Pajak (Pokok) Sewa", "Pajak Sewa (Yang dipotong)", "Harga Setelah Pajak (Ditransfer Bersih)", "Tanggal Habis Kontrak", "Penilaian Harga Sewa (Gross Up)", "Sewa Bersih", "Pajak", "Nama", "Nomor", "", "", "Soft", "Link", "Hard", "Masa Pajak", "Tanggal Bayar", "Nomor NTPN", "Nilai Pajak", "", "Email Negosiator", "Tanggal ACC MCX", "Tanggal ACC Sekretaris", "Tanggal ACC Direktur", "Tanggal Cair", "Cicilan Ke", "Bukti Transfer Sewa (Direktur)"];
    sheetKeu.getRange(1, 1, 1, headerRow1.length).setValues([headerRow1]).setBackground("#dcfce7").setFontWeight("bold");
    sheetKeu.getRange(2, 1, 1, headerRow2.length).setValues([headerRow2]).setBackground("#f1f5f9").setFontWeight("bold");
    // Baris 3 dibiarkan kosong sebagai pemisah visual, data mulai baris 4
    // (konsisten dengan NEW_KEU_HEADER_ROWS=3 / NEW_KEU_DATA_START_ROW=4).
  }

  var sheetRekap = ss.getSheetByName("NEW_REKAP");
  if (!sheetRekap) {
    sheetRekap = ss.insertSheet("NEW_REKAP");
    // Skema ini sudah diverifikasi langsung terhadap struktur produksi
    // (lihat NEW_REKAP_HEADER_START_ROW/NEW_REKAP_HEADER_ROWS/NEW_REKAP_DATA_START_ROW):
    // baris 1 = banner filter periode, baris 2 = kosong, baris 3-5 = header
    // berjenjang, baris 6 = metadata nomor kolom, data mulai baris 7.
    sheetRekap.getRange(1, 1, 1, 3).setValues([["", "PERIODE", ""]]).setFontWeight("bold");
    var headerRekapGroup = ["KODE ID", "WILAYAH / LOKASI", "STATUS SEWA", "INFORMASI SEWA SEBELUMNYA", "", "", "INFORMASI SEWA MENDATANG", "", "", "", "", "Pihak Pembayar Pajak PBB", "INFO REKENING BANK (Pemilik)", "", "", "", ""];
    var headerRekapSub = ["", "", "", "Sewa Lama - Tanggal Habis Kontrak", "Sewa Lama - Masa Sewa (Tahun)", "Sewa Lama - Harga Setelah Pajak (Net)", "Sewa Baru - Harga Sebelum Pajak (Net)", "Pajak Sewa (Yang dipotong)", "Sewa Baru - Harga Setelah Pajak (Net)", "Sewa Baru - Masa Sewa (Tahun)", "Sewa Baru - Tanggal Habis Kontrak", "", "", "Arsip SPK", "", "INFO PAJAK", ""];
    var headerRekapLeaf = ["", "", "", "", "", "", "", "", "", "", "", "", "", "Soft", "Hard", "Nomor NTPN", "Nilai Pajak"];
    sheetRekap.getRange(3, 1, 1, headerRekapGroup.length).setValues([headerRekapGroup]).setBackground("#2e1065").setFontColor("#ffffff").setFontWeight("bold");
    sheetRekap.getRange(4, 1, 1, headerRekapSub.length).setValues([headerRekapSub]).setBackground("#f1f5f9").setFontWeight("bold");
    sheetRekap.getRange(5, 1, 1, headerRekapLeaf.length).setValues([headerRekapLeaf]).setBackground("#f1f5f9").setFontWeight("bold");
    sheetRekap.getRange(6, 1, 1, headerRekapLeaf.length).setBackground("#000000");
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