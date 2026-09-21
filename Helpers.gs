/**
 * ==========================================================================
 * HELPERS.GS — Utility akses kolom berbasis nama header (bukan index angka)
 * ==========================================================================
 * Tujuan: menggantikan pola data[i][25], data[i][37] dst yang rapuh.
 * File ini BERDIRI SENDIRI — tidak mengubah fungsi apapun di Code.gs.
 * Aman untuk langsung ditambahkan ke project tanpa efek samping.
 *
 * CARA PAKAI (akan dipakai bertahap di sesi-sesi berikutnya):
 *
 *   var sheet = SpreadsheetApp.getActive().getSheetByName('NEW_INPUT');
 *   var colMap = getColIndexMap(sheet, 1); // header ada di baris ke-1
 *   var data = sheet.getDataRange().getValues();
 *
 *   for (var i = 2; i < data.length; i++) {
 *     var kodeId   = data[i][colMap['KODE ID']];
 *     var lokasi   = data[i][colMap['LOKASI']];
 *     var nik      = data[i][colMap['NIK']];
 *   }
 *
 * Kalau nanti ada kolom baru disisipkan di tengah sheet, kode di atas
 * TETAP BENAR tanpa perlu diubah — karena index dicari ulang by nama,
 * bukan hardcode angka.
 * ==========================================================================
 */

/**
 * Membaca 1 baris header dari sheet dan mengembalikan Map { namaHeader: indexKolom(0-based) }.
 * Jika ada gabungan header 2 baris (contoh: grup "INFO PAJAK" di baris 2,
 * lalu sub-header "Nomor NTPN" di baris 3), pakai parameter headerRow
 * untuk pilih baris mana yang jadi acuan nama kolom (biasanya baris paling detail).
 *
 * @param {Sheet} sheet - object sheet, contoh: SpreadsheetApp.getActive().getSheetByName('NEW_INPUT')
 * @param {number} headerRow - nomor baris header (1-based). Default 1.
 * @return {Object} peta { 'NAMA HEADER': indexKolom0Based }
 */
function getColIndexMap(sheet, headerRow) {
  if (!sheet) {
    throw new Error('getColIndexMap: sheet tidak ditemukan (null/undefined).');
  }
  headerRow = headerRow || 1;

  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    throw new Error('getColIndexMap: sheet "' + sheet.getName() + '" tidak memiliki kolom sama sekali.');
  }

  var headerValues = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var map = {};
  var duplicateWarnings = [];

  for (var c = 0; c < headerValues.length; c++) {
    var rawName = headerValues[c];
    if (rawName === '' || rawName === null || rawName === undefined) continue;

    var key = normalizeHeaderName(rawName);

    if (map.hasOwnProperty(key)) {
      // Header duplikat (misal ada 2 kolom "Tanggal" tanpa konteks grup).
      // Tidak melempar error karena beberapa sheet SIMPEWA memang punya
      // sub-header berulang (LUNAS/CICILAN 1, 2, 3 dst) — cukup catat
      // peringatan supaya kamu tahu dan bisa disambiguasi manual kalau perlu.
      duplicateWarnings.push(key + ' (kolom ke-' + (c + 1) + ', sebelumnya kolom ke-' + (map[key] + 1) + ')');
    }

    map[key] = c; // 0-based, cocok langsung dengan array getValues()
  }

  if (duplicateWarnings.length > 0) {
    Logger.log('[getColIndexMap] Peringatan header duplikat di sheet "' + sheet.getName() + '": ' + duplicateWarnings.join('; '));
  }

  return map;
}

/**
 * Ambil index kolom tunggal (0-based) berdasarkan nama header, dengan error
 * yang jelas kalau header tidak ditemukan — supaya bug ketahuan langsung
 * saat development, bukan diam-diam mengembalikan undefined.
 *
 * @param {Object} colMap - hasil dari getColIndexMap()
 * @param {string} headerName - nama header yang dicari
 * @return {number} index kolom 0-based
 */
function getColIndex(colMap, headerName) {
  var key = normalizeHeaderName(headerName);
  if (!colMap.hasOwnProperty(key)) {
    throw new Error('getColIndex: header "' + headerName + '" tidak ditemukan di sheet. ' +
      'Cek apakah nama header di spreadsheet sudah sesuai, atau apakah kolom ini memang belum ada.');
  }
  return colMap[key];
}

/**
 * Normalisasi nama header supaya pencocokan tidak gagal karena beda kapital,
 * spasi ganda, atau spasi di awal/akhir. Contoh: "  Kode Id " -> "KODE ID"
 *
 * @param {string} name
 * @return {string}
 */
function normalizeHeaderName(name) {
  return String(name)
    .trim()
    .toUpperCase()
    .replace(/\s+/g, ' ');
}

/**
 * ==========================================================================
 * UPGRADE: header berjenjang (1-3 baris) otomatis
 * ==========================================================================
 * Dipakai untuk sheet seperti NEW_INPUT yang punya campuran kedalaman:
 * - Kolom biasa: nama lengkap sudah ada di 1 baris saja.
 * - Kolom grup (contoh "SKEMA PEMBAYARAN"): baris 1 = judul grup,
 *   baris 2 = sub-grup, baris 3 = field asli (Tanggal/Jumlah, berulang).
 *
 * Fungsi ini otomatis:
 * 1. Mengisi sel kosong akibat merge (forward-fill), tapi dibatasi supaya
 *    tidak "bocor" melewati batas grup di atasnya.
 * 2. Mengambil nama paling detail (paling bawah yang terisi) sebagai key.
 * 3. Kalau nama itu ternyata dipakai berulang di kolom lain (misal "TANGGAL"
 *    muncul 3x untuk 3 skema cicilan berbeda), otomatis ditambah prefix
 *    nama induknya supaya jadi unik, contoh:
 *    "LUNAS / CICILAN 1 - TANGGAL", "LUNAS / CICILAN 2 - TANGGAL", dst.
 *    Kolom yang namanya sudah unik TIDAK diubah, tetap pakai nama pendek.
 *
 * @param {Sheet} sheet
 * @param {number} maxHeaderRows - jumlah baris header yang mau dipindai (mis. 3)
 * @param {number} startRow - baris pertama header (1-based). Default 1.
 *   Berguna kalau sheet punya baris "banner"/filter kontrol di atas header
 *   kolom asli (misal baris judul "PERIODE [dropdown tahun]") -- set
 *   startRow ke baris pertama header KOLOM yang sebenarnya.
 * @return {Object} peta { 'NAMA KOLOM (sudah disambiguasi jika perlu)': indexKolom0Based }
 */
function getColIndexMapAuto(sheet, maxHeaderRows, startRow) {
  if (!sheet) {
    throw new Error('getColIndexMapAuto: sheet tidak ditemukan (null/undefined).');
  }
  maxHeaderRows = maxHeaderRows || 3;
  startRow = startRow || 1;
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    throw new Error('getColIndexMapAuto: sheet "' + sheet.getName() + '" tidak memiliki kolom sama sekali.');
  }

  // Baca grid header mentah (baris x kolom, 0-based di dalam window ini).
  var range = sheet.getRange(startRow, 1, maxHeaderRows, lastCol);
  var grid = range.getValues(); // grid[row][col], row 0..maxHeaderRows-1

  // Tempel ulang nilai merge cell ke SETIAP sel yang tercakup merge itu --
  // ini menggantikan tebakan forward-fill dengan fakta asli dari sheet,
  // jadi kolom tunggal yang "nyempil" di antara grup (misal karena mereka
  // di-merge vertikal 2-3 baris) terbaca benar tanpa perlu ditebak.
  var merges = range.getMergedRanges();
  merges.forEach(function (m) {
    var mRow0 = m.getRow() - startRow;       // baris pertama merge, relatif window
    var mRow1 = m.getLastRow() - startRow;   // baris terakhir merge, relatif window
    var mCol0 = m.getColumn() - 1;           // kolom pertama merge, 0-based
    var mCol1 = m.getLastColumn() - 1;       // kolom terakhir merge, 0-based
    var val = sheet.getRange(m.getRow(), m.getColumn()).getValue();
    for (var rr = Math.max(0, mRow0); rr <= Math.min(maxHeaderRows - 1, mRow1); rr++) {
      for (var cc = Math.max(0, mCol0); cc <= Math.min(lastCol - 1, mCol1); cc++) {
        grid[rr][cc] = val;
      }
    }
  });

  // Tentukan "deepest label" per kolom: nilai paling detail (baris paling
  // bawah yang terisi), dan nama induk = nilai non-kosong terdekat di atas
  // baris itu (melompati baris yang genuinely kosong di antaranya).
  var deepest = [];
  var deepestParent = [];
  for (var c = 0; c < lastCol; c++) {
    var chosen = '', chosenRow = -1;
    for (var r = maxHeaderRows - 1; r >= 0; r--) {
      if (grid[r][c] !== '' && grid[r][c] !== null && grid[r][c] !== undefined) {
        chosen = normalizeHeaderName(grid[r][c]);
        chosenRow = r;
        break;
      }
    }
    deepest.push(chosen);

    var parentName = '';
    for (var pr = chosenRow - 1; pr >= 0; pr--) {
      if (grid[pr][c] !== '' && grid[pr][c] !== null && grid[pr][c] !== undefined) {
        parentName = normalizeHeaderName(grid[pr][c]);
        break;
      }
    }
    deepestParent.push(parentName);
  }

  // Hitung frekuensi tiap deepest label untuk deteksi duplikat.
  var freq = {};
  for (var c3 = 0; c3 < lastCol; c3++) {
    if (deepest[c3] === '') continue;
    freq[deepest[c3]] = (freq[deepest[c3]] || 0) + 1;
  }

  // Bangun map final: pakai nama pendek jika unik, prefix induk jika duplikat.
  var map = {};
  var unresolvedDuplicates = [];
  for (var c4 = 0; c4 < lastCol; c4++) {
    if (deepest[c4] === '') continue;
    var finalKey = deepest[c4];
    if (freq[deepest[c4]] > 1) {
      if (deepestParent[c4]) {
        finalKey = deepestParent[c4] + ' - ' + deepest[c4];
      } else {
        unresolvedDuplicates.push(deepest[c4] + ' (kolom ke-' + (c4 + 1) + ', tidak ada induk untuk disambiguasi)');
      }
    }
    if (map.hasOwnProperty(finalKey)) {
      unresolvedDuplicates.push(finalKey + ' (kolom ke-' + (c4 + 1) + ', bentrok dengan kolom lain)');
    }
    map[finalKey] = c4;
  }

  if (unresolvedDuplicates.length > 0) {
    Logger.log('[getColIndexMapAuto] Masih ada nama kolom bentrok yang tidak bisa auto-disambiguasi di sheet "' +
      sheet.getName() + '": ' + unresolvedDuplicates.join('; ') + '. Cek manual header sheet ini.');
  }

  return map;
}

/**
 * Versi diagnostik untuk getColIndexMapAuto — jalankan manual untuk verifikasi
 * sebelum dipakai di fungsi produksi.
 *
 * @param {string} sheetName
 * @param {number} maxHeaderRows
 * @param {number} startRow - baris pertama header (default 1)
 */
function debugColIndexMapAuto(sheetName, maxHeaderRows, startRow) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('Sheet "' + sheetName + '" tidak ditemukan.');
    return;
  }
  var map = getColIndexMapAuto(sheet, maxHeaderRows || 3, startRow || 1);
  var entries = [];
  for (var key in map) {
    entries.push({ key: key, idx: map[key] });
  }
  entries.sort(function (a, b) { return a.idx - b.idx; });
  var lines = entries.map(function (e) { return 'Kolom ke-' + (e.idx + 1) + ' -> "' + e.key + '"'; });
  Logger.log('=== Peta kolom AUTO sheet "' + sheetName + '" (' + (maxHeaderRows || 3) + ' baris header mulai baris ' + (startRow || 1) + ') ===\n' + lines.join('\n'));
}

/**
 * Fungsi diagnostik cepat — jalankan manual dari editor Apps Script untuk
 * memastikan header sheet terbaca dengan benar sebelum dipakai di fungsi lain.
 * Buka menu: pilih fungsi ini di dropdown "Select function" lalu klik Run,
 * hasilnya cek di menu Executions atau View > Logs (Ctrl+Enter).
 *
 * @param {string} sheetName - nama sheet yang ingin dicek, contoh 'NEW_INPUT'
 * @param {number} headerRow - baris header, default 1
 */
function debugColIndexMap(sheetName, headerRow) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('Sheet "' + sheetName + '" tidak ditemukan. Cek nama sheet-nya.');
    return;
  }
  var map = getColIndexMap(sheet, headerRow || 1);
  var lines = [];
  for (var key in map) {
    lines.push('Kolom ke-' + (map[key] + 1) + ' -> "' + key + '"');
  }
  Logger.log('=== Peta kolom sheet "' + sheetName + '" (header baris ' + (headerRow || 1) + ') ===\n' + lines.join('\n'));
}

function tesHelperRekap() {
  debugColIndexMapAuto('NEW_REKAP', 3, 3);
}

function tesHelperInputUlang() {
  debugColIndexMapAuto('NEW_INPUT', 3);
}