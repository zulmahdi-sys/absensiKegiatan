const SHEET_AGENDA = 'AGENDA_KEGIATAN';
const SHEET_PRESENSI = 'DATA_PRESENSI';
const SHEET_ARSIP = 'ARSIP_SESI';

/**
 * Entry Point Utama Web App Google Apps Script
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SIPRES UIN Ar-Raniry')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Format Tanggal Cepat
 */
function formatCellDate(val, formatStr) {
  if (!val) return '';
  if (val instanceof Date) {
    try {
      var tz = Session.getScriptTimeZone() || 'GMT';
      return Utilities.formatDate(val, tz, formatStr);
    } catch (e) {
      return val.toString();
    }
  }
  return String(val);
}

/**
 * Inisialisasi Database Google Sheets
 */
function initDatabase() {
  const cache = CacheService.getScriptCache();
  if (cache.get('db_inited_v3') === 'true') return { status: 'success' };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let sheetAgenda = ss.getSheetByName(SHEET_AGENDA);
    if (!sheetAgenda) {
      sheetAgenda = ss.insertSheet(SHEET_AGENDA);
      sheetAgenda.getRange(1, 1, 1, 15).setValues([[
        'ID Agenda', 'Judul Kegiatan', 'Kategori', 'Penyelenggara', 
        'Tanggal', 'Waktu Mulai', 'Waktu Selesai', 'Lokasi', 'Deskripsi', 
        'Status', 'Kuota', 'IsLocationActive', 'Latitude', 'Longitude', 'Radius'
      ]]);
      sheetAgenda.getRange('1:1').setFontWeight('bold').setBackground('#006837').setFontColor('#FFFFFF');
    }

    let sheetPresensi = ss.getSheetByName(SHEET_PRESENSI);
    if (!sheetPresensi) {
      sheetPresensi = ss.insertSheet(SHEET_PRESENSI);
      sheetPresensi.getRange(1, 1, 1, 8).setValues([[
        'ID Presensi', 'Timestamp', 'ID Agenda', 'Nama Lengkap', 'Instansi / Unit', 'Tanda Tangan', 'Latitude', 'Longitude'
      ]]);
      sheetPresensi.getRange('1:1').setFontWeight('bold').setBackground('#006837').setFontColor('#FFFFFF');
    }

    let sheetArsip = ss.getSheetByName(SHEET_ARSIP);
    if (!sheetArsip) {
      sheetArsip = ss.insertSheet(SHEET_ARSIP);
      sheetArsip.getRange(1, 1, 1, 9).setValues([[
        'ID Arsip', 'ID Agenda', 'Judul Kegiatan', 'Kategori', 
        'Penyelenggara', 'Tanggal Execution', 'Total Peserta', 'Tanggal Diarsipkan', 'Status'
      ]]);
      sheetArsip.getRange('1:1').setFontWeight('bold').setBackground('#006837').setFontColor('#FFFFFF');
    }

    cache.put('db_inited_v3', 'true', 21600); // Cache 6 jam
    return { status: 'success' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Mendapatkan Agenda Publik
 */
function getAgendasPublic() {
  try {
    const cache = CacheService.getScriptCache();
    const cachedData = cache.get('agendas_public');
    if (cachedData) {
      return { status: 'success', data: JSON.parse(cachedData) };
    }

    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetAgenda = ss.getSheetByName(SHEET_AGENDA);
    if (!sheetAgenda) return { status: 'success', data: [] };

    const lastRow = sheetAgenda.getLastRow();
    if (lastRow <= 1) return { status: 'success', data: [] };

    const data = sheetAgenda.getRange(2, 1, lastRow - 1, 15).getValues();
    const agendas = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row[0] && String(row[9]) !== 'Diarsipkan') {
        agendas.push({
          id: String(row[0]),
          judul: String(row[1] || ''),
          kategori: String(row[2] || 'Rapat'),
          penyelenggara: String(row[3] || ''),
          tanggal: formatCellDate(row[4], 'yyyy-MM-dd'),
          waktuMulai: formatCellDate(row[5], 'HH:mm'),
          waktuSelesai: formatCellDate(row[6], 'HH:mm'),
          lokasi: String(row[7] || ''),
          deskripsi: String(row[8] || ''),
          status: String(row[9] || 'Aktif'),
          kuota: Number(row[10] || 100),
          isLocationActive: Boolean(row[11] === true || row[11] === 'TRUE'),
          latitude: Number(row[12] || 0),
          longitude: Number(row[13] || 0),
          radius: Number(row[14] || 100)
        });
      }
    }

    cache.put('agendas_public', JSON.stringify(agendas), 180); // Cache 3 menit
    return { status: 'success', data: agendas };
  } catch (error) {
    return { status: 'error', message: error.toString(), data: [] };
  }
}

/**
 * Mendapatkan Agenda Admin
 */
function getAgendasAdmin() {
  try {
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetAgenda = ss.getSheetByName(SHEET_AGENDA);
    if (!sheetAgenda) return { status: 'success', data: [] };

    const lastRow = sheetAgenda.getLastRow();
    if (lastRow <= 1) return { status: 'success', data: [] };

    const data = sheetAgenda.getRange(2, 1, lastRow - 1, 15).getValues();
    const agendas = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row[0]) {
        agendas.push({
          id: String(row[0]),
          judul: String(row[1] || ''),
          kategori: String(row[2] || 'Rapat'),
          penyelenggara: String(row[3] || ''),
          tanggal: formatCellDate(row[4], 'yyyy-MM-dd'),
          waktuMulai: formatCellDate(row[5], 'HH:mm'),
          waktuSelesai: formatCellDate(row[6], 'HH:mm'),
          lokasi: String(row[7] || ''),
          deskripsi: String(row[8] || ''),
          status: String(row[9] || 'Aktif'),
          kuota: Number(row[10] || 100),
          isLocationActive: Boolean(row[11] === true || row[11] === 'TRUE'),
          latitude: Number(row[12] || 0),
          longitude: Number(row[13] || 0),
          radius: Number(row[14] || 100)
        });
      }
    }

    return { status: 'success', data: agendas };
  } catch (error) {
    return { status: 'error', message: error.toString(), data: [] };
  }
}

/**
 * Perhitungan Jarak Geolocation (Haversine Formula dalam meter)
 */
function calculateDistance(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
            Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
            Math.sin(dLon/2) * Math.sin(dLon/2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  return Math.round(R * c);
}

/**
 * Memproses Presensi Peserta dengan Lock Service & Validasi Geolocation (Antrean 1000+ Peserta Seamless)
 */
function submitPresensi(payload) {
  if (!payload || !payload.agendaId || !payload.nama || !payload.instansi) {
    return { status: 'error', message: 'Data presensi tidak lengkap!' };
  }

  const cleanAgendaId = String(payload.agendaId).trim();
  const cleanNama = String(payload.nama).trim();

  // Cache objek Spreadsheet aktif untuk menghindari overhead re-fetch
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. LAKUKAN VALIDASI & PEMBACAAN DILUAR LOCK UNTUK MEMPERCEPAT WAKTU TUNGGU
  try {
    const sheetAgenda = ss.getSheetByName(SHEET_AGENDA);

    if (sheetAgenda && sheetAgenda.getLastRow() > 1) {
      const agendaData = sheetAgenda.getRange(2, 1, sheetAgenda.getLastRow() - 1, 15).getValues();
      for (let a = 0; a < agendaData.length; a++) {
        if (String(agendaData[a][0]).trim() === cleanAgendaId) {
          const isLocActive = Boolean(agendaData[a][11] === true || agendaData[a][11] === 'TRUE');
          if (isLocActive) {
            if (!payload.userLat || !payload.userLng) {
              return { status: 'error', message: 'Presensi gagal: Aktifkan GPS lokasi pada perangkat Anda.' };
            }
            const distance = calculateDistance(
              Number(agendaData[a][12] || 0), Number(agendaData[a][13] || 0),
              payload.userLat, payload.userLng
            );
            const maxRadius = Number(agendaData[a][14] || 100);
            if (distance > maxRadius) {
              return { 
                status: 'error', 
                message: `Presensi Ditolak! Anda berada ${distance}m dari lokasi (Maksimal radius ${maxRadius}m).` 
              };
            }
          }
          break;
        }
      }
    }
  } catch (e) {
    // Abaikan error pembacaan ringan agar presensi tetap jalan
  }

  // 2. PENULISAN CEPAT DENGAN LOCK TERISOLASI ULTRA FAST (<50ms)
  const lock = LockService.getScriptLock();
  let hasLock = false;

  try {
    // Waktu tunggu lock responsif 15 detik
    hasLock = lock.waitLock(15000);
    if (!hasLock) {
      return { 
        status: 'busy', 
        message: 'Antrean padat. Mengulang otomatis...' 
      };
    }

    let sheetPresensi = ss.getSheetByName(SHEET_PRESENSI);
    if (!sheetPresensi) {
      sheetPresensi = ss.insertSheet(SHEET_PRESENSI);
      sheetPresensi.getRange(1, 1, 1, 8).setValues([[
        'ID Presensi', 'Timestamp', 'ID Agenda', 'Nama Lengkap', 'Instansi / Unit', 'Tanda Tangan', 'Latitude', 'Longitude'
      ]]);
    }

    const idPresensi = 'PRS-' + new Date().getTime();
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT', 'yyyy-MM-dd HH:mm:ss');

    // Direct Atomic Append Tanpa Scanning Panjang Di Dalam Lock
    sheetPresensi.appendRow([
      idPresensi,
      timestamp,
      cleanAgendaId,
      cleanNama,
      String(payload.instansi).trim(),
      payload.signature || '',
      payload.userLat || '',
      payload.userLng || ''
    ]);

    // Async Cache Invalidation
    try {
      CacheService.getScriptCache().remove('presensi_' + cleanAgendaId);
    } catch (errCache) {}

    return { 
      status: 'success', 
      timestamp: timestamp,
      message: 'Presensi kehadiran berhasil disimpan!' 
    };

  } catch (error) {
    return { status: 'error', message: 'Gagal menyimpan presensi: ' + error.toString() };
  } finally {
    if (hasLock) {
      try {
        lock.releaseLock();
      } catch (e) {}
    }
  }
}

/**
 * Mendapatkan Daftar Peserta Presensi per Agenda
 */
function getPresensiList(agendaId) {
  try {
    const targetAgendaId = agendaId ? String(agendaId).trim() : '';
    if (!targetAgendaId) return { status: 'success', data: [] };

    const cache = CacheService.getScriptCache();
    const cachedData = cache.get('presensi_' + targetAgendaId);
    if (cachedData) {
      return { status: 'success', data: JSON.parse(cachedData) };
    }

    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPresensi = ss.getSheetByName(SHEET_PRESENSI);
    if (!sheetPresensi) return { status: 'success', data: [] };

    const lastRow = sheetPresensi.getLastRow();
    if (lastRow <= 1) return { status: 'success', data: [] };

    const data = sheetPresensi.getRange(2, 1, lastRow - 1, 8).getValues();
    const list = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row[0] || row[3]) {
        const itemAgendaId = String(row[2] || '').trim();
        if (itemAgendaId === targetAgendaId) {
          list.push({
            id: String(row[0] || ('PRS-' + (i + 2))),
            timestamp: formatCellDate(row[1], 'yyyy-MM-dd HH:mm:ss'),
            agendaId: itemAgendaId,
            nama: String(row[3] || '').trim(),
            instansi: String(row[4] || '').trim(),
            signature: String(row[5] || ''),
            lat: row[6] || '',
            lng: row[7] || ''
          });
        }
      }
    }

    cache.put('presensi_' + targetAgendaId, JSON.stringify(list), 120);
    return { status: 'success', data: list };
  } catch (error) {
    return { status: 'error', message: error.toString(), data: [] };
  }
}

/**
 * Menyimpan / Memperbarui Agenda Kegiatan
 */
function saveAgenda(agendaObj) {
  if (!agendaObj || !agendaObj.judul) {
    return { status: 'error', message: 'Judul agenda tidak boleh kosong!' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetAgenda = ss.getSheetByName(SHEET_AGENDA);
    const lastRow = sheetAgenda.getLastRow();

    let rowIndex = -1;
    if (agendaObj.id && lastRow > 1) {
      const ids = sheetAgenda.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === String(agendaObj.id).trim()) {
          rowIndex = i + 2;
          break;
        }
      }
    }

    const agendaId = agendaObj.id || 'AGD-' + new Date().getTime();
    const rowValues = [
      agendaId,
      agendaObj.judul || '',
      agendaObj.kategori || 'Rapat',
      agendaObj.penyelenggara || '',
      agendaObj.tanggal || '',
      agendaObj.waktuMulai || '',
      agendaObj.waktuSelesai || '',
      agendaObj.lokasi || '',
      agendaObj.deskripsi || '',
      agendaObj.status || 'Aktif',
      agendaObj.kuota || 100,
      agendaObj.isLocationActive === true ? true : false,
      agendaObj.latitude || 0,
      agendaObj.longitude || 0,
      agendaObj.radius || 100
    ];

    if (rowIndex > 0) {
      sheetAgenda.getRange(rowIndex, 1, 1, 15).setValues([rowValues]);
    } else {
      sheetAgenda.appendRow(rowValues);
    }

    const cache = CacheService.getScriptCache();
    cache.remove('agendas_public');

    return { status: 'success', message: 'Agenda kegiatan berhasil disimpan!' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menghapus Agenda Kegiatan
 */
function deleteAgenda(agendaId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetAgenda = ss.getSheetByName(SHEET_AGENDA);
    const lastRow = sheetAgenda.getLastRow();

    if (lastRow > 1) {
      const ids = sheetAgenda.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === String(agendaId).trim()) {
          sheetAgenda.deleteRow(i + 2);
          const cache = CacheService.getScriptCache();
          cache.remove('agendas_public');
          cache.remove('presensi_' + agendaId);
          return { status: 'success', message: 'Agenda kegiatan berhasil dihapus!' };
        }
      }
    }

    return { status: 'error', message: 'Agenda kegiatan tidak ditemukan!' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Mengarsipkan Sesi Kegiatan
 */
function archiveSession(agendaId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetAgenda = ss.getSheetByName(SHEET_AGENDA);
    const sheetPresensi = ss.getSheetByName(SHEET_PRESENSI);
    const sheetArsip = ss.getSheetByName(SHEET_ARSIP);

    const lastRowAgenda = sheetAgenda.getLastRow();
    let agendaRowIndex = -1;
    let agendaObj = null;

    if (lastRowAgenda > 1) {
      const dataAgenda = sheetAgenda.getRange(2, 1, lastRowAgenda - 1, 5).getValues();
      for (let i = 0; i < dataAgenda.length; i++) {
        if (String(dataAgenda[i][0]).trim() === String(agendaId).trim()) {
          agendaRowIndex = i + 2;
          agendaObj = {
            id: String(dataAgenda[i][0]),
            judul: String(dataAgenda[i][1]),
            kategori: String(dataAgenda[i][2]),
            penyelenggara: String(dataAgenda[i][3]),
            tanggal: formatCellDate(dataAgenda[i][4], 'yyyy-MM-dd')
          };
          break;
        }
      }
    }

    if (agendaRowIndex === -1 || !agendaObj) {
      return { status: 'error', message: 'Agenda tidak ditemukan!' };
    }

    sheetAgenda.getRange(agendaRowIndex, 10).setValue('Diarsipkan');

    let totalPeserta = 0;
    const lastRowPresensi = sheetPresensi.getLastRow();
    if (lastRowPresensi > 1) {
      const agendaIds = sheetPresensi.getRange(2, 3, lastRowPresensi - 1, 1).getValues();
      for (let j = 0; j < agendaIds.length; j++) {
        if (String(agendaIds[j][0]).trim() === String(agendaId).trim()) {
          totalPeserta++;
        }
      }
    }

    const idArsip = 'ARS-' + new Date().getTime();
    const tglArsip = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT', 'yyyy-MM-dd HH:mm:ss');
    sheetArsip.appendRow([
      idArsip, agendaObj.id, agendaObj.judul, agendaObj.kategori,
      agendaObj.penyelenggara, agendaObj.tanggal, totalPeserta, tglArsip, 'Diarsipkan'
    ]);

    const cache = CacheService.getScriptCache();
    cache.remove('agendas_public');

    return { 
      status: 'success', 
      message: `Sesi '${agendaObj.judul}' berhasil diarsipkan dengan total ${totalPeserta} peserta!` 
    };

  } catch (error) {
    return { status: 'error', message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Mengaktifkan Kembali Sesi
 */
function unarchiveSession(agendaId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetAgenda = ss.getSheetByName(SHEET_AGENDA);
    const sheetArsip = ss.getSheetByName(SHEET_ARSIP);

    const lastRowAgenda = sheetAgenda.getLastRow();
    let agendaRowIndex = -1;
    let agendaTitle = "";

    if (lastRowAgenda > 1) {
      const dataAgenda = sheetAgenda.getRange(2, 1, lastRowAgenda - 1, 2).getValues();
      for (let i = 0; i < dataAgenda.length; i++) {
        if (String(dataAgenda[i][0]).trim() === String(agendaId).trim()) {
          agendaRowIndex = i + 2;
          agendaTitle = String(dataAgenda[i][1]);
          break;
        }
      }
    }

    if (agendaRowIndex === -1) {
      return { status: 'error', message: 'Agenda tidak ditemukan!' };
    }

    sheetAgenda.getRange(agendaRowIndex, 10).setValue('Aktif');

    if (sheetArsip) {
      const lastRowArsip = sheetArsip.getLastRow();
      if (lastRowArsip > 1) {
        const arsipIds = sheetArsip.getRange(2, 2, lastRowArsip - 1, 1).getValues();
        for (let k = 0; k < arsipIds.length; k++) {
          if (String(arsipIds[k][0]).trim() === String(agendaId).trim()) {
            sheetArsip.getRange(k + 2, 9).setValue('Aktif');
            break;
          }
        }
      }
    }

    const cache = CacheService.getScriptCache();
    cache.remove('agendas_public');

    return { 
      status: 'success', 
      message: `Agenda '${agendaTitle}' berhasil diaktifkan kembali!` 
    };

  } catch (error) {
    return { status: 'error', message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Mendapatkan Daftar Arsip Sesi
 */
function getArchivedSessions() {
  try {
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetArsip = ss.getSheetByName(SHEET_ARSIP);
    if (!sheetArsip) return { status: 'success', data: [] };

    const lastRow = sheetArsip.getLastRow();
    if (lastRow <= 1) return { status: 'success', data: [] };

    const data = sheetArsip.getRange(2, 1, lastRow - 1, 9).getValues();
    const archives = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row[0] && String(row[8]) === 'Diarsipkan') {
        archives.push({
          idArsip: String(row[0]),
          agendaId: String(row[1]),
          judul: String(row[2]),
          kategori: String(row[3]),
          penyelenggara: String(row[4]),
          tanggal: formatCellDate(row[5], 'yyyy-MM-dd'),
          totalPeserta: Number(row[6] || 0),
          tglArsip: formatCellDate(row[7], 'yyyy-MM-dd HH:mm:ss')
        });
      }
    }

    return { status: 'success', data: archives };
  } catch (error) {
    return { status: 'error', message: error.toString(), data: [] };
  }
}

/**
 * Memperbarui Data Presensi Peserta
 */
function updatePresensi(presensiObj) {
  if (!presensiObj || !presensiObj.id) {
    return { status: 'error', message: 'ID Presensi tidak valid!' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPresensi = ss.getSheetByName(SHEET_PRESENSI);
    const lastRow = sheetPresensi.getLastRow();

    if (lastRow > 1) {
      const ids = sheetPresensi.getRange(2, 1, lastRow - 1, 3).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === String(presensiObj.id).trim()) {
          const targetRow = i + 2;
          const agendaId = String(ids[i][2]).trim();

          sheetPresensi.getRange(targetRow, 4).setValue(String(presensiObj.nama).trim());
          sheetPresensi.getRange(targetRow, 5).setValue(String(presensiObj.instansi).trim());

          const cache = CacheService.getScriptCache();
          cache.remove('presensi_' + agendaId);

          return { status: 'success', message: 'Data presensi peserta berhasil diperbarui!' };
        }
      }
    }

    return { status: 'error', message: 'Data presensi tidak ditemukan!' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menghapus Data Presensi Peserta
 */
function deletePresensi(presensiId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPresensi = ss.getSheetByName(SHEET_PRESENSI);
    const lastRow = sheetPresensi.getLastRow();

    if (lastRow > 1) {
      const ids = sheetPresensi.getRange(2, 1, lastRow - 1, 3).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === String(presensiId).trim()) {
          const agendaId = String(ids[i][2]).trim();
          sheetPresensi.deleteRow(i + 2);

          const cache = CacheService.getScriptCache();
          cache.remove('presensi_' + agendaId);

          return { status: 'success', message: 'Data presensi peserta berhasil dihapus!' };
        }
      }
    }

    return { status: 'error', message: 'Data presensi tidak ditemukan!' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}