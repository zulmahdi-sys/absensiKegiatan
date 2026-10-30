/**
 * Sistem Informasi Presensi Kegiatan - UIN Ar-Raniry Banda Aceh
 * Backend Script (code.gs) - Compatible with Vercel Web Hosting & GAS
 */

// ID Google Spreadsheet Resmi Tujuan
const SPREADSHEET_ID = '1ZehIX6YgrrYy6HL8bh3APB_od72blKF_Aj5talrc6go';

function doGet(e) {
  // Jika dipanggil via REST API HTTP (misal dari Vercel / External Web)
  if (e && e.parameter && e.parameter.action) {
    let result = [];
    if (e.parameter.action === 'getEvents') {
      result = getEvents();
    } else if (e.parameter.action === 'getAttendees') {
      result = getAttendees(e.parameter.eventId);
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Jika dibuka langsung di Google Apps Script HTML Service
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Sistem Presensi Kegiatan - UIN Ar-Raniry')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let result = { success: false, error: "Invalid Request" };
  try {
    let postData = {};
    if (e.postData && e.postData.contents) {
      postData = JSON.parse(e.postData.contents);
    }
    const action = postData.action;

    if (action === 'getEvents') {
      result = getEvents();
    } else if (action === 'addNewEvent') {
      result = addNewEvent(postData.data);
    } else if (action === 'updateEvent') {
      result = updateEvent(postData.eventId, postData.data);
    } else if (action === 'deleteEvent') {
      result = deleteEvent(postData.eventId);
    } else if (action === 'submitAttendance') {
      result = submitAttendance(postData.data);
    } else if (action === 'getAttendees') {
      result = getAttendees(postData.eventId);
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

function getEvents() {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('Events');
    if (!sheet) {
      sheet = ss.insertSheet('Events');
      sheet.appendRow(['id', 'title', 'type', 'date', 'time', 'location', 'status', 'lat', 'lng', 'radius']);
      return [];
    }
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const events = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        events.push({
          id: String(data[i][0]),
          title: String(data[i][1] || ''),
          type: String(data[i][2] || 'RAPAT'),
          date: String(data[i][3] || ''),
          time: String(data[i][4] || ''),
          location: String(data[i][5] || ''),
          status: String(data[i][6] || 'Active'),
          lat: parseCoord(data[i][7], 5.5714),
          lng: parseCoord(data[i][8], 95.3703),
          radius: parseInt(data[i][9]) || 200
        });
      }
    }
    return events;
  } catch (e) {
    return [];
  }
}

// Handler Khusus: Mencegah & Memperbaiki Pergeseran Desimal akibat Format Regional Google Sheets
function parseCoord(val, defaultVal) {
  if (val === null || val === undefined || val === '') return defaultVal;
  
  let strVal = String(val).replace(',', '.').trim();
  let num = typeof val === 'number' ? val : parseFloat(strVal);
  if (isNaN(num)) return defaultVal;

  if (Math.abs(num) > 10000) num = num / 10000;
  else if (Math.abs(num) > 1000) num = num / 1000;
  else if (Math.abs(num) > 100) num = num / 100;

  if (defaultVal < 10 && Math.abs(num) > 90) {
    num = num / 10;
  }

  return num;
}

function addNewEvent(eventData) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('Events');
    if (!sheet) {
      sheet = ss.insertSheet('Events');
      sheet.appendRow(['id', 'title', 'type', 'date', 'time', 'location', 'status', 'lat', 'lng', 'radius']);
    }
    
    const newId = 'EVT-' + new Date().getTime();
    const cleanLat = parseCoord(eventData.lat, 5.5714);
    const cleanLng = parseCoord(eventData.lng, 95.3703);

    sheet.appendRow([
      newId,
      eventData.title || '',
      eventData.type || 'RAPAT',
      eventData.date || '',
      eventData.time || '',
      eventData.location || '',
      eventData.status || 'Active',
      "'" + cleanLat.toFixed(6),
      "'" + cleanLng.toFixed(6),
      parseInt(eventData.radius) || 200
    ]);
    
    return { success: true, id: newId };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function updateEvent(eventId, eventData) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('Events');
    if (!sheet) return { success: false, error: "Tabel Events tidak ditemukan." };
    
    const data = sheet.getDataRange().getValues();
    const cleanLat = parseCoord(eventData.lat, 5.5714);
    const cleanLng = parseCoord(eventData.lng, 95.3703);

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(eventId)) {
        sheet.getRange(i + 1, 2, 1, 9).setValues([[
          eventData.title || '',
          eventData.type || 'RAPAT',
          eventData.date || '',
          eventData.time || '',
          eventData.location || '',
          eventData.status || 'Active',
          "'" + cleanLat.toFixed(6),
          "'" + cleanLng.toFixed(6),
          parseInt(eventData.radius) || 200
        ]]);
        return { success: true };
      }
    }
    return { success: false, error: "Agenda tidak ditemukan." };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function deleteEvent(eventId) {
  try {
    const ss = getSpreadsheet();
    
    let sheetEvents = ss.getSheetByName('Events');
    if (sheetEvents) {
      const dataEvents = sheetEvents.getDataRange().getValues();
      for (let i = dataEvents.length - 1; i >= 1; i--) {
        if (String(dataEvents[i][0]) === String(eventId)) {
          sheetEvents.deleteRow(i + 1);
          break;
        }
      }
    }
    
    let sheetAttendance = ss.getSheetByName('Attendance');
    if (sheetAttendance) {
      const dataAttendance = sheetAttendance.getDataRange().getValues();
      for (let i = dataAttendance.length - 1; i >= 1; i--) {
        if (String(dataAttendance[i][0]) === String(eventId)) {
          sheetAttendance.deleteRow(i + 1);
        }
      }
    }
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function submitAttendance(attendanceData) {
  const inputName = (attendanceData.name || '').toString().toLowerCase().trim();
  const targetEventId = (attendanceData.eventId || '').toString();
  
  if (!inputName || !targetEventId) {
    return { success: false, error: "Data presensi tidak lengkap." };
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = 'att_' + targetEventId + '_' + Utilities.base64EncodeWebSafe(inputName).substring(0, 40);
  
  if (cache.get(cacheKey) === '1') {
    return { success: false, error: "Nama Anda sudah terdaftar dalam presensi untuk kegiatan ini." };
  }

  cache.put(cacheKey, '1', 21600);

  const timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy, HH:mm:ss");
  const rowData = [
    targetEventId,
    attendanceData.name || '',
    attendanceData.instansi || '',
    attendanceData.status || '',
    timestamp
  ];

  const lock = LockService.getScriptLock();
  let savedSuccess = false;

  for (let attempt = 0; attempt < 10; attempt++) {
    try {
      const hasLock = lock.tryLock(3000);
      if (hasLock) {
        try {
          const ss = getSpreadsheet();
          let sheet = ss.getSheetByName('Attendance');
          if (!sheet) {
            sheet = ss.insertSheet('Attendance');
            sheet.appendRow(['eventId', 'name', 'instansi', 'status', 'timestamp']);
          }
          sheet.appendRow(rowData);
          savedSuccess = true;
        } finally {
          lock.releaseLock();
        }
        break;
      } else {
        Utilities.sleep(150 + Math.floor(Math.random() * 200));
      }
    } catch (err) {
      Utilities.sleep(200 + Math.floor(Math.random() * 300));
    }
  }

  if (savedSuccess) {
    return { success: true };
  } else {
    try {
      const ss = getSpreadsheet();
      let sheet = ss.getSheetByName('Attendance');
      if (!sheet) {
        sheet = ss.insertSheet('Attendance');
        sheet.appendRow(['eventId', 'name', 'instansi', 'status', 'timestamp']);
      }
      sheet.appendRow(rowData);
      return { success: true };
    } catch (fallbackErr) {
      try { cache.remove(cacheKey); } catch(cErr) {}
      return { success: false, error: "Gagal menyimpan presensi. Silakan coba lagi." };
    }
  }
}

function getAttendees(eventId) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('Attendance');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const attendees = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(eventId)) {
        attendees.push({
          eventId: String(data[i][0]),
          name: String(data[i][1] || ''),
          instansi: String(data[i][2] || ''),
          status: String(data[i][3] || 'Hadir'),
          timestamp: String(data[i][4] || '')
        });
      }
    }
    return attendees;
  } catch (e) {
    return [];
  }
}