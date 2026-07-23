// LPP Search API v9 (v8 + updateHive/updateSwarm can move pins: lat/lng/address write-back)
//
// ⚠️ BACKUP COPY — the LIVE copy of this script lives inside the Google Sheet:
//    open the Master Sheet → Extensions → Apps Script.
//    This file is kept in the repo so the script is never lost.
//    Editing THIS file does NOT change the live app.
//
// To install/update the live script:
//   1. Open the Master Sheet → Extensions → Apps Script
//   2. Select all the code there and replace it with this entire file
//   3. Save, then: Deploy → Manage deployments → pencil → New version → Deploy
//
// If you change the live script, paste the new version back into this file too.

const SHEET_ID = '1UED2EPSARAUhNua4mUGSgxkO2q18v3PkzkrVWmg5Ngk';

function doGet(e) {
  const action = e.parameter.action;
  try {
    if (action === 'getRegistry')    return getRegistry();
    if (action === 'getSwarms')      return getSwarms();
    if (action === 'getHives')       return getHives();
    if (action === 'getContactLogs') return getContactLogs();
    if (action === 'getSearchLogs')  return getSearchLogs();
    if (action === 'getAll')         return getAll();
    if (action === 'getUngeocodedRecords') return getUngeocodedRecords();
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
  return jsonResponse({ error: 'Unknown action: ' + action });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    if (action === 'logSearch')          return logSearch(data);
    if (action === 'logContact')         return logContact(data);
    if (action === 'addSwarm')           return addSwarm(data);
    if (action === 'updateSwarm')        return updateSwarm(data);
    if (action === 'deleteSwarm')        return deleteSwarm(data);
    if (action === 'addHive')            return addHive(data);
    if (action === 'updateHive')         return updateHive(data);
    if (action === 'deleteHive')         return deleteHive(data);
    if (action === 'updateRecordCoords') return updateRecordCoords(data);
    if (action === 'updateRecordLocation') return updateRecordLocation(data);
    if (action === 'verifyRecord')      return verifyRecord(data);
    if (action === 'logLocationFix')    return logLocationFix(data);
    if (action === 'censusGeocodeBatch') return censusGeocodeBatch(data);
    return jsonResponse({ error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ── REGISTRY ────────────────────────────────────────────
function getRegistry() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const records = [];
  ['Apiary', 'Hypersensitive'].forEach(tabName => {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    const type = tabName === 'Apiary' ? 'Apiary' : 'Hyper';
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const rec = rowToRecord(data[i], type);
      rec.tab = tabName;
      rec.row = i + 1;
      records.push(rec);
    }
  });
  return jsonResponse({ records: records, count: records.length });
}

function rowToRecord(row, type) {
  const rec = {
    name: row[0] || '',
    addressType: row[1] || '',
    address: row[2] || '',
    phoneDay: row[3] || '',
    phoneNight: row[4] || '',
    phoneAlt: row[5] || '',
    email: row[6] || '',
    township: row[7] || '',
    county: row[8] || '',
    flag: row[9] || '',           // Column J - GEOCODE FAILED marker
    verified: row[10] || '',      // Column K - verification status (VERIFIED, NEEDS_REVIEW, blank)
    confidence: row[11] || '',    // Column L - geocoder confidence (ROOFTOP, PARCEL, STREET, ZIP)
    type: type,
    lat: null,
    lng: null
  };
  const match = String(rec.address).match(/(-?\d+\.\d+),\s*(-?\d+\.\d+)/);
  if (match) {
    rec.lat = parseFloat(match[1]);
    rec.lng = parseFloat(match[2]);
  }
  return rec;
}

// Get ONLY records without GPS coords (for geocoder)
function getUngeocodedRecords() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const records = [];
  ['Apiary', 'Hypersensitive'].forEach(tabName => {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const address = String(data[i][2] || '');
      const hasCoords = address.match(/(-?\d+\.\d+),\s*(-?\d+\.\d+)/);
      if (!hasCoords) {
        records.push({
          tab: tabName,
          row: i + 1,
          name: data[i][0],
          address: address,
          township: data[i][7] || '',
          county: data[i][8] || ''
        });
      }
    }
  });
  return jsonResponse({ records: records, count: records.length });
}

// Update a single record's coords (append "lat, lng" to address)
function updateRecordCoords(data) {
  if (!data.tab || !data.row) return jsonResponse({ error: 'Missing tab/row' });
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(data.tab);
  if (!sheet) return jsonResponse({ error: 'Tab not found: ' + data.tab });

  const cell = sheet.getRange(data.row, 3); // Address column (C)
  let address = String(cell.getValue() || '');

  // Skip if already has coords
  if (address.match(/(-?\d+\.\d+),\s*(-?\d+\.\d+)/)) {
    return jsonResponse({ success: true, skipped: true });
  }

  if (data.lat && data.lng) {
    // Append coords to address
    address = address.trim().replace(/,\s*$/, '');
    const newAddress = `${address}, ${data.lat}, ${data.lng}`;
    cell.setValue(newAddress);
    return jsonResponse({ success: true });
  } else if (data.failed) {
    // Mark as failed for review (in column J, beyond our normal columns)
    sheet.getRange(data.row, 10).setValue('GEOCODE FAILED');
    return jsonResponse({ success: true, marked_failed: true });
  }

  return jsonResponse({ error: 'No coords provided' });
}

// Mark a record as verified (column K)
function verifyRecord(data) {
  if (!data.tab || !data.row) return jsonResponse({ error: 'Missing tab/row' });
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(data.tab);
  if (!sheet) return jsonResponse({ error: 'Tab not found' });
  sheet.getRange(data.row, 11).setValue(data.status || 'VERIFIED');
  return jsonResponse({ success: true });
}

// Replace coords for a record (used when correcting bad geocoding)
// Strips old coords from address text, appends new ones
function updateRecordLocation(data) {
  if (!data.tab || !data.row || data.lat == null || data.lng == null) {
    return jsonResponse({ error: 'Missing tab/row/lat/lng' });
  }
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(data.tab);
  if (!sheet) return jsonResponse({ error: 'Tab not found: ' + data.tab });

  const cell = sheet.getRange(data.row, 3); // Column C - Address
  let address = String(cell.getValue() || '');

  // Strip existing coords pattern (lat, lng) — handles both with and without leading comma/space
  address = address.replace(/,?\s*-?\d+\.\d+,\s*-?\d+\.\d+\s*$/, '').trim();
  address = address.replace(/,\s*$/, ''); // trailing comma

  const newAddress = `${address}, ${data.lat}, ${data.lng}`;
  cell.setValue(newAddress);

  // Also clear "GEOCODE FAILED" marker if present
  const flagCell = sheet.getRange(data.row, 10);
  if (String(flagCell.getValue() || '').toUpperCase().includes('FAILED')) {
    flagCell.setValue('FIXED');
  }

  // If marked verified (data.verified=true), mark column K as VERIFIED
  if (data.verified) {
    sheet.getRange(data.row, 11).setValue('VERIFIED');
  }

  // Confidence level from auto-geocoders (column L)
  if (data.confidence) {
    sheet.getRange(data.row, 12).setValue(data.confidence);
  }

  return jsonResponse({ success: true, newAddress: newAddress });
}

// Log a location correction so admins can audit changes
function logLocationFix(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Location Fixes');
  if (!sheet) {
    sheet = ss.insertSheet('Location Fixes');
    sheet.appendRow(['Timestamp','User Email','User Name','Tab','Row','Record Name','Old Lat','Old Lng','New Lat','New Lng','Method','Note']);
    sheet.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#3a8dde').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    new Date(), data.userEmail||'', data.userName||'', data.tab||'', data.row||'',
    data.recordName||'', data.oldLat||'', data.oldLng||'', data.newLat||'', data.newLng||'',
    data.method||'', data.note||''
  ]);
  return jsonResponse({ success: true });
}

// Bulk Census geocoder - processes up to ~25 records per call, writes back to sheet
function censusGeocodeBatch(data) {
  const items = data.items || [];
  if (items.length === 0) return jsonResponse({ error: 'No items' });
  if (items.length > 30) return jsonResponse({ error: 'Max 30 per batch' });

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const results = [];

  for (const item of items) {
    try {
      const clean = String(item.address || '').replace(/,?\s*-?\d+\.\d+,\s*-?\d+\.\d+\s*$/, '').trim();
      if (!clean) {
        results.push({ tab: item.tab, row: item.row, status: 'fail', reason: 'no address' });
        continue;
      }

      const url = 'https://geocoding.geo.census.gov/geocoder/locations/onelineaddress?address='
        + encodeURIComponent(clean) + '&benchmark=Public_AR_Current&format=json';

      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (response.getResponseCode() !== 200) {
        results.push({ tab: item.tab, row: item.row, status: 'fail', reason: 'http ' + response.getResponseCode() });
        continue;
      }

      const json = JSON.parse(response.getContentText());
      const matches = json?.result?.addressMatches || [];
      if (matches.length === 0) {
        results.push({ tab: item.tab, row: item.row, status: 'fail', reason: 'no match' });
        continue;
      }

      const m = matches[0];
      const lat = parseFloat(m.coordinates.y);
      const lng = parseFloat(m.coordinates.x);

      let confidence = 'LOW';
      const matched = String(m.matchedAddress || '').toUpperCase();
      const side = String(m.tigerLine?.side || '').toUpperCase();
      const hasNumber = /^\d+\s/.test(matched);
      if (hasNumber && (side === 'L' || side === 'R')) confidence = 'HIGH';
      else if (hasNumber) confidence = 'MEDIUM';

      const sheet = ss.getSheetByName(item.tab);
      if (sheet) {
        const cell = sheet.getRange(item.row, 3);
        let addr = String(cell.getValue() || '');
        addr = addr.replace(/,?\s*-?\d+\.\d+,\s*-?\d+\.\d+\s*$/, '').trim().replace(/,\s*$/, '');
        cell.setValue(`${addr}, ${lat}, ${lng}`);
        sheet.getRange(item.row, 12).setValue(confidence);
      }

      results.push({ tab: item.tab, row: item.row, status: 'ok', lat: lat, lng: lng, confidence: confidence, matched: m.matchedAddress });
    } catch(e) {
      results.push({ tab: item.tab, row: item.row, status: 'fail', reason: e.message });
    }
  }

  return jsonResponse({ success: true, results: results });
}

// ── SWARMS ──────────────────────────────────────────────
function getSwarms() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Swarms');
  if (!sheet) sheet = createSwarmSheet(ss);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonResponse({ swarms: [] });
  const swarms = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    swarms.push({
      id: i + 1,
      date: row[0] ? new Date(row[0]).toISOString() : '',
      address: row[1] || '',
      lat: row[2] ? parseFloat(row[2]) : null,
      lng: row[3] ? parseFloat(row[3]) : null,
      status: row[4] || 'Active',
      notes: row[5] || '',
      loggedBy: row[6] || '',
      loggedByEmail: row[7] || '',
      likelySource: row[8] || ''
    });
  }
  swarms.sort((a, b) => b.date.localeCompare(a.date));
  return jsonResponse({ swarms: swarms });
}

function createSwarmSheet(ss) {
  const sheet = ss.insertSheet('Swarms');
  sheet.appendRow(['Date', 'Address', 'Lat', 'Lng', 'Status', 'Notes', 'Logged By', 'Logged By Email', 'Likely Source Apiaries']);
  sheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#E8A020').setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
  return sheet;
}

function addSwarm(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Swarms');
  if (!sheet) sheet = createSwarmSheet(ss);
  sheet.appendRow([
    data.date ? new Date(data.date) : new Date(),
    data.address || '', data.lat || '', data.lng || '',
    data.status || 'Active', data.notes || '',
    data.loggedBy || '', data.loggedByEmail || '', data.likelySource || ''
  ]);
  return jsonResponse({ success: true });
}

function updateSwarm(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Swarms');
  if (!sheet || !data.id) return jsonResponse({ error: 'Swarm not found' });
  // Move / fix location: update Lat (col 3) & Lng (col 4), and Address (col 2) for GPS-based swarms
  if (data.lat != null && data.lng != null) {
    sheet.getRange(data.id, 3).setValue(data.lat);
    sheet.getRange(data.id, 4).setValue(data.lng);
  }
  if (data.address !== undefined && data.address !== '') sheet.getRange(data.id, 2).setValue(data.address);
  if (data.status) sheet.getRange(data.id, 5).setValue(data.status);
  if (data.notes !== undefined) sheet.getRange(data.id, 6).setValue(data.notes);
  return jsonResponse({ success: true });
}

function deleteSwarm(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Swarms');
  if (!sheet || !data.id) return jsonResponse({ error: 'Swarm not found' });
  sheet.deleteRow(data.id);
  return jsonResponse({ success: true });
}

// ── UNREGISTERED HIVES ──────────────────────────────────
function getHives() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Hives');
  if (!sheet) sheet = createHiveSheet(ss);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonResponse({ hives: [] });
  const hives = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[3] === '' || row[4] === '') continue; // need lat & lng
    hives.push({
      id: i + 1,
      date: row[0] ? new Date(row[0]).toISOString() : '',
      label: row[1] || '',
      address: row[2] || '',
      lat: row[3] ? parseFloat(row[3]) : null,
      lng: row[4] ? parseFloat(row[4]) : null,
      status: row[5] || 'Active',
      notes: row[6] || '',
      loggedBy: row[7] || '',
      loggedByEmail: row[8] || '',
      county: row[9] || ''
    });
  }
  hives.sort((a, b) => b.date.localeCompare(a.date));
  return jsonResponse({ hives: hives });
}

function createHiveSheet(ss) {
  const sheet = ss.insertSheet('Hives');
  sheet.appendRow(['Date', 'Label', 'Address', 'Lat', 'Lng', 'Status', 'Notes', 'Logged By', 'Logged By Email', 'County']);
  sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#9333EA').setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
  return sheet;
}

function addHive(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Hives');
  if (!sheet) sheet = createHiveSheet(ss);
  sheet.appendRow([
    data.date ? new Date(data.date) : new Date(),
    data.label || '', data.address || '',
    data.lat || '', data.lng || '',
    data.status || 'Active', data.notes || '',
    data.loggedBy || '', data.loggedByEmail || '', data.county || ''
  ]);
  return jsonResponse({ success: true });
}

function updateHive(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Hives');
  if (!sheet || !data.id) return jsonResponse({ error: 'Hive not found' });
  // Move / fix location: update Lat (col 4) & Lng (col 5), and Address (col 3) for GPS-based hives
  if (data.lat != null && data.lng != null) {
    sheet.getRange(data.id, 4).setValue(data.lat);
    sheet.getRange(data.id, 5).setValue(data.lng);
  }
  if (data.address !== undefined && data.address !== '') sheet.getRange(data.id, 3).setValue(data.address);
  if (data.status) sheet.getRange(data.id, 6).setValue(data.status);
  if (data.notes !== undefined) sheet.getRange(data.id, 7).setValue(data.notes);
  return jsonResponse({ success: true });
}

function deleteHive(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Hives');
  if (!sheet || !data.id) return jsonResponse({ error: 'Hive not found' });
  sheet.deleteRow(data.id);
  return jsonResponse({ success: true });
}

// ── ALL (registry + swarms + hives) ─────────────────────
function getAll() {
  const reg = JSON.parse(getRegistry().getContent());
  const sw  = JSON.parse(getSwarms().getContent());
  const hv  = JSON.parse(getHives().getContent());
  return jsonResponse({
    registry: reg.records || [],
    swarms: sw.swarms || [],
    hives: hv.hives || []
  });
}

// ── LOGS ────────────────────────────────────────────────
function logSearch(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Search Log');
  if (!sheet) {
    sheet = ss.insertSheet('Search Log');
    sheet.appendRow(['Timestamp', 'User Email', 'User Name', 'Searched Address', 'Lat', 'Lng', 'Hypers Found', 'Apiaries Found', 'Result']);
    sheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#2E7D32').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([new Date(), data.userEmail||'', data.userName||'', data.address||'',
    data.lat||'', data.lng||'', data.hyperCount||0, data.apiaryCount||0, data.result||'']);
  return jsonResponse({ success: true });
}

function logContact(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheetName = data.contactType === 'Hyper' ? 'Hyper Contact Log' : 'Apiary Contact Log';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['Timestamp', 'User Email', 'User Name', 'Contact Name', 'Contact Address', 'Phone Called', 'Note', 'Searched Address']);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#2E7D32').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([new Date(), data.userEmail||'', data.userName||'', data.contactName||'',
    data.contactAddress||'', data.phoneCalled||'', data.note||'', data.searchedAddress||'']);
  return jsonResponse({ success: true });
}

function getContactLogs() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const logs = [];
  ['Hyper Contact Log', 'Apiary Contact Log'].forEach(tabName => {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;
    const type = tabName.includes('Hyper') ? 'Hyper' : 'Apiary';
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      logs.push({
        timestamp: row[0] ? new Date(row[0]).toISOString() : '',
        userEmail: row[1]||'', userName: row[2]||'', contactName: row[3]||'',
        contactAddress: row[4]||'', phoneCalled: row[5]||'', note: row[6]||'',
        searchedAddress: row[7]||'', type: type
      });
    }
  });
  logs.sort((a, b) => b.timestamp.localeCompare(a.timestamp));
  return jsonResponse({ logs: logs });
}

function getSearchLogs() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Search Log');
  if (!sheet) return jsonResponse({ logs: [] });
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonResponse({ logs: [] });
  const logs = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    logs.push({
      timestamp: row[0] ? new Date(row[0]).toISOString() : '',
      userEmail: row[1]||'', userName: row[2]||'', address: row[3]||'',
      lat: row[4]||'', lng: row[5]||'', hyperCount: row[6]||0,
      apiaryCount: row[7]||0, result: row[8]||''
    });
  }
  logs.sort((a, b) => b.timestamp.localeCompare(a.timestamp));
  return jsonResponse({ logs: logs });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
