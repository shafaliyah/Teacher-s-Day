/**
 * Google Apps Script Backend - Guestbook Ucapan Hari Guru Nasional 2026
 */

const SPREADSHEET_ID = '10muhTrP2Wa_NsGhKa6wUNO5jHEuCPM8LDVCT493OBXQ';
const SHEET_NAME = 'Sheet1';
const MAX_NAME_LENGTH = 45;
const MAX_MESSAGE_LENGTH = 320;
const MAX_RESULTS = 200;

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Timestamp', 'Nama', 'Ucapan']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * GET: Mengambil semua ucapan (terbaru di atas)
 */
function doGet() {
  try {
    const sheet = getOrCreateSheet_();
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      return jsonResponse_({ ok: true, wishes: [] });
    }

    const numRows = lastRow - 1;
    const values = sheet.getRange(2, 1, numRows, 3).getValues();
    const wishes = [];

    for (let i = values.length - 1; i >= 0; i--) {
      const row = values[i];
      const name = String(row[1] || '').trim();
      const text = String(row[2] || '').trim();
      if (!name || !text) continue;

      wishes.push({
        n: name,
        t: text,
        ts: row[0] instanceof Date ? row[0].toISOString() : String(row[0] || '')
      });

      if (wishes.length >= MAX_RESULTS) break;
    }

    return jsonResponse_({ ok: true, wishes: wishes });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message || 'Server error' });
  }
}

/**
 * POST: Menyimpan ucapan baru ke Google Sheets
 * Menggunakan text/plain JSON agar tidak terkena CORS preflight di browser.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    let data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (_) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    const name = String(data.n || data.name || '').trim().slice(0, MAX_NAME_LENGTH);
    const text = String(data.t || data.message || '').trim().slice(0, MAX_MESSAGE_LENGTH);

    if (!name || !text) {
      return jsonResponse_({ ok: false, error: 'Nama dan ucapan wajib diisi.' });
    }

    const now = new Date();
    const sheet = getOrCreateSheet_();
    sheet.appendRow([now, name, text]);

    return jsonResponse_({
      ok: true,
      wish: {
        n: name,
        t: text,
        ts: now.toISOString()
      }
    });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message || 'Gagal menyimpan ucapan.' });
  } finally {
    lock.releaseLock();
  }
}
