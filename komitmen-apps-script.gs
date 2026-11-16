/**
 * Penerima komitmen EBIC'26 → Google Sheet
 *
 * Cara pasang:
 * 1. Buka Google Sheet baharu (cth. "Komitmen EBIC26").
 * 2. Extensions → Apps Script. Padam kod sedia ada, tampal fail ini, Save.
 * 3. Deploy → New deployment → jenis "Web app".
 *      Execute as:     Me
 *      Who has access: Anyone
 *    Klik Deploy, benarkan akses (Authorize), salin URL Web App (berakhir /exec).
 * 4. Tampal URL itu ke COMMITMENT_ENDPOINT dalam index.html, commit & push.
 *
 * Jika kod ini diubah kemudian: Deploy → Manage deployments → Edit → Version: New version
 * (URL kekal sama).
 */

const SHEET_NAME = 'Komitmen';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const sheet = getSheet_();
    sheet.appendRow([
      new Date(),                         // Masa diterima (pelayan)
      clean_(data.nama),
      clean_(data.tindakan),
      clean_(data.ukuran),
      clean_(data.masa),                  // Masa di browser peserta
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Untuk uji URL dalam browser: patut papar "OK".
function doGet() {
  return ContentService.createTextOutput('OK');
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Diterima', 'Nama / pasukan', 'Tindakan', 'Ukuran', 'Masa (browser)']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Hadkan panjang & elak formula injection dalam Sheet.
function clean_(v) {
  let s = String(v == null ? '' : v).slice(0, 500);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}
