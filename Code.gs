// =============================================
// Google Apps Script — Code.gs
// StudyTrack: Penjadwalan PR & Tugas Sekolah
// =============================================
//
// LANGKAH SETUP:
// 1. Buka https://script.google.com → New Project
// 2. Hapus semua kode default, paste seluruh file ini
// 3. Isi SPREADSHEET_ID di bawah (ambil dari URL spreadsheet)
// 4. Klik "Deploy" → "New Deployment"
//    - Type        : Web App
//    - Execute as  : Me
//    - Who has access: Anyone
// 5. Klik Deploy → Authorize → Copy URL
// 6. Paste URL ke variabel GOOGLE_APP_SCRIPT_URL di index.html
//
// STRUKTUR KOLOM SPREADSHEET (otomatis dibuat jika kosong):
// A: Mata Pelajaran | B: Nama Tugas | C: Tenggat Waktu | D: Catatan | E: Timestamp
// =============================================

// ── KONFIGURASI ──────────────────────────────
// Ambil ID dari URL: https://docs.google.com/spreadsheets/d/[ID_INI]/edit
const SPREADSHEET_ID = 'ISI_SPREADSHEET_ID_DISINI';
const SHEET_NAME     = 'Sheet1';
const NUM_COLS       = 5;

// =============================================
// doGet — Membaca semua data tugas
// Endpoint: GET <WebAppURL>?action=get
// =============================================
function doGet(e) {
  try {
    const sheet    = getSheet();
    const lastRow  = sheet.getLastRow();

    if (lastRow <= 1) {
      return respond({ status: 'success', data: [], total: 0 });
    }

    const values = sheet.getRange(2, 1, lastRow - 1, NUM_COLS).getValues();

    const data = values
      .filter(row => String(row[0]).trim() !== '')   // skip baris kosong
      .map((row, i) => ({
        id:            i + 2,                         // nomor baris di sheet (untuk delete)
        mataPelajaran: String(row[0]).trim(),
        namaTugas:     String(row[1]).trim(),
        tenggat:       toDateString(row[2]),
        catatan:       String(row[3]).trim(),
        timestamp:     row[4] ? new Date(row[4]).toISOString() : '',
      }))
      .sort((a, b) => a.tenggat.localeCompare(b.tenggat)); // urutkan by tenggat terdekat

    return respond({ status: 'success', data: data, total: data.length });

  } catch (err) {
    return respond({ status: 'error', message: err.message });
  }
}

// =============================================
// doPost — Menulis atau menghapus data tugas
//
// Tambah tugas:
//   body: { "action": "post", "data": { mataPelajaran, namaTugas, tenggat, catatan } }
//
// Hapus tugas (by row id):
//   body: { "action": "delete", "id": <nomor baris> }
// =============================================
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    if (payload.action === 'post')   return handlePost(payload.data);
    if (payload.action === 'delete') return handleDelete(payload.id);

    return respond({ status: 'error', message: 'Action tidak dikenali. Gunakan "post" atau "delete".' });

  } catch (err) {
    return respond({ status: 'error', message: 'Gagal parse request: ' + err.message });
  }
}

// ── Handler: Tambah tugas ─────────────────────
function handlePost(d) {
  if (!d) return respond({ status: 'error', message: 'Data tidak ditemukan.' });

  // Validasi field wajib
  const missing = ['mataPelajaran', 'namaTugas', 'tenggat'].filter(k => !d[k] || !String(d[k]).trim());
  if (missing.length > 0) {
    return respond({ status: 'error', message: 'Field wajib kosong: ' + missing.join(', ') });
  }

  // Validasi format tanggal
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d.tenggat)) {
    return respond({ status: 'error', message: 'Format tenggat harus YYYY-MM-DD.' });
  }

  const sheet = getSheet();
  ensureHeader(sheet);

  sheet.appendRow([
    String(d.mataPelajaran).trim(),
    String(d.namaTugas).trim(),
    d.tenggat,
    String(d.catatan || '').trim(),
    new Date(),
  ]);

  // Auto-resize kolom agar rapi
  sheet.autoResizeColumns(1, NUM_COLS);

  return respond({ status: 'success', message: 'Tugas berhasil disimpan.' });
}

// ── Handler: Hapus tugas by row number ────────
function handleDelete(rowId) {
  if (!rowId || isNaN(rowId)) {
    return respond({ status: 'error', message: 'ID baris tidak valid.' });
  }

  const id = parseInt(rowId);
  if (id <= 1) return respond({ status: 'error', message: 'Tidak bisa menghapus baris header.' });

  const sheet = getSheet();
  if (id > sheet.getLastRow()) {
    return respond({ status: 'error', message: 'Baris tidak ditemukan.' });
  }

  sheet.deleteRow(id);
  return respond({ status: 'success', message: `Baris ${id} berhasil dihapus.` });
}

// =============================================
// HELPER FUNCTIONS
// =============================================

/**
 * Ambil sheet, buat header otomatis jika sheet kosong.
 */
function getSheet() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" tidak ditemukan di spreadsheet.`);
  ensureHeader(sheet);
  return sheet;
}

/**
 * Buat baris header jika sheet masih kosong.
 */
function ensureHeader(sheet) {
  if (sheet.getLastRow() > 0) return;

  sheet.appendRow(['Mata Pelajaran', 'Nama Tugas', 'Tenggat Waktu', 'Catatan', 'Timestamp']);

  const header = sheet.getRange(1, 1, 1, NUM_COLS);
  header.setFontWeight('bold');
  header.setBackground('#9FB3DF');
  header.setFontColor('#FFFFFF');
  header.setHorizontalAlignment('center');
  sheet.setFrozenRows(1); // freeze header row
  sheet.autoResizeColumns(1, NUM_COLS);
}

/**
 * Konversi nilai Date dari Sheets ke string YYYY-MM-DD.
 * Jika sudah string, kembalikan apa adanya.
 */
function toDateString(val) {
  if (!val) return '';
  if (typeof val === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(val)) return val;
  const d = new Date(val);
  if (isNaN(d.getTime())) return String(val);
  const yyyy = d.getFullYear();
  const mm   = String(d.getMonth() + 1).padStart(2, '0');
  const dd   = String(d.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

/**
 * Buat JSON response. Google Apps Script tidak support
 * custom CORS headers, tapi mode "Anyone" + deployment
 * publik sudah cukup untuk akses dari browser.
 */
function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================
// FUNGSI TEST — Jalankan manual dari editor GAS
// untuk memastikan koneksi ke spreadsheet benar
// sebelum deploy.
// =============================================

/** Test: baca semua data */
function testGet() {
  const result = doGet({});
  Logger.log(result.getContent());
}

/** Test: tambah 1 baris dummy */
function testPost() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        action: 'post',
        data: {
          mataPelajaran: 'TEST - Matematika',
          namaTugas:     'TEST - Soal Latihan',
          tenggat:       '2026-05-15',
          catatan:       'Ini data test, boleh dihapus',
        }
      })
    }
  };
  const result = doPost(fakeEvent);
  Logger.log(result.getContent());
}

/** Test: hapus baris terakhir */
function testDelete() {
  const sheet  = getSheet();
  const lastId = sheet.getLastRow();
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({ action: 'delete', id: lastId })
    }
  };
  const result = doPost(fakeEvent);
  Logger.log(result.getContent());
}
