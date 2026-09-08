const CONFIG = Object.freeze({
  spreadsheetId: '17E79JTYAY9qY5EYtiG-VNyfAFqhKIxtHDGIbv-5tdJQ',
  sheetName: 'Form Responses 1'
});

function doGet(e) {
  const action = String((e && e.parameter.action) || '');
  if (!action) {
    return HtmlService.createHtmlOutput('API Autofill Karyawan aktif.');
  }

  try {
    if (action === 'search') {
      return jsonp_({ ok: true, data: searchEmployees(e.parameter.q) }, e.parameter.callback);
    }
    if (action === 'get') {
      return jsonp_({ ok: true, data: getEmployee(e.parameter.name) }, e.parameter.callback);
    }
    return jsonp_({ ok: false, message: 'Aksi tidak dikenali.' }, e.parameter.callback);
  } catch (error) {
    return jsonp_({ ok: false, message: error.message || 'Terjadi kesalahan.' }, e.parameter.callback);
  }
}

function searchEmployees(query) {
  const keyword = normalize_(query);
  if (keyword.length < 2) return [];
  return employeeRows_()
    .filter(row => normalize_(row.nama).includes(keyword))
    .slice(0, 12)
    .map(row => ({ nama: row.nama }));
}

function getEmployee(name) {
  const target = normalize_(name);
  const employee = employeeRows_().find(row => normalize_(row.nama) === target);
  if (!employee) throw new Error('Data karyawan tidak ditemukan.');
  return employee;
}

function employeeRows_() {
  const sheet = SpreadsheetApp.openById(CONFIG.spreadsheetId).getSheetByName(CONFIG.sheetName);
  if (!sheet) throw new Error('Sheet sumber tidak ditemukan.');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const range = sheet.getRange(2, 2, lastRow - 1, 5);
  const displayRows = range.getDisplayValues();
  const rawRows = range.getValues();
  return displayRows
    .map((row, index) => ({ row, raw: rawRows[index] }))
    .filter(item => item.row[0])
    .map(item => ({
      nama: String(item.row[0]).trim(),
      nik: String(item.row[1]).trim(),
      bagian: String(item.row[2]).trim(),
      jabatan: String(item.row[3]).trim(),
      tanggalMasuk: item.raw[4] instanceof Date
        ? Utilities.formatDate(item.raw[4], Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : dateToIso_(item.row[4])
    }));
}

function dateToIso_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!match) return '';
  return match[3] + '-' + String(match[1]).padStart(2, '0') + '-' + String(match[2]).padStart(2, '0');
}

function jsonp_(value, callback) {
  const safeCallback = /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(String(callback || '')) ? callback : 'callback';
  return ContentService.createTextOutput(safeCallback + '(' + JSON.stringify(value) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function normalize_(value) {
  return String(value || '').trim().toLocaleLowerCase('id-ID');
}
