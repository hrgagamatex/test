const CONFIG = Object.freeze({
  spreadsheetId: '17E79JTYAY9qY5EYtiG-VNyfAFqhKIxtHDGIbv-5tdJQ',
  sheetName: 'Form Responses 1',
  formResponseUrl: 'https://docs.google.com/forms/d/e/1FAIpQLSeKmvRS4eHSkPywcVfxEVltvCJgO4IVnwpJapJ5LHnB8ks9yg/formResponse',
  fields: {
    nama: 'entry.2005620554',
    nik: 'entry.1166974658',
    bagian: 'entry.839337160',
    jabatan: 'entry.134342005'
  }
});

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Data Karyawan')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function searchEmployees(query) {
  const keyword = normalize_(query);
  if (keyword.length < 2) return [];

  return employeeRows_()
    .filter((row) => normalize_(row.nama).includes(keyword))
    .slice(0, 12)
    .map((row) => ({ nama: row.nama }));
}

function getEmployee(name) {
  const target = normalize_(name);
  const employee = employeeRows_().find((row) => normalize_(row.nama) === target);
  if (!employee) throw new Error('Data karyawan tidak ditemukan. Silakan cari ulang nama Anda.');
  return employee;
}

function submitEmployee(payload) {
  if (!payload || !payload.nama) throw new Error('Silakan pilih nama karyawan terlebih dahulu.');
  const employee = getEmployee(payload.nama);

  const formData = {};
  formData[CONFIG.fields.nama] = employee.nama;
  formData[CONFIG.fields.nik] = employee.nik;
  formData[CONFIG.fields.bagian] = employee.bagian;
  formData[CONFIG.fields.jabatan] = employee.jabatan;

  const response = UrlFetchApp.fetch(CONFIG.formResponseUrl, {
    method: 'post',
    payload: formData,
    followRedirects: false,
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 400) {
    throw new Error('Google Form menolak pengiriman. Periksa apakah formulir masih menerima jawaban.');
  }
  return { success: true, nama: employee.nama };
}

function employeeRows_() {
  const sheet = SpreadsheetApp.openById(CONFIG.spreadsheetId).getSheetByName(CONFIG.sheetName);
  if (!sheet) throw new Error('Sheet sumber data tidak ditemukan.');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 2, lastRow - 1, 4).getDisplayValues()
    .filter((row) => row[0])
    .map((row) => ({
      nama: String(row[0]).trim(),
      nik: String(row[1]).trim(),
      bagian: String(row[2]).trim(),
      jabatan: String(row[3]).trim()
    }));
}

function normalize_(value) {
  return String(value || '').trim().toLocaleLowerCase('id-ID');
}
