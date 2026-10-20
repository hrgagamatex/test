/**
 * Tambahkan sebagai file Dashboard.gs pada proyek yang sudah memiliki FORM_ID dan FIELDS.
 * Jalankan perbaruiDashboardPenilaian() untuk membuat/memperbarui dashboard.
 * Data respons asli dan fungsi doPost tidak diubah.
 */
function perbaruiDashboardPenilaian() {
  var book = SpreadsheetApp.openById('1eMQcwebDDh0uXQRdj-JYm7uqRjrmnJTyirNBPM299hc');
  var form = FormApp.openById(FORM_ID);
  var scaleFields = FIELDS.filter(function(f) { return f[2] === 'scale'; });
  if (scaleFields.length !== 24) throw new Error('Dashboard membutuhkan 24 pertanyaan skala.');
  var ids = scaleFields.map(function(f) { return String(f[1]); });
  var titles = {};
  form.getItems().forEach(function(item) { titles[String(item.getId())] = item.getTitle(); });
  var buckets = {};
  function bucket(name) {
    if (!buckets[name]) buckets[name] = {count:0, employees:{}, sums:ids.map(function(){return 0;}), counts:ids.map(function(){return 0;}), dist:ids.map(function(){return [0,0,0,0,0];})};
    return buckets[name];
  }
  bucket('Semua divisi');
  form.getResponses().forEach(function(response) {
    var answers = {};
    response.getItemResponses().forEach(function(a) { answers[String(a.getItem().getId())] = a.getResponse(); });
    var department = String(answers['1770822543'] || '').trim() || 'Divisi belum diisi';
    if (/^(finance|accounting|finance\s*&\s*accounting)$/i.test(department)) department = 'Finance & Accounting';
    var nik = String(answers['227649005'] || '').trim();
    [bucket('Semua divisi'), bucket(department)].forEach(function(b) {
      b.count++;
      if (nik) b.employees[nik] = true;
      ids.forEach(function(id, i) {
        var raw = answers[id];
        if (raw === '' || raw === null || typeof raw === 'undefined') return;
        var n = Number(raw);
        if (!isFinite(n) || n < 1 || n > 5 || Math.floor(n) !== n) return;
        b.sums[i] += n; b.counts[i]++; b.dist[i][n-1]++;
      });
    });
  });
  var names = ['Semua divisi'].concat(Object.keys(buckets).filter(function(n){return n !== 'Semua divisi';}).sort());
  // Hanya dua tab khusus milik dashboard ini yang diperbarui.
  function ownedSheet(name) {
    var s = book.getSheetByName(name);
    if (s && s.getRange('A1').getNote() !== 'GAMATEX_DASHBOARD_V1') throw new Error('Tab '+name+' sudah ada dan bukan milik dashboard. Ubah nama tab tersebut terlebih dahulu.');
    if (!s) s = book.insertSheet(name);
    return s;
  }
  var data = ownedSheet('Rekap Diagram 3C');
  var dash = ownedSheet('Dashboard Penilaian');
  var chosen = dash.getRange('B2').getDisplayValue();
  if (names.indexOf(chosen) < 0) chosen = 'Semua divisi';
  [data,dash].forEach(function(s) { s.getCharts().forEach(function(c){s.removeChart(c);}); s.clear(); s.getRange('A1').setNote('GAMATEX_DASHBOARD_V1'); });
  var rows = [['Divisi','Jumlah respons','Karyawan unik (NIK)','Pertanyaan','Aspek','Rata-rata','Jawaban valid','Nilai 1','Nilai 2','Nilai 3','Nilai 4','Nilai 5']];
  var aspects = ['Competence','Character','Caring','Candor'];
  names.forEach(function(name) {
    var b = buckets[name];
    ids.forEach(function(id,i) {
      rows.push([name,b.count,Object.keys(b.employees).length,(i+1)+'. '+titles[id],aspects[Math.floor(i/6)],b.counts[i] ? b.sums[i]/b.counts[i] : '',b.counts[i]].concat(b.dist[i]));
    });
  });
  data.getRange(1,1,rows.length,12).setValues(rows);
  data.getRange(2,6,rows.length-1,1).setNumberFormat('0.00');
  data.setFrozenRows(1);
  data.getRange(1,1,1,12).setBackground('#062f5f').setFontColor('white').setFontWeight('bold');
  data.setColumnWidth(4,440);
  dash.getRange('A1').setValue('DASHBOARD PENILAIAN KARYAWAN').setFontWeight('bold').setFontSize(18);
  dash.getRange('A2').setValue('Filter divisi');
  dash.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(names,true).setAllowInvalid(false).build()).setValue(chosen).setBackground('#eaf4ff');
  dash.getRange('A3').setValue('Diperbarui');
  dash.getRange('B3').setValue(new Date()).setNumberFormat('dd/MM/yyyy HH:mm');
  dash.getRange('A4').setValue('Jumlah penilaian');
  dash.getRange('B4').setFormula('=IFERROR(INDEX(\'Rekap Diagram 3C\'!B:B,MATCH(B2,\'Rekap Diagram 3C\'!A:A,0)),0)');
  dash.getRange('A5').setValue('Karyawan sudah dinilai (NIK unik)');
  dash.getRange('B5').setFormula('=IFERROR(INDEX(\'Rekap Diagram 3C\'!C:C,MATCH(B2,\'Rekap Diagram 3C\'!A:A,0)),0)');
  dash.getRange('A7').setValue('Nilai kosong diabaikan. Rata-rata dihitung per jawaban valid; penilaian berulang tetap dihitung.');
  dash.getRange('A8').setValue('Ganti pilihan B2 untuk memfilter diagram. Jalankan perbaruiDashboardPenilaian untuk mengambil jawaban terbaru.');
  dash.getRange('A10:B10').setValues([['Aspek','Rata-rata (1–5)']]);
  aspects.forEach(function(a,i){
    dash.getRange(11+i,1).setValue(a);
    // Pembobotan berdasarkan jumlah jawaban valid, bukan rata-rata dari rata-rata pertanyaan.
    dash.getRange(11+i,2).setFormula('=IFERROR(SUMPRODUCT((\'Rekap Diagram 3C\'!A2:A'+rows.length+'=$B$2)*(\'Rekap Diagram 3C\'!E2:E'+rows.length+'=A'+(11+i)+'),\'Rekap Diagram 3C\'!F2:F'+rows.length+',\'Rekap Diagram 3C\'!G2:G'+rows.length+')/SUMIFS(\'Rekap Diagram 3C\'!G2:G'+rows.length+',\'Rekap Diagram 3C\'!A2:A'+rows.length+',$B$2,\'Rekap Diagram 3C\'!E2:E'+rows.length+',A'+(11+i)+'),"")');
  });
  dash.getRange('A17:H17').setValues([['Pertanyaan','Rata-rata','Jawaban valid','Nilai 1','Nilai 2','Nilai 3','Nilai 4','Nilai 5']]);
  dash.getRange('A18').setFormula('=FILTER(\'Rekap Diagram 3C\'!D2:D'+rows.length+',\'Rekap Diagram 3C\'!A2:A'+rows.length+'=$B$2)');
  dash.getRange('B18').setFormula('=FILTER(\'Rekap Diagram 3C\'!F2:L'+rows.length+',\'Rekap Diagram 3C\'!A2:A'+rows.length+'=$B$2)');
  dash.getRange('B11:B14').setNumberFormat('0.00');
  dash.getRange('B18:B41').setNumberFormat('0.00');
  [10,17].forEach(function(r){dash.getRange(r,1,1,r===10?2:8).setBackground('#062f5f').setFontColor('white').setFontWeight('bold');});
  dash.setColumnWidth(1,480); dash.setColumnWidth(2,190);
  dash.setColumnWidths(3,6,110);
  dash.getRange('A18:A41').setWrap(true);
  dash.setFrozenRows(2);
  sesuaikanRumusDashboard_(dash);
  dash.insertChart(dash.newChart().setChartType(Charts.ChartType.COLUMN).addRange(dash.getRange('A10:B14')).setPosition(10,10,0,0).setOption('title','Rata-rata per aspek — sesuai filter divisi').setOption('vAxis',{viewWindow:{min:0,max:5}}).setOption('legend',{position:'none'}).setOption('colors',['#0b73e0']).setOption('width',780).setOption('height',350).build());
  dash.insertChart(dash.newChart().setChartType(Charts.ChartType.BAR).addRange(dash.getRange('A17:A41')).addRange(dash.getRange('D17:H41')).setPosition(29,10,0,0).setOption('title','Distribusi jawaban 1–5 per pertanyaan').setOption('isStacked',true).setOption('width',1000).setOption('height',900).setOption('legend',{position:'top'}).setOption('colors',['#d92d20','#f79009','#eac54f','#32a879','#0b73e0']).build());
  SpreadsheetApp.flush();
  return 'Dashboard diperbarui: '+buckets['Semua divisi'].count+' penilaian.';
}

/** Jalankan ini untuk memperbaiki dashboard yang sudah dibuat. */
function perbaikiRumusDashboard() {
  var book = SpreadsheetApp.openById('1eMQcwebDDh0uXQRdj-JYm7uqRjrmnJTyirNBPM299hc');
  var dash = book.getSheetByName('Dashboard Penilaian');
  if (!dash || dash.getRange('A1').getNote() !== 'GAMATEX_DASHBOARD_V1') throw new Error('Jalankan perbaruiDashboardPenilaian terlebih dahulu.');
  sesuaikanRumusDashboard_(dash);
}
function sesuaikanRumusDashboard_(dash) {
  var probe = dash.getRange('Z1');
  var separator = ',';
  try {
    probe.setFormula('=IF(TRUE,1,0)');
    SpreadsheetApp.flush();
    if (probe.getValue() !== 1) {
      separator = ';';
      probe.setFormula('=IF(TRUE;1;0)');
      SpreadsheetApp.flush();
      if (probe.getValue() !== 1) throw new Error('Tidak dapat menentukan pemisah rumus Sheet.');
    }
  } finally { probe.clearContent(); }
  var addresses = ['B4','B5','B11','B12','B13','B14','A18','B18'];
  addresses.forEach(function(address) {
    var cell = dash.getRange(address), formula = cell.getFormula();
    if (formula) cell.setFormula(separator === ';' ? formula.replace(/,/g,';') : formula.replace(/;/g,','));
  });
  SpreadsheetApp.flush();
  var errors = addresses.filter(function(address) {
    return dash.getRange(address).getDisplayValue().charAt(0) === '#';
  });
  if (errors.length) throw new Error('Masih ada kesalahan rumus pada '+errors.join(', ')+'. Periksa pesan pada sel tersebut.');
}
