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
    if (/^(finance and tax|accounting|finance\s*&\s*accounting)$/i.test(department)) department = 'Finance & Accounting';
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
  if (dash.getMaxRows() < 800) dash.insertRowsAfter(dash.getMaxRows(),800-dash.getMaxRows());
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
  ids.forEach(function(id,i) {
    var start=45+i*30, sourceRow=18+i;
    dash.getRange(start,1).setValue((i+1)+'. '+titles[id]).setFontWeight('bold').setWrap(true);
    dash.getRange(start+1,1,1,4).setValues([['Nilai / jumlah / persentase','Jumlah jawaban','Persentase','Nilai']]);
    var formulas=[];
    for(var score=1;score<=5;score++){
      var r=start+1+score, countColumn=String.fromCharCode(67+score);
      formulas.push(['=D'+r+'&" — "&B'+r+'&" ("&ROUND(C'+r+'*1000)/10&"%)"','='+countColumn+sourceRow,'=IFERROR(B'+r+'/$C$'+sourceRow+',0)','='+score]);
    }
    dash.getRange(start+2,1,5,4).setFormulas(formulas);
    dash.getRange(start+2,3,5,1).setNumberFormat('0.0%');
  });
  sesuaikanRumusDashboard_(dash);
  ids.forEach(function(id,i){
    var start=45+i*30;
    siapkanSumberDiagram_(dash,start);
    dash.insertChart(dash.newChart().setChartType(Charts.ChartType.COLUMN)
      .addRange(dash.getRange(start+1,5,6,2)).setNumHeaders(1).setPosition(start+8,1,0,0)
      .setOption('title',(i+1)+'. '+titles[id]).setOption('legend',{position:'none'})
      .setOption('colors',['#34a853'])
      .setOption('hAxis',{title:'Nilai — jumlah jawaban (persentase)',slantedText:false})
      .setOption('vAxis',{title:'Persentase jawaban',format:'0%',viewWindow:{min:0,max:1}})
      .setOption('titleTextStyle',{fontSize:14})
      .setOption('chartArea',{left:65,top:90,width:'84%',height:'60%'})
      .setOption('width',980).setOption('height',400).build());
  });
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
  for(var i=0;i<24;i++) {
    for(var j=0;j<5;j++){
      var r=47+i*30+j;
      ['A','B','C','D'].forEach(function(col){addresses.push(col+r);});
    }
  }
  // Baca sekaligus agar tidak melakukan ratusan panggilan sel satu per satu.
  var formulas = dash.getRange('A1:D800').getFormulas();
  function adjusted(formula) {
    return separator === ';' ? formula.replace(/,/g,';') : formula.replace(/;/g,',');
  }
  ['B4','B5','B11','B12','B13','B14','A18','B18'].forEach(function(address) {
    var cell=dash.getRange(address), formula=cell.getFormula();
    if(formula) cell.setFormula(adjusted(formula));
  });
  for(var n=0;n<24;n++){
    var first=47+n*30;
    var block=formulas.slice(first-1,first+4);
    if(block.some(function(row){return row.some(Boolean);})) {
      dash.getRange(first,1,5,4).setFormulas(block.map(function(row){return row.map(adjusted);}));
    }
  }
  SpreadsheetApp.flush();
  var values=dash.getRange('A1:D800').getDisplayValues();
  var errors=addresses.filter(function(address){
    var match=address.match(/^([A-D])([0-9]+)$/);
    return values[Number(match[2])-1][match[1].charCodeAt(0)-65].charAt(0)==='#';
  });
  if(errors.length) throw new Error('Kesalahan rumus pada '+errors.join(', '));
}

/** Perbaiki diagram saja tanpa menarik ulang seluruh respons Google Form. */
function buatUlangDiagramPenilaian() {
  var book=SpreadsheetApp.openById('1eMQcwebDDh0uXQRdj-JYm7uqRjrmnJTyirNBPM299hc');
  var dash=book.getSheetByName('Dashboard Penilaian');
  if(!dash || dash.getRange('A1').getNote()!=='GAMATEX_DASHBOARD_V1') throw new Error('Dashboard belum tersedia.');
  var titles=dash.getRange('A1:A800').getDisplayValues();
  if(!titles[44][0] || !titles[734][0]) throw new Error('Tabel pertanyaan belum lengkap. Jalankan perbaruiDashboardPenilaian.');
  dash.getRange('A6').setValue('Status diagram');
  dash.getRange('B6').setValue('Membuat diagram…');
  dash.getCharts().forEach(function(chart){dash.removeChart(chart);});
  for(var i=0;i<24;i++){
    var row=45+i*30;
    siapkanSumberDiagram_(dash,row);
    dash.insertChart(dash.newChart().setChartType(Charts.ChartType.COLUMN)
      .addRange(dash.getRange(row+1,5,6,2))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setNumHeaders(1).setPosition(row+8,1,0,0)
      .setOption('title',titles[row-1][0]).setOption('legend',{position:'none'})
      .setOption('colors',['#34a853'])
      .setOption('vAxis',{title:'Persentase jawaban',format:'0%',viewWindow:{min:0,max:1}})
      .setOption('hAxis',{title:'Nilai — jumlah jawaban (persentase)'})
      .setOption('chartArea',{left:70,top:85,width:'83%',height:'60%'})
      .setOption('width',980).setOption('height',400).build());
  }
  SpreadsheetApp.flush();
  var count=dash.getCharts().length;
  dash.getRange('B6').setValue(count+' diagram berhasil dibuat');
  if(count!==24) throw new Error('Jumlah diagram tidak sesuai: '+count);
  console.log(count+' diagram berhasil dibuat; filter divisi tetap memakai B2.');
}

function siapkanSumberDiagram_(dash,start) {
  // Satu rentang bersebelahan: domain teks dan seri angka.
  dash.getRange(start+1,5,1,2).setValues([['Nilai / jumlah / persentase','Persentase jawaban']]);
  var refs=[];
  for(var i=0;i<5;i++){
    var row=start+2+i;
    refs.push(['=A'+row,'=C'+row]);
  }
  dash.getRange(start+2,5,5,2).setFormulas(refs);
  dash.getRange(start+2,6,5,1).setNumberFormat('0.0%');
  SpreadsheetApp.flush();
  var values=dash.getRange(start+2,6,5,1).getValues();
  if(values.some(function(row){return typeof row[0]!=='number' || !isFinite(row[0]);})) {
    throw new Error('Sumber persentase bukan angka pada baris '+(start+2)+'. Jalankan perbaruiDashboardPenilaian.');
  }
}
