const SPREADSHEET_ID = ''; // Kosongkan jika script dibuat dari Google Sheets (bound script).
const SHEET_NAME = 'DATA_DAPODIK';

const SCHOOLS = [
  'SMPN 19 HST','SMPN 8 HST','SMPN 33 HST','SMPN 5 HST','SMPN 15 HST',
  'SMPN 26 HST','SMPN 31 HST','SMPN 2 HST','SMPN 22 HST','SMPN 1 HST',
  'SMPN 12 HST','SMPN 16 HST','SMPN 3 HST','SMPN 17 HST','SMPN 27 HST',
  'SMPIT AL KHAIR','SMPN 18 HST','SMPN 9 HST','SMPN 29 HST','SMPN 13 HST',
  'SMPN 20 HST','SMPN 4 HST','SMPN 32 SATAP HST','SMP ISLAM MUHAJIRIN',
  'SMPN 21 HST','SMPN 6 HST','SMPN 28 HST','SMPN 23 HST','SMPN 11 HST',
  'SMPN 24 HST','SMPN 25 SATAP HST','SMPN 30 SATAP HST','SMPN 34 SATAP HST',
  'SMPN 14 HST'
];

function doGet() {
  setupSheet_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Statistik Indikator Kualitas Dapodik SMP HST')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSchools() {
  return SCHOOLS;
}

function setupSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  if (sh.getLastRow() === 0) {
    sh.appendRow([
      'Timestamp','Bulan Sinkronisasi','Nama Sekolah',
      'Kelengkapan Data (%)','Validitas Data (%)','Kebaruan Data (%)'
    ]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,6).setFontWeight('bold');
  }
  return sh;
}

function getSpreadsheet_() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim()) {
    return SpreadsheetApp.openById(SPREADSHEET_ID.trim());
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function saveData(payload) {
  if (!payload) throw new Error('Data tidak ditemukan.');
  const month = String(payload.month || '').trim();
  const school = String(payload.school || '').trim();
  const completeness = Number(payload.completeness);
  const validity = Number(payload.validity);
  const freshness = Number(payload.freshness);

  if (!/^\d{4}-\d{2}$/.test(month)) throw new Error('Bulan sinkronisasi harus dipilih.');
  if (!SCHOOLS.includes(school)) throw new Error('Nama sekolah tidak valid.');
  [completeness, validity, freshness].forEach(v => {
    if (!Number.isFinite(v) || v < 0 || v > 100) {
      throw new Error('Nilai indikator harus berada pada rentang 0 sampai 100.');
    }
  });

  const sh = setupSheet_();
  const values = sh.getDataRange().getValues();

  // Satu baris per sekolah per bulan. Jika sudah ada, diperbarui.
  for (let i = 1; i < values.length; i++) {
    const rowMonth = formatMonth_(values[i][1]);
    const rowSchool = String(values[i][2] || '').trim();
    if (rowMonth === month && rowSchool === school) {
      sh.getRange(i + 1, 1, 1, 6).setValues([[
        new Date(), month, school, completeness, validity, freshness
      ]]);
      return {ok:true, action:'updated', message:'Data sekolah dan bulan tersebut berhasil diperbarui.'};
    }
  }

  sh.appendRow([new Date(), month, school, completeness, validity, freshness]);
  return {ok:true, action:'inserted', message:'Data berhasil disimpan.'};
}

function getData(filters) {
  const sh = setupSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return [];

  const month = String((filters && filters.month) || '').trim();
  const school = String((filters && filters.school) || '').trim();

  return values.slice(1).filter(r => r[2]).map(r => ({
    timestamp: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
    month: formatMonth_(r[1]),
    school: String(r[2]),
    completeness: Number(r[3]) || 0,
    validity: Number(r[4]) || 0,
    freshness: Number(r[5]) || 0
  })).filter(r => (!month || r.month === month) && (!school || r.school === school));
}

function getSummary(month) {
  const data = getData({month: month || ''});
  const bySchool = {};
  SCHOOLS.forEach(s => bySchool[s] = null);

  data.forEach(r => {
    bySchool[r.school] = {
      school: r.school,
      month: r.month,
      completeness: r.completeness,
      validity: r.validity,
      freshness: r.freshness,
      average: Number(((r.completeness + r.validity + r.freshness) / 3).toFixed(2))
    };
  });

  const rows = SCHOOLS.map(s => bySchool[s]).filter(Boolean);
  const overall = rows.length ? {
    completeness: avg_(rows.map(x => x.completeness)),
    validity: avg_(rows.map(x => x.validity)),
    freshness: avg_(rows.map(x => x.freshness)),
    average: avg_(rows.map(x => x.average))
  } : null;

  return {rows, overall, months: [...new Set(data.map(x => x.month))].sort()};
}

function getMonthlyTrend(school) {
  const data = getData({school: school || ''});
  const byMonth = {};
  data.forEach(r => byMonth[r.month] = r);
  return Object.keys(byMonth).sort().map(m => {
    const r = byMonth[m];
    return {
      month: m,
      completeness: r.completeness,
      validity: r.validity,
      freshness: r.freshness,
      average: Number(((r.completeness + r.validity + r.freshness) / 3).toFixed(2))
    };
  });
}

function avg_(arr) {
  if (!arr.length) return 0;
  return Number((arr.reduce((a,b) => a + Number(b || 0), 0) / arr.length).toFixed(2));
}

function formatMonth_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  const s = String(value || '');
  return /^\d{4}-\d{2}/.test(s) ? s.substring(0,7) : s;
}
