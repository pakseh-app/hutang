const SHEET_NAME = 'Hutang';
const HEADERS = [
  'id','nama','bulan','tanggal_jatuh_tempo','total_hutang',
  'jumlah_angsuran_total','jumlah_angsuran_dibayar','nominal_angsuran',
  'angsuran_terakhir','status','catatan','created_at','updated_at'
];

// 1) Buat Google Sheet kosong.
// 2) Extensions > Apps Script.
// 3) Hapus isi Code.gs, tempel kode ini.
// 4) Jalankan setup() sekali dan izinkan akses.
// 5) Atur password admin lewat setAdminPassword('PASSWORD_ANDA') lalu jalankan sekali.
// 6) Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone.
// 7) Salin URL /exec ke API_URL di index.html.

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  sh.clear();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,HEADERS.length).setFontWeight('bold');
}

function setAdminPassword(password) {
  if (!password || String(password).length < 8) throw new Error('Password minimal 8 karakter.');
  const hash = sha256_(String(password));
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD_HASH', hash);
}

function doGet(e) {
  try {
    const p = e.parameter || {};
    const action = p.action || '';
    let result;
    switch (action) {
      case 'user':
        result = {ok:true, data:getUser_(p.name)};
        break;
      case 'adminLogin':
        result = adminLogin_(p.passwordHash);
        break;
      case 'adminData':
        requireToken_(p.token);
        result = {ok:true, data:getAll_()};
        break;
      case 'create':
        requireToken_(p.token);
        result = {ok:true, data:save_(JSON.parse(p.data), false)};
        break;
      case 'update':
        requireToken_(p.token);
        result = {ok:true, data:save_(JSON.parse(p.data), true)};
        break;
      case 'delete':
        requireToken_(p.token);
        result = {ok:true, data:delete_(p.id)};
        break;
      default:
        result = {ok:false,error:'Action tidak dikenal.'};
    }
    return jsonp_(e, result);
  } catch (err) {
    return jsonp_(e, {ok:false,error:String(err.message || err)});
  }
}

function adminLogin_(passwordHash) {
  const saved = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD_HASH');
  if (!saved) throw new Error('Password admin belum diset. Jalankan setAdminPassword().');
  if (!passwordHash || passwordHash !== saved) throw new Error('Password admin salah.');
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('ADMIN_' + token, '1', 21600);
  return {ok:true,token:token};
}

function requireToken_(token) {
  if (!token || CacheService.getScriptCache().get('ADMIN_' + token) !== '1') {
    throw new Error('Sesi admin tidak valid atau sudah berakhir. Silakan login lagi.');
  }
}

function getSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Sheet Hutang belum ada. Jalankan setup().');
  return sh;
}

function getAll_() {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return [];
  return values.slice(1).filter(r => r[0]).map(rowToObject_);
}

function getUser_(name) {
  const wanted = normalize_(name);
  if (!wanted) return [];
  return getAll_().filter(x => normalize_(x.nama) === wanted);
}

function rowToObject_(r) {
  const x = {};
  HEADERS.forEach((h,i)=>x[h]=formatValue_(r[i],h));
  x.total_hutang = Number(x.total_hutang)||0;
  x.jumlah_angsuran_total = Number(x.jumlah_angsuran_total)||0;
  x.jumlah_angsuran_dibayar = Number(x.jumlah_angsuran_dibayar)||0;
  x.nominal_angsuran = Number(x.nominal_angsuran)||0;
  x.terbayar = Math.min(x.total_hutang, x.jumlah_angsuran_dibayar * x.nominal_angsuran);
  x.sisa = Math.max(0, x.total_hutang - x.terbayar);
  return x;
}

function save_(d, isUpdate) {
  const sh = getSheet_();
  const now = new Date();
  if (!d.nama || !d.bulan || !d.tanggal_jatuh_tempo || Number(d.total_hutang) < 0) {
    throw new Error('Data wajib belum lengkap.');
  }
  const obj = {
    id: d.id || Utilities.getUuid(),
    nama: String(d.nama).trim(),
    bulan: String(d.bulan),
    tanggal_jatuh_tempo: String(d.tanggal_jatuh_tempo),
    total_hutang: Number(d.total_hutang)||0,
    jumlah_angsuran_total: Number(d.jumlah_angsuran_total)||1,
    jumlah_angsuran_dibayar: Number(d.jumlah_angsuran_dibayar)||0,
    nominal_angsuran: Number(d.nominal_angsuran)||0,
    angsuran_terakhir: String(d.angsuran_terakhir||''),
    status: String(d.status||'Berjalan'),
    catatan: String(d.catatan||''),
    created_at: d.created_at || now.toISOString(),
    updated_at: now.toISOString()
  };
  if (isUpdate) {
    const row = findRowById_(obj.id);
    if (!row) throw new Error('Data tidak ditemukan.');
    sh.getRange(row,1,1,HEADERS.length).setValues([objectToRow_(obj)]);
  } else {
    sh.appendRow(objectToRow_(obj));
  }
  return obj;
}

function delete_(id) {
  const row = findRowById_(id);
  if (!row) throw new Error('Data tidak ditemukan.');
  getSheet_().deleteRow(row);
  return {id:id};
}

function findRowById_(id) {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return null;
  const ids = sh.getRange(2,1,last-1,1).getValues().flat();
  const i = ids.findIndex(x => String(x) === String(id));
  return i < 0 ? null : i + 2;
}

function objectToRow_(o) {
  return HEADERS.map(h => o[h] ?? '');
}

function formatValue_(v,h) {
  if (v instanceof Date) {
    if (h === 'bulan') return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
    if (h === 'tanggal_jatuh_tempo' || h === 'angsuran_terakhir') return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return v.toISOString();
  }
  return v;
}

function normalize_(s) {
  return String(s||'').trim().replace(/\s+/g,' ').toLowerCase();
}

function sha256_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function jsonp_(e, obj) {
  const cb = String(e.parameter.callback || 'callback').replace(/[^\w.$]/g,'');
  const body = cb + '(' + JSON.stringify(obj) + ');';
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JAVASCRIPT);
}
