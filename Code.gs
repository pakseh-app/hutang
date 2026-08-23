const SHEET_NAME = 'Hutang';
const SETTINGS_SHEET_NAME = 'Pengaturan';
const USERS_SHEET_NAME = 'Pengguna';
const PAYMENTS_SHEET_NAME = 'Pembayaran';
const USER_HEADERS = ['nama','pin_hash','created_at','updated_at'];
const PAYMENT_HEADERS = ['id','hutang_id','nama','angsuran_ke','nominal','tanggal_pembayaran','created_at'];
const REQUESTS_SHEET_NAME = 'Permintaan Pembayaran';
const REQUEST_HEADERS = ['id','hutang_id','nama','angsuran_ke','nominal','tanggal_pengajuan','tanggal_pembayaran','bukti_file_id','bukti_url','bukti_nama','status','alasan_admin','created_at','updated_at'];
const DEBT_ITEM_HEADER = 'barang';
const HEADERS = [
  'id','nama','barang','tanggal_mulai_angsuran','tanggal_jatuh_tempo_day','tanggal_jatuh_tempo_pertama',
  'jumlah_angsuran','angsuran_per_bulan','angsuran_dibayar',
  'tanggal_angsuran_terakhir','status','catatan','created_at','updated_at'
];

// SETUP:
// 1) Buat Google Sheet kosong.
// 2) Extensions > Apps Script.
// 3) Hapus isi Code.gs, tempel kode ini.
// 4) Jalankan setup() sekali dan izinkan akses.
// 5) Deploy > New deployment > Web app > Execute as: Me > Anyone.
// 6) Salin URL /exec ke API_URL di index.html.
// 7) Password admin dapat diganti langsung dari aplikasi.

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
  } else {
    ensureDebtSchema_(sh);
  }
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,HEADERS.length).setFontWeight('bold');

  let users = ss.getSheetByName(USERS_SHEET_NAME);
  if (!users) users = ss.insertSheet(USERS_SHEET_NAME);
  users.getRange(1,1,1,USER_HEADERS.length).setValues([USER_HEADERS]);
  users.setFrozenRows(1);
  users.getRange(1,1,1,USER_HEADERS.length).setFontWeight('bold');

  let payments = ss.getSheetByName(PAYMENTS_SHEET_NAME);
  if (!payments) payments = ss.insertSheet(PAYMENTS_SHEET_NAME);
  payments.getRange(1,1,1,PAYMENT_HEADERS.length).setValues([PAYMENT_HEADERS]);
  payments.setFrozenRows(1);
  payments.getRange(1,1,1,PAYMENT_HEADERS.length).setFontWeight('bold');

  let requests = ss.getSheetByName(REQUESTS_SHEET_NAME);
  if (!requests) requests = ss.insertSheet(REQUESTS_SHEET_NAME);
  requests.getRange(1,1,1,REQUEST_HEADERS.length).setValues([REQUEST_HEADERS]);
  requests.setFrozenRows(1);
  requests.getRange(1,1,1,REQUEST_HEADERS.length).setFontWeight('bold');

  let settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settings) settings = ss.insertSheet(SETTINGS_SHEET_NAME);
  if (settings.getLastRow() === 0) {
    settings.getRange('A1:B1').setValues([['Kunci','Nilai']]);
    settings.getRange('A2:B2').setValues([['ADMIN_PASSWORD','1234']]);
  } else if (!settings.getRange('A1').getValue()) {
    settings.getRange('A1:B1').setValues([['Kunci','Nilai']]);
    settings.getRange('A2:B2').setValues([['ADMIN_PASSWORD','1234']]);
  }
  if (!settings.getRange('A3').getValue()) {
    const folder = DriveApp.createFolder('Bukti Transfer Catatan Hutang');
    settings.getRange('A3:B3').setValues([['BUKTI_FOLDER_ID', folder.getId()]]);
  }
  settings.getRange('A1:B1').setFontWeight('bold');
  settings.setFrozenRows(1);
  settings.autoResizeColumns(1,2);
}

function ensureDebtSchema_(sh) {
  const lastCol = sh.getLastColumn();
  const old = lastCol ? sh.getRange(1,1,1,lastCol).getValues()[0].map(String) : [];
  const isNew = HEADERS.every((h,i)=>old[i]===h) && old.length===HEADERS.length;
  if (isNew) return;
  const oldData = sh.getLastRow() > 1 ? sh.getRange(2,1,sh.getLastRow()-1,Math.max(lastCol,1)).getValues() : [];
  const oldIndex = {}; old.forEach((h,i)=>oldIndex[h]=i);
  const getOld = (r,h) => oldIndex[h] === undefined ? '' : r[oldIndex[h]];
  const rows = oldData.filter(r=>r.some(v=>v!=='' && v!==null)).map(r=>{
    const installments = Number(getOld(r,'jumlah_angsuran_total')) || Number(getOld(r,'jumlah_angsuran')) || 1;
    const amount = Number(getOld(r,'nominal_angsuran')) || Number(getOld(r,'angsuran_per_bulan')) || 0;
    const due = getOld(r,'tanggal_jatuh_tempo') || getOld(r,'tanggal_jatuh_tempo_pertama');
    const month = getOld(r,'bulan');
    const start = getOld(r,'tanggal_mulai_angsuran') || (month ? month + '-01' : due);
    const dueDay = Number(getOld(r,'tanggal_jatuh_tempo_day')) || dayFromDate_(due) || 1;
    const firstDue = getOld(r,'tanggal_jatuh_tempo_pertama') || due || firstDueFromStart_(start,dueDay);
    return [
      getOld(r,'id') || Utilities.getUuid(), getOld(r,'nama'), getOld(r,'barang') || getOld(r,'nama_barang') || '', start, dueDay, firstDue,
      installments, amount, Number(getOld(r,'jumlah_angsuran_dibayar')) || Number(getOld(r,'angsuran_dibayar')) || 0,
      getOld(r,'angsuran_terakhir') || getOld(r,'tanggal_angsuran_terakhir') || '', getOld(r,'status') || 'Berjalan', getOld(r,'catatan'),
      getOld(r,'created_at') || new Date(), getOld(r,'updated_at') || new Date()
    ];
  });
  sh.clear();
  sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
  if (rows.length) sh.getRange(2,1,rows.length,HEADERS.length).setValues(rows);
}

function doGet(e) {
  try {
    const p = e.parameter || {};
    const action = p.action || '';
    let result;
    switch (action) {
      case 'userLogin': result = userLogin_(p.name, p.pinHash); break;
      case 'userData': result = {ok:true, data:getUserByToken_(p.token)}; break;
      case 'users': requireToken_(p.token); result = {ok:true, data:getUsers_()}; break;
      case 'saveUser': requireToken_(p.token); result = {ok:true, data:saveUser_(JSON.parse(p.data))}; break;
      case 'deleteUser': requireToken_(p.token); result = {ok:true, data:deleteUser_(p.name)}; break;
      case 'user': result = {ok:true, data:getUser_(p.name)}; break;
      case 'adminLogin': result = adminLogin_(p.passwordHash); break;
      case 'adminData': requireToken_(p.token); result = {ok:true, data:getAll_(), paymentRequests:getPaymentRequests_()}; break;
      case 'changePassword': requireToken_(p.token); result = changeAdminPassword_(p.newPassword); break;
      case 'create': requireToken_(p.token); result = {ok:true, data:save_(JSON.parse(p.data), false)}; break;
      case 'update': requireToken_(p.token); result = {ok:true, data:save_(JSON.parse(p.data), true)}; break;
      case 'payInstallment': requireToken_(p.token); result = {ok:true, data:payInstallment_(p.id,p.tanggalPembayaran)}; break;
      case 'approvePaymentRequest': requireToken_(p.token); result = {ok:true, data:approvePaymentRequest_(p.id)}; break;
      case 'rejectPaymentRequest': requireToken_(p.token); result = {ok:true, data:reviewPaymentRequest_(p.id,'Tidak',p.alasan||'')}; break;
      case 'requestStatus': result = {ok:true, data:getRequestByIdForUser_(p.token,p.id)}; break;
      case 'delete': requireToken_(p.token); result = {ok:true, data:delete_(p.id)}; break;
      default: result = {ok:false,error:'Action tidak dikenal.'};
    }
    return jsonp_(e, result);
  } catch (err) {
    return jsonp_(e, {ok:false,error:String(err.message || err)});
  }
}

function doPost(e) {
  try {
    const p = e.parameter || {};
    if (p.action === 'submitPaymentRequest') {
      const result = submitPaymentRequest_(p);
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:'Action tidak dikenal.'})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err.message || err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

function adminLogin_(passwordHash) {
  const password = getAdminPassword_();
  if (!password) throw new Error('Password admin belum diset.');
  if (!passwordHash || passwordHash !== sha256_(password)) throw new Error('Password admin salah.');
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('ADMIN_' + token, '1', 21600);
  return {ok:true,token:token};
}

function getSettingsSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
  if (!sh) throw new Error('Sheet Pengaturan belum ada. Jalankan setup().');
  return sh;
}
function getAdminPassword_() { return String(getSettingsSheet_().getRange('B2').getDisplayValue() || '').trim(); }
function changeAdminPassword_(newPassword) {
  const password = String(newPassword || '').trim();
  if (!/^\d{4}$/.test(password)) throw new Error('Password admin harus tepat 4 angka.');
  getSettingsSheet_().getRange('B2').setValue(password);
  return {ok:true,message:'Password admin berhasil diubah.'};
}
function requireToken_(token) {
  if (!token || CacheService.getScriptCache().get('ADMIN_' + token) !== '1') throw new Error('Sesi admin tidak valid atau sudah berakhir. Silakan login lagi.');
}

function getSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Sheet Hutang belum ada. Jalankan setup().');
  ensureDebtSchema_(sh);
  return sh;
}
function getAll_() {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return [];
  return values.slice(1).filter(r=>r[0]).map(rowToObject_);
}

function userLogin_(name,pinHash) {
  const wanted = normalize_(name);
  if (!wanted || !pinHash) throw new Error('Nama dan PIN wajib diisi.');
  const user = findUser_(wanted);
  if (!user || String(user.pin_hash)!==String(pinHash)) throw new Error('Nama atau PIN salah.');
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('USER_'+token,wanted,21600);
  return {ok:true,token:token,name:user.nama,data:getUser_(user.nama)};
}
function getUserByToken_(token) {
  const name = token ? CacheService.getScriptCache().get('USER_'+token) : '';
  if (!name) throw new Error('Sesi pengguna sudah berakhir. Silakan login lagi.');
  return {debts:getUser_(name),requests:getPaymentRequestsForUser_(name)};
}
function getUser_(name) { const wanted=normalize_(name); return wanted?getAll_().filter(x=>normalize_(x.nama)===wanted):[]; }

function getUsers_() {
  const sh=getUsersSheet_(), last=sh.getLastRow(); if(last<2)return [];
  return sh.getRange(2,1,last-1,USER_HEADERS.length).getValues().filter(r=>r[0]).map(r=>({nama:String(r[0]),created_at:dateOut_(r[2]),updated_at:dateOut_(r[3])}));
}
function getUsersSheet_() {
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET_NAME);
  if(!sh) throw new Error('Sheet Pengguna belum ada. Jalankan setup().'); return sh;
}
function findUser_(wanted) {
  const sh=getUsersSheet_(), last=sh.getLastRow(); if(last<2)return null;
  const rows=sh.getRange(2,1,last-1,USER_HEADERS.length).getValues();
  for(const r of rows) if(normalize_(r[0])===wanted) return {nama:String(r[0]),pin_hash:String(r[1]||'')};
  return null;
}
function saveUser_(d) {
  const name=String(d.nama||'').trim(), pinHash=String(d.pinHash||'').trim();
  if(!name)throw new Error('Nama pengguna wajib diisi.');
  if(!/^[a-f0-9]{64}$/i.test(pinHash))throw new Error('PIN tidak valid.');
  const sh=getUsersSheet_(),last=sh.getLastRow(),wanted=normalize_(name);
  if(last>=2){const rows=sh.getRange(2,1,last-1,USER_HEADERS.length).getValues();for(let i=0;i<rows.length;i++)if(normalize_(rows[i][0])===wanted){sh.getRange(i+2,2).setValue(pinHash);sh.getRange(i+2,4).setValue(new Date());return {nama:String(rows[i][0]),message:'PIN pengguna berhasil disimpan.'};}}
  const now=new Date();sh.appendRow([name,pinHash,now,now]);return {nama:name,message:'Pengguna berhasil ditambahkan.'};
}
function deleteUser_(name) {
  const sh=getUsersSheet_(),last=sh.getLastRow(),wanted=normalize_(name); if(last<2)throw new Error('Pengguna tidak ditemukan.');
  const rows=sh.getRange(2,1,last-1,USER_HEADERS.length).getValues(),i=rows.findIndex(r=>normalize_(r[0])===wanted); if(i<0)throw new Error('Pengguna tidak ditemukan.');
  sh.deleteRow(i+2);return {nama:name};
}

function rowToObject_(r) {
  const x={};HEADERS.forEach((h,i)=>x[h]=formatValue_(r[i],h));
  x.jumlah_angsuran=Number(x.jumlah_angsuran)||0;
  x.angsuran_per_bulan=Number(x.angsuran_per_bulan)||0;
  x.tanggal_jatuh_tempo_day=Number(x.tanggal_jatuh_tempo_day)||dayFromDate_(x.tanggal_jatuh_tempo_pertama)||1;
  x.tanggal_jatuh_tempo_pertama=x.tanggal_jatuh_tempo_pertama||firstDueFromStart_(x.tanggal_mulai_angsuran,x.tanggal_jatuh_tempo_day);
  x.angsuran_dibayar=Math.max(0,Math.min(x.jumlah_angsuran,Number(x.angsuran_dibayar)||0));
  x.total_hutang=x.jumlah_angsuran*x.angsuran_per_bulan;
  x.terbayar=Math.min(x.total_hutang,x.angsuran_dibayar*x.angsuran_per_bulan);
  x.sisa=Math.max(0,x.total_hutang-x.terbayar);
  x.angsuran_ke_berikutnya=x.angsuran_dibayar+1;
  x.jatuh_tempo_berikutnya=x.sisa>0?addMonthsSafe_(x.tanggal_jatuh_tempo_pertama,x.angsuran_dibayar,x.tanggal_jatuh_tempo_day):'';
  if(x.angsuran_dibayar>=x.jumlah_angsuran && x.jumlah_angsuran>0)x.status='Lunas';
  return x;
}

function save_(d,isUpdate) {
  const sh=getSheet_(),now=new Date();
  const nama=String(d.nama||'').trim();
  const barang=String(d.barang||'').trim();
  const start=String(d.tanggal_mulai_angsuran||'').trim();
  const dueDay=Math.floor(Number(d.tanggal_jatuh_tempo_day)||0);
  const firstDue=firstDueFromStart_(start,dueDay);
  const jumlah=Math.floor(Number(d.jumlah_angsuran)||0);
  const amount=Number(d.angsuran_per_bulan)||0;
  if(!nama||!barang||!start||dueDay<1||dueDay>31||!firstDue||jumlah<1||amount<=0) throw new Error('Nama, nama barang, tanggal mulai, tanggal jatuh tempo tiap bulan, jumlah angsuran, dan angsuran per bulan wajib diisi.');

  let existing=null;
  if(isUpdate){
    const row=findRowById_(String(d.id||''));
    if(!row) throw new Error('Data tidak ditemukan.');
    existing=rowToObject_(sh.getRange(row,1,1,HEADERS.length).getValues()[0]);
  }

  // Jumlah pembayaran TIDAK lagi diambil dari form.
  // Saat membuat data baru selalu 0. Saat mengedit, jumlah pembayaran lama dipertahankan.
  const paid=Math.max(0,Math.min(jumlah, existing ? Number(existing.angsuran_dibayar)||0 : 0));
  const lastPaid=existing ? String(existing.tanggal_angsuran_terakhir||'') : '';
  const obj={
    id:d.id||Utilities.getUuid(),nama,barang,tanggal_mulai_angsuran:start,tanggal_jatuh_tempo_day:dueDay,tanggal_jatuh_tempo_pertama:firstDue,
    jumlah_angsuran:jumlah,angsuran_per_bulan:amount,angsuran_dibayar:paid,
    tanggal_angsuran_terakhir:lastPaid,
    status:paid>=jumlah?'Lunas':String(d.status||'Berjalan'),catatan:String(d.catatan||''),
    created_at:(existing&&existing.created_at)||d.created_at||now.toISOString(),updated_at:now.toISOString()
  };
  if(isUpdate){
    const row=findRowById_(obj.id);
    sh.getRange(row,1,1,HEADERS.length).setValues([objectToRow_(obj)]);
  } else {
    sh.appendRow(objectToRow_(obj));
  }
  return rowToObject_(objectToRow_(obj));
}
function payInstallment_(id,paymentDate){
  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  try { return applyPayment_(id,paymentDate); } finally { lock.releaseLock(); }
}
function applyPayment_(id,paymentDate){
  const row=findRowById_(id); if(!row) throw new Error('Data hutang tidak ditemukan.');
  const sh=getSheet_(); const obj=rowToObject_(sh.getRange(row,1,1,HEADERS.length).getValues()[0]);
  const paid=Number(obj.angsuran_dibayar)||0, total=Number(obj.jumlah_angsuran)||0;
  if(paid>=total) throw new Error('Semua angsuran sudah lunas.');
  const date=String(paymentDate||'').trim();
  if(!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Tanggal pembayaran tidak valid.');
  const parts=date.split('-').map(Number), dt=new Date(parts[0],parts[1]-1,parts[2]);
  if(dt.getFullYear()!==parts[0]||dt.getMonth()!==parts[1]-1||dt.getDate()!==parts[2]) throw new Error('Tanggal pembayaran tidak valid.');
  const next=paid+1, paymentSh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PAYMENTS_SHEET_NAME);
  if(!paymentSh) throw new Error('Sheet Pembayaran belum ada. Jalankan setup().');
  paymentSh.appendRow([Utilities.getUuid(),obj.id,obj.nama,next,Number(obj.angsuran_per_bulan)||0,date,new Date()]);
  obj.angsuran_dibayar=next; obj.tanggal_angsuran_terakhir=date; obj.status=next>=total?'Lunas':'Berjalan'; obj.updated_at=new Date().toISOString();
  sh.getRange(row,1,1,HEADERS.length).setValues([objectToRow_(obj)]); return rowToObject_(objectToRow_(obj));
}

function getRequestsSheet_() {
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQUESTS_SHEET_NAME);
  if(!sh) throw new Error('Sheet Permintaan Pembayaran belum ada. Jalankan setup().');
  return sh;
}
function requestRowToObject_(r){
  const x={}; REQUEST_HEADERS.forEach((h,i)=>x[h]=formatValue_(r[i],h));
  x.nominal=Number(x.nominal)||0; x.angsuran_ke=Number(x.angsuran_ke)||0;
  return x;
}
function getPaymentRequests_(){
  const sh=getRequestsSheet_(), last=sh.getLastRow(); if(last<2)return [];
  return sh.getRange(2,1,last-1,REQUEST_HEADERS.length).getValues().filter(r=>r[0]).map(requestRowToObject_).reverse();
}
function getPaymentRequestsForUser_(name){
  const wanted=normalize_(name); return getPaymentRequests_().filter(x=>normalize_(x.nama)===wanted);
}
function getRequestByIdForUser_(token,id){
  const name=token ? CacheService.getScriptCache().get('USER_'+token) : ''; if(!name)throw new Error('Sesi pengguna sudah berakhir.');
  const x=getPaymentRequestsForUser_(name).find(r=>String(r.id)===String(id)); if(!x)throw new Error('Permintaan tidak ditemukan.'); return x;
}
function getDriveFolder_(){
  const settings=getSettingsSheet_(); let id=String(settings.getRange('B3').getDisplayValue()||'').trim();
  if(!id){const folder=DriveApp.createFolder('Bukti Transfer Catatan Hutang');id=folder.getId();settings.getRange('A3:B3').setValues([['BUKTI_FOLDER_ID',id]]);}
  return DriveApp.getFolderById(id);
}
function submitPaymentRequest_(p){
  const token=String(p.token||'').trim(); const name=token?CacheService.getScriptCache().get('USER_'+token):'';
  if(!name)throw new Error('Sesi pengguna tidak valid atau sudah berakhir.');
  const id=String(p.requestId||'').trim()||Utilities.getUuid(); const debtId=String(p.hutangId||'').trim();
  const row=findRowById_(debtId); if(!row)throw new Error('Data hutang tidak ditemukan.');
  const sh=getSheet_(); const obj=rowToObject_(sh.getRange(row,1,1,HEADERS.length).getValues()[0]);
  if(normalize_(obj.nama)!==normalize_(name))throw new Error('Data hutang bukan milik pengguna ini.');
  const paid=Number(obj.angsuran_dibayar)||0,total=Number(obj.jumlah_angsuran)||0; if(paid>=total)throw new Error('Semua angsuran sudah lunas.');
  const existing=getPaymentRequests_().find(r=>String(r.id)===id); if(existing) return {ok:true,id:id,status:existing.status};
  const pending=getPaymentRequestsForUser_(name).find(r=>String(r.hutang_id)===debtId && r.status==='Menunggu Persetujuan' && Number(r.angsuran_ke)===paid+1);
  if(pending)throw new Error('Permintaan pembayaran untuk angsuran ini sudah menunggu persetujuan admin.');
  const date=String(p.tanggalPembayaran||'').trim(); if(!/^\d{4}-\d{2}-\d{2}$/.test(date))throw new Error('Tanggal pembayaran tidak valid.');
  const mime=String(p.buktiMime||''); const data=String(p.buktiData||''); const fileName=String(p.buktiName||'bukti-transfer.jpg');
  if(!/^image\/(jpeg|jpg|png|webp|heic|heif)$/i.test(mime))throw new Error('Bukti transfer harus berupa gambar.');
  if(!data)throw new Error('Bukti transfer wajib diupload.');
  if(data.length>2600000)throw new Error('Ukuran bukti terlalu besar. Gunakan gambar yang lebih kecil.');
  const bytes=Utilities.base64Decode(data); if(bytes.length>1900000)throw new Error('Ukuran gambar maksimal sekitar 1,9 MB.');
  const blob=Utilities.newBlob(bytes,mime,fileName); const file=getDriveFolder_().createFile(blob); file.setName('Bukti - '+obj.nama+' - Angsuran '+(paid+1));
  const now=new Date(); const req={id, hutang_id:obj.id,nama:obj.nama,angsuran_ke:paid+1,nominal:Number(obj.angsuran_per_bulan)||0,tanggal_pengajuan:now,tanggal_pembayaran:date,bukti_file_id:file.getId(),bukti_url:file.getUrl(),bukti_nama:file.getName(),status:'Menunggu Persetujuan',alasan_admin:'',created_at:now,updated_at:now};
  getRequestsSheet_().appendRow(REQUEST_HEADERS.map(h=>req[h]??''));
  return {ok:true,id,status:req.status};
}
function approvePaymentRequest_(id){
  const lock=LockService.getScriptLock();
  lock.waitLock(15000);
  try{
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const reqSh=getRequestsSheet_();
    const last=reqSh.getLastRow();
    if(last<2) throw new Error('Permintaan pembayaran tidak ditemukan.');
    const rows=reqSh.getRange(2,1,last-1,REQUEST_HEADERS.length).getValues();
    const idx=rows.findIndex(r=>String(r[0])===String(id));
    if(idx<0) throw new Error('Permintaan pembayaran tidak ditemukan.');
    const req=requestRowToObject_(rows[idx]);
    if(req.status!=='Menunggu Persetujuan') throw new Error('Permintaan ini sudah diproses.');

    const debtRow=findRowById_(req.hutang_id);
    if(!debtRow) throw new Error('Data kredit untuk pembayaran ini tidak ditemukan.');
    const debtSh=getSheet_();
    const obj=rowToObject_(debtSh.getRange(debtRow,1,1,HEADERS.length).getValues()[0]);
    const paid=Number(obj.angsuran_dibayar)||0;
    const total=Number(obj.jumlah_angsuran)||0;
    if(paid>=total) throw new Error('Kredit ini sudah lunas.');
    if(Number(req.angsuran_ke)!==paid+1) throw new Error('Urutan angsuran tidak sesuai. Refresh dashboard lalu coba lagi.');

    const paymentSh=ss.getSheetByName(PAYMENTS_SHEET_NAME);
    if(!paymentSh) throw new Error('Sheet Pembayaran belum ada. Jalankan setup().');
    const paymentId=Utilities.getUuid();
    paymentSh.appendRow([paymentId,obj.id,obj.nama,paid+1,Number(obj.angsuran_per_bulan)||0,req.tanggal_pembayaran,new Date()]);

    obj.angsuran_dibayar=paid+1;
    obj.tanggal_angsuran_terakhir=req.tanggal_pembayaran;
    obj.status=(paid+1)>=total?'Lunas':'Berjalan';
    obj.updated_at=new Date().toISOString();
    debtSh.getRange(debtRow,1,1,HEADERS.length).setValues([objectToRow_(obj)]);

    rows[idx][REQUEST_HEADERS.indexOf('status')]='Disetujui';
    rows[idx][REQUEST_HEADERS.indexOf('alasan_admin')]='Disetujui admin';
    rows[idx][REQUEST_HEADERS.indexOf('updated_at')]=new Date();
    reqSh.getRange(idx+2,1,1,REQUEST_HEADERS.length).setValues([rows[idx]]);
    return {request:requestRowToObject_(rows[idx]),debt:rowToObject_(debtSh.getRange(debtRow,1,1,HEADERS.length).getValues()[0])};
  }finally{
    lock.releaseLock();
  }
}

function rejectPaymentRequest_(id,reason){
  const lock=LockService.getScriptLock();
  lock.waitLock(15000);
  try{
    const sh=getRequestsSheet_(),last=sh.getLastRow();
    if(last<2) throw new Error('Permintaan pembayaran tidak ditemukan.');
    const rows=sh.getRange(2,1,last-1,REQUEST_HEADERS.length).getValues();
    const idx=rows.findIndex(r=>String(r[0])===String(id));
    if(idx<0) throw new Error('Permintaan pembayaran tidak ditemukan.');
    const req=requestRowToObject_(rows[idx]);
    if(req.status!=='Menunggu Persetujuan') throw new Error('Permintaan ini sudah diproses.');
    rows[idx][REQUEST_HEADERS.indexOf('status')]='Ditolak';
    rows[idx][REQUEST_HEADERS.indexOf('alasan_admin')]=String(reason||'Ditolak admin');
    rows[idx][REQUEST_HEADERS.indexOf('updated_at')]=new Date();
    sh.getRange(idx+2,1,1,REQUEST_HEADERS.length).setValues([rows[idx]]);
    return {request:requestRowToObject_(rows[idx])};
  }finally{
    lock.releaseLock();
  }
}

function delete_(id){const row=findRowById_(id);if(!row)throw new Error('Data tidak ditemukan.');getSheet_().deleteRow(row);return {id:id};}
function findRowById_(id){const sh=getSheet_(),last=sh.getLastRow();if(last<2)return null;const ids=sh.getRange(2,1,last-1,1).getValues().flat(),i=ids.findIndex(x=>String(x)===String(id));return i<0?null:i+2;}
function objectToRow_(o){return HEADERS.map(h=>o[h]??'');}

function dayFromDate_(dateValue){
  if(!dateValue)return 0;
  const p=String(dateValue).split('-').map(Number);
  return p.length===3?Number(p[2])||0:0;
}
function firstDueFromStart_(start,day){
  const p=String(start||'').split('-').map(Number); if(p.length!==3)return '';
  let y=p[0],m=p[1]-1,d=new Date(y,m,1);
  const startDate=new Date(y,m,p[2]);
  if(startDate.getDate()<=Number(day||1)) d=new Date(y,m,Math.min(Number(day||1),new Date(y,m+1,0).getDate()));
  else {m++; d=new Date(y,m,Math.min(Number(day||1),new Date(y,m+1,0).getDate()));}
  return Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd');
}
function addMonthsSafe_(dateValue,months,forcedDay){
  if(!dateValue)return '';
  const parts=String(dateValue).split('-').map(Number); if(parts.length<3||!parts[0])return '';
  const baseMonth=new Date(Date.UTC(parts[0],parts[1]-1,1));
  baseMonth.setUTCMonth(baseMonth.getUTCMonth()+Number(months||0));
  const y=baseMonth.getUTCFullYear(), m=baseMonth.getUTCMonth();
  const targetDay=Number(forcedDay)||parts[2];
  const lastDay=new Date(Date.UTC(y,m+1,0)).getUTCDate();
  const day=Math.min(targetDay,lastDay);
  return `${y}-${String(m+1).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
}
function formatValue_(v,h){
  if(v instanceof Date){
    if(['tanggal_mulai_angsuran','tanggal_jatuh_tempo_pertama','tanggal_angsuran_terakhir'].includes(h))return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');
    return v.toISOString();
  }
  return v;
}
function dateOut_(v){return v instanceof Date?v.toISOString():String(v||'');}
function normalize_(s){return String(s||'').trim().replace(/\s+/g,' ').toLowerCase();}
function sha256_(text){const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,text,Utilities.Charset.UTF_8);return bytes.map(b=>('0'+(b&0xFF).toString(16)).slice(-2)).join('');}
function jsonp_(e,obj){const cb=String(e.parameter.callback||'callback').replace(/[^\w.$]/g,'');return ContentService.createTextOutput(cb+'('+JSON.stringify(obj)+');').setMimeType(ContentService.MimeType.JAVASCRIPT);}
