/** ============================================================================
 *  نظام منارات المتحدة للمقايسات والتسعير — النسخة الاحترافية v3.1 (نسخة النشر)
 *  Google Apps Script مرتبط بجدول Google Sheets
 *  متوافق تماماً مع البيانات القديمة: لا يحذف ولا يعدّل أعمدة أو صفوفاً موجودة،
 *  ويضيف فقط أعمدة التكلفة التفصيلية في نهاية ورقة LINES عند الحاجة.
 *
 *  الجديد في v3.1:
 *   - صلاحيات دخول (قائمة مستخدمين مصرح لهم) على كل دالة يستدعيها المتصفح.
 *   - قفل LockService على كل عمليات الكتابة لمنع تضارب مستخدمين في اللحظة نفسها.
 *   - تحديد صف الكتابة من بيانات الورقة الخام (لا يتأثر بوجود صفوف فارغة).
 *   - حفظ رقم البند كنص حتى لا يحوّله الجدول إلى تاريخ (مثل 3/5).
 *   - دالة bootstrap واحدة لتحميل الواجهة بدل ست استدعاءات.
 * ==========================================================================*/
const S = { CAT:'دليل البنود', PROJECTS:'PROJECTS', LINES:'LINES', LOG:'سجل التعديلات', SETTINGS:'SETTINGS' };
const H = {
  CAT:      ['معرف','رقم الباب','اسم الباب','التصنيف الفرعي','رقم البند','وصف البند','الوصف الفرعي','الوحدة','السعر الأساسي','نشط'],
  PROJECTS: ['معرف المشروع','اسم المشروع','اسم العميل','تاريخ الإنشاء','الحالة','نسبة المصاريف %','نسبة الربح %','نسبة الضريبة %','ملاحظات'],
  LINES:    ['معرف السطر','معرف المشروع','معرف البند','رقم الباب','اسم الباب','التصنيف الفرعي','رقم البند','الوصف','الوحدة','الكمية','سعر الوحدة','الإجمالي','تاريخ التعديل','تكلفة الوحدة','إجمالي التكلفة','نسبة المصاريف الإدارية %','نسبة الأرباح %','السعر الاسترشادي','نظام التسعير','تكلفة المواد','تكلفة العمالة','تكلفة المعدات','تكلفة مقاول الباطن','نسبة الهالك %','نسبة الخصم %','ملاحظات البند'],
  LOG:      ['التاريخ','المستخدم','العملية','التفاصيل'],
  SETTINGS: ['المفتاح','القيمة']
};
/** الأعمدة التي تُحفظ كنص (رقم البند) — ترقيم يبدأ من 1 */
const TEXT_COLS = { CAT:[5], LINES:[7] };
const PROJECT_STATUSES = ['مسودة','قيد التسعير','معتمد','قيد التنفيذ','منفذ','ملغي'];
const COMPANY_DEFAULT_NAME = 'شركة منارات المتحدة للمقاولات العامة للمباني';
const APP_TITLE = 'منارات المتحدة | نظام المقايسات والتسعير';

/* ------------------------------ أدوات أساسية ------------------------------ */
function db_(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function tab_(n){
  const sh = db_().getSheetByName(n);
  if(!sh) throw new Error('ورقة العمل غير موجودة: ' + n);
  return sh;
}
function props_(){ return PropertiesService.getScriptProperties(); }
function who_(){ return me_() || 'مستخدم'; }
function log_(a,d){ tab_(S.LOG).appendRow([new Date(), who_(), a, safe_(d)]); }
function rows_(n){ const v = tab_(n).getDataRange().getValues(); return v.slice(1).filter(r=>r[0]!==''&&r[0]!==null); }
function id_(){ return Utilities.getUuid(); }
function check_(x){ if(!x || typeof x!=='string') throw Error('بيانات غير صحيحة'); }
function obj_(x){ if(!x || typeof x!=='object') throw Error('بيانات غير صحيحة'); return x; }
function clean_(x){ return String(x ?? '').trim(); }
/** نص آمن للكتابة في الجدول: يمنع تحوّل النص الذي يبدأ بـ = أو + إلى معادلة */
function safe_(x){ const s = clean_(x); return /^[=+]/.test(s) ? "'" + s : s; }
function number_(x){ const n = Number(x); if(!Number.isFinite(n) || n<0) throw Error('القيمة يجب أن تكون رقماً غير سالب'); return n; }
function round3_(n){ return Math.round((n + Number.EPSILON) * 1000) / 1000; }
function blank_(v){ return v===undefined || v===null || v===''; }
function numOr_(v, fallback){ if(blank_(v)) return fallback; const n=Number(v); return Number.isFinite(n)?n:fallback; }
function today_(){ return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'); }
function dateStr_(v){
  if(v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  return String(v == null ? '' : v);
}
/** رقم البند كما يراه المستخدم (الخلايا القديمة التي حوّلها الجدول إلى تاريخ تُعرض يوم/شهر) */
function formatItemCode_(value){
  if(value instanceof Date && !isNaN(value.getTime())){
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'd/M');
  }
  return String(value == null ? '' : value);
}

/* ------------------------- القفل والكتابة الآمنة -------------------------- */
/** ينفّذ عملية كتابة تحت قفل حتى لا يتضارب مستخدمان في اللحظة نفسها */
function locked_(fn){
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(20000)) throw Error('النظام مشغول بعملية أخرى؛ أعد المحاولة بعد لحظات');
  try{ return fn(); }
  finally{ try{ SpreadsheetApp.flush(); }catch(e){} lock.releaseLock(); }
}
/** فهرس الصف (داخل مصفوفة القيم الخام) المطابق للمعرّف، أو -1. رقم الصف في الورقة = الفهرس + 1 */
function rowOf_(v, id, projectId){
  for(let i=1;i<v.length;i++){
    if(String(v[i][0])===String(id) && (projectId===undefined || String(v[i][1])===String(projectId))) return i;
  }
  return -1;
}
function nextRow_(sh){ return Math.max(sh.getLastRow(), 1) + 1; }
/** كتابة صفوف دفعة واحدة مع ضبط أعمدة النص */
function writeRows_(sh, startRow, rows, textCols){
  if(!rows.length) return;
  const width = rows[0].length, lastNeeded = startRow + rows.length - 1;
  if(sh.getMaxColumns() < width) sh.insertColumnsAfter(sh.getMaxColumns(), width - sh.getMaxColumns());
  if(sh.getMaxRows() < lastNeeded) sh.insertRowsAfter(sh.getMaxRows(), lastNeeded - sh.getMaxRows());
  (textCols||[]).forEach(c=>sh.getRange(startRow, c, rows.length, 1).setNumberFormat('@'));
  sh.getRange(startRow, 1, rows.length, width).setValues(rows);
}
/** حذف صفوف مع ضمان بقاء صف واحد غير مجمّد على الأقل (شرط Google Sheets) */
function deleteRows_(sh, startRow, count){
  if(count<=0) return;
  if(sh.getMaxRows() - count < 2) sh.insertRowsAfter(sh.getMaxRows(), 1);
  sh.deleteRows(startRow, count);
}

/* ------------------------------- الصلاحيات -------------------------------- */
function me_(){ try{ return String(Session.getActiveUser().getEmail()||'').toLowerCase(); }catch(e){ return ''; } }
/** مدير النظام: خاصية ADMIN_EMAIL إن وُجدت، وإلا مالك ملف الجدول */
function admin_(){
  const p = clean_(props_().getProperty('ADMIN_EMAIL')).toLowerCase();
  if(p) return p;
  try{ const o = db_().getOwner(); return o ? String(o.getEmail()||'').toLowerCase() : ''; }catch(e){ return ''; }
}
function allowed_(){
  return String(props_().getProperty('ALLOWED_USERS')||'').split(/[,;\s]+/).map(x=>x.trim().toLowerCase()).filter(Boolean);
}
function isAdmin_(){ const me = me_(); return !!me && me === admin_(); }
/**
 * يمنع غير المصرح لهم. إذا كانت قائمة ALLOWED_USERS فارغة فالنظام مفتوح لكل من يملك الرابط
 * (السلوك القديم) — لذلك يجب تعبئة القائمة قبل النشر من شاشة الإعدادات أو بدالة setAllowedUsers.
 */
function guard_(){
  const list = allowed_();
  if(!list.length) return;
  const me = me_();
  if(me && (list.indexOf(me)>=0 || me === admin_())) return;
  throw Error('غير مصرح لهذا الحساب باستخدام النظام' + (me ? ' ('+me+')' : '') + '. راجع مدير النظام.');
}
function accessInfo_(){
  const admin = isAdmin_();
  return { me: me_(), isAdmin: admin, enabled: allowed_().length>0, allowed: admin ? allowed_() : [] };
}
function accessSettings(){ guard_(); return accessInfo_(); }
/** تحديد المستخدمين المصرح لهم (إيميلات مفصولة بفاصلة أو سطر) — للمدير فقط. قائمة فارغة = النظام مفتوح */
function setAllowedUsers(list){
  guard_();
  if(!isAdmin_()) throw Error('هذا الإجراء لمدير النظام (مالك ملف الجدول) فقط');
  const emails = String(list||'').split(/[,;\s]+/).map(x=>x.trim().toLowerCase()).filter(Boolean);
  const bad = emails.filter(e=>!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e));
  if(bad.length) throw Error('بريد إلكتروني غير صحيح: ' + bad.join('، '));
  const uniq = emails.filter((e,i)=>emails.indexOf(e)===i);
  props_().setProperty('ALLOWED_USERS', uniq.join(','));
  locked_(()=>log_('تعديل صلاحيات الدخول', uniq.length ? 'عدد المصرح لهم: '+uniq.length : 'القائمة فارغة — النظام مفتوح'));
  return accessInfo_();
}

/* --------------------------- الإعداد وواجهة الويب -------------------------- */
function setup(){
  guard_();
  const ss = db_();
  Object.keys(S).forEach(key=>{
    const name = S[key];
    let sh = ss.getSheetByName(name);
    if(!sh){ sh = ss.insertSheet(name); sh.setRightToLeft(true); }
    if(sh.getLastRow()===0){ sh.getRange(1,1,1,H[key].length).setValues([H[key]]); sh.setFrozenRows(1); }
  });
  ensurePricingColumns_();
  return 'تم التحقق من أوراق النظام دون حذف البيانات';
}
function doGet(){
  try{ guard_(); }
  catch(e){
    const msg = String(e.message||e).replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
    return HtmlService.createHtmlOutput(
      '<div dir="rtl" style="font-family:Tahoma,Arial;max-width:520px;margin:80px auto;padding:30px;border:1px solid #dde5f0;border-radius:14px;text-align:center;color:#172c4c">'+
      '<h2 style="color:#b3372f">غير مصرح بالدخول</h2><p style="line-height:1.9">'+msg+'</p></div>'
    ).setTitle(APP_TITLE);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_TITLE)
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}
/** تحميل كل بيانات الواجهة في استدعاء واحد */
function bootstrap(){
  guard_();
  return {
    projects: projects_(),
    catalogAll: catalogAll_(),
    settings: companySettings_(),
    defaults: defaultRates_(),
    stats: projectsOverview_(),
    access: accessInfo_()
  };
}

function classify_(name,desc){
  const s=(name+' '+desc).toLowerCase();
  const groups=[['هدم وإزالة',/هدم|ازال|إزال|تكسير|فك /],['حفر وردم',/حفر|ردم|دفان|تربة/],['خرسانة وتسليح',/خرسان|تسليح|اسمنت|أسمنت/],['أبواب ونوافذ',/باب|أبواب|نافذ|شباك/],['أرضيات',/أرضي|بلاط|سيراميك|رخام|جرانيت|بورسلان/],['حوائط وكسوة',/حائط|حوائط|جدار|قواطع|تكسي/],['أسقف',/سقف|أسقف|جبس|سقوف/],['دهانات وتشطيبات',/دهان|صبغ|معجون|طلاء/],['تمديدات كهربائية',/كهرب|كابل|كيبل|أسلاك|مفتاح|مقبس/],['إنارة',/إنار|إضاء|لمبة|كشاف/],['تكييف وتهوية',/تكييف|تبريد|مكيف|دكت|تهوية/],['تغذية وصرف',/صرف|مياه|أنبوب|مواسير|خلاط|مغسل|مرحاض/],['عزل',/عزل|مائي|حراري/],['حدائق وري',/زراع|ري |نبات|نخيل/],['مكافحة الحريق',/حريق|اطفاء|إطفاء|رشاش/]];
  return (groups.find(g=>g[1].test(s))||['بنود عامة'])[0];
}

/** استيراد من تبويب "جدول وصف البنود واسعارها" بعد رفع ملف المستخدم إلى Google Sheets — يُشغَّل من المحرر */
function importOriginal(){
  setup();
  const ss = db_();
  const normalizeName_ = x=>String(x).replace(/[أإآ]/g,'ا').replace(/ى/g,'ي').replace(/\s+/g,'').trim();
  const source = ss.getSheets().find(sh=>normalizeName_(sh.getName())===normalizeName_('جدول وصف البنود وأسعارها'));
  if(!source) throw Error('لم أجد ورقة «جدول وصف البنود وأسعارها» داخل ملف البرنامج. تأكد من استيرادها في الملف نفسه.');
  const vals=source.getDataRange().getValues(), out=[], seen=new Set();
  for(let i=1;i<vals.length;i++){
    const r=vals[i], no=clean_(formatItemCode_(r[1])), chapter=clean_(r[3]), desc=clean_(r[4]);
    if(!no || !chapter || !desc || no==='رقم البند') continue;
    const key=clean_(r[2])+'|'+no+'|'+desc;
    if(seen.has(key)) continue;
    seen.add(key);
    const rawPrice=r[7];
    const parsedPrice = typeof rawPrice==='number' ? rawPrice : Number(String(rawPrice??'').replace(/,/g,'').trim());
    out.push([id_(), clean_(r[2]), safe_(chapter), classify_(chapter,desc), no, safe_(desc), safe_(r[5]), safe_(r[6]), Number.isFinite(parsedPrice)&&parsedPrice>=0?parsedPrice:0, true]);
  }
  if(!out.length) throw Error('لم يتم العثور على بنود صالحة للاستيراد؛ لم تُغيّر أي بيانات');
  return locked_(()=>{
    const sh = tab_(S.CAT);
    if(sh.getLastRow()>1) throw Error('دليل البنود يحتوي على بيانات بالفعل؛ أوقفنا الاستيراد لحمايتها. لا تُعد تشغيله قبل مراجعة البيانات.');
    for(let i=0;i<out.length;i+=500) writeRows_(sh, i+2, out.slice(i,i+500), TEXT_COLS.CAT);
    log_('استيراد دليل البنود','عدد البنود: '+out.length);
    return {count:out.length, chapters:new Set(out.map(x=>x[2])).size};
  });
}

/* ------------------------------ دليل البنود ------------------------------- */
function catalogAll_(){
  return rows_(S.CAT).map(r=>({
    id:String(r[0]), chapter:String(r[2]), sub:String(r[3]||'بنود عامة'),
    code:formatItemCode_(r[4]), description:String(r[5]), detail:String(r[6]||''),
    unit:String(r[7]||''), price:Number(r[8])||0, active:r[9]!==false
  }));
}
function catalog(){ guard_(); return catalogAll_().filter(x=>x.active); }
/** كل البنود بما فيها الموقوفة — لشاشة إدارة الدليل */
function catalogAll(){ guard_(); return catalogAll_(); }
function saveCatalogItem(item){
  guard_(); obj_(item);
  check_(item.code); check_(item.description); check_(item.chapter);
  return locked_(()=>{
    const sh=tab_(S.CAT), v=sh.getDataRange().getValues();
    const id=clean_(item.id)||id_();
    const i=item.id ? rowOf_(v,id) : -1;
    if(item.id && i<0) throw Error('البند غير موجود');
    const row=[id, i>0?v[i][1]:'',                       // الحفاظ على رقم الباب الأصلي
               safe_(item.chapter), safe_(item.sub)||classify_(item.chapter,item.description),
               clean_(item.code), safe_(item.description), safe_(item.detail), safe_(item.unit),
               number_(item.price||0), i>0 ? v[i][9]!==false : true];   // الحفاظ على حالة التفعيل
    writeRows_(sh, i>0?i+1:nextRow_(sh), [row], TEXT_COLS.CAT);
    log_(i>0?'تعديل بند بالدليل':'إضافة بند للدليل', item.code);
    return id;
  });
}
function setCatalogActive(itemId, active){
  guard_(); check_(itemId);
  return locked_(()=>{
    const sh=tab_(S.CAT), i=rowOf_(sh.getDataRange().getValues(), itemId);
    if(i<0) throw Error('البند غير موجود');
    sh.getRange(i+1,10).setValue(active===true);
    log_(active===true?'تفعيل بند':'إيقاف بند', itemId);
    return true;
  });
}
function updateCatalogPrice(itemId,price){
  guard_(); check_(itemId);
  const value=number_(price);
  return locked_(()=>{
    const sh=tab_(S.CAT), i=rowOf_(sh.getDataRange().getValues(), itemId);
    if(i<0) throw Error('البند غير موجود');
    sh.getRange(i+1,9).setValue(value);
    log_('تعديل سعر دليل البنود', itemId);
    return true;
  });
}

/* ------------------------------- المشروعات -------------------------------- */
function projects_(){
  return rows_(S.PROJECTS).map(r=>({
    id:String(r[0]), name:String(r[1]||''), client:String(r[2]||''), date:dateStr_(r[3]),
    status:String(r[4]||'مسودة'), overhead:Number(r[5])||0, profit:Number(r[6])||0,
    tax:Number(r[7])||0, notes:String(r[8]||'')
  }));
}
function projects(){ guard_(); return projects_(); }
/** الحفظ الفعلي — يُستدعى دائماً من داخل locked_ */
function saveProject_(p){
  const sh=tab_(S.PROJECTS), v=sh.getDataRange().getValues();
  const id=clean_(p.id)||id_();
  const i=p.id ? rowOf_(v,id) : -1;
  if(p.id && i<0) throw Error('المشروع غير موجود');
  const status = PROJECT_STATUSES.indexOf(clean_(p.status))>=0 ? clean_(p.status) : 'مسودة';
  const row=[id, safe_(p.name), safe_(p.client), (i>0 && v[i][3]) ? v[i][3] : new Date(),
             status, number_(p.overhead||0), number_(p.profit||0), number_(p.tax||0), safe_(p.notes)];
  writeRows_(sh, i>0?i+1:nextRow_(sh), [row]);
  log_(i>0?'تعديل مشروع':'إنشاء مشروع', p.name);
  return id;
}
function saveProject(p){
  guard_(); obj_(p); check_(p.name);
  if(!clean_(p.name)) throw Error('أدخل اسم المشروع');
  return locked_(()=>saveProject_(p));
}
function setProjectStatus(projectId, status){
  guard_(); check_(projectId);
  const st=clean_(status);
  if(PROJECT_STATUSES.indexOf(st)<0) throw Error('حالة المشروع غير معروفة');
  return locked_(()=>{
    const sh=tab_(S.PROJECTS), i=rowOf_(sh.getDataRange().getValues(), projectId);
    if(i<0) throw Error('المشروع غير موجود');
    sh.getRange(i+1,5).setValue(st);
    log_('تغيير حالة مشروع', st);
    return true;
  });
}
function deleteProject(projectId){
  guard_(); check_(projectId);
  return locked_(()=>{
    const shP=tab_(S.PROJECTS), vP=shP.getDataRange().getValues(), ip=rowOf_(vP, projectId);
    if(ip<0) throw Error('المشروع غير موجود');
    const name=String(vP[ip][1]||'');
    // حذف بنود المشروع على دفعات متجاورة من أسفل لأعلى
    const shL=tab_(S.LINES), vL=shL.getDataRange().getValues();
    let end=-1, removed=0;
    for(let i=vL.length-1;i>=0;i--){
      const hit = i>=1 && String(vL[i][1])===String(projectId);
      if(hit){ if(end<0) end=i; continue; }
      if(end>=0){ deleteRows_(shL, i+2, end-i); removed+=end-i; end=-1; }
    }
    deleteRows_(shP, ip+1, 1);
    log_('حذف مشروع', name+' — بنود محذوفة: '+removed);
    return true;
  });
}
function duplicateProject(projectId){
  guard_(); check_(projectId);
  return locked_(()=>{
    const p=projects_().find(x=>x.id===String(projectId));
    if(!p) throw Error('المشروع غير موجود');
    const newId=saveProject_({name:p.name+' (نسخة)', client:p.client, overhead:p.overhead, profit:p.profit, tax:p.tax, status:'مسودة', notes:p.notes});
    ensurePricingColumns_();
    const sh=tab_(S.LINES), v=sh.getDataRange().getValues(), rows=[];
    for(let i=1;i<v.length;i++) if(String(v[i][1])===String(projectId)) rows.push(copyLineRow_(v[i], newId));
    writeRows_(sh, nextRow_(sh), rows, TEXT_COLS.LINES);
    log_('نسخ مشروع', p.name+' → نسخة جديدة ('+rows.length+' بند)');
    return newId;
  });
}

/* ---------------------------- بنود المشروع (LINES) ------------------------- */
/** إضافة أعمدة التسعير الجديدة دون حذف أو تعديل أي صف قديم */
function ensurePricingColumns_(){
  const sh=tab_(S.LINES), headers=H.LINES;
  if(sh.getMaxColumns()<headers.length) sh.insertColumnsAfter(sh.getMaxColumns(), headers.length-sh.getMaxColumns());
  const current=sh.getRange(1,1,1,headers.length).getValues()[0];
  for(let i=0;i<headers.length;i++) if(!current[i]) sh.getRange(1,i+1).setValue(headers[i]);
}
/** نسخة من صف بند بمعرّف جديد (ومشروع جديد اختيارياً) مع تثبيت رقم البند كنص */
function copyLineRow_(src, newProjectId){
  const row=src.slice(0, H.LINES.length);
  while(row.length<H.LINES.length) row.push('');
  row[0]=id_(); if(newProjectId) row[1]=newProjectId;
  row[6]=formatItemCode_(row[6]); row[12]=new Date();
  return row;
}
function lines_(projectId){
  return rows_(S.LINES).filter(r=>String(r[1])===String(projectId)).map(r=>{
    const qty=Number(r[9])||0, price=Number(r[10])||0;
    const cost = blank_(r[13]) ? price : (Number(r[13])||0);
    const costTotal = blank_(r[14]) ? round3_(qty*cost) : (Number(r[14])||0);
    return {
      id:String(r[0]), itemId:String(r[2]), chapter:String(r[4]||''), sub:String(r[5]||''),
      code:formatItemCode_(r[6]), description:String(r[7]||''), unit:String(r[8]||''),
      qty:qty, price:price, total:Number(r[11])||0, cost:cost, costTotal:costTotal,
      overhead:Number(r[15])||0, profit:Number(r[16])||0, reference:Number(r[17])||0,
      pricingMode:r[18]==='v3'?'v3':(r[18]==='v2'?'v2':'legacy'),
      materials:Number(r[19])||0, labor:Number(r[20])||0, equipment:Number(r[21])||0,
      subcontractor:Number(r[22])||0, waste:Number(r[23])||0, discount:Number(r[24])||0,
      note:String(r[25]||'')
    };
  });
}
function lines(projectId){ guard_(); check_(projectId); return lines_(projectId); }
/**
 * حفظ بند — يدعم الطريقتين:
 *  1) كائن واحد: saveLine({projectId,itemId,qty,lineId,manual,cost,materials,labor,equipment,subcontractor,waste,overhead,profit,discount,note})
 *  2) الوسائط القديمة: saveLine(projectId,itemId,qty,price,lineId,cost,overhead,profit)
 */
function saveLine(a,b,c,d,e,f,g,h){
  guard_();
  const p=(a && typeof a==='object')
    ? a
    : {projectId:a, itemId:b, qty:c, lineId:e, cost:f, overhead:g, profit:h, legacyCall:true};
  check_(p.projectId); check_(p.itemId);
  const q=number_(p.qty); if(q<=0) throw Error('أدخل كمية أكبر من صفر');
  const oh=number_(p.overhead||0), pf=number_(p.profit||0);
  const disc=number_(p.discount||0); if(disc>=100) throw Error('نسبة الخصم يجب أن تكون أقل من 100%');
  const waste=number_(p.waste||0);

  const bd={materials:0, labor:0, equipment:0, subcontractor:0};
  let hasBD=false;
  ['materials','labor','equipment','subcontractor'].forEach(k=>{
    if(!blank_(p[k])){ bd[k]=number_(p[k]); if(bd[k]>0) hasBD=true; }
  });

  let unitCost, mode;
  if(p.legacyCall){
    unitCost=number_(p.cost); mode='v2';
  }else{
    if(hasBD && p.manual!==true){
      unitCost=round3_((bd.materials+bd.labor+bd.equipment+bd.subcontractor)*(1+waste/100));
    }else{
      unitCost=number_(p.cost||0);
    }
    mode='v3';
  }
  const unitPrice=round3_(unitCost*(1+(oh+pf)/100)*(1-disc/100));

  return locked_(()=>{
    const project=projects_().find(x=>x.id===String(p.projectId));
    if(!project) throw Error('المشروع غير موجود');
    const item=catalogAll_().find(x=>x.id===String(p.itemId));
    if(!item) throw Error('البند غير موجود');
    ensurePricingColumns_();
    const sh=tab_(S.LINES), v=sh.getDataRange().getValues();
    const id=clean_(p.lineId)||id_();
    const i=p.lineId ? rowOf_(v, id, p.projectId) : -1;
    if(p.lineId && i<0) throw Error('السطر غير موجود');
    const row=[id, String(p.projectId), String(p.itemId), '', safe_(item.chapter), safe_(item.sub), clean_(item.code),
               safe_(item.description+(item.detail?' — '+item.detail:'')), safe_(item.unit), q, unitPrice,
               round3_(q*unitPrice), new Date(), unitCost, round3_(q*unitCost), oh, pf, item.price,
               mode, bd.materials, bd.labor, bd.equipment, bd.subcontractor, waste, disc, safe_(p.note)];
    writeRows_(sh, i>0?i+1:nextRow_(sh), [row], TEXT_COLS.LINES);
    log_('حفظ بند', project.name+' / '+item.code);
    return {id:id, total:round3_(q*unitPrice)};
  });
}
function removeLine(projectId,lineId){
  guard_(); check_(projectId); check_(lineId);
  return locked_(()=>{
    const sh=tab_(S.LINES), i=rowOf_(sh.getDataRange().getValues(), lineId, projectId);
    if(i<0) throw Error('السطر غير موجود');
    deleteRows_(sh, i+1, 1);
    log_('حذف بند', lineId);
    return true;
  });
}
function duplicateLine(projectId,lineId){
  guard_(); check_(projectId); check_(lineId);
  return locked_(()=>{
    ensurePricingColumns_();
    const sh=tab_(S.LINES), v=sh.getDataRange().getValues(), i=rowOf_(v, lineId, projectId);
    if(i<0) throw Error('السطر غير موجود');
    writeRows_(sh, nextRow_(sh), [copyLineRow_(v[i])], TEXT_COLS.LINES);
    log_('نسخ بند', lineId);
    return true;
  });
}

/* ------------------------------ الإحصاءات --------------------------------- */
function summary_(p, ls){
  const sales=ls.reduce((a,x)=>a+x.total,0), cost=ls.reduce((a,x)=>a+x.costTotal,0);
  // الأسعار الجديدة تشمل المصاريف والأرباح داخل كل بند. البنود القديمة تبقى كما كانت محفوظة.
  const legacySales=ls.filter(x=>x.pricingMode==='legacy').reduce((a,x)=>a+x.total,0);
  const overhead=legacySales*p.overhead/100, profit=(legacySales+overhead)*p.profit/100;
  const net=sales+overhead+profit, tax=net*p.tax/100;
  return {count:ls.length, sales:sales, cost:cost, overhead:overhead, profit:profit, tax:tax,
          final:net+tax, difference:net-cost,
          margin:net>0 ? round3_((net-cost)/net*100) : 0};
}
function projectSummary(projectId){
  guard_();
  const p=projects_().find(x=>x.id===String(projectId)); if(!p) throw Error('المشروع غير موجود');
  return summary_(p, lines_(projectId));
}
/** نظرة شاملة لكل المشاريع — تغذي لوحة التحكم وشاشة المشروعات */
function projectsOverview_(){
  const ps=projects_();
  const all=rows_(S.LINES), byProject={};
  all.forEach(r=>{ const pid=String(r[1]); (byProject[pid]=byProject[pid]||[]).push(r); });
  const totals={sales:0,cost:0,net:0,final:0,lines:0};
  const per=ps.map(p=>{
    const rs=byProject[p.id]||[];
    const sales=rs.reduce((a,r)=>a+(Number(r[11])||0),0);
    const cost=rs.reduce((a,r)=>{ const q=Number(r[9])||0,c=blank_(r[13])?(Number(r[10])||0):(Number(r[13])||0); return a+(blank_(r[14])?q*c:(Number(r[14])||0)); },0);
    const legacy=rs.filter(r=>r[18]!=='v2'&&r[18]!=='v3').reduce((a,r)=>a+(Number(r[11])||0),0);
    const over=legacy*p.overhead/100, prof=(legacy+over)*p.profit/100;
    const net=sales+over+prof, tax=net*p.tax/100, fin=net+tax;
    totals.sales+=sales; totals.cost+=cost; totals.net+=net; totals.final+=fin; totals.lines+=rs.length;
    return {id:p.id, name:p.name, client:p.client, status:p.status, date:p.date,
            overhead:p.overhead, profit:p.profit, tax:p.tax,
            count:rs.length, sales:sales, cost:cost, net:net, final:fin,
            margin:net>0?round3_((net-cost)/net*100):0};   // الهامش من البيع قبل الضريبة
  });
  const byStatus={}; ps.forEach(p=>{ byStatus[p.status]=(byStatus[p.status]||0)+1; });
  const cat=rows_(S.CAT), chapters={};
  cat.forEach(r=>{ const c=String(r[2]||'عام'); chapters[c]=(chapters[c]||0)+1; });
  return {
    projectsCount:ps.length, byStatus:byStatus, totals:totals,
    catalogCount:cat.length, chapters:chapters,
    topProjects:per.filter(x=>x.final>0).sort((a,b)=>b.final-a.final).slice(0,5),
    recent:per.slice().reverse().slice(0,5),
    perProject:per
  };
}
function projectsOverview(){ guard_(); return projectsOverview_(); }
function getLog(limit){
  guard_();
  const v=rows_(S.LOG);
  return v.slice(-(Number(limit)||100)).reverse().map(r=>({
    date: r[0] ? Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : '',
    user:String(r[1]||''), action:String(r[2]||''), details:String(r[3]||'')
  }));
}

/* ------------------------- إعدادات الشركة والطباعة ------------------------- */
function companySettings_(){
  const props=props_();
  return {
    name:props.getProperty('COMPANY_NAME')||COMPANY_DEFAULT_NAME,
    stampId:props.getProperty('COMPANY_STAMP_FILE_ID')||'',
    address:props.getProperty('COMPANY_ADDRESS')||'',
    phone:props.getProperty('COMPANY_PHONE')||'',
    currency:'د.ك'
  };
}
function companySettings(){ guard_(); return companySettings_(); }
/** يقبل معرّف الملف أو رابط Drive كاملاً. ترك الحقل فارغاً يلغي الختم المسجل */
function setCompanyDetails(address,phone,stampFileId,name){
  guard_();
  const props=props_();
  const raw=clean_(stampFileId);
  if(raw){
    const m=raw.match(/[-\w]{25,}/);
    if(!m) throw Error('معرّف ملف الختم غير صحيح');
    const file=DriveApp.getFileById(m[0]);
    if(!/^image\/(png|jpeg|webp)$/.test(file.getMimeType())) throw Error('الختم يجب أن يكون صورة PNG أو JPG أو WEBP');
    props.setProperty('COMPANY_STAMP_FILE_ID', file.getId());
  }else{
    props.deleteProperty('COMPANY_STAMP_FILE_ID');
  }
  props.setProperty('COMPANY_ADDRESS', clean_(address));
  props.setProperty('COMPANY_PHONE', clean_(phone));
  if(!blank_(name) && clean_(name)) props.setProperty('COMPANY_NAME', clean_(name));
  locked_(()=>log_('تعديل بيانات الشركة','تحديث بيانات الطباعة والختم'));
  return companySettings_();
}
function defaultRates_(){
  const props=props_();
  return {
    overhead:Number(props.getProperty('DEFAULT_OVERHEAD'))||0,
    profit:Number(props.getProperty('DEFAULT_PROFIT'))||0,
    tax:Number(props.getProperty('DEFAULT_TAX'))||0
  };
}
function defaultRates(){ guard_(); return defaultRates_(); }
function setDefaultRates(o,p,t){
  guard_();
  const props=props_();
  const vals=[number_(o||0), number_(p||0), number_(t||0)];
  props.setProperty('DEFAULT_OVERHEAD', String(vals[0]));
  props.setProperty('DEFAULT_PROFIT',  String(vals[1]));
  props.setProperty('DEFAULT_TAX',     String(vals[2]));
  locked_(()=>log_('تعديل النسب الافتراضية', vals.join(' / ')));
  return defaultRates_();
}

/* --------------------------- إخراج جدول الكميات --------------------------- */
/** ملف داخلي: يحتوي على أسعار التكلفة ونسب الأرباح — لا يُرسل للعميل كما هو */
function createBOQ(projectId){
  guard_();
  const pr=projects_().find(p=>p.id===String(projectId)); if(!pr) throw Error('المشروع غير موجود');
  const ls=lines_(projectId); if(!ls.length) throw Error('أضف بنداً واحداً على الأقل قبل إصدار جدول الكميات');
  const ss=SpreadsheetApp.create('جدول كميات - '+pr.name), sh=ss.getActiveSheet();
  sh.setName('جدول الكميات'); sh.setRightToLeft(true);
  sh.getRange(1,1,1,10).merge().setValue('جدول كميات وتسعير مشروع: '+pr.name).setFontSize(16).setFontWeight('bold').setBackground('#173b70').setFontColor('white');
  sh.getRange(2,1).setValue('العميل: '+pr.client);
  sh.getRange(3,1).setValue('نسخة داخلية — تحتوي على التكلفة ونسب الأرباح').setFontColor('#b3372f').setFontWeight('bold');
  const headers=['م','الباب','التصنيف الفرعي','رقم البند / وصف الأعمال','الوحدة','الكمية','السعر شامل','إجمالي البيع','سعر التكلفة','إجمالي التكلفة','مصاريف إدارية %','أرباح %','خصم %','السعر الاسترشادي'];
  sh.getRange(4,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e8eff8');
  const data=ls.map((x,i)=>[i+1, safe_(x.chapter), safe_(x.sub), safe_(x.code+' - '+x.description), safe_(x.unit), x.qty, x.price,
    '=F'+(i+5)+'*G'+(i+5), x.cost, '=F'+(i+5)+'*I'+(i+5),
    x.pricingMode==='legacy'?'قديم':x.overhead, x.pricingMode==='legacy'?'قديم':x.profit,
    x.pricingMode==='legacy'?'':x.discount, x.pricingMode==='legacy'?'':x.reference]);
  sh.getRange(5,1,data.length,headers.length).setValues(data);
  const end=5+ls.length;
  sh.getRange(end,7).setValue('إجمالي سعر البنود شامل المصاريف والأرباح'); sh.getRange(end,8).setFormula('=SUM(H5:H'+(end-1)+')');
  const oldSales=ls.filter(x=>x.pricingMode==='legacy').reduce((a,x)=>a+x.total,0), oldOver=oldSales*pr.overhead/100, oldProfit=(oldSales+oldOver)*pr.profit/100;
  sh.getRange(end+1,7).setValue('إضافات البنود القديمة فقط'); sh.getRange(end+1,8).setValue(oldOver+oldProfit);
  sh.getRange(end+2,7).setValue('الضريبة '+pr.tax+'%'); sh.getRange(end+2,8).setFormula('=SUM(H'+end+':H'+(end+1)+')*'+pr.tax+'/100');
  sh.getRange(end+3,7).setValue('الإجمالي النهائي'); sh.getRange(end+3,8).setFormula('=SUM(H'+end+':H'+(end+2)+')');
  sh.getRange(end+5,7).setValue('إجمالي تكلفة التنفيذ'); sh.getRange(end+5,8).setFormula('=SUM(J5:J'+(end-1)+')');
  sh.getRange(end+6,7).setValue('الفرق بين البيع والتكلفة قبل الضريبة'); sh.getRange(end+6,8).setFormula('=SUM(H'+end+':H'+(end+1)+')-H'+(end+5));
  sh.getRange(5,6,ls.length+7,5).setNumberFormat('#,##0.000'); sh.getRange(5,14,ls.length,1).setNumberFormat('#,##0.000');
  sh.setColumnWidth(4,620); sh.setColumnWidth(2,210); sh.setColumnWidth(3,170);
  for(let c=7;c<=14;c++) sh.setColumnWidth(c,140);
  sh.setFrozenRows(4);
  const folderId=props_().getProperty('BOQ_FOLDER_ID');
  if(folderId) DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(folderId));
  locked_(()=>log_('إصدار جدول كميات', pr.name));
  return ss.getUrl();
}
/** اختياري: استدعِ هذه الدالة من المحرر بعد إنشاء مجلد مخصص في Drive */
function setOutputFolder(folderId){
  guard_();
  if(!isAdmin_()) throw Error('هذا الإجراء لمدير النظام فقط');
  DriveApp.getFolderById(folderId);
  props_().setProperty('BOQ_FOLDER_ID', folderId);
  return true;
}

/* ------------------------- بيانات الطباعة والتقارير ------------------------ */
function quoteData(projectId){
  guard_();
  const p=projects_().find(x=>x.id===String(projectId)); if(!p) throw Error('المشروع غير موجود');
  const ls=lines_(projectId); if(!ls.length) throw Error('أضف بنداً واحداً على الأقل قبل إصدار المقايسة');
  const settings=companySettings_(); let stamp='';
  if(settings.stampId){
    const f=DriveApp.getFileById(settings.stampId), b=f.getBlob();
    if(b.getBytes().length>800000) throw Error('صورة الختم كبيرة؛ استخدم صورة أقل من 800 كيلوبايت');
    stamp='data:'+b.getContentType()+';base64,'+Utilities.base64Encode(b.getBytes());
  }
  // نسخة العميل: لا تُرسل للمتصفح أي بيانات تكلفة أو هوامش
  const clientLines=ls.map(x=>({chapter:x.chapter, sub:x.sub, code:x.code, description:x.description,
                               unit:x.unit, qty:x.qty, price:x.price, total:x.total}));
  const s=summary_(p, ls);
  return {company:settings, stamp:stamp, project:p, lines:clientLines,
          summary:{count:s.count, sales:s.sales, overhead:s.overhead, profit:s.profit, tax:s.tax, final:s.final},
          date:today_()};
}
/** تقرير التكلفة الداخلي (للإدارة فقط — لا يُطبع للعميل) */
function costReportData(projectId){
  guard_();
  const p=projects_().find(x=>x.id===String(projectId)); if(!p) throw Error('المشروع غير موجود');
  const ls=lines_(projectId); if(!ls.length) throw Error('أضف بنداً واحداً على الأقل قبل إصدار التقرير');
  return {company:companySettings_(), project:p, lines:ls, summary:summary_(p, ls), date:today_()};
}

/* ------------------------------- اختبارات ---------------------------------- */
function testProjectsRead(){
  guard_();
  const data=projects_();
  console.log('عدد المشروعات المقروءة: '+data.length);
  console.log(JSON.stringify(data.map(p=>({id:p.id, name:p.name}))));
}
