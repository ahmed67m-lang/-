/** ============================================================================
 *  نظام منارات المتحدة للمقايسات والتسعير — النسخة الاحترافية v3.8 (نسخة النشر)
 *  Google Apps Script مرتبط بجدول Google Sheets
 *  متوافق تماماً مع البيانات القديمة: لا يحذف ولا يعدّل أعمدة أو صفوفاً موجودة،
 *  ويضيف فقط أوراقاً وأعمدة جديدة عند الحاجة.
 *
 *  الجديد في v3.4: البنود التي بلا رقم في ملف الأسعار تُضاف داخل أبوابها بتصنيف فرعي وترقيم رئيسي/فرعي (س/باب، ف/س/باب).
 *  الجديد في v3.3:
 *   - حسابات المشروع: ميزانية المقايسة، مصروفات (مواد/عمالة/معدات/تشغيل/مقاول باطن)،
 *     دفعات العميل (مستحقة/محصلة)، ورفع نسخة العقد إلى Google Drive.
 *   - أوراق جديدة تُنشأ تلقائياً عند setup: EXPENSES و PAYMENTS و CONTRACTS.
 *  الجديد في v3.2:
 *   - تسجيل دخول باسم مستخدم وكلمة مرور (rpc) — لا تُنفَّذ أي دالة بيانات بدون جلسة صالحة.
 *   - إدارة المستخدمين من شاشة الإعدادات (مدير / مستخدم)، وكلمات المرور تُحفظ مشفّرة (hash) فقط.
 *  من v3.1:
 *   - قفل LockService على كل عمليات الكتابة لمنع تضارب مستخدمين في اللحظة نفسها.
 *   - تحديد صف الكتابة من بيانات الورقة الخام (لا يتأثر بوجود صفوف فارغة).
 *   - حفظ رقم البند كنص حتى لا يحوّله الجدول إلى تاريخ (مثل 3/5).
 *   - دالة bootstrap واحدة لتحميل الواجهة بدل ست استدعاءات.
 * ==========================================================================*/
const S = { CAT:'دليل البنود', PROJECTS:'PROJECTS', LINES:'LINES', LOG:'سجل التعديلات', SETTINGS:'SETTINGS',
            EXPENSES:'EXPENSES', PAYMENTS:'PAYMENTS', CONTRACTS:'CONTRACTS' };
const H = {
  CAT:      ['معرف','رقم الباب','اسم الباب','التصنيف الفرعي','رقم البند','وصف البند','الوصف الفرعي','الوحدة','السعر الأساسي','نشط','طريقة التنفيذ'],
  PROJECTS: ['معرف المشروع','اسم المشروع','اسم العميل','تاريخ الإنشاء','الحالة','نسبة المصاريف %','نسبة الربح %','نسبة الضريبة %','ملاحظات'],
  LINES:    ['معرف السطر','معرف المشروع','معرف البند','رقم الباب','اسم الباب','التصنيف الفرعي','رقم البند','الوصف','الوحدة','الكمية','سعر الوحدة','الإجمالي','تاريخ التعديل','تكلفة الوحدة','إجمالي التكلفة','نسبة المصاريف الإدارية %','نسبة الأرباح %','السعر الاسترشادي','نظام التسعير','تكلفة المواد','تكلفة العمالة','تكلفة المعدات','تكلفة مقاول الباطن','نسبة الهالك %','نسبة الخصم %','ملاحظات البند'],
  LOG:      ['التاريخ','المستخدم','العملية','التفاصيل'],
  SETTINGS: ['المفتاح','القيمة'],
  EXPENSES: ['معرف','معرف المشروع','التاريخ','الفئة','البيان','المبلغ','طريقة الصرف','ملاحظات'],
  PAYMENTS: ['معرف','معرف المشروع','التاريخ','النوع','البيان','المبلغ','ملاحظات','تاريخ الاستحقاق','طريقة الدفع','رقم السند'],
  CONTRACTS:['معرف','معرف المشروع','اسم الملف','معرف الملف','الرابط','تاريخ الرفع','المستخدم']
};
const EXPENSE_CATS = ['مواد','عمالة','معدات','تشغيل','مقاول باطن','أخرى'];
const PAYMENT_TYPES = ['مستحقة','محصلة'];
/** الأعمدة التي تُحفظ كنص (رقم البند) — ترقيم يبدأ من 1 */
const TEXT_COLS = { CAT:[5], LINES:[7] };
const PROJECT_STATUSES = ['مسودة','قيد التسعير','معتمد','قيد التنفيذ','منفذ','ملغي'];
const COMPANY_DEFAULT_NAME = 'شركة منارات المتحدة للمقاولات العامة للمباني';
const APP_TITLE = 'منارات المتحدة | نظام المقايسات والتسعير';

/* ------------------------------ أدوات أساسية ------------------------------ */
function db_(){ return SpreadsheetApp.getActiveSpreadsheet(); }
/** يجلب الورقة، وينشئها بعناوينها إن كانت من أوراق النظام وغير موجودة بعد (إصلاح ذاتي بدون setup يدوي) */
function tab_(n){
  let sh = db_().getSheetByName(n);
  if(!sh){
    const key = Object.keys(S).find(k=>S[k]===n);
    if(!key) throw new Error('ورقة العمل غير موجودة: ' + n);
    sh = db_().insertSheet(n);
    sh.setRightToLeft(true);
    sh.getRange(1,1,1,H[key].length).setValues([H[key]]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function props_(){ return PropertiesService.getScriptProperties(); }
function who_(){ return CURRENT_ ? CURRENT_.username : (me_() || 'النظام'); }
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

/* --------------------------- الدخول والصلاحيات ---------------------------- */
/** المستخدم الحالي — لا يُضبط إلا داخل rpc بعد التحقق من الجلسة */
let CURRENT_ = null;
const HASH_ROUNDS = 200, SESSION_SECONDS = 21600, MAX_FAILS = 5, LOCK_SECONDS = 900;
const ROLES = ['admin','user'];

function me_(){ try{ return String(Session.getActiveUser().getEmail()||'').toLowerCase(); }catch(e){ return ''; } }
/** مالك ملف الجدول (أو خاصية ADMIN_EMAIL إن وُجدت) — هو الوحيد الذي ينشئ حساب المدير الأول */
function owner_(){
  const p = clean_(props_().getProperty('ADMIN_EMAIL')).toLowerCase();
  if(p) return p;
  try{ const o = db_().getOwner(); return o ? String(o.getEmail()||'').toLowerCase() : ''; }catch(e){ return ''; }
}
function isOwner_(){ const me = me_(); return !!me && me === owner_(); }
/** كل دالة بيانات تبدأ بهذا الفحص: الاستدعاء المباشر من المتصفح (بدون rpc) يُرفض */
function guard_(){ if(!CURRENT_) throw Error('يجب تسجيل الدخول أولاً'); }
function adminOnly_(){ guard_(); if(CURRENT_.role!=='admin') throw Error('هذا الإجراء لمدير النظام فقط'); }

function userKey_(name){ return 'USER_' + clean_(name).toLowerCase(); }
function getUser_(name){
  const raw = props_().getProperty(userKey_(name));
  if(!raw) return null;
  try{ return JSON.parse(raw); }catch(e){ return null; }
}
function putUser_(u){ props_().setProperty(userKey_(u.username), JSON.stringify(u)); }
function allUsers_(){
  const p = props_();
  return p.getKeys().filter(k=>k.indexOf('USER_')===0).map(k=>{ try{ return JSON.parse(p.getProperty(k)); }catch(e){ return null; } })
    .filter(Boolean).sort((a,b)=>String(a.username).localeCompare(String(b.username)));
}
function publicUser_(u){ return {username:u.username, role:u.role, active:u.active!==false, created:u.created||'', lastLogin:u.lastLogin||''}; }
function activeAdmins_(list){ return list.filter(u=>u.role==='admin' && u.active!==false).length; }
function checkUsername_(name){
  const n = clean_(name);
  if(!/^[A-Za-z0-9؀-ۿ._-]{3,30}$/.test(n)) throw Error('اسم المستخدم من 3 إلى 30 حرفاً أو رقماً بدون مسافات');
  return n;
}
function checkPassword_(pw){
  if(typeof pw!=='string' || pw.length<8 || pw.length>100) throw Error('كلمة المرور يجب ألا تقل عن 8 أحرف');
  return pw;
}
function hash_(password, salt){
  const saltBytes = Utilities.newBlob(salt).getBytes();
  let b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + password, Utilities.Charset.UTF_8);
  for(let i=0;i<HASH_ROUNDS;i++) b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, b.concat(saltBytes));
  return Utilities.base64Encode(b);
}
function newUser_(username, password, role){
  const salt = id_();
  return {username:username, salt:salt, hash:hash_(password, salt), role:role, active:true, ver:1, created:today_()};
}
function setPassword_(u, password){ u.salt = id_(); u.hash = hash_(password, u.salt); u.ver = (Number(u.ver)||1) + 1; }
function cache_(){ return CacheService.getScriptCache(); }
function openSession_(u){
  const token = (id_() + id_()).replace(/-/g,'');
  cache_().put('SESS_'+token, JSON.stringify({u:u.username, v:u.ver}), SESSION_SECONDS);
  return token;
}
/** يتحقق من رمز الجلسة ويعيد المستخدم، ويمدّ صلاحية الجلسة */
function session_(token){
  const expired = 'انتهت الجلسة؛ سجّل الدخول من جديد';
  if(!token || typeof token!=='string' || token.length>100) throw Error(expired);
  const raw = cache_().get('SESS_'+token);
  if(!raw) throw Error(expired);
  let s; try{ s = JSON.parse(raw); }catch(e){ throw Error(expired); }
  const u = getUser_(s.u);
  if(!u || u.active===false || u.ver!==s.v){ cache_().remove('SESS_'+token); throw Error(expired); }
  cache_().put('SESS_'+token, raw, SESSION_SECONDS);
  return {username:u.username, role:u.role, token:token};
}

/** حالة النظام لشاشة الدخول (لا تتطلب جلسة) */
function authState(){
  const has = allUsers_().length>0;
  return {hasUsers:has, canSetup:!has && isOwner_(), company:companySettings_().name};
}
/** إنشاء حساب المدير الأول — مرة واحدة فقط، ومن حساب Google المالك لملف الجدول */
function setupAdmin(username, password){
  const name = checkUsername_(username); checkPassword_(password);
  return locked_(()=>{
    if(allUsers_().length) throw Error('تم إعداد حساب المدير من قبل؛ استخدم شاشة الدخول');
    if(!isOwner_()) throw Error('إنشاء حساب المدير الأول يتم فقط من حساب Google المالك لملف الجدول');
    const u = newUser_(name, password, 'admin');
    putUser_(u);
    CURRENT_ = {username:u.username, role:u.role};
    log_('إنشاء حساب المدير', u.username);
    return {token:openSession_(u), user:{username:u.username, role:u.role}};
  });
}
function login(username, password){
  const bad = 'اسم المستخدم أو كلمة المرور غير صحيحة';
  const name = clean_(username).toLowerCase();
  if(!name || typeof password!=='string' || !password || name.length>40 || password.length>100) throw Error(bad);
  const failKey = 'FAIL_'+name, fails = Number(cache_().get(failKey))||0;
  if(fails>=MAX_FAILS) throw Error('تم إيقاف الدخول مؤقتاً بعد محاولات خاطئة متكررة؛ حاول بعد 15 دقيقة');
  const u = getUser_(name);
  if(!u || u.active===false || hash_(password, u.salt)!==u.hash){
    cache_().put(failKey, String(fails+1), LOCK_SECONDS);
    throw Error(bad);
  }
  cache_().remove(failKey);
  u.lastLogin = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  putUser_(u);
  return {token:openSession_(u), user:{username:u.username, role:u.role}};
}
function logout(token){
  if(token && typeof token==='string' && token.length<=100) cache_().remove('SESS_'+token);
  return true;
}
/** المدخل الوحيد لدوال البيانات من المتصفح */
function rpc(token, name, args){
  CURRENT_ = session_(token);
  if(typeof name!=='string' || !Object.prototype.hasOwnProperty.call(API_, name)) throw Error('عملية غير معروفة');
  return API_[name].apply(null, Array.isArray(args) ? args : []);
}

/* ----------------------------- إدارة المستخدمين ---------------------------- */
function listUsers(){ adminOnly_(); return allUsers_().map(publicUser_); }
/** إضافة أو تعديل مستخدم: {username, password, role, active, isNew} — كلمة المرور اختيارية عند التعديل */
function saveUser(d){
  adminOnly_(); obj_(d);
  const name = checkUsername_(d.username);
  const role = ROLES.indexOf(d.role)>=0 ? d.role : 'user';
  const hasPw = !blank_(d.password);
  if(hasPw) checkPassword_(d.password);
  return locked_(()=>{
    const all = allUsers_(), old = getUser_(name);
    let u;
    if(d.isNew===true){
      if(old) throw Error('اسم المستخدم موجود بالفعل');
      if(!hasPw) throw Error('أدخل كلمة مرور للمستخدم الجديد');
      u = newUser_(name, d.password, role);
    }else{
      if(!old) throw Error('المستخدم غير موجود');
      u = old;
      const active = d.active!==false;
      const others = all.filter(x=>userKey_(x.username)!==userKey_(name));
      if((role!=='admin' || !active) && activeAdmins_(others)===0) throw Error('يجب أن يبقى مدير واحد مفعّل على الأقل');
      if(u.role!==role || (u.active!==false)!==active) u.ver = (Number(u.ver)||1) + 1;
      u.role = role; u.active = active;
      if(hasPw) setPassword_(u, d.password);
    }
    putUser_(u);
    log_(d.isNew===true?'إضافة مستخدم':'تعديل مستخدم', u.username);
    return allUsers_().map(publicUser_);
  });
}
function deleteUser(username){
  adminOnly_(); check_(username);
  return locked_(()=>{
    const u = getUser_(username);
    if(!u) throw Error('المستخدم غير موجود');
    if(userKey_(u.username)===userKey_(CURRENT_.username)) throw Error('لا يمكنك حذف حسابك الحالي');
    props_().deleteProperty(userKey_(username));
    log_('حذف مستخدم', u.username);
    return allUsers_().map(publicUser_);
  });
}
function changeMyPassword(oldPassword, newPassword){
  guard_(); checkPassword_(newPassword);
  return locked_(()=>{
    const u = getUser_(CURRENT_.username);
    if(!u || typeof oldPassword!=='string' || hash_(oldPassword, u.salt)!==u.hash) throw Error('كلمة المرور الحالية غير صحيحة');
    setPassword_(u, newPassword);
    putUser_(u);
    cache_().remove('SESS_'+CURRENT_.token);          // إنهاء كل الجلسات القديمة وفتح جلسة جديدة
    log_('تغيير كلمة المرور', u.username);
    return {token:openSession_(u)};
  });
}

/* --------------------------- الإعداد وواجهة الويب -------------------------- */
function setup(){
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
    user: {username:CURRENT_.username, role:CURRENT_.role}
  };
}

/** قواعد التصنيف الفرعي للبنود بحسب طبيعتها (الأقدم أولوية) — الأوصاف العربية */
const SUB_RULES_AR = [
  ['مكافحة الحريق', /حريق|اطفاء|إطفاء|انذار|إنذار/],
  ['هدم وإزالة', /هدم|إزال|ازال|تكسير|نزع|قلع|فك /],
  ['حفر وردم', /حفر|ردم|دفان|تربة|اعمال ترابية|أعمال ترابية/],
  ['خرسانة وتسليح', /خرسان|تسليح|اسمنت|أسمنت|شدات|اساسات|أساسات|قواعد/],
  ['حدائق وري', /زراع|نبات|نخيل|عشب|شجر|أشجار|اشجار|شبكة الري|شبكه الري|رشاش|تنقيط|(^|\s)(ال)?ري(\s|$)/],
  ['مباني وبياض', /مباني|بلوك|طوب|بناء|لياسة|لياسه|محارة|بياض/],
  ['أبواب ونوافذ', /باب|أبواب|ابواب|نافذ|شباك|ألومنيوم|الومنيوم|زجاج|واجهات/],
  ['أرضيات', /أرضي|ارضي|بلاط|سيراميك|رخام|جرانيت|بورسلان|موزاييك|باركيه|انترلوك|إنترلوك|كربستون/],
  ['حوائط وكسوة', /حائط|حوائط|جدار|قواطع|تكسي|كسوة|كسوه|كلادينج/],
  ['أسقف', /سقف|أسقف|اسقف|سقوف|جبس/],
  ['دهانات وتشطيبات', /دهان|صبغ|معجون|طلاء|بوية|بويه/],
  ['عزل', /عزل|رطوب/],
  ['إنارة', /إنار|انار|إضاء|اضاء|كشاف|لمبة|ثريا/],
  ['تمديدات كهربائية', /كهرب|كابل|كيبل|أسلاك|اسلاك|مفتاح|مقبس|لوحة|لوحه/],
  ['تكييف وتهوية', /تكييف|تبريد|مكيف|دكت|تهوية|تهويه|فريون|شيلر|مروحة|مروحه/],
  ['تغذية وصرف', /صرف|مياه|أنبوب|انبوب|مواسير|خلاط|مغسل|مرحاض|خزان|مضخ|سباك|تغذية|تغذيه/]
];
/** قواعد الأوصاف الإنجليزية (أبواب الكهرباء والإنارة والتيار الخفيف وغيرها) */
const SUB_RULES_EN = [
  ['إنارة الأحجار والأرضيات', /stone|kerb|brick light|inground|in-ground|paver/i],
  ['إنارة تحت الماء', /underwater|submersible|pond/i],
  ['إنارة LED ديكورية', /pixel|rgb|led strip|linear light|led tube|flat led|facade|neon|wall wash/i],
  ['إنارة خارجية وأعمدة', /street|pole|post light|bollard|garden|landscape|flood|spot|projector/i],
  ['إنارة داخلية', /downlight|panel light|recessed|pendant|chandelier|ceiling light|wall light|wall lamp|emergency|exit sign/i],
  ['محولات وأنظمة تحكم', /driver|transformer|dimmer|controller|dmx|power supply|ballast/i],
  ['إنارة', /\blight|lamp|\bled\b|luminaire|fixture/i],
  ['سخانات', /heater|boiler/i],
  ['لوحات توزيع وقواطع', /panel|mdb|smdb|\bdb\b|breaker|mccb|mcb|switchgear|busbar|feeder|isolator/i],
  ['كابلات ومواسير', /cable|wire|conduit|trunking|xlpe|silicone|pipe|tray/i],
  ['ماكينات ومعدات كهربائية', /generator|motor|pump|ups\b|inverter|compressor/i],
  ['تيار خفيف واتصالات', /cctv|camera|telephone|data|network|fire alarm|intercom|access control|speaker|pa system|\btv\b|fiber|fibre|rack/i],
  ['تكييف وتهوية', /air.?condition|hvac|duct|chiller|\bfan\b|split unit|fcu/i],
  ['مكافحة الحريق', /fire|sprinkler|extinguisher/i],
  ['أبواب ونوافذ', /door|window|glass|aluminium|aluminum/i],
  ['دهانات وتشطيبات', /paint|coating|primer/i],
  ['أرضيات', /tile|floor|carpet|marble|granite|turf|rubber/i],
  ['أثاث وتجهيزات', /bench|table|chair|\bbin\b|shade|canopy|furniture|pergola/i]
];
/** كل أسماء التصنيفات التي يولّدها النظام آلياً (لا يُعاد تصنيف بند عدّله المستخدم باسم آخر) */
const AUTO_SUBS = (function(){ const x={'بنود عامة':1,'ألعاب أطفال':1,'معدات رياضية':1,'هدم وإزالة':1,'حفر وردم':1,'خرسانة وتسليح':1,'أبواب ونوافذ':1,'أرضيات':1,'حوائط وكسوة':1,'أسقف':1,'دهانات وتشطيبات':1,'تمديدات كهربائية':1,'إنارة':1,'تكييف وتهوية':1,'تغذية وصرف':1,'عزل':1,'حدائق وري':1,'مكافحة الحريق':1};
  SUB_RULES_AR.concat(SUB_RULES_EN).forEach(r=>{x[r[0]]=1;}); return x; })();
/** التصنيف الفرعي لبند بحسب طبيعته: الباب (للألعاب) ثم الوصف العربي ثم الإنجليزي */
function classifyItem_(chName, desc, detail){
  const text=String(desc||'')+' '+String(detail||'');
  if(/العاب|ألعاب|ملاعب/.test(String(chName))){
    return /fitness|gym|exercise|sport|basketball|football|tennis|volley|goal post|treadmill|رياض/i.test(text) ? 'معدات رياضية' : 'ألعاب أطفال';
  }
  const ar=SUB_RULES_AR.find(r=>r[1].test(text)); if(ar) return ar[0];
  const en=SUB_RULES_EN.find(r=>r[1].test(text)); return en?en[0]:'بنود عامة';
}
function classify_(name,desc){ return classifyItem_(name,desc,''); }

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


/* ---------------- البنود غير المرقّمة في ملف الأسعار الأصلي (v3.4) ---------------- */
/** يحوّل نص سعر قد يحتوي أرقاماً عربية وفواصل ٬ ٫ ورمز العملة إلى رقم (أو -1 إن تعذّر) */
function parsePrice_(raw){
  if(typeof raw==='number') return Number.isFinite(raw)&&raw>=0 ? raw : -1;
  let t=String(raw==null?'':raw).trim();
  if(!t) return -1;
  t=t.replace(/[٠-٩]/g,d=>String(d.charCodeAt(0)-1632)).replace(/[۰-۹]/g,d=>String(d.charCodeAt(0)-1776))
     .replace(/[٬,\s\u200e\u200f]/g,'').replace(/٫/g,'.');
  const m=t.match(/\d+(?:\.\d+)?/);   // أول رقم فقط، فيُهمل رمز العملة «د.ك.»
  const n=m?Number(m[0]):NaN;
  return Number.isFinite(n)&&n>=0 ? n : -1;
}
/** يقرأ من ورقة الأسعار الأصلية البنود التي لا تحمل «رقم بند» (الأوصاف الطويلة بالإنجليزية غالباً) */
function unnumberedFromSource_(){
  const norm=x=>String(x).replace(/[أإآ]/g,'ا').replace(/ى/g,'ي').replace(/\s+/g,'').trim();
  const source=db_().getSheets().find(sh=>norm(sh.getName())===norm('جدول وصف البنود وأسعارها'));
  if(!source) throw Error('لم أجد ورقة «جدول وصف البنود وأسعارها» داخل ملف البرنامج');
  const vals=source.getDataRange().getValues(), out=[], seen=new Set();
  for(let i=1;i<vals.length;i++){
    const r=vals[i], no=clean_(formatItemCode_(r[1])), chNo=clean_(r[2]), chName=clean_(r[3]), desc=clean_(r[4]);
    if(no || !chNo || !chName || !desc || desc==='وصف المادة') continue;   // أي صف له رقم بند يُستورد من importOriginal
    const detail=clean_(r[5]), unit=clean_(r[6]), price=parsePrice_(r[7]);
    const key=chNo+'|'+desc+'|'+detail+'|'+price;
    if(seen.has(key)) continue;   // تكرار حرفي لنفس البند
    seen.add(key);
    out.push({chNo:chNo, chName:chName, desc:desc, detail:detail, unit:unit, price:price<0?0:price});
  }
  return out;
}

/** إعادة تصنيف الدليل بحسب طبيعة البنود (يغيّر عمود «التصنيف الفرعي» فقط، ولا يلمس ما عدّله المستخدم باسم آخر) */
function reclassifyCatalog(){
  adminOnly_();
  return locked_(()=>{
    const sh=tab_(S.CAT), v=sh.getDataRange().getValues();
    if(v.length<2) return {changed:0,total:0};
    const col=[]; let changed=0, total=0;
    for(let i=1;i<v.length;i++){
      const cur=String(v[i][3]||'');
      if(!v[i][0]){ col.push([cur]); continue; }
      total++;
      if(!cur || AUTO_SUBS[cur]){
        const nv=classifyItem_(v[i][2], v[i][5], v[i][6]);
        if(nv!==cur){ changed++; col.push([nv]); continue; }
      }
      col.push([cur]);
    }
    if(changed) sh.getRange(2,4,col.length,1).setValues(col);
    log_('إعادة تصنيف الدليل','عدد البنود المعدّلة: '+changed);
    return {changed:changed,total:total};
  });
}

/** الرقم الرئيسي الذي يحمله رقم بند موجود: «س/باب» ← س ، «ف/س/باب» ← س ، وإلا -1 */
function mainNoOf_(code, chNo){
  const p=String(code||'').split('/');
  if(p.length<2 || p[p.length-1]!==String(chNo)) return -1;
  const m=Number(p.length===2 ? p[0] : p[1]);
  return Number.isFinite(m) ? m : -1;
}
/** إضافة البنود غير المرقّمة داخل أبوابها الأصلية (مع تصنيف فرعي بحسب الطبيعة) بترقيم يتّبع أسلوب الدليل:
 *  رقم رئيسي لكل بند «س/الباب»، وإن تعدّدت مقاساته أو درجاته فترقيم فرعي «ف/س/الباب».
 *  آمنة للتكرار: لا تُضيف ما أُضيف من قبل ولا تغيّر أي بند موجود. */
function importUnnumbered_(){
  const items=unnumberedFromSource_();
  if(!items.length) return {added:0, total:0, skipped:0, message:'لا توجد بنود بدون رقم في ورقة الأسعار'};
  return locked_(()=>{
    const sh=tab_(S.CAT), v=sh.getDataRange().getValues();
    const un=t=>String(t==null?'':t).trim().replace(/^'/,'');   // يُهمل بادئة الحماية من المعادلات
    const keyOf=(chNo,desc,detail,price)=>[chNo,un(desc),un(detail),round3_(price)].join('|');
    const have=new Set(), maxMain={}, chName={}, groups={};   // groups: chNo|وصف ← {main, count}
    for(let i=1;i<v.length;i++){
      const chNo=String(v[i][1]); if(!chNo) continue;
      if(!chName[chNo]) chName[chNo]=String(v[i][2]);
      have.add(keyOf(chNo,v[i][5],v[i][6],Number(v[i][8])||0));
      const m=mainNoOf_(formatItemCode_(v[i][4]), chNo);
      if(m>=0){ maxMain[chNo]=Math.max(maxMain[chNo]||0,m);
        const gk=chNo+'|'+un(v[i][5]); if(!groups[gk]) groups[gk]={main:m,count:0}; groups[gk].count++; }
    }
    // تجميع البنود الجديدة بحسب (الباب + الوصف)
    const fresh=[], order=[], byGroup={};
    items.forEach(x=>{
      const k=keyOf(x.chNo,x.desc,x.detail,x.price);
      if(have.has(k)) return;
      have.add(k); fresh.push(x);
      const gk=x.chNo+'|'+x.desc;
      if(!byGroup[gk]){ byGroup[gk]=[]; order.push(gk); }
      byGroup[gk].push(x);
    });
    const rows=[], bySub={}, perChapter={};
    order.forEach(gk=>{
      const list=byGroup[gk], chNo=list[0].chNo;
      let g=groups[gk], main;
      if(g){ main=g.main; }
      else { main=(maxMain[chNo]||0)+1; maxMain[chNo]=main; g={main:main,count:0}; groups[gk]=g; }
      const multi = list.length>1 || g.count>0;
      list.forEach(x=>{
        g.count++;
        const code = multi ? (g.count+'/'+main+'/'+chNo) : (main+'/'+chNo);
        const chapter = chName[chNo] || x.chName;
        const sub = classifyItem_(chapter, x.desc, x.detail);
        rows.push([id_(), chNo, safe_(chapter), safe_(sub), code, safe_(x.desc), safe_(x.detail), safe_(x.unit), x.price, true]);
        perChapter[chapter]=(perChapter[chapter]||0)+1;
        bySub[chapter+' ← '+sub]=(bySub[chapter+' ← '+sub]||0)+1;
      });
    });
    for(let i=0;i<rows.length;i+=200) writeRows_(sh, nextRow_(sh), rows.slice(i,i+200), TEXT_COLS.CAT);
    if(rows.length) log_('استيراد بنود غير مرقمة','عدد البنود: '+rows.length);
    return {added:rows.length, skipped:items.length-fresh.length, total:items.length, perChapter:perChapter, bySub:bySub};
  });
}
function importUnnumbered(){ adminOnly_(); return importUnnumbered_(); }
/** للتشغيل اليدوي من محرر Apps Script */
function importUnnumberedItems(){ setup(); return importUnnumbered_(); }

/* ------------------------------ دليل البنود ------------------------------- */
function catalogAll_(){
  return rows_(S.CAT).map(r=>({
    id:String(r[0]), chapter:String(r[2]), sub:String(r[3]||'بنود عامة'),
    code:formatItemCode_(r[4]), description:String(r[5]), detail:String(r[6]||''),
    unit:String(r[7]||''), price:Number(r[8])||0, active:r[9]!==false, method:String(r[10]||'')
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
               number_(item.price||0), i>0 ? v[i][9]!==false : true,   // الحفاظ على حالة التفعيل
               safe_(item.method)];
    if(!clean_(v[0][10])) sh.getRange(1,11).setValue('طريقة التنفيذ');
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
    let removed=deleteByProject_(S.LINES, projectId);
    deleteByProject_(S.EXPENSES, projectId);
    deleteByProject_(S.PAYMENTS, projectId);
    deleteContractsByProject_(projectId);
    deleteRows_(shP, ip+1, 1);
    log_('حذف مشروع', name+' — بنود محذوفة: '+removed);
    return true;
  });
}
/** حذف كل صفوف مشروع من ورقة ما (على دفعات متجاورة من أسفل لأعلى) ويعيد عددها */
function deleteByProject_(sheetName, projectId){
  const sh=tab_(sheetName), v=sh.getDataRange().getValues();
  let end=-1, removed=0;
  for(let i=v.length-1;i>=0;i--){
    const hit = i>=1 && String(v[i][1])===String(projectId);
    if(hit){ if(end<0) end=i; continue; }
    if(end>=0){ deleteRows_(sh, i+2, end-i); removed+=end-i; end=-1; }
  }
  return removed;
}
/** حذف عقود المشروع من الورقة مع نقل ملفاتها إلى سلة المهملات */
function deleteContractsByProject_(projectId){
  const sh=tab_(S.CONTRACTS), v=sh.getDataRange().getValues();
  for(let i=v.length-1;i>=1;i--){
    if(String(v[i][1])===String(projectId)){
      const fid=String(v[i][3]||'');
      if(fid){ try{ DriveApp.getFileById(fid).setTrashed(true); }catch(e){} }
      deleteRows_(sh, i+1, 1);
    }
  }
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
  // إجمالي المستحقات والمحصّل من العملاء عبر كل المشاريع
  let dueTotal=0, paidTotal=0;
  rows_(S.PAYMENTS).forEach(r=>{
    const amt=Number(r[5])||0;
    if(String(r[3])==='محصلة') paidTotal+=amt; else dueTotal+=amt;
  });
  totals.receivableDue=dueTotal; totals.receivedPaid=paidTotal;
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
  adminOnly_();
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
  adminOnly_();
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
/** اختياري (للمدير): تحديد مجلد Drive الذي تُحفظ فيه جداول الكميات */
function setOutputFolder(folderId){
  adminOnly_();
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
  const data=projects_();
  console.log('عدد المشروعات المقروءة: '+data.length);
  console.log(JSON.stringify(data.map(p=>({id:p.id, name:p.name}))));
}

/* --------------------- الدوال المتاحة للمتصفح عبر rpc ---------------------- */
const API_ = {
  bootstrap:bootstrap, catalog:catalog, catalogAll:catalogAll, saveCatalogItem:saveCatalogItem,
  setCatalogActive:setCatalogActive, updateCatalogPrice:updateCatalogPrice,
  projects:projects, saveProject:saveProject, setProjectStatus:setProjectStatus,
  deleteProject:deleteProject, duplicateProject:duplicateProject,
  lines:lines, saveLine:saveLine, removeLine:removeLine, duplicateLine:duplicateLine,
  projectSummary:projectSummary, projectsOverview:projectsOverview, getLog:getLog,
  companySettings:companySettings, setCompanyDetails:setCompanyDetails,
  defaultRates:defaultRates, setDefaultRates:setDefaultRates,
  createBOQ:createBOQ, setOutputFolder:setOutputFolder, quoteData:quoteData, costReportData:costReportData,
  listUsers:listUsers, saveUser:saveUser, deleteUser:deleteUser, changeMyPassword:changeMyPassword,
  listExpenses:listExpenses, saveExpense:saveExpense, removeExpense:removeExpense,
  listPayments:listPayments, savePayment:savePayment, removePayment:removePayment,
  listContracts:listContracts, uploadContract:uploadContract, removeContract:removeContract,
  projectAccounts:projectAccounts, importUnnumbered:importUnnumbered, reclassifyCatalog:reclassifyCatalog
};

/* ---------- حسابات المشاريع: مصروفات ودفعات وعقود (v3.3) ---------- */
function financeRows_(sheetName, projectId){
  return rows_(sheetName).filter(r=>String(r[1])===String(projectId));
}
/* ------------------------------- المصروفات --------------------------------- */
function listExpenses(projectId){
  guard_(); check_(projectId);
  return financeRows_(S.EXPENSES, projectId).map(r=>({
    id:String(r[0]), projectId:String(r[1]), date:dateStr_(r[2]), category:String(r[3]||'أخرى'),
    description:String(r[4]||''), amount:Number(r[5])||0, method:String(r[6]||''), notes:String(r[7]||'')
  }));
}
function saveExpense(d){
  guard_(); obj_(d); check_(d.projectId);
  const amount=number_(d.amount); if(amount<=0) throw Error('أدخل مبلغاً أكبر من صفر');
  const cat=EXPENSE_CATS.indexOf(clean_(d.category))>=0?clean_(d.category):'أخرى';
  return locked_(()=>{
    const sh=tab_(S.EXPENSES), v=sh.getDataRange().getValues();
    const id=clean_(d.id)||id_();
    const i=d.id?rowOf_(v,id,d.projectId):-1;
    if(d.id&&i<0) throw Error('السطر غير موجود');
    const row=[id, String(d.projectId), (i>0&&v[i][2])?v[i][2]:new Date(), cat,
               safe_(d.description), amount, safe_(d.method), safe_(d.notes)];
    writeRows_(sh, i>0?i+1:nextRow_(sh), [row]);
    log_(i>0?'تعديل مصروف':'إضافة مصروف', cat+' / '+amount+' د.ك');
    return id;
  });
}
function removeExpense(projectId, id){
  guard_(); check_(projectId); check_(id);
  return locked_(()=>{
    const sh=tab_(S.EXPENSES), i=rowOf_(sh.getDataRange().getValues(), id, projectId);
    if(i<0) throw Error('السطر غير موجود');
    deleteRows_(sh, i+1, 1);
    log_('حذف مصروف', id);
    return true;
  });
}
/* --------------------------- دفعات العميل (التحصيل) ------------------------- */
/** يحوّل نص تاريخ (yyyy-MM-dd أو dd/MM/yyyy) أو Date إلى Date، أو null */
function parseDate_(v){
  if(v instanceof Date && !isNaN(v.getTime())) return v;
  const t=clean_(v); if(!t) return null;
  let m=t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if(m) return new Date(Number(m[1]), Number(m[2])-1, Number(m[3]));
  m=t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if(m) return new Date(Number(m[3]), Number(m[2])-1, Number(m[1]));
  return null;
}
function isoDate_(d){ return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : ''; }
function startOfDay_(d){ return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
/** يضيف عناوين أعمدة الدفعات الجديدة (الاستحقاق/الطريقة/السند) في الأوراق القديمة */
function ensurePaymentColumns_(sh){
  const want=H.PAYMENTS;
  if(sh.getLastColumn()<want.length || clean_(sh.getRange(1,8).getValue())!==want[7]){
    sh.getRange(1,1,1,want.length).setValues([want]);
  }
}
function listPayments(projectId){
  guard_(); check_(projectId);
  return financeRows_(S.PAYMENTS, projectId).map(r=>{
    const dd=parseDate_(r[2]), due=parseDate_(r[7]);
    return {
      id:String(r[0]), projectId:String(r[1]), date:dateStr_(r[2]), dateISO:isoDate_(dd),
      type:String(r[3]||'مستحقة'), description:String(r[4]||''), amount:Number(r[5])||0,
      notes:String(r[6]||''), dueDate:due?dateStr_(due):'', dueISO:isoDate_(due),
      method:String(r[8]||''), ref:String(r[9]||'')
    };
  });
}
function savePayment(d){
  guard_(); obj_(d); check_(d.projectId);
  const amount=number_(d.amount); if(amount<=0) throw Error('أدخل مبلغاً أكبر من صفر');
  const type=PAYMENT_TYPES.indexOf(clean_(d.type))>=0?clean_(d.type):'مستحقة';
  const when=parseDate_(d.date), due=parseDate_(d.dueDate);
  if(clean_(d.dueDate)&&!due) throw Error('تاريخ الاستحقاق غير صحيح');
  if(clean_(d.date)&&!when) throw Error('تاريخ الحركة غير صحيح');
  return locked_(()=>{
    const sh=tab_(S.PAYMENTS); ensurePaymentColumns_(sh);
    const v=sh.getDataRange().getValues();
    const id=clean_(d.id)||id_();
    const i=d.id?rowOf_(v,id,d.projectId):-1;
    if(d.id&&i<0) throw Error('السطر غير موجود');
    const row=[id, String(d.projectId), when||((i>0&&v[i][2])?v[i][2]:new Date()), type,
               safe_(d.description), amount, safe_(d.notes), due||'', safe_(d.method), safe_(d.ref)];
    writeRows_(sh, i>0?i+1:nextRow_(sh), [row]);
    log_(i>0?'تعديل دفعة':'إضافة دفعة', type+' / '+amount+' د.ك');
    return id;
  });
}
function removePayment(projectId, id){
  guard_(); check_(projectId); check_(id);
  return locked_(()=>{
    const sh=tab_(S.PAYMENTS), i=rowOf_(sh.getDataRange().getValues(), id, projectId);
    if(i<0) throw Error('السطر غير موجود');
    deleteRows_(sh, i+1, 1);
    log_('حذف دفعة', id);
    return true;
  });
}
/* -------------------------------- العقود ----------------------------------- */
function listContracts(projectId){
  guard_(); check_(projectId);
  return financeRows_(S.CONTRACTS, projectId).map(r=>({
    id:String(r[0]), projectId:String(r[1]), name:String(r[2]||''), fileId:String(r[3]||''),
    url:String(r[4]||''), date:dateStr_(r[5]), user:String(r[6]||'')
  }));
}
/** رفع نسخة العقد إلى مجلد Drive خاص بالنظام (PDF أو صورة — حتى 10MB تقريباً) */
function uploadContract(projectId, fileName, mimeType, base64Data){
  guard_(); check_(projectId); check_(fileName); check_(base64Data);
  if(String(base64Data).length>14000000) throw Error('حجم الملف كبير؛ الحد الأقصى نحو 10 ميجابايت');
  const mime=clean_(mimeType)||'application/pdf';
  if(!/^(application\/pdf|image\/(png|jpeg|webp))$/.test(mime)) throw Error('العقد يجب أن يكون PDF أو صورة PNG/JPG');
  return locked_(()=>{
    let folder=null;
    const fid=props_().getProperty('CONTRACTS_FOLDER_ID');
    if(fid){ try{ folder=DriveApp.getFolderById(fid); }catch(e){ folder=null; } }
    if(!folder){
      folder=DriveApp.createFolder('عقود مشاريع منارات المتحدة');
      props_().setProperty('CONTRACTS_FOLDER_ID', folder.getId());
    }
    const blob=Utilities.newBlob(Utilities.base64Decode(base64Data), mime, clean_(fileName));
    const file=folder.createFile(blob);
    try{ file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(e){}
    const id=id_();
    writeRows_(tab_(S.CONTRACTS), nextRow_(tab_(S.CONTRACTS)),
      [[id, String(projectId), safe_(fileName), file.getId(), file.getUrl(), new Date(), who_()]]);
    log_('رفع عقد', fileName);
    return {id:id, url:file.getUrl()};
  });
}
function removeContract(projectId, contractId){
  guard_(); check_(projectId); check_(contractId);
  return locked_(()=>{
    const sh=tab_(S.CONTRACTS), v=sh.getDataRange().getValues(), i=rowOf_(v, contractId, projectId);
    if(i<0) throw Error('العقد غير موجود');
    const fid=String(v[i][3]||'');
    const name=String(v[i][2]||'');
    deleteRows_(sh, i+1, 1);
    if(fid){ try{ DriveApp.getFileById(fid).setTrashed(true); }catch(e){} }
    log_('حذف عقد', name);
    return true;
  });
}
/* --------------------- ملخص حسابات المشروع (شاشة الحسابات) ------------------ */
/** جدول الاستحقاقات: يوزّع المحصَّل على المستحقات بالأقدم استحقاقاً (FIFO) ويحدد المتأخر */
function dueSchedule_(pays){
  const t0=startOfDay_(new Date()).getTime(), DAY=86400000;
  const dues=pays.filter(x=>x.type!=='محصلة').map(x=>{
    const d=parseDate_(x.dueISO||x.dateISO);
    return Object.assign({}, x, {_t:d?d.getTime():0});
  }).sort((a,b)=>a._t-b._t);
  let pool=pays.filter(x=>x.type==='محصلة').reduce((a,x)=>a+x.amount,0);
  const out=dues.map(x=>{
    const settled=Math.min(pool, x.amount); pool=round3_(pool-settled);
    const remaining=round3_(x.amount-settled);
    let status, daysLate=0, daysLeft=0;
    if(remaining<=0.0005) status='مسددة';
    else if(x._t && x._t<t0){ status='متأخرة'; daysLate=Math.round((t0-x._t)/DAY); }
    else { status=settled>0?'جزئية':'مستحقة'; if(x._t) daysLeft=Math.round((x._t-t0)/DAY); }
    return {id:x.id, date:x.date, dueDate:x.dueDate||x.date, description:x.description, amount:x.amount,
            settled:round3_(settled), remaining:remaining, status:status, daysLate:daysLate, daysLeft:daysLeft,
            partialLate:(remaining>0.0005 && x._t && x._t<t0 && settled>0)};
  });
  return {list:out, credit:pool};
}
/** كشف حساب العميل بنفس تنسيق كشف الحساب في برنامج حسابات المكتب: مدين/دائن/رصيد تراكمي */
function clientStatement_(pays){
  const rows=pays.map(x=>{
    const d=parseDate_(x.dateISO);
    return {t:d?d.getTime():0, x:x};
  }).sort((a,b)=>a.t-b.t || (a.x.type==='محصلة'?1:0)-(b.x.type==='محصلة'?1:0));
  let bal=0;
  return rows.map(r=>{
    const x=r.x, debit=x.type==='محصلة'?0:x.amount, credit=x.type==='محصلة'?x.amount:0;
    bal=round3_(bal+debit-credit);
    return {date:x.date, ref:x.ref||'', kind:x.type==='محصلة'?'سداد':'استحقاق',
            description:x.description+(x.type!=='محصلة'&&x.dueDate?' — مستحقة بتاريخ '+x.dueDate:''),
            method:x.method, debit:debit, credit:credit, balance:bal, id:x.id};
  });
}
function projectAccounts(projectId){
  guard_(); check_(projectId);
  const p=projects_().find(x=>x.id===String(projectId));
  if(!p) throw Error('المشروع غير موجود');
  const ls=lines_(projectId);
  const s=summary_(p, ls);
  const ex=listExpenses(projectId), pays=listPayments(projectId);
  const byCat={}; EXPENSE_CATS.forEach(c=>{byCat[c]=0;});
  let spent=0;
  ex.forEach(x=>{ spent+=x.amount; byCat[x.category]=round3_((byCat[x.category]||0)+x.amount); });
  let due=0, paid=0;
  pays.forEach(x=>{ if(x.type==='محصلة') paid+=x.amount; else due+=x.amount; });
  const sched=dueSchedule_(pays);
  const overdue=sched.list.filter(x=>x.status==='متأخرة');
  const overdueAmt=round3_(overdue.reduce((a,x)=>a+x.remaining,0));
  const upcoming=sched.list.filter(x=>x.status==='مستحقة'||x.status==='جزئية');
  const upcomingAmt=round3_(upcoming.reduce((a,x)=>a+x.remaining,0));
  const chap={};
  ls.forEach(x=>{ const c=chap[x.chapter]||(chap[x.chapter]={chapter:x.chapter,count:0,sales:0,cost:0});
    c.count++; c.sales=round3_(c.sales+x.total); c.cost=round3_(c.cost+x.costTotal); });
  const contractValue=s.final;
  return {
    project:p, summary:s,
    byChapter:Object.keys(chap).map(k=>chap[k]),
    expenses:{total:spent, byCategory:byCat, count:ex.length, list:ex},
    payments:{due:due, paid:paid, count:pays.length, list:pays},
    schedule:sched.list,
    overdue:{count:overdue.length, amount:overdueAmt, maxDays:overdue.reduce((a,x)=>Math.max(a,x.daysLate),0)},
    upcoming:{count:upcoming.length, amount:upcomingAmt},
    statement:clientStatement_(pays),
    contracts:listContracts(projectId),
    accounts:{
      contractValue:contractValue,          // قيمة المقايسة النهائية (العقد التقديري)
      totalReceived:paid,                   // إجمالي المحصَّل من العميل
      totalDue:due,                         // إجمالي الدفعات المستحقة (المطالَب بها)
      dueBalance:round3_(Math.max(due-paid,0)), // المستحق غير المسدَّد فعلاً
      notYetDue:round3_(contractValue-due), // جزء العقد الذي لم تُصدر له مطالبة بعد
      outstanding:round3_(contractValue-paid), // المتبقي على العميل
      expectedNet:round3_(contractValue-spent), // صافي الربح المتوقع بعد المصروفات
      costVsSpent:round3_(s.cost-spent),    // فرق تكلفة المقايسة عن المصروف الفعلي
      cashPosition:round3_(paid-spent)      // السيولة: المحصَّل - المصروف
    }
  };
}
