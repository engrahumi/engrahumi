/**
 * نظام الكوزات الإلكترونية
 * ------------------------------------------------------------
 * - كل طالب ياخذ أسئلة مختلفة من بنك الأسئلة وبترتيب مختلف.
 * - خيارات كل سؤال تتخلط لكل طالب.
 * - لكل سؤال وقت محدد، والوقت يُحسب على السيرفر (ما يتلاعب بيه الطالب).
 * - السيرفر يرسل سؤال واحد بس بكل مرة، والجواب الصحيح ما يوصل للمتصفح أبداً.
 * - تسجيل مرات خروج الطالب من صفحة الكوز.
 *
 * لا تحتاج تعدّل أي شي بهذا الملف.
 */

// مهلة إضافية (بالملي ثانية) لتعويض بطء الإنترنت عند انتهاء وقت السؤال
var GRACE_MS = 8000;

var SH = { Q: 'الأسئلة', QZ: 'الكوزات', ST: 'الطلاب', AT: 'المحاولات' };

var HEAD = {
  Q: ['المعرف', 'نوع السؤال', 'السؤال', 'الخيار أ', 'الخيار ب', 'الخيار ج', 'الخيار د',
      'الجواب الصحيح', 'الصعوبة', 'الموضوع', 'الوقت (ثانية)'],
  QZ: ['الرمز', 'اسم الكوز', 'عدد السهل', 'عدد المتوسط', 'عدد الصعب', 'الموضوع',
       'وقت السؤال الافتراضي (ثانية)', 'يفتح', 'يغلق', 'إظهار الدرجة للطالب', 'مفعّل', 'تاريخ الإنشاء'],
  ST: ['رقم الطالب', 'اسم الطالب', 'رمز الدخول'],
  AT: ['المعرف', 'رمز الكوز', 'رقم الطالب', 'اسم الطالب', 'وقت البدء', 'وقت الانتهاء', 'الحالة',
       'الدرجة', 'عدد الأسئلة', 'النسبة %', 'مرات مغادرة الصفحة', 'متوسط وقت الإجابة (ثانية)', 'التفاصيل']
};

var DIFFS = ['سهل', 'متوسط', 'صعب'];
var LETTERS = ['أ', 'ب', 'ج', 'د'];
var TF = 'صح وخطأ';
var MC = 'اختيار';
var DONE = 'منتهي';
var RUNNING = 'جاري';

// ============================================================
// نقطة الدخول للموقع
// ============================================================
function doGet(e) {
  var p = (e && e.parameter) || {};
  var t = HtmlService.createTemplateFromFile('Index');
  t.mode = p.admin ? 'admin' : 'student';
  t.quiz = normCode_(p.quiz);
  return t.evaluate()
    .setTitle(p.admin ? 'لوحة الأستاذ - نظام الكوزات' : 'الكوز الإلكتروني')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// قائمة داخل Google Sheets
function onOpen() {
  SpreadsheetApp.getUi().createMenu('نظام الكوزات')
    .addItem('تجهيز الأوراق', 'setupSheets')
    .addItem('إعادة تعيين كلمة مرور الأستاذ', 'resetAdminPassword')
    .addToUi();
}

function setupSheets() {
  Object.keys(SH).forEach(sheet_);
  SpreadsheetApp.getUi().alert('تم تجهيز الأوراق بنجاح.');
}

function resetAdminPassword() {
  PropertiesService.getScriptProperties().deleteProperty('ADMIN_HASH');
  SpreadsheetApp.getUi().alert('تم حذف كلمة المرور. افتح رابط لوحة الأستاذ وأنشئ كلمة مرور جديدة.');
}

// ============================================================
// أدوات مساعدة
// ============================================================
function err_(msg) { throw new Error('§' + msg); }

function sheet_(key) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH[key]);
  if (!sh) {
    sh = ss.insertSheet(SH[key]);
    var h = HEAD[key];
    sh.getRange(1, 1, 1, h.length).setValues([h])
      .setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    // نخلي الأعمدة نصية حتى ما تتحول الأسئلة اللي تبدي بـ = إلى معادلات، ولا تنحذف الأصفار من الأرقام
    if (key === 'Q' || key === 'ST') sh.getRange(2, 1, sh.getMaxRows() - 1, h.length).setNumberFormat('@');
    try { sh.setRightToLeft(true); } catch (e) { /* غير مدعوم */ }
  }
  return sh;
}

function rows_(key) {
  var sh = sheet_(key);
  var n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, HEAD[key].length).getValues();
}

function str_(v) { return v === null || v === undefined ? '' : String(v).trim(); }

function ms_(v) {
  if (v instanceof Date) return v.getTime();
  if (v === '' || v === null || v === undefined) return null;
  var n = Number(v);
  if (!isNaN(n) && n > 1e11) return n;
  var d = new Date(v);
  return isNaN(d.getTime()) ? null : d.getTime();
}

function bool_(v) { return v === true || /^(true|نعم|1)$/i.test(str_(v)); }

function normCode_(c) { return str_(c).toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 10); }

function normName_(s) {
  return str_(s).replace(/[ً-ْ]/g, '').replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه')
    .replace(/ى/g, 'ي').replace(/\s+/g, ' ').toLowerCase();
}

function shuffle_(arr) {
  var a = arr.slice();
  for (var i = a.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var t = a[i]; a[i] = a[j]; a[j] = t;
  }
  return a;
}

function randCode_(alphabet, len) {
  var s = '';
  for (var i = 0; i < len; i++) s += alphabet.charAt(Math.floor(Math.random() * alphabet.length));
  return s;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// ============================================================
// صلاحيات الأستاذ
// ============================================================
function hash_(s) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'quiz-system:' + s, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(bytes);
}

function adminStatus() {
  return { hasPassword: !!PropertiesService.getScriptProperties().getProperty('ADMIN_HASH') };
}

function adminSetPassword(pw) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('ADMIN_HASH')) err_('كلمة المرور موجودة مسبقاً.');
  pw = str_(pw);
  if (pw.length < 4) err_('كلمة المرور لازم تكون 4 أحرف على الأقل.');
  props.setProperty('ADMIN_HASH', hash_(pw));
  Object.keys(SH).forEach(sheet_);
  return adminGetDashboard(pw);
}

function auth_(pw) {
  var h = PropertiesService.getScriptProperties().getProperty('ADMIN_HASH');
  if (!h) err_('لم يتم إنشاء كلمة مرور بعد.');
  if (hash_(str_(pw)) !== h) {
    Utilities.sleep(800);
    err_('كلمة المرور غير صحيحة.');
  }
}

// ============================================================
// بنك الأسئلة
// ============================================================
function readBank_() {
  return rows_('Q').filter(function (r) { return str_(r[2]); }).map(function (r) {
    var opts = [r[3], r[4], r[5], r[6]].map(str_);
    var ans = LETTERS.indexOf(str_(r[7]).replace('ا', 'أ'));
    var time = Number(r[10]);
    return {
      id: str_(r[0]),
      type: str_(r[1]) === TF ? TF : MC,
      text: str_(r[2]),
      opts: opts,
      ans: ans,
      diff: DIFFS.indexOf(str_(r[8])) >= 0 ? str_(r[8]) : 'متوسط',
      topic: str_(r[9]),
      time: time >= 5 ? Math.round(time) : null
    };
  }).filter(function (q) {
    return q.ans >= 0 && q.opts[q.ans] && q.opts.filter(Boolean).length >= 2;
  });
}

function bankStats_(bank) {
  var stats = { total: bank.length, byDiff: { 'سهل': 0, 'متوسط': 0, 'صعب': 0 }, topics: {} };
  bank.forEach(function (q) {
    stats.byDiff[q.diff]++;
    var t = q.topic || '';
    if (!stats.topics[t]) stats.topics[t] = { 'سهل': 0, 'متوسط': 0, 'صعب': 0 };
    stats.topics[t][q.diff]++;
  });
  return stats;
}

function adminUploadQuestions(pw, items, replace) {
  auth_(pw);
  if (!Array.isArray(items) || !items.length) err_('لا توجد أسئلة للرفع.');
  if (items.length > 5000) err_('الحد الأقصى 5000 سؤال بالمرة الوحدة.');
  var base = 'Q' + Date.now().toString(36).toUpperCase();
  var out = items.map(function (it, i) {
    var n = i + 1;
    var type = it.type === TF ? TF : MC;
    var text = str_(it.text);
    if (!text) err_('السؤال رقم ' + n + ' فارغ.');
    var opts = (it.opts || []).slice(0, 4).map(str_);
    while (opts.length < 4) opts.push('');
    if (opts.filter(Boolean).length < 2) err_('السؤال رقم ' + n + ' يحتاج خيارين على الأقل.');
    var ans = Number(it.ans);
    if (!(ans >= 0 && ans < 4 && opts[ans])) err_('الجواب الصحيح للسؤال رقم ' + n + ' غير صحيح.');
    if (DIFFS.indexOf(it.diff) < 0) err_('صعوبة السؤال رقم ' + n + ' غير صحيحة.');
    var time = '';
    if (it.time !== '' && it.time !== null && it.time !== undefined) {
      time = Math.round(Number(it.time));
      if (!(time >= 5 && time <= 600)) err_('وقت السؤال رقم ' + n + ' لازم يكون بين 5 و 600 ثانية.');
    }
    return [base + '-' + n, type, text, opts[0], opts[1], opts[2], opts[3],
            LETTERS[ans], it.diff, str_(it.topic), time === '' ? '' : String(time)];
  });
  withLock_(function () {
    var sh = sheet_('Q');
    var last = sh.getLastRow();
    if (replace && last > 1) sh.getRange(2, 1, last - 1, HEAD.Q.length).clearContent();
    var start = replace ? 2 : Math.max(sh.getLastRow(), 1) + 1;
    var rng = sh.getRange(start, 1, out.length, HEAD.Q.length);
    rng.setNumberFormat('@');
    rng.setValues(out);
  });
  return adminGetDashboard(pw);
}

// ============================================================
// الكوزات
// ============================================================
function readQuizzes_() {
  return rows_('QZ').filter(function (r) { return str_(r[0]); }).map(function (r) {
    return {
      code: str_(r[0]),
      name: str_(r[1]),
      easy: Number(r[2]) || 0,
      med: Number(r[3]) || 0,
      hard: Number(r[4]) || 0,
      topic: str_(r[5]),
      defTime: Number(r[6]) >= 5 ? Number(r[6]) : 30,
      opens: ms_(r[7]),
      closes: ms_(r[8]),
      showScore: bool_(r[9]),
      active: bool_(r[10]),
      created: ms_(r[11])
    };
  });
}

function findQuiz_(code) {
  code = normCode_(code);
  if (!code) return null;
  var list = readQuizzes_();
  for (var i = 0; i < list.length; i++) if (list[i].code === code) return list[i];
  return null;
}

function adminCreateQuiz(pw, q) {
  auth_(pw);
  q = q || {};
  var name = str_(q.name);
  if (!name) err_('اكتب اسم الكوز.');
  var easy = Math.max(0, parseInt(q.easy, 10) || 0);
  var med = Math.max(0, parseInt(q.med, 10) || 0);
  var hard = Math.max(0, parseInt(q.hard, 10) || 0);
  if (easy + med + hard < 1) err_('لازم يكون بالكوز سؤال واحد على الأقل.');
  if (easy + med + hard > 200) err_('الحد الأقصى 200 سؤال بالكوز.');
  var defTime = parseInt(q.defTime, 10);
  if (!(defTime >= 5 && defTime <= 600)) err_('وقت السؤال لازم يكون بين 5 و 600 ثانية.');
  var opens = q.opens ? Number(q.opens) : null;
  var closes = q.closes ? Number(q.closes) : null;
  if (opens && closes && closes <= opens) err_('وقت الغلق لازم يكون بعد وقت الفتح.');
  var topic = str_(q.topic);

  var stats = bankStats_(readBank_());
  var avail = topic ? stats.topics[topic] : stats.byDiff;
  if (!avail) err_('هذا الموضوع غير موجود في بنك الأسئلة.');
  [['سهل', easy], ['متوسط', med], ['صعب', hard]].forEach(function (p) {
    if (p[1] > avail[p[0]]) err_('عدد الأسئلة (' + p[0] + ') المطلوبة ' + p[1] + ' والمتوفر بالبنك ' + avail[p[0]] + ' بس.');
  });

  return withLock_(function () {
    var existing = readQuizzes_().map(function (x) { return x.code; });
    var code;
    do { code = randCode_('ABCDFGHJKMNPQRSTUVWXYZ23456789', 6); } while (existing.indexOf(code) >= 0);
    var sh = sheet_('QZ');
    var row = Math.max(sh.getLastRow(), 1) + 1;
    var rng = sh.getRange(row, 1, 1, HEAD.QZ.length);
    rng.setNumberFormat('@');
    sh.getRange(row, 8, 1, 2).setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange(row, 12).setNumberFormat('yyyy-mm-dd hh:mm');
    rng.setValues([[code, name, String(easy), String(med), String(hard), topic, String(defTime),
                    opens ? new Date(opens) : '', closes ? new Date(closes) : '',
                    q.showScore ? 'نعم' : 'لا', 'نعم', new Date()]]);
    return adminGetDashboard(pw);
  });
}

function quizRow_(code) {
  var sh = sheet_('QZ');
  var n = sh.getLastRow();
  if (n < 2) return -1;
  var codes = sh.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < codes.length; i++) if (str_(codes[i][0]) === code) return i + 2;
  return -1;
}

function adminSetQuizActive(pw, code, active) {
  auth_(pw);
  var row = quizRow_(normCode_(code));
  if (row < 0) err_('الكوز غير موجود.');
  sheet_('QZ').getRange(row, 11).setValue(active ? 'نعم' : 'لا');
  return adminGetDashboard(pw);
}

function adminDeleteQuiz(pw, code) {
  auth_(pw);
  withLock_(function () {
    var row = quizRow_(normCode_(code));
    if (row < 0) err_('الكوز غير موجود.');
    sheet_('QZ').deleteRow(row);
  });
  return adminGetDashboard(pw);
}

// ============================================================
// الطلاب (اختياري) — إذا رفعت قائمة طلاب، الدخول يصير برقم الطالب ورمز الدخول
// ============================================================
function readStudents_() {
  return rows_('ST').filter(function (r) { return str_(r[0]); }).map(function (r) {
    return { id: str_(r[0]), name: str_(r[1]), code: str_(r[2]) };
  });
}

function adminGetStudents(pw) {
  auth_(pw);
  return readStudents_();
}

function adminUploadStudents(pw, items, replace) {
  auth_(pw);
  items = Array.isArray(items) ? items : [];
  if (!replace && !items.length) err_('لا يوجد طلاب للرفع.');
  return withLock_(function () {
    var list = replace ? [] : readStudents_();
    var byId = {};
    list.forEach(function (s) { byId[s.id] = s; });
    items.forEach(function (it, i) {
      var id = str_(it.id), name = str_(it.name);
      if (!id) err_('الطالب بالسطر ' + (i + 1) + ' ماعنده رقم.');
      if (!name) err_('الطالب بالسطر ' + (i + 1) + ' ماعنده اسم.');
      if (byId[id]) { byId[id].name = name; return; }
      var s = { id: id, name: name, code: '' };
      byId[id] = s;
      list.push(s);
    });
    var used = {};
    list.forEach(function (s) { if (s.code) used[s.code] = 1; });
    list.forEach(function (s) {
      if (s.code) return;
      var c;
      do { c = randCode_('0123456789', 6); } while (used[c]);
      used[c] = 1;
      s.code = c;
    });
    var sh = sheet_('ST');
    var last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, 3).clearContent();
    if (list.length) {
      var rng = sh.getRange(2, 1, list.length, 3);
      rng.setNumberFormat('@');
      rng.setValues(list.map(function (s) { return [s.id, s.name, s.code]; }));
    }
    return list;
  });
}

// ============================================================
// اللوحة الرئيسية للأستاذ
// ============================================================
function adminGetDashboard(pw) {
  auth_(pw);
  var counts = {};
  var sh = sheet_('AT');
  var n = sh.getLastRow();
  if (n > 1) {
    sh.getRange(2, 2, n - 1, 6).getValues().forEach(function (r) {
      var c = str_(r[0]);
      if (!counts[c]) counts[c] = { attempts: 0, finished: 0 };
      counts[c].attempts++;
      if (str_(r[5]) === DONE) counts[c].finished++;
    });
  }
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { /* قبل النشر */ }
  return {
    url: url,
    now: Date.now(),
    bank: bankStats_(readBank_()),
    students: readStudents_().length,
    quizzes: readQuizzes_().map(function (q) {
      var c = counts[q.code] || { attempts: 0, finished: 0 };
      q.attempts = c.attempts;
      q.finished = c.finished;
      return q;
    }).reverse()
  };
}

// ============================================================
// النتائج
// ============================================================
function adminGetResults(pw, code) {
  auth_(pw);
  code = normCode_(code);
  return rows_('AT').filter(function (r) { return str_(r[1]) === code; }).map(function (r) {
    var unanswered = 0;
    try {
      var d = JSON.parse(r[12]);
      d.a.forEach(function (a) { if (a && a.ch === null) unanswered++; });
    } catch (e) { /* تجاهل */ }
    return {
      id: str_(r[0]), sid: str_(r[2]), name: str_(r[3]),
      start: ms_(r[4]), end: ms_(r[5]), status: str_(r[6]),
      score: r[7] === '' ? null : Number(r[7]), total: Number(r[8]) || 0,
      pct: r[9] === '' ? null : Number(r[9]),
      blur: Number(r[10]) || 0,
      avg: r[11] === '' ? null : Number(r[11]),
      unanswered: unanswered
    };
  });
}

function adminGetAttemptDetail(pw, id) {
  auth_(pw);
  var ctx = loadAttempt_(id);
  return {
    name: ctx.name, sid: ctx.sid,
    items: ctx.d.qs.map(function (q, i) {
      var a = ctx.d.a[i];
      return {
        text: q.t, options: q.o, correct: q.c,
        chosen: a ? a.ch : null,
        seconds: a && a.ms !== null && a.ms !== undefined ? Math.round(a.ms / 100) / 10 : null,
        reached: !!a || i === ctx.d.i
      };
    })
  };
}

function adminResetAttempt(pw, id) {
  auth_(pw);
  withLock_(function () {
    var row = attemptRow_(id);
    if (row < 0) err_('المحاولة غير موجودة.');
    sheet_('AT').deleteRow(row);
  });
  return true;
}

// ============================================================
// واجهة الطالب
// ============================================================
function quizWindowError_(quiz, now) {
  if (!quiz) return 'رمز الكوز غير صحيح.';
  if (!quiz.active) return 'هذا الكوز غير متاح حالياً.';
  if (quiz.opens && now < quiz.opens) return 'الكوز لم يبدأ بعد.';
  if (quiz.closes && now >= quiz.closes) return 'انتهى وقت هذا الكوز.';
  return '';
}

function studentGetQuizInfo(code) {
  var quiz = findQuiz_(code);
  if (!quiz) err_('رمز الكوز غير صحيح.');
  var now = Date.now();
  return {
    code: quiz.code, name: quiz.name,
    count: quiz.easy + quiz.med + quiz.hard,
    defTime: quiz.defTime,
    opens: quiz.opens, closes: quiz.closes, now: now,
    needsCode: readStudents_().length > 0,
    problem: quizWindowError_(quiz, now)
  };
}

function buildQuestions_(quiz) {
  var bank = readBank_().filter(function (q) { return !quiz.topic || q.topic === quiz.topic; });
  var picked = [];
  [['سهل', quiz.easy], ['متوسط', quiz.med], ['صعب', quiz.hard]].forEach(function (p) {
    if (!p[1]) return;
    var pool = shuffle_(bank.filter(function (q) { return q.diff === p[0]; }));
    if (pool.length < p[1]) err_('لا توجد أسئلة كافية في البنك لهذا الكوز. راجع الأستاذ.');
    picked = picked.concat(pool.slice(0, p[1]));
  });
  return shuffle_(picked).map(function (q) {
    var idx = [];
    q.opts.forEach(function (o, i) { if (o) idx.push(i); });
    if (q.type !== TF) idx = shuffle_(idx);
    return {
      id: q.id, t: q.text, y: q.type,
      o: idx.map(function (i) { return q.opts[i]; }),
      c: idx.indexOf(q.ans),
      tm: q.time || quiz.defTime
    };
  });
}

function studentStart(code, sid, name, access) {
  var quiz = findQuiz_(code);
  var now = Date.now();
  var problem = quizWindowError_(quiz, now);
  sid = str_(sid).slice(0, 50);
  name = str_(name).slice(0, 100);
  access = str_(access);

  var students = readStudents_();
  if (students.length) {
    var s = null;
    for (var i = 0; i < students.length; i++) if (students[i].id === sid) { s = students[i]; break; }
    if (!s || s.code !== access) err_('رقم الطالب أو رمز الدخول غير صحيح.');
    name = s.name;
  } else {
    if (!sid) err_('اكتب رقم الطالب.');
    if (name.length < 3) err_('اكتب اسمك الكامل.');
  }

  var token = withLock_(function () {
    var sh = sheet_('AT');
    var n = sh.getLastRow();
    if (n > 1) {
      var rows = sh.getRange(2, 1, n - 1, 7).getValues();
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        if (str_(r[1]) !== quiz.code || str_(r[2]) !== sid) continue;
        if (str_(r[6]) === DONE) err_('لقد أنهيت هذا الكوز مسبقاً، لا يمكن الدخول مرة ثانية.');
        if (!students.length && normName_(r[3]) !== normName_(name)) err_('رقم الطالب هذا مستخدم من طالب آخر.');
        return str_(r[0]); // يكمل من وين ما وقف
      }
    }
    if (problem) err_(problem);
    var d = { qs: buildQuestions_(quiz), i: 0, s: null, a: [] };
    if (JSON.stringify(d).length > 45000) err_('أسئلة هذا الكوز طويلة جداً. قلل عدد الأسئلة.');
    var id = Utilities.getUuid().replace(/-/g, '');
    var row = Math.max(sh.getLastRow(), 1) + 1;
    var rng = sh.getRange(row, 1, 1, HEAD.AT.length);
    rng.setNumberFormat('@');
    sh.getRange(row, 5, 1, 2).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    rng.setValues([[id, quiz.code, sid, name, new Date(now), '', RUNNING, '', String(d.qs.length), '', '0', '', JSON.stringify(d)]]);
    return id;
  });
  return { token: token, name: name, state: studentState(token) };
}

function studentState(token) {
  return withAttempt_(token, null);
}

function studentAnswer(token, index, choice) {
  return withAttempt_(token, function (ctx, quiz, now) {
    var d = ctx.d;
    if (Number(index) !== d.i || d.s === null || d.i >= d.qs.length) return;
    var q = d.qs[d.i];
    if (now > d.s + q.tm * 1000 + GRACE_MS) return;
    if (quiz && quiz.closes && now > quiz.closes + GRACE_MS) return;
    var ch = choice === null || choice === undefined || choice === '' ? null : Number(choice);
    if (ch !== null && !(ch % 1 === 0 && ch >= 0 && ch < q.o.length)) ch = null;
    d.a[d.i] = { ch: ch, ms: now - d.s };
    d.i++;
    d.s = null;
  });
}

function studentBlur(token) {
  var row = attemptRow_(token);
  if (row < 0) return 0;
  var sh = sheet_('AT');
  var vals = sh.getRange(row, 7, 1, 5).getValues()[0];
  var count = Number(vals[4]) || 0;
  if (str_(vals[0]) === DONE) return count;
  count++;
  sh.getRange(row, 11).setValue(String(count));
  return count;
}

// ============================================================
// منطق المحاولة والوقت
// ============================================================
function attemptRow_(token) {
  token = str_(token);
  if (!/^[a-f0-9]{32}$/.test(token)) return -1;
  var sh = sheet_('AT');
  var n = sh.getLastRow();
  if (n < 2) return -1;
  var cell = sh.getRange(2, 1, n - 1, 1).createTextFinder(token).matchEntireCell(true).findNext();
  return cell ? cell.getRow() : -1;
}

function loadAttempt_(token) {
  var row = attemptRow_(token);
  if (row < 0) err_('الجلسة غير موجودة. ادخل للكوز مرة ثانية.');
  var r = sheet_('AT').getRange(row, 1, 1, HEAD.AT.length).getValues()[0];
  return {
    row: row, id: str_(r[0]), quiz: str_(r[1]), sid: str_(r[2]), name: str_(r[3]),
    status: str_(r[6]), score: r[7] === '' ? null : Number(r[7]),
    blur: Number(r[10]) || 0, d: JSON.parse(r[12])
  };
}

// يتجاوز الأسئلة اللي خلص وقتها، ويبدأ وقت السؤال الجاي لحظة إرساله للطالب
function progress_(d, quiz, now) {
  var close = quiz && quiz.closes ? quiz.closes : Infinity;
  while (d.i < d.qs.length) {
    var q = d.qs[d.i];
    if (d.s === null) {
      if (now >= close) { d.a[d.i] = { ch: null, ms: null }; d.i++; continue; }
      d.s = now;
      return;
    }
    if (now > d.s + q.tm * 1000 + GRACE_MS || now > close + GRACE_MS) {
      d.a[d.i] = { ch: null, ms: null };
      d.i++;
      d.s = null;
      continue;
    }
    return;
  }
}

function withAttempt_(token, mutate) {
  var ctx = loadAttempt_(token);
  var quiz = findQuiz_(ctx.quiz);
  var now = Date.now();
  if (ctx.status !== DONE) {
    var before = ctx.d.i + '|' + ctx.d.s;
    if (mutate) mutate(ctx, quiz, now);
    progress_(ctx.d, quiz, now);
    if (ctx.d.i + '|' + ctx.d.s !== before) save_(ctx);
  }
  return view_(ctx, quiz, now);
}

function save_(ctx) {
  var d = ctx.d;
  var sh = sheet_('AT');
  var done = d.i >= d.qs.length;
  var score = 0, sum = 0, answered = 0;
  d.qs.forEach(function (q, i) {
    var a = d.a[i];
    if (!a) return;
    if (a.ch !== null && a.ch === q.c) score++;
    if (a.ch !== null) { sum += a.ms; answered++; }
  });
  var avg = answered ? String(Math.round(sum / answered / 100) / 10) : '';
  if (done) {
    ctx.status = DONE;
    ctx.score = score;
    sh.getRange(ctx.row, 6, 1, 5).setValues([[new Date(), DONE, String(score), String(d.qs.length),
      String(Math.round(score * 1000 / d.qs.length) / 10)]]);
  }
  sh.getRange(ctx.row, 12, 1, 2).setValues([[avg, JSON.stringify(d)]]);
}

function view_(ctx, quiz, now) {
  var d = ctx.d;
  if (ctx.status === DONE || d.i >= d.qs.length) {
    var r = { status: 'done', total: d.qs.length };
    if (quiz && quiz.showScore) r.score = ctx.score;
    return r;
  }
  var q = d.qs[d.i];
  return {
    status: 'question', index: d.i, total: d.qs.length, blur: ctx.blur,
    question: {
      text: q.t, type: q.y, options: q.o, time: q.tm,
      remainingMs: Math.max(0, d.s + q.tm * 1000 - now)
    }
  };
}
