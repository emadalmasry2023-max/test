/**
 * منصة المهارات الرقمية - Google Apps Script
 * 1) أنشئ Google Sheet فارغاً.
 * 2) من Extensions > Apps Script الصق هذا الملف.
 * 3) شغّل setup() مرة واحدة ووافق على الصلاحيات.
 * 4) شغّل setup()؛ سيتم ضبط TEACHER_PASSWORD تلقائياً على كلمة المرور المتفق عليها.
 * 5) Deploy > New deployment > Web app
 *    Execute as: Me | Who has access: Anyone
 * 6) انسخ رابط Web App وضعه في CONFIG.API_URL داخل index.html.
 */
const SHEET_NAME = "نتائج الاختبارات";
const TEACHER_DEFAULT_PASSWORD = "01004422591";
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(["وقت التسجيل","اسم الطالب","الفصل","البريد الإلكتروني","الدرجة","الدرجة الكلية","الصحيحة","الخاطئة","وقت البداية","وقت النهاية","المدة بالثواني","تفاصيل الإجابات"]);
    sh.setFrozenRows(1);
  }
  // تعيين كلمة مرور المعلم المطلوبة في Script Properties عند تشغيل setup.
  PropertiesService.getScriptProperties().setProperty("TEACHER_PASSWORD", TEACHER_DEFAULT_PASSWORD);
  return "تم تجهيز ورقة النتائج وضبط كلمة مرور المعلم";
}
function doPost(e) {
  try {
    const raw = e && e.parameter && e.parameter.payload;
    if (!raw) throw new Error("بيانات الطلب غير موجودة");
    const p = JSON.parse(raw);
    if (p.action !== "save") throw new Error("طلب غير معروف");
    ["studentName","className","email","score","startTime","endTime","details"].forEach(k => {
      if (p[k] === undefined || p[k] === null || p[k] === "") throw new Error("بيانات ناقصة: " + k);
    });
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(p.email))) throw new Error("البريد الإلكتروني غير صحيح");
    const sh = getSheet_();
    sh.appendRow([
      new Date(), safe_(p.studentName), safe_(p.className), safe_(p.email),
      Number(p.score), 20, Number(p.correctCount), Number(p.wrongCount),
      p.startTime, p.endTime, Number(p.durationSeconds || 0),
      JSON.stringify(p.details)
    ]);
    return json_({ok:true});
  } catch (err) {
    return json_({ok:false,error:String(err.message || err)});
  }
}
function doGet(e) {
  const params = (e && e.parameter) || {};
  let result;
  try {
    if (params.action !== "dashboard") throw new Error("طلب غير معروف");
    const expected = PropertiesService.getScriptProperties().getProperty("TEACHER_PASSWORD");
    if (!expected) throw new Error("لم يتم ضبط كلمة مرور المعلم في Script Properties");
    if (!params.password || params.password !== expected) throw new Error("كلمة المرور غير صحيحة");
    const sh = getSheet_(), values = sh.getDataRange().getValues();
    const rows = values.length < 2 ? [] : values.slice(1).map(r => ({
      studentName:r[1], className:r[2], email:r[3], score:r[4], totalScore:r[5],
      correctCount:r[6], wrongCount:r[7], startTime:r[8], endTime:r[9],
      durationSeconds:r[10], details:parseDetails_(r[11])
    }));
    result = {ok:true,rows:rows};
  } catch(err) { result = {ok:false,error:String(err.message || err)}; }
  const callback = String(params.callback || "");
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + "(" + JSON.stringify(result) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(result);
}
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { setup(); sh = ss.getSheetByName(SHEET_NAME); }
  return sh;
}
function parseDetails_(value) {
  try { return JSON.parse(value || "[]"); } catch(e) { return []; }
}
function safe_(value) {
  // Prevent spreadsheet formula injection.
  const s = String(value);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}