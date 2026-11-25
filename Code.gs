const SHEET_NAME = "Registrations";

function doGet(e) {
  var page = e && e.parameter && e.parameter.page
    ? e.parameter.page
    : 'index';

  if (page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('admin')
      .setTitle('پنل مدیریت');
  }

  if (page === 'dashboard') {
    return HtmlService.createHtmlOutputFromFile('dashboard')
      .setTitle('داشبورد مدیریت');
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('انجمن علمی رباتیک دانشگاه ولیعصر');
}

function getSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (!id) throw new Error("ابتدا SPREADSHEET_ID را در Script Properties تنظیم کنید.");
  const ss = SpreadsheetApp.openById(id);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["زمان ثبت", "نام", "نام خانوادگی", "شماره همراه", "شماره دانشجویی", "رشته تحصیلی"]);
  }
  return sheet;
}

function saveRegistration(data) {
  if (!data || !data.firstName || !data.lastName || !data.phone || !data.studentId || !data.major) {
    throw new Error("لطفاً همه فیلدها را تکمیل کنید.");
  }
  if (!/^09\d{9}$/.test(String(data.phone))) {
    throw new Error("شماره همراه باید ۱۱ رقم و با 09 شروع شود.");
  }
  getSheet_().appendRow([
    new Date(), String(data.firstName).trim(), String(data.lastName).trim(),
    String(data.phone).trim(), String(data.studentId).trim(), String(data.major).trim()
  ]);
  return { success: true, message: "ثبت‌نام شما با موفقیت انجام شد." };
}

function loginAdmin(username, password) {
  const props = PropertiesService.getScriptProperties();
  const savedUser = props.getProperty("ADMIN_USERNAME");
  const savedPass = props.getProperty("ADMIN_PASSWORD");
  if (!savedUser || !savedPass) return { success: false, message: "اطلاعات مدیر در Script Properties تنظیم نشده است." };
  if (String(username) !== savedUser || String(password) !== savedPass) {
    return { success: false, message: "نام کاربری یا رمز عبور اشتباه است." };
  }
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put("ADMIN_" + token, "ok", 1800);
  return { success: true, token: token };
}

function requireAdmin_(token) {
  if (!token || CacheService.getScriptCache().get("ADMIN_" + token) !== "ok") {
    throw new Error("نشست مدیر معتبر نیست؛ دوباره وارد شوید.");
  }
}

function getRegistrations(token) {
  requireAdmin_(token);
  const values = getSheet_().getDataRange().getValues();
  if (values.length <= 1) return { rows: [], total: 0 };
  const rows = values.slice(1).map((r, i) => ({
    rowNumber: i + 2,
    date: r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm") : String(r[0]),
    firstName: String(r[1]), lastName: String(r[2]), phone: String(r[3]),
    studentId: String(r[4]), major: String(r[5])
  })).reverse();
  return { rows: rows, total: rows.length };
}

function deleteRegistration(token, rowNumber) {
  requireAdmin_(token);
  const row = Number(rowNumber);
  const sheet = getSheet_();
  if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) throw new Error("ردیف انتخاب‌شده معتبر نیست.");
  sheet.deleteRow(row);
  return { success: true };
}

function logoutAdmin(token) {
  if (token) CacheService.getScriptCache().remove("ADMIN_" + token);
  return true;
}

function exportRegistrationsCsv(token) {
  requireAdmin_(token);
  const values = getSheet_().getDataRange().getDisplayValues();
  const csv = values.map(row => row.map(cell => '"' + String(cell).replace(/"/g, '""') + '"').join(",")).join("\r\n");
  return Utilities.base64Encode("\uFEFF" + csv, Utilities.Charset.UTF_8);
}
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}