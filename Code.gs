const SHEET_NAME = 'Registrations';

// نمایش صفحات
function doGet(e) {
  const page = e && e.parameter && e.parameter.page
    ? e.parameter.page
    : 'index';

  const allowed = ['index', 'admin', 'dashboard'];
  const file = allowed.includes(page) ? page : 'index';

  const template = HtmlService.createTemplateFromFile(file);
  template.appUrl = ScriptApp.getService().getUrl();

  return template.evaluate()
    .setTitle('Robotics Association')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// دریافت اطلاعات از سایت GitHub
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No data received');
    }

    const data = JSON.parse(e.postData.contents);
    const result = saveRegistration(data);

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: error.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// دریافت فایل Google Sheets
function getSheet_() {
  const id = PropertiesService
    .getScriptProperties()
    .getProperty('SPREADSHEET_ID');

  if (!id) {
    throw new Error('SPREADSHEET_ID is missing');
  }

  const ss = SpreadsheetApp.openById(id);
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Date',
      'First Name',
      'Last Name',
      'Phone',
      'Student ID',
      'Major'
    ]);
  }

  return sheet;
}

// ثبت اطلاعات
function saveRegistration(data) {
  if (!data) throw new Error('اطلاعات دریافت نشد.');

  const firstName = String(data.firstName || '').trim();
  const lastName = String(data.lastName || '').trim();
  const phone = String(data.phone || '').trim();
  const studentId = String(data.studentId || '').trim();
  const major = String(data.major || '').trim();

  if (!firstName || !lastName || !phone ||
      !studentId || !major) {
    throw new Error('تمام فیلدها را تکمیل کنید.');
  }

  if (!/^09\d{9}$/.test(phone)) {
    throw new Error('شماره موبایل معتبر نیست.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getSheet_();

    sheet.appendRow([
      new Date(),
      firstName,
      lastName,
      phone,
      studentId,
      major
    ]);

    return {
      success: true,
      message: 'اطلاعات با موفقیت ثبت شد.'
    };

  } finally {
    lock.releaseLock();
  }
}

// ورود مدیر
function loginAdmin(username, password) {
  const props = PropertiesService.getScriptProperties();

  const correctUser = props.getProperty('ADMIN_USERNAME');
  const correctPass = props.getProperty('ADMIN_PASSWORD');

  if (!correctUser || !correctPass) {
    throw new Error('اطلاعات مدیر تنظیم نشده است.');
  }

  if (username !== correctUser || password !== correctPass) {
    throw new Error('نام کاربری یا رمز عبور اشتباه است.');
  }

  const token = Utilities.getUuid();

  CacheService.getScriptCache().put(
    'admin_' + token,
    'valid',
    21600
  );

  return token;
}

// بررسی اعتبار مدیر
function requireAdmin_(token) {
  if (!token ||
      CacheService.getScriptCache().get('admin_' + token) !== 'valid') {
    throw new Error('نشست شما منقضی شده است.');
  }
}

// دریافت اطلاعات اعضا
function getRegistrations(token) {
  requireAdmin_(token);

  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();

  const rows = values.slice(1).map((r, i) => ({
    rowNumber: i + 2,
    date: r[0] instanceof Date
      ? Utilities.formatDate(
          r[0],
          Session.getScriptTimeZone(),
          'yyyy/MM/dd HH:mm'
        )
      : String(r[0]),
    firstName: String(r[1]),
    lastName: String(r[2]),
    phone: String(r[3]),
    studentId: String(r[4]),
    major: String(r[5])
  }));

  return {
    total: rows.length,
    rows: rows.reverse()
  };
}

// حذف عضو
function deleteRegistration(token, rowNumber) {
  requireAdmin_(token);

  const sheet = getSheet_();
  const row = Number(rowNumber);

  if (!Number.isInteger(row) ||
      row < 2 ||
      row > sheet.getLastRow()) {
    throw new Error('ردیف نامعتبر است.');
  }

  sheet.deleteRow(row);
  return true;
}

// خروجی CSV برای Excel
function exportRegistrationsCsv(token) {
  requireAdmin_(token);

  const values = getSheet_().getDataRange().getDisplayValues();

  const csv = values.map(row =>
    row.map(value =>
      '"' + String(value).replace(/"/g, '""') + '"'
    ).join(',')
  ).join('\r\n');

  return Utilities.base64Encode(
    Utilities.newBlob(csv).getBytes()
  );
}

// خروج مدیر
function logoutAdmin(token) {
  if (token) {
    CacheService.getScriptCache().remove('admin_' + token);
  }

  return true;
}