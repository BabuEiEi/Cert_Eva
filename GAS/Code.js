/**
 * ============================================================
 *  ระบบประเมินความพึงพอใจและออกเกียรติบัตรออนไลน์
 *  เฟส 1: สร้างฐานข้อมูล (Database Setup)
 * ============================================================
 *  วิธีใช้:
 *  1. เปิด Google Sheet ที่ต้องการใช้เป็นฐานข้อมูล
 *  2. เมนู Extensions > Apps Script
 *  3. วางโค้ดนี้ แล้วบันทึก
 *  4. เลือกฟังก์ชัน setupDatabase แล้วกด Run (อนุญาตสิทธิ์ครั้งแรก)
 *  5. ตรวจสอบว่ามี 5 ชีตถูกสร้างขึ้น พร้อมข้อมูลตัวอย่าง
 * ============================================================
 */

// ====== ค่าคงที่: ชื่อชีตทั้งหมด ======
const SHEETS = {
  CERTIFICATES: 'Certificates',
  QUESTIONS: 'Questions',
  RESPONSES: 'Responses',
  SETTINGS: 'Settings',
  USERS: 'Users'
};

/**
 * ฟังก์ชันหลัก: สร้างฐานข้อมูลทั้งหมด
 * รันฟังก์ชันนี้ครั้งเดียวเพื่อติดตั้งระบบ
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // สร้างทีละชีต
  setupCertificatesSheet_(ss);
  setupQuestionsSheet_(ss);
  setupResponsesSheet_(ss);
  setupSettingsSheet_(ss);
  setupUsersSheet_(ss);

  // ลบชีตเปล่าเริ่มต้น (Sheet1) ถ้ามี
  removeDefaultSheet_(ss);

  SpreadsheetApp.getUi().alert(
    '✅ ติดตั้งฐานข้อมูลสำเร็จ!\n\n' +
    'สร้างครบ 5 ชีต: Certificates, Questions, Responses, Settings, Users\n\n' +
    '⚠️ อย่าลืมไปกรอกค่า folderId, templateId, sheetId ในชีต Settings'
  );
}

/**
 * ============================================================
 *  ชีต 1: Certificates (ข้อมูลเกียรติบัตร)
 * ============================================================
 */
function setupCertificatesSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEETS.CERTIFICATES);
  sheet.clear();

  // หัวคอลัมน์
  const headers = ['runNo', 'certNo', 'fullName', 'school', 'status', 'fileUrl', 'createdAt'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ข้อมูลตัวอย่าง 3 รายการ (เก็บเป็นภาษาไทยปกติ)
  const sampleData = [
    [1, 'เลขที่ สพม.พลอต ๐๐๐๑/๒๕๖๙', 'สมชาย ใจดี', 'โรงเรียนวัดบางปลา', 'ยังไม่ประเมิน', '', ''],
    [2, 'เลขที่ สพม.พลอต ๐๐๐๒/๒๕๖๙', 'สมหญิง รักเรียน', 'โรงเรียนบ้านหนองหอย', 'ยังไม่ประเมิน', '', ''],
    [3, 'เลขที่ สพม.พลอต ๐๐๐๓/๒๕๖๙', 'วีระ ขยันยิ่ง', 'โรงเรียนวัดบางปลา', 'ยังไม่ประเมิน', '', '']
  ];
  sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);

  formatHeader_(sheet, headers.length);
  sheet.setColumnWidth(2, 180); // certNo กว้างหน่อย
}

/**
 * ============================================================
 *  ชีต 2: Questions (คำถามแบบประเมิน 3 ตอน 20 ข้อ)
 * ============================================================
 */
function setupQuestionsSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEETS.QUESTIONS);
  sheet.clear();

  // หัวคอลัมน์ (10 คอลัมน์ตาม schema)
  const headers = ['questionId', 'part', 'section', 'category', 'questionText', 'questionType', 'choices', 'required', 'orderNo', 'active'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const RATING_CHOICES = '5=มากที่สุด|4=มาก|3=ปานกลาง|2=น้อย|1=น้อยที่สุด';

  // คำถามทั้ง 20 ข้อ
  const questions = [
    // ----- ตอนที่ 1: ข้อมูลทั่วไป (radio) -----
    ['Q001', 'ตอนที่ 1', 'ทั่วไป', 'ข้อมูลทั่วไป', 'เพศ', 'radio', 'ชาย|หญิง', true, 1, true],
    ['Q002', 'ตอนที่ 1', 'ทั่วไป', 'ข้อมูลทั่วไป', 'อายุ', 'radio', '20 - 24 ปี|25 - 35 ปี|มากกว่า 36 ปี', true, 2, true],
    ['Q003', 'ตอนที่ 1', 'ทั่วไป', 'ข้อมูลทั่วไป', 'การศึกษา', 'radio', 'ปริญญาตรี|ปริญญาโท|ปริญญาเอก', true, 3, true],

    // ----- ตอนที่ 2: ระดับความคิดเห็น (rating) -----
    // ด้านการฝึกอบรม
    ['Q004', 'ตอนที่ 2', 'ระดับความคิดเห็นในการฝึกอบรม', 'การฝึกอบรม', 'เนื้อหาในการฝึกอบรมตรงกับวัตถุประสงค์', 'rating', RATING_CHOICES, true, 4, true],
    ['Q005', 'ตอนที่ 2', 'ระดับความคิดเห็นในการฝึกอบรม', 'การฝึกอบรม', 'ระยะเวลาในการฝึกอบรมมีความเหมาะสม', 'rating', RATING_CHOICES, true, 5, true],
    ['Q006', 'ตอนที่ 2', 'ระดับความคิดเห็นในการฝึกอบรม', 'การฝึกอบรม', 'รูปแบบและวิธีการฝึกอบรมมีความเหมาะสมกับสถานการณ์ปัจจุบัน', 'rating', RATING_CHOICES, true, 6, true],
    ['Q007', 'ตอนที่ 2', 'ระดับความคิดเห็นในการฝึกอบรม', 'การฝึกอบรม', 'หลักสูตรเอื้ออำนวยต่อการเรียนรู้และพัฒนาความสามารถของท่าน', 'rating', RATING_CHOICES, true, 7, true],
    ['Q008', 'ตอนที่ 2', 'ระดับความคิดเห็นในการฝึกอบรม', 'การฝึกอบรม', 'ท่านสามารถนำสิ่งที่ได้รับจากโครงการนี้ไปใช้ในการปฏิบัติงาน', 'rating', RATING_CHOICES, true, 8, true],
    // ด้านวิทยากร
    ['Q009', 'ตอนที่ 2', 'ระดับความคิดเห็นในการฝึกอบรม', 'วิทยากร', 'ความสามารถในการถ่ายทอด/สื่อสาร/ความเข้าใจ', 'rating', RATING_CHOICES, true, 9, true],
    ['Q010', 'ตอนที่ 2', 'ระดับความคิดเห็นในการฝึกอบรม', 'วิทยากร', 'การเรียงลำดับบรรยายเนื้อหาได้ครบถ้วน', 'rating', RATING_CHOICES, true, 10, true],
    ['Q011', 'ตอนที่ 2', 'ระดับความคิดเห็นในการฝึกอบรม', 'วิทยากร', 'การเปิดโอกาสให้ซักถามและแสดงความคิดเห็น', 'rating', RATING_CHOICES, true, 11, true],
    ['Q012', 'ตอนที่ 2', 'ระดับความคิดเห็นในการฝึกอบรม', 'วิทยากร', 'การตอบคำถามได้ตรงประเด็นและชัดเจน', 'rating', RATING_CHOICES, true, 12, true],
    ['Q013', 'ตอนที่ 2', 'ระดับความคิดเห็นในการฝึกอบรม', 'วิทยากร', 'ใช้เวลาเหมาะสมมาก/น้อย เพียงใด', 'rating', RATING_CHOICES, true, 13, true],
    // ด้านความรู้ความเข้าใจ
    ['Q014', 'ตอนที่ 2', 'ระดับความคิดเห็นในการฝึกอบรม', 'ความรู้ความเข้าใจที่ได้รับจากการฝึกอบรม', 'ความรู้ก่อนฝึกอบรม', 'rating', RATING_CHOICES, true, 14, true],
    ['Q015', 'ตอนที่ 2', 'ระดับความคิดเห็นในการฝึกอบรม', 'ความรู้ความเข้าใจที่ได้รับจากการฝึกอบรม', 'ความรู้หลังการฝึกอบรม', 'rating', RATING_CHOICES, true, 15, true],
    // ด้านอุปกรณ์/ระยะเวลา
    ['Q016', 'ตอนที่ 2', 'ระดับความคิดเห็นในการฝึกอบรม', 'อุปกรณ์เทคโนโลยี / ระยะเวลา', 'ความพร้อมของอุปกรณ์โสตทัศนูปกรณ์', 'rating', RATING_CHOICES, true, 16, true],
    ['Q017', 'ตอนที่ 2', 'ระดับความคิดเห็นในการฝึกอบรม', 'อุปกรณ์เทคโนโลยี / ระยะเวลา', 'ระยะเวลาในการอบรม / สัมมนามีความเหมาะสม', 'rating', RATING_CHOICES, true, 17, true],

    // ----- ตอนที่ 3: ข้อเสนอแนะ (textarea) -----
    ['Q018', 'ตอนที่ 3', 'ข้อเสนอแนะ', 'ข้อเสนอแนะ', 'สิ่งที่ท่านประทับใจในการอบรมครั้งนี้', 'textarea', '', true, 18, true],
    ['Q019', 'ตอนที่ 3', 'ข้อเสนอแนะ', 'ข้อเสนอแนะ', 'ท่านจะนำความรู้นี้ไปใช้ไปประยุกต์ในการทำงานของท่านอย่างไร', 'textarea', '', true, 19, true],
    ['Q020', 'ตอนที่ 3', 'ข้อเสนอแนะ', 'ข้อเสนอแนะ', 'ข้อเสนอแนะเพิ่มเติม (ถ้ามี)', 'textarea', '', false, 20, true]
  ];
  sheet.getRange(2, 1, questions.length, headers.length).setValues(questions);

  formatHeader_(sheet, headers.length);
  sheet.setColumnWidth(5, 320); // questionText กว้าง
  sheet.setColumnWidth(7, 250); // choices กว้าง
}

/**
 * ============================================================
 *  ชีต 3: Responses (คำตอบจากผู้ใช้ - ระบบเติมอัตโนมัติ)
 * ============================================================
 */
function setupResponsesSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEETS.RESPONSES);
  sheet.clear();

  const headers = ['timestamp', 'certNo', 'fullName', 'school', 'answersJSON'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  formatHeader_(sheet, headers.length);
  sheet.setColumnWidth(5, 400); // answersJSON กว้าง
}

/**
 * ============================================================
 *  ชีต 4: Settings (ตั้งค่าระบบ)
 * ============================================================
 */
function setupSettingsSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEETS.SETTINGS);
  sheet.clear();

  const headers = ['key', 'value', 'description'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const settings = [
    ['startNumber', 1, 'เลขรันนิ่งเริ่มต้นของเกียรติบัตร'],
    ['certCount', 100, 'จำนวนเกียรติบัตรทั้งหมด'],
    ['certPrefix', 'เลขที่ สพม.พลอต {NO_TH:๐๐๐๐}/{YEAR_TH}', 'รูปแบบเลขที่เกียรติบัตร ({NO}=เลขอารบิก, {NO_TH}=เลขไทย, {NO_TH:๐๐๐๐}=เลขไทยเติม 0, {YEAR_TH}=ปีเลขไทย)'],
    ['folderId', '', '🔴 TODO: ใส่ ID โฟลเดอร์ Google Drive สำหรับเก็บไฟล์ PDF'],
    ['templateId', '', '🔴 TODO: ใส่ ID ไฟล์ Google Slides Template เกียรติบัตร'],
    ['sheetId', ss.getId(), 'ID ของ Spreadsheet นี้ (เติมให้อัตโนมัติแล้ว)'],
    ['projectName', 'โครงการฝึกอบรม ปี 2569', 'ชื่อโครงการ'],
    ['certDate', '21 มิถุนายน 2569', 'วันที่บนเกียรติบัตร'],
    ['ratingYear', '2569', 'ปีที่ใช้ในรูปแบบ {YEAR} หรือ {YEAR_TH}']
  ];
  sheet.getRange(2, 1, settings.length, headers.length).setValues(settings);

  formatHeader_(sheet, headers.length);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 400);
}

/**
 * ============================================================
 *  ชีต 5: Users (ผู้ใช้ระบบ - รหัสผ่าน plain text)
 * ============================================================
 */
function setupUsersSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, SHEETS.USERS);
  sheet.clear();

  const headers = ['role', 'password', 'displayName'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const users = [
    ['admin', 'Admin1234', 'ผู้ดูแลระบบ'],
    ['staff', 'Staff1234', 'เจ้าหน้าที่']
  ];
  sheet.getRange(2, 1, users.length, headers.length).setValues(users);

  formatHeader_(sheet, headers.length);
}

/**
 * ============================================================
 *  ฟังก์ชันช่วยเหลือ (Helper Functions)
 * ============================================================
 */

// ดึงชีตที่มีอยู่ หรือสร้างใหม่ถ้ายังไม่มี
function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

// จัดรูปแบบหัวตาราง (พื้นหลังสี + ตัวหนา + ตรึงแถว)
function formatHeader_(sheet, numCols) {
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

// ลบชีตเริ่มต้น (Sheet1 / ชีต1) ถ้ามีและว่างเปล่า
function removeDefaultSheet_(ss) {
  const defaultNames = ['Sheet1', 'ชีต1'];
  defaultNames.forEach(function (name) {
    const sheet = ss.getSheetByName(name);
    if (sheet && ss.getSheets().length > 1) {
      ss.deleteSheet(sheet);
    }
  });
}

/**
 * ============================================================
 *  ฟังก์ชันถอดรหัส Base64 สำหรับรองรับข้อมูลเก่า
 *  ระบบใหม่เก็บ fullName/school เป็นภาษาไทยปกติในชีต
 * ============================================================
 */

// เข้ารหัสข้อความเป็น Base64 (เก็บไว้เพื่อความเข้ากันได้กับข้อมูล/สคริปต์เก่า)
function encodeText_(text) {
  if (text === '' || text === null || text === undefined) return '';
  const bytes = Utilities.newBlob(text).getBytes();
  return Utilities.base64Encode(bytes);
}

// ถอดรหัส Base64 กลับเป็นข้อความ ถ้าไม่ใช่ Base64 จะคืนค่าเดิม
function decodeText_(value) {
  if (value === '' || value === null || value === undefined) return '';
  const text = String(value);
  const compact = text.trim();

  if (!isLikelyBase64_(compact)) return text;

  try {
    const bytes = Utilities.base64Decode(compact);
    const decoded = Utilities.newBlob(bytes).getDataAsString('UTF-8');
    const encodedAgain = Utilities.base64Encode(Utilities.newBlob(decoded).getBytes()).replace(/=+$/, '');
    const original = compact.replace(/=+$/, '');

    if (encodedAgain !== original) return text;
    if (/\uFFFD|[\x00-\x08\x0B\x0C\x0E-\x1F]/.test(decoded)) return text;
    return decoded;
  } catch (e) {
    return text;
  }
}

function isLikelyBase64_(text) {
  if (!text || text.length < 4 || text.length % 4 !== 0) return false;
  return /^[A-Za-z0-9+/]+={0,2}$/.test(text);
}

/**
 * ย้ายข้อมูลเก่า fullName/school จาก Base64 ให้เป็นภาษาไทยปกติ
 * วิธีใช้: รันฟังก์ชันนี้ครั้งเดียวใน Apps Script หลังอัปเดตโค้ด
 */
function migrateBase64TextToThai() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targets = [
    { sheetName: SHEETS.CERTIFICATES, headers: ['fullName', 'school'] },
    { sheetName: SHEETS.RESPONSES, headers: ['fullName', 'school'] }
  ];
  let changed = 0;

  targets.forEach(function (target) {
    const sheet = ss.getSheetByName(target.sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const cols = target.headers
      .map(function (header) { return headers.indexOf(header) + 1; })
      .filter(function (col) { return col > 0; });

    cols.forEach(function (col) {
      const range = sheet.getRange(2, col, sheet.getLastRow() - 1, 1);
      const values = range.getValues();
      const updated = values.map(function (row) {
        const current = row[0];
        const decoded = decodeText_(current);
        if (String(current || '') !== String(decoded || '')) changed++;
        return [decoded];
      });
      range.setValues(updated);
    });
  });

  SpreadsheetApp.getUi().alert('แปลงข้อมูล Base64 เป็นภาษาไทยปกติแล้ว ' + changed + ' ช่อง');
}

/**
 * ย้าย schema Certificates เดิมที่มี prefix ให้เป็น schema ใหม่ที่ไม่มี prefix
 * วิธีใช้: รันฟังก์ชันนี้ครั้งเดียวหลังอัปเดตโค้ด ถ้าชีตเดิมยังมีคอลัมน์ prefix
 */
function migrateCertificatesRemovePrefix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.CERTIFICATES);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('ไม่พบชีต Certificates');
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const prefixIndex = headers.indexOf('prefix');
  if (prefixIndex === -1) {
    SpreadsheetApp.getUi().alert('ชีต Certificates ไม่มีคอลัมน์ prefix อยู่แล้ว');
    return;
  }

  sheet.deleteColumn(prefixIndex + 1);
  SpreadsheetApp.flush();
  if (typeof clearSheetCache_ === 'function') clearSheetCache_(SHEETS.CERTIFICATES);
  SpreadsheetApp.getUi().alert('ลบคอลัมน์ prefix จากชีต Certificates เรียบร้อยแล้ว');
}

/**
 * ============================================================
 *  ฟังก์ชันทดสอบ: ตรวจสอบการเข้ารหัส/ถอดรหัส
 *  (รันเพื่อทดสอบว่า Base64 ทำงานถูกต้องกับภาษาไทย)
 * ============================================================
 */
function testEncodeDecode() {
  const original = 'สมชาย ใจดี';
  const encoded = encodeText_(original);
  const decoded = decodeText_(encoded);
  Logger.log('ต้นฉบับ: ' + original);
  Logger.log('เข้ารหัส: ' + encoded);
  Logger.log('ถอดรหัส: ' + decoded);
  Logger.log('ตรงกัน: ' + (original === decoded ? '✅ ใช่' : '❌ ไม่'));
}

/**
 * ============================================================
 *  ฟังก์ชันสร้างเลขที่เกียรติบัตรจากรูปแบบ (certPrefix)
 *  รองรับ {NO}, {NO:0000}, {NO_TH}, {NO_TH:๐๐๐๐}, {YEAR}, {YEAR_TH}
 *  ตัวอย่าง: 'เลขที่ สพม.พลอต {NO_TH:๐๐๐๐}/{YEAR_TH}' + runNo=5 => 'เลขที่ สพม.พลอต ๐๐๐๕/๒๕๖๙'
 * ============================================================
 */
function buildCertNo_(pattern, runNo, year) {
  let result = pattern;

  // แทน {NO_TH:๐๐๐๐} หรือ {NO_TH:0000} (เลขไทย เติม 0 ตามจำนวนหลัก)
  result = result.replace(/\{NO_TH:([0๐]+)\}/g, function (match, zeros) {
    const width = zeros.length;
    return toThaiDigits_(String(runNo).padStart(width, '0'));
  });

  // แทน {NO_TH} (เลขไทย ไม่เติม 0)
  result = result.replace(/\{NO_TH\}/g, toThaiDigits_(runNo));

  // แทน {NO:0000} (เติม 0 ตามจำนวนหลัก)
  result = result.replace(/\{NO:(0+)\}/g, function (match, zeros) {
    return String(runNo).padStart(zeros.length, '0');
  });

  // แทน {NO} (เลขธรรมดา)
  result = result.replace(/\{NO\}/g, String(runNo));

  // แทน {YEAR}
  result = result.replace(/\{YEAR\}/g, String(year || ''));

  // แทน {YEAR_TH}
  result = result.replace(/\{YEAR_TH\}/g, toThaiDigits_(year || ''));

  return result;
}

function toThaiDigits_(value) {
  const map = {
    '0': '๐', '1': '๑', '2': '๒', '3': '๓', '4': '๔',
    '5': '๕', '6': '๖', '7': '๗', '8': '๘', '9': '๙'
  };
  return String(value).replace(/[0-9]/g, function (digit) {
    return map[digit];
  });
}

/**
 * รันเลขที่เกียรติบัตรใหม่จาก Settings: startNumber, certCount, certPrefix, ratingYear
 * วิธีใช้: รันฟังก์ชันนี้จาก Apps Script หรือเรียกผ่านหน้า Admin
 */
function regenerateCertificateNumbers() {
  const result = regenerateCertificateNumbers_();
  SpreadsheetApp.getUi().alert(result.message);
  return result;
}

function regenerateCertificateNumbers_() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const settings = getSettingsMapForNumbering_();
    const startNumber = Number(settings.startNumber) || 1;
    const certCount = Number(settings.certCount) || 0;
    const certPrefix = settings.certPrefix || 'เลขที่ สพม.พลอต {NO_TH:๐๐๐๐}/{YEAR_TH}';
    const year = settings.ratingYear || '';

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.CERTIFICATES);
    if (!sheet) return { success: false, message: 'ไม่พบชีต Certificates' };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'ยังไม่มีรายชื่อเกียรติบัตรให้รันเลข' };

    const existingRowCount = lastRow - 1;
    const rowCount = certCount > 0 ? Math.min(existingRowCount, certCount) : existingRowCount;
    const values = [];

    for (let i = 0; i < rowCount; i++) {
      const runNo = startNumber + i;
      const certNo = buildCertNo_(certPrefix, runNo, year);
      values.push([runNo, certNo]);
    }

    sheet.getRange(2, 1, values.length, 2).setValues(values);
    SpreadsheetApp.flush();

    if (typeof clearSheetCache_ === 'function') clearSheetCache_(SHEETS.CERTIFICATES);

    const skipped = existingRowCount - rowCount;
    return {
      success: true,
      message: 'รันเลขเกียรติบัตรใหม่สำเร็จ ' + rowCount + ' รายการ' + (skipped > 0 ? ' (ไม่ได้รัน ' + skipped + ' รายการ เพราะเกิน certCount)' : ''),
      updated: rowCount,
      skipped: skipped,
      startNumber: startNumber,
      certPrefix: certPrefix,
      ratingYear: year
    };

  } catch (err) {
    return { success: false, message: 'รันเลขเกียรติบัตรไม่สำเร็จ: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

function getSettingsMapForNumbering_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.SETTINGS);
  if (!sheet || sheet.getLastRow() < 2) return {};

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const map = {};
  values.forEach(function (row) {
    map[row[0]] = row[1];
  });
  return map;
}

/**
 * ฟังก์ชันทดสอบ: สร้างเลขที่เกียรติบัตร
 */
function testBuildCertNo() {
  Logger.log(buildCertNo_('เลขที่ สพม.พลอต {NO:0000}/2569', 5));   // เลขที่ สพม.พลอต 0005/2569
  Logger.log(buildCertNo_('เลขที่ สพม.พลอต {NO}/2569', 5));        // เลขที่ สพม.พลอต 5/2569
  Logger.log(buildCertNo_('เลขที่ สพม.พลอต {NO_TH:๐๐๐๐}/{YEAR_TH}', 5, 2569)); // เลขที่ สพม.พลอต ๐๐๐๕/๒๕๖๙
  Logger.log(buildCertNo_('CERT-{NO:000}-{YEAR}', 12, 2569)); // CERT-012-2569
}