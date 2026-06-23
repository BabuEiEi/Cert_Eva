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
  const headers = ['runNo', 'certNo', 'prefix', 'fullName', 'school', 'status', 'fileUrl', 'createdAt'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ข้อมูลตัวอย่าง 3 รายการ (fullName/school เข้ารหัส Base64)
  const sampleData = [
    [1, 'สพม.พลอต 0001/2569', 'นาย', encodeText_('สมชาย ใจดี'), encodeText_('โรงเรียนวัดบางปลา'), 'ยังไม่ประเมิน', '', ''],
    [2, 'สพม.พลอต 0002/2569', 'นางสาว', encodeText_('สมหญิง รักเรียน'), encodeText_('โรงเรียนบ้านหนองหอย'), 'ยังไม่ประเมิน', '', ''],
    [3, 'สพม.พลอต 0003/2569', 'นาย', encodeText_('วีระ ขยันยิ่ง'), encodeText_('โรงเรียนวัดบางปลา'), 'ยังไม่ประเมิน', '', '']
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
    ['certPrefix', 'สพม.พลอต {NO:0000}/2569', 'รูปแบบเลขที่เกียรติบัตร ({NO}=เลขรันนิ่ง, {NO:0000}=เติม 0, {YEAR}=ปี)'],
    ['folderId', '', '🔴 TODO: ใส่ ID โฟลเดอร์ Google Drive สำหรับเก็บไฟล์ PDF'],
    ['templateId', '', '🔴 TODO: ใส่ ID ไฟล์ Google Slides Template เกียรติบัตร'],
    ['sheetId', ss.getId(), 'ID ของ Spreadsheet นี้ (เติมให้อัตโนมัติแล้ว)'],
    ['projectName', 'โครงการฝึกอบรม ปี 2569', 'ชื่อโครงการ'],
    ['certDate', '21 มิถุนายน 2569', 'วันที่บนเกียรติบัตร'],
    ['ratingYear', '2569', 'ปีที่ใช้ในรูปแบบ {YEAR}']
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
  defaultNames.forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    if (sheet && ss.getSheets().length > 1) {
      ss.deleteSheet(sheet);
    }
  });
}

/**
 * ============================================================
 *  ฟังก์ชันเข้ารหัส/ถอดรหัส Base64 (ระดับ "ซ่อนตา")
 *  ใช้กับข้อมูลส่วนตัว เช่น fullName, school
 * ============================================================
 */

// เข้ารหัสข้อความเป็น Base64 (รองรับภาษาไทย UTF-8)
function encodeText_(text) {
  if (text === '' || text === null || text === undefined) return '';
  const bytes = Utilities.newBlob(text).getBytes();
  return Utilities.base64Encode(bytes);
}

// ถอดรหัส Base64 กลับเป็นข้อความ
function decodeText_(encoded) {
  if (encoded === '' || encoded === null || encoded === undefined) return '';
  try {
    const bytes = Utilities.base64Decode(encoded);
    return Utilities.newBlob(bytes).getDataAsString();
  } catch (e) {
    return encoded; // ถ้าถอดไม่ได้ คืนค่าเดิม (กันข้อมูลเก่าที่ไม่ได้เข้ารหัส)
  }
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
 *  รองรับ {NO}, {NO:0000}, {YEAR}
 *  ตัวอย่าง: 'สพม.พลอต {NO:0000}/2569' + runNo=5 => 'สพม.พลอต 0005/2569'
 * ============================================================
 */
function buildCertNo_(pattern, runNo, year) {
  let result = pattern;

  // แทน {NO:0000} (เติม 0 ตามจำนวนหลัก)
  result = result.replace(/\{NO:(0+)\}/g, function(match, zeros) {
    return String(runNo).padStart(zeros.length, '0');
  });

  // แทน {NO} (เลขธรรมดา)
  result = result.replace(/\{NO\}/g, String(runNo));

  // แทน {YEAR}
  result = result.replace(/\{YEAR\}/g, String(year || ''));

  return result;
}

/**
 * ฟังก์ชันทดสอบ: สร้างเลขที่เกียรติบัตร
 */
function testBuildCertNo() {
  Logger.log(buildCertNo_('สพม.พลอต {NO:0000}/2569', 5));   // สพม.พลอต 0005/2569
  Logger.log(buildCertNo_('สพม.พลอต {NO}/2569', 5));        // สพม.พลอต 5/2569
  Logger.log(buildCertNo_('CERT-{NO:000}-{YEAR}', 12, 2569)); // CERT-012-2569
}