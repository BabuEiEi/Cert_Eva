/**
 * ============================================================
 *  เฟส 2: Backend API (Web App)
 *  *** เพิ่มโค้ดนี้ต่อท้าย Code.gs เดิม (ไฟล์เดียวกัน) ***
 * ============================================================
 *  หลังวางโค้ดนี้แล้ว ให้ Deploy เป็น Web App:
 *  1. Deploy > New deployment > เลือก Web app
 *  2. Execute as: Me
 *  3. Who has access: Anyone
 *  4. ก๊อป Web App URL ไปใส่ใน index.html / admin.html (เฟส 3-4)
 * ============================================================
 */

// ====== ตัวจัดการคำขอแบบ GET (อ่านข้อมูล) ======
function doGet(e) {
  return handleRequest_(e);
}

// ====== ตัวจัดการคำขอแบบ POST (เขียนข้อมูล) ======
function doPost(e) {
  return handleRequest_(e);
}

/**
 * ตัวจัดการคำขอกลาง (Router)
 * แยกการทำงานตามพารามิเตอร์ action
 */
function handleRequest_(e) {
  let result;
  try {
    // รวมพารามิเตอร์จาก GET (e.parameter) และ POST body (JSON)
    let params = e && e.parameter ? e.parameter : {};
    if (e && e.postData && e.postData.contents) {
      try {
        const body = JSON.parse(e.postData.contents);
        params = Object.assign({}, params, body);
      } catch (err) {
        // ถ้า body ไม่ใช่ JSON ก็ใช้ parameter อย่างเดียว
      }
    }

    const action = params.action || '';

    // ===== เส้นทาง (Routing) ตาม action =====
    switch (action) {
      case 'login': result = apiLogin_(params); break;
      case 'searchCertificate': result = apiSearchCertificate_(params); break;
      case 'getQuestions': result = apiGetQuestions_(); break;
      case 'checkStatus': result = apiCheckStatus_(params); break;
      case 'submitResponse': result = apiSubmitResponse_(params); break;
      case 'getCertificates': result = apiGetCertificates_(); break;
      case 'generateCertificate': result = apiGenerateCertificate_(params); break;
      case 'generateAll': result = apiGenerateAll_(); break;
      case 'repairCertificateFiles': result = apiRepairCertificateFiles_(); break;
      case 'saveQuestion': result = apiSaveQuestion_(params); break;
      case 'deleteQuestion': result = apiDeleteQuestion_(params); break;
      case 'getSettings': result = apiGetSettings_(); break;
      case 'saveSettings': result = apiSaveSettings_(params); break;
      case 'getResults': result = apiGetResults_(); break;
      case 'getDashboard': result = apiGetDashboard_(); break;
      case 'getUsers': result = apiGetUsers_(); break;
      case 'saveUser': result = apiSaveUser_(params); break;
      default:
        result = { success: false, message: 'ไม่พบ action: ' + action };
    }
  } catch (err) {
    result = { success: false, message: 'เกิดข้อผิดพลาด: ' + err.message };
  }

  // คืนค่าเป็น JSON เสมอ
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ============================================================
 *  ตัวช่วยอ่านข้อมูลชีตแบบ batch + cache
 * ============================================================
 */

// อ่านข้อมูลทั้งชีตเป็น array ของ object (ใช้แถวแรกเป็น key)
function readSheetAsObjects_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = values[i][j];
    }
    obj._rowIndex = i + 1; // เก็บเลขแถวจริง (สำหรับแก้ไข)
    rows.push(obj);
  }
  return rows;
}

// อ่านข้อมูลแบบ cache (สำหรับข้อมูลที่เปลี่ยนไม่บ่อย เช่น Questions, Settings)
function readSheetCached_(sheetName, cacheSeconds) {
  const cache = CacheService.getScriptCache();
  const key = 'sheet_' + sheetName;
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { }
  }
  const data = readSheetAsObjects_(sheetName);
  try {
    cache.put(key, JSON.stringify(data), cacheSeconds || 60);
  } catch (e) {
    // ถ้าข้อมูลใหญ่เกิน cache limit ก็ข้ามไป
  }
  return data;
}

// ล้าง cache ของชีต (เรียกหลังเขียนข้อมูล)
function clearSheetCache_(sheetName) {
  CacheService.getScriptCache().remove('sheet_' + sheetName);
}

/**
 * ============================================================
 *  API: login (ตรวจสอบรหัสผ่าน Admin/Staff)
 * ============================================================
 */
function apiLogin_(params) {
  const password = params.password || '';
  if (!password) return { success: false, message: 'กรุณากรอกรหัสผ่าน' };

  const users = readSheetAsObjects_(SHEETS.USERS);
  const matched = users.find(function (u) {
    return String(u.password) === String(password);
  });

  if (matched) {
    return {
      success: true,
      role: matched.role,
      displayName: matched.displayName,
      message: 'เข้าสู่ระบบสำเร็จ'
    };
  }
  return { success: false, message: 'รหัสผ่านไม่ถูกต้อง' };
}

/**
 * ============================================================
 *  API: searchCertificate (ค้นหาเกียรติบัตร — สำหรับ User)
 *  ค้นด้วยชื่อ-นามสกุล หรือ ชื่อโรงเรียน (ถอดรหัส Base64 ก่อนเทียบ)
 * ============================================================
 */
function apiSearchCertificate_(params) {
  const keyword = (params.keyword || '').trim().toLowerCase();
  if (!keyword) return { success: false, message: 'กรุณากรอกคำค้นหา' };

  const certs = readSheetCached_(SHEETS.CERTIFICATES, 30);
  const results = [];

  certs.forEach(function (c) {
    const fullName = decodeText_(c.fullName);
    const school = decodeText_(c.school);
    // เทียบทั้งชื่อและโรงเรียน
    if (fullName.toLowerCase().indexOf(keyword) !== -1 ||
      school.toLowerCase().indexOf(keyword) !== -1) {
      results.push({
        certNo: c.certNo,
        prefix: c.prefix,
        fullName: fullName,
        school: school,
        status: c.status,
        fileUrl: c.fileUrl || ''
      });
    }
  });

  return {
    success: true,
    count: results.length,
    data: results
  };
}

/**
 * ============================================================
 *  API: getQuestions (ดึงคำถามทั้งหมด — เรียงตาม orderNo)
 *  คืนเฉพาะคำถามที่ active=true จัดกลุ่มตาม part
 * ============================================================
 */
function apiGetQuestions_() {
  const questions = readSheetCached_(SHEETS.QUESTIONS, 60);

  const active = questions
    .filter(function (q) { return q.active === true || String(q.active).toUpperCase() === 'TRUE'; })
    .sort(function (a, b) { return Number(a.orderNo) - Number(b.orderNo); })
    .map(function (q) {
      // แปลง choices string เป็น array
      let choices = [];
      if (q.choices) {
        choices = String(q.choices).split('|').map(function (opt) {
          // รองรับรูปแบบ "5=มากที่สุด" => {value:5, label:'มากที่สุด'}
          if (opt.indexOf('=') !== -1) {
            const parts = opt.split('=');
            return { value: parts[0].trim(), label: parts[1].trim() };
          }
          return { value: opt.trim(), label: opt.trim() };
        });
      }
      return {
        questionId: q.questionId,
        part: q.part,
        section: q.section,
        category: q.category,
        questionText: q.questionText,
        questionType: q.questionType,
        choices: choices,
        required: (q.required === true || String(q.required).toUpperCase() === 'TRUE'),
        orderNo: Number(q.orderNo)
      };
    });

  return { success: true, data: active };
}

/**
 * ============================================================
 *  API: checkStatus (เช็คสถานะเกียรติบัตรใบเดียว)
 * ============================================================
 */
function apiCheckStatus_(params) {
  const certNo = params.certNo || '';
  const certs = readSheetCached_(SHEETS.CERTIFICATES, 30);
  const cert = certs.find(function (c) { return String(c.certNo) === String(certNo); });

  if (!cert) return { success: false, message: 'ไม่พบเกียรติบัตร' };

  return {
    success: true,
    certNo: cert.certNo,
    status: cert.status,
    fileUrl: cert.fileUrl || ''
  };
}

/**
 * ============================================================
 *  API: submitResponse (บันทึกคำตอบแบบประเมิน — สำหรับ User)
 *  ใช้ LockService กันเขียนชนกัน
 * ============================================================
 */
function apiSubmitResponse_(params) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // รอ lock สูงสุด 10 วินาที

    const certNo = params.certNo || '';
    const answers = params.answers; // คาดหวังเป็น object {Q001:..., Q004:5, ...}

    if (!certNo) return { success: false, message: 'ไม่พบเลขที่เกียรติบัตร' };
    if (!answers) return { success: false, message: 'ไม่พบคำตอบ' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // หาเกียรติบัตร
    const certs = readSheetAsObjects_(SHEETS.CERTIFICATES);
    const cert = certs.find(function (c) { return String(c.certNo) === String(certNo); });
    if (!cert) return { success: false, message: 'ไม่พบเกียรติบัตรนี้ในระบบ' };

    // กันประเมินซ้ำ
    const responses = readSheetAsObjects_(SHEETS.RESPONSES);
    const already = responses.find(function (r) { return String(r.certNo) === String(certNo); });
    if (already) {
      return { success: false, message: 'เกียรติบัตรนี้ทำแบบประเมินไปแล้ว', alreadyDone: true };
    }

    // แปลง answers เป็น object ถ้าส่งมาเป็น string
    let answersObj = answers;
    if (typeof answers === 'string') {
      answersObj = JSON.parse(answers);
    }

    // บันทึกลงชีต Responses เป็นภาษาไทยปกติ
    const respSheet = ss.getSheetByName(SHEETS.RESPONSES);
    respSheet.appendRow([
      new Date(),
      certNo,
      decodeText_(cert.fullName),
      decodeText_(cert.school),
      JSON.stringify(answersObj)
    ]);

    // อัปเดตสถานะเกียรติบัตร => "พร้อมดาวน์โหลด"
    const certSheet = ss.getSheetByName(SHEETS.CERTIFICATES);
    const statusCol = 6; // คอลัมน์ status (ลำดับที่ 6)
    certSheet.getRange(cert._rowIndex, statusCol).setValue('พร้อมดาวน์โหลด');

    clearSheetCache_(SHEETS.RESPONSES);
    clearSheetCache_(SHEETS.CERTIFICATES);

    return {
      success: true,
      message: 'บันทึกแบบประเมินสำเร็จ',
      certNo: certNo,
      newStatus: 'พร้อมดาวน์โหลด',
      fileUrl: cert.fileUrl || ''
    };

  } catch (err) {
    return { success: false, message: 'บันทึกไม่สำเร็จ: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================
 *  API: getCertificates (ดึงรายชื่อเกียรติบัตรทั้งหมด — สำหรับ Admin)
 * ============================================================
 */
function apiGetCertificates_() {
  const certs = readSheetCached_(SHEETS.CERTIFICATES, 30);
  const data = certs.map(function (c) {
    return {
      runNo: c.runNo,
      certNo: c.certNo,
      prefix: c.prefix,
      fullName: decodeText_(c.fullName),
      school: decodeText_(c.school),
      status: c.status,
      fileUrl: c.fileUrl || ''
    };
  });
  return { success: true, count: data.length, data: data };
}

/**
 * ============================================================
 *  API: generateCertificate (สร้างเกียรติบัตร 1 ใบจาก Slides Template)
 * ============================================================
 */
function apiGenerateCertificate_(params) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const certNo = params.certNo || '';
    const certs = readSheetAsObjects_(SHEETS.CERTIFICATES);
    const cert = certs.find(function (c) { return String(c.certNo) === String(certNo); });
    if (!cert) return { success: false, message: 'ไม่พบเกียรติบัตร' };

    const result = createCertificateFile_(cert);
    return result;

  } catch (err) {
    return { success: false, message: 'สร้างเกียรติบัตรไม่สำเร็จ: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================
 *  API: generateAll (สร้างเกียรติบัตรทั้งหมดที่ยังไม่มีไฟล์)
 * ============================================================
 */
function apiGenerateAll_() {
  const certs = readSheetAsObjects_(SHEETS.CERTIFICATES);
  let created = 0, skipped = 0, failed = 0;
  const errors = [];

  certs.forEach(function (cert) {
    if (cert.fileUrl) { skipped++; return; } // มีไฟล์แล้ว ข้าม
    const r = createCertificateFile_(cert);
    if (r.success) created++;
    else { failed++; errors.push(cert.certNo + ': ' + r.message); }
  });

  return {
    success: true,
    message: 'สร้างเสร็จ ' + created + ' ใบ, ข้าม ' + skipped + ' ใบ, ล้มเหลว ' + failed + ' ใบ',
    created: created, skipped: skipped, failed: failed, errors: errors
  };
}

/**
 * ฟังก์ชันสร้างไฟล์เกียรติบัตรจริง (ใช้ร่วมกัน)
 * คัดลอก Slides Template > แทนค่า placeholder > export PDF > เก็บใน Drive
 */
function createCertificateFile_(cert) {
  const settings = getSettingsMap_();
  const templateId = settings.templateId;
  const folderId = settings.folderId;

  // ตรวจสอบการตั้งค่า
  if (!templateId) return { success: false, message: 'ยังไม่ได้ตั้งค่า templateId ใน Settings' };
  if (!folderId) return { success: false, message: 'ยังไม่ได้ตั้งค่า folderId ใน Settings' };

  let copyFile = null;

  try {
    const folder = DriveApp.getFolderById(folderId);
    const fullName = decodeText_(cert.fullName);

    // 1. คัดลอกไฟล์ template
    const templateFile = DriveApp.getFileById(templateId);
    const copyName = 'เกียรติบัตร_' + cert.certNo.replace(/[\/\\]/g, '-');
    copyFile = templateFile.makeCopy(copyName, folder);

    // 2. เปิด Slides แล้วแทนค่า placeholder
    const slides = SlidesApp.openById(copyFile.getId());
    slides.replaceAllText('{{prefix}}', cert.prefix || '');
    slides.replaceAllText('{{fullName}}', fullName);
    slides.replaceAllText('{{certNo}}', cert.certNo);
    slides.replaceAllText('{{date}}', settings.certDate || '');
    slides.replaceAllText('{{projectName}}', settings.projectName || '');
    slides.saveAndClose();

    // 3. Export เป็น PDF
    const pdfBlob = copyFile.getAs('application/pdf').setName(copyName + '.pdf');
    const pdfFile = folder.createFile(pdfBlob);
    let sharingMessage = '';
    try {
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (sharingErr) {
      sharingMessage = ' แต่ตั้งค่าแชร์ PDF ไม่สำเร็จ: ' + sharingErr.message;
    }

    // 4. บันทึก URL กลับลงชีตทันทีหลังสร้าง PDF สำเร็จ
    const pdfUrl = 'https://drive.google.com/file/d/' + pdfFile.getId() + '/view';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const certSheet = ss.getSheetByName(SHEETS.CERTIFICATES);
    certSheet.getRange(cert._rowIndex, 7).setValue(pdfUrl);          // fileUrl (คอลัมน์ 7)
    certSheet.getRange(cert._rowIndex, 8).setValue(new Date());      // createdAt (คอลัมน์ 8)
    SpreadsheetApp.flush();
    clearSheetCache_(SHEETS.CERTIFICATES);

    // 5. ลบไฟล์ Slides ชั่วคราว (เก็บแค่ PDF) โดยไม่ให้กระทบการบันทึกลิงก์
    let cleanupMessage = '';
    try {
      copyFile.setTrashed(true);
    } catch (cleanupErr) {
      cleanupMessage = ' แต่ลบไฟล์ Slides ชั่วคราวไม่สำเร็จ: ' + cleanupErr.message;
    }

    return { success: true, message: 'สร้างสำเร็จ' + sharingMessage + cleanupMessage, certNo: cert.certNo, fileUrl: pdfUrl };

  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * ซ่อมข้อมูลไฟล์ที่เคยสร้างแล้ว แต่ยังไม่ได้บันทึกลิงก์ลงชีต
 * ค้นหา PDF ตามชื่อเกียรติบัตรใน folderId แล้วเติม fileUrl/createdAt พร้อมลบ Slides ชั่วคราวชื่อเดียวกัน
 */
function apiRepairCertificateFiles_() {
  const settings = getSettingsMap_();
  const folderId = settings.folderId;
  if (!folderId) return { success: false, message: 'ยังไม่ได้ตั้งค่า folderId ใน Settings' };

  try {
    const folder = DriveApp.getFolderById(folderId);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const certSheet = ss.getSheetByName(SHEETS.CERTIFICATES);
    const certs = readSheetAsObjects_(SHEETS.CERTIFICATES);
    let linked = 0;
    let trashed = 0;
    let shareFailed = 0;
    let trashFailed = 0;

    certs.forEach(function (cert) {
      const copyName = 'เกียรติบัตร_' + String(cert.certNo).replace(/[\/\\]/g, '-');
      const pdfName = copyName + '.pdf';

      if (!cert.fileUrl) {
        const pdfFiles = folder.getFilesByName(pdfName);
        if (pdfFiles.hasNext()) {
          const pdfFile = pdfFiles.next();
          const pdfUrl = 'https://drive.google.com/file/d/' + pdfFile.getId() + '/view';
          try {
            pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          } catch (sharingErr) {
            shareFailed++;
          }
          certSheet.getRange(cert._rowIndex, 7).setValue(pdfUrl);
          certSheet.getRange(cert._rowIndex, 8).setValue(pdfFile.getDateCreated());
          linked++;
        }
      }

      const slideFiles = folder.getFilesByName(copyName);
      while (slideFiles.hasNext()) {
        const slideFile = slideFiles.next();
        try {
          slideFile.setTrashed(true);
          trashed++;
        } catch (trashErr) {
          trashFailed++;
        }
      }
    });

    SpreadsheetApp.flush();
    clearSheetCache_(SHEETS.CERTIFICATES);
    return {
      success: true,
      message: 'ซ่อมข้อมูลสำเร็จ: เติมลิงก์ ' + linked + ' รายการ, ลบไฟล์ Slides ชั่วคราว ' + trashed + ' ไฟล์, แชร์ PDF ไม่สำเร็จ ' + shareFailed + ' ไฟล์, ลบ Slides ไม่สำเร็จ ' + trashFailed + ' ไฟล์',
      linked: linked,
      trashed: trashed,
      shareFailed: shareFailed,
      trashFailed: trashFailed
    };
  } catch (err) {
    return { success: false, message: 'ซ่อมข้อมูลไม่สำเร็จ: ' + err.message };
  }
}

function repairCertificateFiles() {
  const result = apiRepairCertificateFiles_();
  SpreadsheetApp.getUi().alert(result.message);
}

function authorizeRequiredServices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettingsMap_();
  DriveApp.getRootFolder().getName();

  if (settings.templateId) {
    SlidesApp.openById(settings.templateId).getName();
  }

  SpreadsheetApp.getUi().alert('อนุญาตสิทธิ์เรียบร้อยแล้วสำหรับ Spreadsheet, Drive และ Slides\nSpreadsheet: ' + ss.getName());
}

function checkDriveSettings() {
  const settings = getSettingsMap_();
  const messages = [];

  if (!settings.folderId) {
    messages.push('folderId: ยังไม่ได้ตั้งค่า');
  } else {
    try {
      const folder = DriveApp.getFolderById(settings.folderId);
      messages.push('folderId: เข้าถึงได้ - ' + folder.getName());
    } catch (err) {
      messages.push('folderId: เข้าถึงไม่ได้ - ' + err.message);
    }
  }

  if (!settings.templateId) {
    messages.push('templateId: ยังไม่ได้ตั้งค่า');
  } else {
    try {
      const templateFile = DriveApp.getFileById(settings.templateId);
      messages.push('templateId: เข้าถึงไฟล์ได้ - ' + templateFile.getName());

      try {
        SlidesApp.openById(settings.templateId).getName();
        messages.push('templateId: เปิดเป็น Google Slides ได้');
      } catch (slidesErr) {
        messages.push('templateId: เปิดเป็น Google Slides ไม่ได้ - ' + slidesErr.message);
      }
    } catch (err) {
      messages.push('templateId: เข้าถึงไม่ได้ - ' + err.message);
    }
  }

  SpreadsheetApp.getUi().alert(messages.join('\n'));
}

/**
 * ============================================================
 *  API: saveQuestion (เพิ่ม/แก้ไขคำถาม — สำหรับ Admin/Staff)
 * ============================================================
 */
function apiSaveQuestion_(params) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.QUESTIONS);

    const q = {
      questionId: params.questionId || '',
      part: params.part || '',
      section: params.section || '',
      category: params.category || '',
      questionText: params.questionText || '',
      questionType: params.questionType || 'radio',
      choices: params.choices || '',
      required: (params.required === true || String(params.required).toUpperCase() === 'TRUE'),
      orderNo: Number(params.orderNo) || 0,
      active: (params.active === undefined) ? true : (params.active === true || String(params.active).toUpperCase() === 'TRUE')
    };

    const questions = readSheetAsObjects_(SHEETS.QUESTIONS);
    const existing = questions.find(function (item) { return String(item.questionId) === String(q.questionId); });

    const row = [q.questionId, q.part, q.section, q.category, q.questionText, q.questionType, q.choices, q.required, q.orderNo, q.active];

    if (existing) {
      // แก้ไขแถวเดิม
      sheet.getRange(existing._rowIndex, 1, 1, row.length).setValues([row]);
    } else {
      // เพิ่มใหม่ (สร้าง questionId อัตโนมัติถ้าไม่ได้ส่งมา)
      if (!q.questionId) {
        const nextNum = questions.length + 1;
        row[0] = 'Q' + String(nextNum).padStart(3, '0');
      }
      sheet.appendRow(row);
    }

    clearSheetCache_(SHEETS.QUESTIONS);
    return { success: true, message: 'บันทึกคำถามสำเร็จ', questionId: row[0] };

  } catch (err) {
    return { success: false, message: 'บันทึกไม่สำเร็จ: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================
 *  API: deleteQuestion (ลบคำถาม — จริงๆ คือตั้ง active=false)
 * ============================================================
 */
function apiDeleteQuestion_(params) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.QUESTIONS);
    const questions = readSheetAsObjects_(SHEETS.QUESTIONS);
    const existing = questions.find(function (q) { return String(q.questionId) === String(params.questionId); });

    if (!existing) return { success: false, message: 'ไม่พบคำถาม' };

    // ตั้ง active = false (คอลัมน์ 10) — เก็บข้อมูลไว้ ไม่ลบจริง
    sheet.getRange(existing._rowIndex, 10).setValue(false);
    clearSheetCache_(SHEETS.QUESTIONS);
    return { success: true, message: 'ลบคำถามสำเร็จ' };

  } catch (err) {
    return { success: false, message: 'ลบไม่สำเร็จ: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================
 *  API: getSettings / saveSettings (ตั้งค่าระบบ — Admin)
 * ============================================================
 */
function apiGetSettings_() {
  return { success: true, data: getSettingsMap_() };
}

function apiSaveSettings_(params) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.SETTINGS);
    const settings = readSheetAsObjects_(SHEETS.SETTINGS);

    // params.settings = object {key: value, ...}
    let updates = params.settings;
    if (typeof updates === 'string') updates = JSON.parse(updates);

    Object.keys(updates).forEach(function (key) {
      const existing = settings.find(function (s) { return String(s.key) === String(key); });
      if (existing) {
        sheet.getRange(existing._rowIndex, 2).setValue(updates[key]);
      } else {
        sheet.appendRow([key, updates[key], '']);
      }
    });

    clearSheetCache_(SHEETS.SETTINGS);
    return { success: true, message: 'บันทึกการตั้งค่าสำเร็จ' };

  } catch (err) {
    return { success: false, message: 'บันทึกไม่สำเร็จ: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

// อ่าน Settings เป็น object {key: value}
function getSettingsMap_() {
  const settings = readSheetCached_(SHEETS.SETTINGS, 60);
  const map = {};
  settings.forEach(function (s) { map[s.key] = s.value; });
  return map;
}

/**
 * ============================================================
 *  API: getResults (ผลการประเมิน + คำนวณสถิติ)
 *  ตอนที่ 1: จำนวน + ร้อยละ
 *  ตอนที่ 2: n + Mean + S.D. (รายข้อ + รายด้าน + ภาพรวม) + แปลผล
 *  ตอนที่ 3: แสดงข้อความรายข้อ
 * ============================================================
 */
function apiGetResults_() {
  const questions = readSheetCached_(SHEETS.QUESTIONS, 60)
    .filter(function (q) { return q.active === true || String(q.active).toUpperCase() === 'TRUE'; })
    .sort(function (a, b) { return Number(a.orderNo) - Number(b.orderNo); });

  const responses = readSheetCached_(SHEETS.RESPONSES, 30);

  // แปลง answersJSON เป็น object
  const allAnswers = responses.map(function (r) {
    try { return JSON.parse(r.answersJSON); } catch (e) { return {}; }
  });

  const totalResponses = allAnswers.length;

  // ===== ตอนที่ 1: choice (radio) — จำนวน + ร้อยละ =====
  const part1 = [];
  questions.filter(function (q) { return q.questionType === 'radio'; }).forEach(function (q) {
    const counts = {};
    let answered = 0;
    allAnswers.forEach(function (ans) {
      const v = ans[q.questionId];
      if (v !== undefined && v !== '') {
        counts[v] = (counts[v] || 0) + 1;
        answered++;
      }
    });
    // สร้างผลลัพธ์แต่ละตัวเลือก
    const options = String(q.choices).split('|').map(function (opt) {
      const label = opt.indexOf('=') !== -1 ? opt.split('=')[1].trim() : opt.trim();
      const cnt = counts[label] || 0;
      return {
        label: label,
        count: cnt,
        percent: answered > 0 ? Math.round((cnt / answered) * 1000) / 10 : 0
      };
    });
    part1.push({
      questionId: q.questionId,
      questionText: q.questionText,
      answered: answered,
      options: options
    });
  });

  // ===== ตอนที่ 2: rating — n + Mean + S.D. + แปลผล =====
  const part2Items = [];
  const categoryGroups = {}; // เก็บคะแนนแยกตาม category

  questions.filter(function (q) { return q.questionType === 'rating'; }).forEach(function (q) {
    const scores = [];
    allAnswers.forEach(function (ans) {
      const v = Number(ans[q.questionId]);
      if (!isNaN(v) && v >= 1 && v <= 5) scores.push(v);
    });

    const stat = calcMeanSD_(scores);
    const item = {
      questionId: q.questionId,
      category: q.category,
      questionText: q.questionText,
      n: scores.length,
      mean: stat.mean,
      sd: stat.sd,
      interpret: interpretMean_(stat.mean)
    };
    part2Items.push(item);

    // เก็บเข้ากลุ่ม category
    if (!categoryGroups[q.category]) categoryGroups[q.category] = [];
    categoryGroups[q.category] = categoryGroups[q.category].concat(scores);
  });

  // สรุปรายด้าน (category)
  const part2ByCategory = Object.keys(categoryGroups).map(function (cat) {
    const stat = calcMeanSD_(categoryGroups[cat]);
    return {
      category: cat,
      n: categoryGroups[cat].length,
      mean: stat.mean,
      sd: stat.sd,
      interpret: interpretMean_(stat.mean)
    };
  });

  // สรุปภาพรวมทั้งตอนที่ 2
  let allRatingScores = [];
  Object.keys(categoryGroups).forEach(function (cat) {
    allRatingScores = allRatingScores.concat(categoryGroups[cat]);
  });
  const overallStat = calcMeanSD_(allRatingScores);

  // ===== ตอนที่ 3: textarea — แสดงข้อความรายข้อ =====
  const part3 = [];
  questions.filter(function (q) { return q.questionType === 'textarea'; }).forEach(function (q) {
    const texts = [];
    allAnswers.forEach(function (ans) {
      const v = ans[q.questionId];
      if (v !== undefined && String(v).trim() !== '') {
        texts.push(String(v).trim());
      }
    });
    part3.push({
      questionId: q.questionId,
      questionText: q.questionText,
      answers: texts
    });
  });

  return {
    success: true,
    totalResponses: totalResponses,
    part1: part1,
    part2: {
      items: part2Items,
      byCategory: part2ByCategory,
      overall: {
        n: allRatingScores.length,
        mean: overallStat.mean,
        sd: overallStat.sd,
        interpret: interpretMean_(overallStat.mean)
      }
    },
    part3: part3
  };
}

/**
 * คำนวณค่าเฉลี่ย (Mean) และส่วนเบี่ยงเบนมาตรฐาน (Sample S.D. หาร n-1)
 */
function calcMeanSD_(scores) {
  const n = scores.length;
  if (n === 0) return { mean: 0, sd: 0 };

  const sum = scores.reduce(function (a, b) { return a + b; }, 0);
  const mean = sum / n;

  if (n === 1) return { mean: Math.round(mean * 100) / 100, sd: 0 };

  // Sample S.D. (หาร n-1)
  const variance = scores.reduce(function (acc, x) {
    return acc + Math.pow(x - mean, 2);
  }, 0) / (n - 1);
  const sd = Math.sqrt(variance);

  return {
    mean: Math.round(mean * 100) / 100,  // ทศนิยม 2 ตำแหน่ง
    sd: Math.round(sd * 100) / 100
  };
}

/**
 * แปลผลค่าเฉลี่ยเป็นระดับความพึงพอใจ (เกณฑ์มาตรฐาน)
 */
function interpretMean_(mean) {
  if (mean >= 4.51) return 'มากที่สุด';
  if (mean >= 3.51) return 'มาก';
  if (mean >= 2.51) return 'ปานกลาง';
  if (mean >= 1.51) return 'น้อย';
  if (mean >= 1.00) return 'น้อยที่สุด';
  return '-';
}

/**
 * ============================================================
 *  API: getDashboard (สรุปข้อมูลหน้า Dashboard)
 * ============================================================
 */
function apiGetDashboard_() {
  const certs = readSheetCached_(SHEETS.CERTIFICATES, 30);
  const responses = readSheetCached_(SHEETS.RESPONSES, 30);

  let evaluated = 0, notEvaluated = 0, ready = 0, generated = 0;
  certs.forEach(function (c) {
    if (c.status === 'พร้อมดาวน์โหลด') ready++;
    if (c.status === 'ยังไม่ประเมิน') notEvaluated++;
    if (c.status === 'ประเมินแล้ว') evaluated++;
    if (c.fileUrl) generated++;
  });

  // คะแนนเฉลี่ยรวม (จาก rating ทั้งหมด)
  const questions = readSheetCached_(SHEETS.QUESTIONS, 60)
    .filter(function (q) { return q.questionType === 'rating'; });
  const ratingIds = questions.map(function (q) { return q.questionId; });

  let allScores = [];
  responses.forEach(function (r) {
    try {
      const ans = JSON.parse(r.answersJSON);
      ratingIds.forEach(function (id) {
        const v = Number(ans[id]);
        if (!isNaN(v) && v >= 1 && v <= 5) allScores.push(v);
      });
    } catch (e) { }
  });
  const stat = calcMeanSD_(allScores);

  return {
    success: true,
    totalCertificates: certs.length,
    totalResponses: responses.length,
    readyToDownload: ready,
    notEvaluated: notEvaluated,
    generated: generated,
    avgScore: stat.mean,
    avgInterpret: interpretMean_(stat.mean)
  };
}

/**
 * ============================================================
 *  API: getUsers / saveUser (จัดการผู้ใช้ — เฉพาะ Admin)
 * ============================================================
 */
function apiGetUsers_() {
  const users = readSheetAsObjects_(SHEETS.USERS);
  const data = users.map(function (u) {
    return { role: u.role, password: u.password, displayName: u.displayName };
  });
  return { success: true, data: data };
}

function apiSaveUser_(params) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.USERS);
    const users = readSheetAsObjects_(SHEETS.USERS);
    const existing = users.find(function (u) { return String(u.role) === String(params.role); });

    const row = [params.role, params.password, params.displayName || ''];
    if (existing) {
      sheet.getRange(existing._rowIndex, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    return { success: true, message: 'บันทึกผู้ใช้สำเร็จ' };
  } catch (err) {
    return { success: false, message: 'บันทึกไม่สำเร็จ: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================
 *  ฟังก์ชันทดสอบ API (รันใน Editor เพื่อทดสอบก่อน Deploy)
 * ============================================================
 */
function testApi() {
  // ทดสอบ login
  Logger.log('--- login ---');
  Logger.log(JSON.stringify(apiLogin_({ password: 'Admin1234' })));

  // ทดสอบค้นหา
  Logger.log('--- search ---');
  Logger.log(JSON.stringify(apiSearchCertificate_({ keyword: 'สมชาย' })));

  // ทดสอบดึงคำถาม
  Logger.log('--- questions ---');
  const q = apiGetQuestions_();
  Logger.log('จำนวนคำถาม: ' + q.data.length);

  // ทดสอบ dashboard
  Logger.log('--- dashboard ---');
  Logger.log(JSON.stringify(apiGetDashboard_()));
}
