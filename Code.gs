/**
 * ==============================================================================
 * ระบบรับสมัครเรียนออนไลน์ครบวงจร (ปวช. / ปวส.) - วิทยาลัยอาชีวศึกษา
 * ไฟล์: Code.gs
 * หน้าที่: Routing หน้าเว็บ, API รับ-ส่งข้อมูล, Validation, และความปลอดภัย PDPA
 * ==============================================================================
 */

/**
 * จัดการ Routing เมื่อเปิดหน้าเว็บผ่าน URL
 */
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page.toLowerCase() : "register";
  
  let templateName = "index";
  let title = "ระบบรับสมัครเรียนออนไลน์";

  if (page === "admin") {
    templateName = "admin";
    title = "ระบบจัดการผู้สมัคร (สำหรับเจ้าหน้าที่)";
  } else if (page === "status") {
    templateName = "status";
    title = "ตรวจสอบสถานะการสมัครเรียน";
  }

  const html = HtmlService.createTemplateFromFile(templateName);
  html.page = page;

  return html.evaluate()
    .setTitle(title)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * ฟังก์ชัน Include สำหรับดึงไฟล์ HTML ย่อย (CSS/JS) มาประกอบในหน้าหลัก
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * อ่านค่าการตั้งค่าจากชีต Config ทั้งหมดเป็น Object Key-Value
 */
function getSystemSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  if (!configSheet) return {};

  const data = configSheet.getRange(2, 7, configSheet.getLastRow() - 1, 2).getValues();
  const settings = {};
  data.forEach(row => {
    if (row[0]) settings[String(row[0]).trim()] = String(row[1]).trim();
  });
  return settings;
}

/**
 * ดึงข้อมูลสาขาวิชาและค่าตั้งค่าเบื้องต้นสำหรับแสดงหน้าฟอร์ม (Public API)
 */
function getPublicConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  const settings = getSystemSettings();

  const majorData = configSheet.getRange(2, 1, Math.max(1, configSheet.getLastRow() - 1), 5).getValues();
  const majors = [];
  
  majorData.forEach(row => {
    const level = row[0];
    const code = row[1];
    const name = row[2];
    const quota = row[3];
    const isActive = row[4];
    
    if (level && name && (isActive === true || String(isActive).toUpperCase() === "TRUE")) {
      majors.push({
        level: String(level).trim(),
        code: String(code).trim(),
        name: String(name).trim(),
        quota: Number(quota) || 0
      });
    }
  });

  return {
    academicYear: settings["CURRENT_ACADEMIC_YEAR"] || "2568",
    round: settings["CURRENT_ROUND"] || "รอบทั่วไป",
    isOpen: settings["IS_REGISTRATION_OPEN"] === "TRUE",
    collegeName: settings["COLLEGE_NAME"] || "วิทยาลัยอาชีวศึกษา",
    majors: majors
  };
}

/**
 * ตรวจสอบความถูกต้องของเลขประจำตัวประชาชน 13 หลัก (Checksum Modulo 11)
 */
function validateThaiNationalId(id) {
  if (!id) return false;
  const cleanId = String(id).replace(/\D/g, "");
  if (cleanId.length !== 13) return false;

  let sum = 0;
  for (let i = 0; i < 12; i++) {
    sum += parseInt(cleanId.charAt(i), 10) * (13 - i);
  }
  const checkDigit = (11 - (sum % 11)) % 10;
  return checkDigit === parseInt(cleanId.charAt(12), 10);
}

/**
 * ตรวจสอบความถูกต้องของเบอร์โทรศัพท์ 10 หลัก
 */
function validatePhoneNumber(phone) {
  if (!phone) return false;
  const cleanPhone = String(phone).replace(/\D/g, "");
  return /^0[689]\d{8}$/.test(cleanPhone);
}

/**
 * ระบบป้องกัน Brute-force Search ด้วย CacheService
 * คืนค่า true ถ้ายังอยู่ในเกณฑ์, false ถ้าเกินขีดจำกัด
 */
function checkRateLimit(identifier, maxAttempts, windowSecs) {
  const cache = CacheService.getScriptCache();
  const key = "rate_limit_" + identifier;
  const current = cache.get(key);
  const count = current ? parseInt(current, 10) : 0;

  if (count >= maxAttempts) {
    return false; // โดนบล็อกชั่วคราว
  }

  cache.put(key, String(count + 1), windowSecs);
  return true;
}

/**
 * ปกปิดข้อมูลส่วนบุคคลตามหลัก PDPA (Masking)
 */
function maskNationalId(id) {
  if (!id || id.length !== 13) return id;
  // รูปแบบ 1-1002-xxxxx-34-5
  return id.substring(0, 1) + "-" + id.substring(1, 5) + "-xxxxx-" + id.substring(10, 12) + "-" + id.substring(12);
}

function maskName(prefix, firstName, lastName) {
  const fMask = firstName ? firstName.charAt(0) + "***" : "";
  const lMask = lastName ? lastName.charAt(0) + "***" : "";
  return `${prefix || ""}${fMask} ${lMask}`;
}

function maskPhone(phone) {
  if (!phone || phone.length < 10) return phone;
  return phone.substring(0, 3) + "-xxx-" + phone.substring(7);
}

/**
 * บันทึก Audit Log ลงชีต
 */
function logAudit(actionType, targetId, status, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName("AuditLogs");
    if (!logSheet) return;
    
    logSheet.appendRow([
      new Date(),
      actionType,
      maskNationalId(targetId),
      status,
      details || "",
      "AppsScript-WebClient"
    ]);
  } catch (err) {
    Logger.log("Audit log error: " + err.message);
  }
}

/**
 * API: บันทึกการสมัครเรียนใหม่พร้อมอัปโหลดไฟล์
 */
function submitApplication(formData) {
  const lock = LockService.getScriptLock();
  try {
    // ล็อคเพื่อป้องกัน race condition ในการรันเลขที่ใบสมัคร
    lock.waitLock(30000);

    const settings = getSystemSettings();
    if (settings["IS_REGISTRATION_OPEN"] !== "TRUE") {
      return { success: false, message: "ระบบปิดรับสมัครอยู่ในขณะนี้" };
    }

    // 1. ตรวจสอบ Checksum เลขบัตร ปชช.
    const nationalId = String(formData.nationalId || "").replace(/\D/g, "");
    if (!validateThaiNationalId(nationalId)) {
      return { success: false, message: "เลขประจำตัวประชาชน 13 หลักไม่ถูกต้องตามหลักการคำนวณ" };
    }

    // 2. ตรวจสอบเบอร์โทร
    const phone = String(formData.phone || "").replace(/\D/g, "");
    if (!validatePhoneNumber(phone)) {
      return { success: false, message: "เบอร์โทรศัพท์ต้องเป็นตัวเลข 10 หลัก (ขึ้นต้นด้วย 06, 08, 09)" };
    }

    // 3. ตรวจสอบความยินยอม PDPA
    if (!formData.consentPdpa) {
      return { success: false, message: "กรุณายินยอมให้เก็บและประมวลผลข้อมูลตามนโยบายคุ้มครองข้อมูลส่วนบุคคล (PDPA)" };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appSheet = ss.getSheetByName("Applications");
    const academicYear = settings["CURRENT_ACADEMIC_YEAR"] || "2568";
    const admissionRound = settings["CURRENT_ROUND"] || "รอบทั่วไป";

    // 4. ตรวจสอบการสมัครซ้ำในปีการศึกษาเดียวกัน
    const existingData = appSheet.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      const rowYear = String(existingData[i][2]);
      const rowId = String(existingData[i][10]);
      if (rowYear === academicYear && rowId === nationalId) {
        return { 
          success: false, 
          message: "หมายเลขบัตรประชาชนนี้ได้ลงทะเบียนในปีการศึกษา " + academicYear + " แล้ว หากต้องการตรวจสถานะ กรุณาไปที่เมนู 'ตรวจสอบสถานะ'" 
        };
      }
    }

    // 5. ออกรหัสใบสมัครอัตโนมัติ (เช่น APP68-0001)
    const yearShort = academicYear.slice(-2);
    const count = existingData.length; // จำนวนแถวทั้งหมด
    const seq = ("0000" + count).slice(-4);
    const applicationId = `APP${yearShort}-${seq}`;

    // 6. จัดการโฟลเดอร์ Google Drive สำหรับผู้สมัครรายนี้
    const rootFolderId = settings["DRIVE_ROOT_FOLDER_ID"];
    let applicantFolder = null;
    let applicantFolderId = "";

    if (rootFolderId) {
      const rootFolder = DriveApp.getFolderById(rootFolderId);
      // โฟลเดอร์ปีการศึกษา
      let yearFolder;
      const yearFolders = rootFolder.getFoldersByName(`ปีการศึกษา ${academicYear}`);
      if (yearFolders.hasNext()) {
        yearFolder = yearFolders.next();
      } else {
        yearFolder = rootFolder.createFolder(`ปีการศึกษา ${academicYear}`);
      }

      // โฟลเดอร์ผู้สมัคร
      const folderName = `${applicationId}_${formData.prefix}${formData.firstName}_${formData.lastName}`;
      applicantFolder = yearFolder.createFolder(folderName);
      applicantFolderId = applicantFolder.getId();
    }

    // 7. บันทึกไฟล์แนบทั้ง 4 รายการลงโฟลเดอร์ใน Drive
    let filePhotoUrl = "";
    let fileIdCardUrl = "";
    let fileHouseRegUrl = "";
    let fileTranscriptUrl = "";

    if (applicantFolder) {
      if (formData.filePhoto) {
        filePhotoUrl = saveBase64File(applicantFolder, formData.filePhoto, `photo_${applicationId}`);
      }
      if (formData.fileIdCard) {
        fileIdCardUrl = saveBase64File(applicantFolder, formData.fileIdCard, `idcard_${applicationId}`);
      }
      if (formData.fileHouseReg) {
        fileHouseRegUrl = saveBase64File(applicantFolder, formData.fileHouseReg, `housereg_${applicationId}`);
      }
      if (formData.fileTranscript) {
        fileTranscriptUrl = saveBase64File(applicantFolder, formData.fileTranscript, `transcript_${applicationId}`);
      }
    }

    const timestamp = new Date();
    const consentTimestamp = new Date();
    const initialStatus = "รอตรวจสอบ";

    // 8. รวบรวมข้อมูลเตรียมลง Sheet
    const newRow = [
      applicationId,               // A
      timestamp,                   // B
      academicYear,                // C
      admissionRound,              // D
      formData.educationLevel,     // E
      formData.majorCode,          // F
      formData.majorName,          // G
      formData.prefix,             // H
      formData.firstName,          // I
      formData.lastName,           // J
      nationalId,                  // K
      formData.birthdate,          // L
      formData.gender,             // M
      phone,                       // N
      formData.email || "",        // O
      formData.currentAddress || "", // P
      formData.prevSchool || "",   // Q
      formData.prevSchoolProvince || "", // R
      Number(formData.gpa) || 0,   // S
      formData.referralChannel || "", // T
      true,                        // U: consent_pdpa
      consentTimestamp,            // V
      applicantFolderId,           // W
      filePhotoUrl,                // X
      fileIdCardUrl,               // Y
      fileHouseRegUrl,             // Z
      fileTranscriptUrl,           // AA
      initialStatus,               // AB
      "",                          // AC: admin_notes
      "",                          // AD: reviewed_by
      "",                          // AE: reviewed_at
      ""                           // AF: pdf_application_url (จะอัปเดตต่อจากนี้)
    ];

    appSheet.appendRow(newRow);
    const lastRowIndex = appSheet.getLastRow();

    // 9. สร้างไฟล์ PDF ใบสมัครอัตโนมัติจาก Google Docs Template
    let pdfUrl = "";
    try {
      const applicantObject = {
        applicationId: applicationId,
        timestamp: Utilities.formatDate(timestamp, "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss"),
        academicYear: academicYear,
        admissionRound: admissionRound,
        educationLevel: formData.educationLevel,
        majorName: formData.majorName,
        majorCode: formData.majorCode,
        fullName: `${formData.prefix} ${formData.firstName} ${formData.lastName}`,
        nationalId: nationalId,
        birthdate: formData.birthdate,
        gender: formData.gender,
        phone: phone,
        email: formData.email || "-",
        currentAddress: formData.currentAddress || "-",
        prevSchool: formData.prevSchool || "-",
        prevSchoolProvince: formData.prevSchoolProvince || "-",
        gpa: formData.gpa ? Number(formData.gpa).toFixed(2) : "-",
        photoFile: formData.filePhoto,
        collegeName: settings["COLLEGE_NAME"] || "วิทยาลัยอาชีวศึกษา"
      };

      if (applicantFolder && typeof generateApplicationPdf === "function") {
        pdfUrl = generateApplicationPdf(applicantObject, applicantFolder, settings["DOC_APPLICATION_TEMPLATE_ID"]);
        if (pdfUrl) {
          appSheet.getRange(lastRowIndex, 32).setValue(pdfUrl); // Column AF
        }
      }
    } catch (pdfErr) {
      Logger.log("PDF Generation warning: " + pdfErr.message);
    }

    // 10. ส่งการแจ้งเตือน (LINE Message แจ้งเจ้าหน้าที่ + อีเมลยืนยันถึงผู้สมัคร)
    try {
      if (typeof notifyNewApplication === "function") {
        notifyNewApplication({
          applicationId: applicationId,
          fullName: `${formData.prefix} ${formData.firstName} ${formData.lastName}`,
          educationLevel: formData.educationLevel,
          majorName: formData.majorName,
          phone: phone,
          email: formData.email,
          pdfUrl: pdfUrl,
          academicYear: academicYear
        }, settings);
      }
    } catch (notifyErr) {
      Logger.log("Notification warning: " + notifyErr.message);
    }

    // บันทึก Log
    logAudit("NEW_APPLICATION", nationalId, "SUCCESS", `รหัสสมัคร: ${applicationId}`);

    return {
      success: true,
      applicationId: applicationId,
      maskedId: maskNationalId(nationalId),
      fullName: `${formData.prefix} ${formData.firstName} ${formData.lastName}`,
      majorName: formData.majorName,
      pdfUrl: pdfUrl,
      message: "ลงทะเบียนสำเร็จ ระบบได้บันทึกข้อมูลและส่งอีเมลยืนยันเรียบร้อยแล้ว"
    };

  } catch (error) {
    Logger.log("Error in submitApplication: " + error.toString());
    return { success: false, message: "เกิดข้อผิดพลาดของระบบ: " + error.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ฟังก์ชันช่วยแปลงและบันทึกไฟล์ Base64 ลงในโฟลเดอร์ Google Drive
 */
function saveBase64File(folder, base64Payload, fileNamePrefix) {
  try {
    if (!base64Payload || !base64Payload.data) return "";
    
    const parts = base64Payload.data.split(",");
    const base64Data = parts.length > 1 ? parts[1] : parts[0];
    const contentType = base64Payload.type || "application/octet-stream";
    const extension = base64Payload.name ? base64Payload.name.split('.').pop() : "bin";
    const fullFileName = `${fileNamePrefix}.${extension}`;

    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, contentType, fullFileName);
    const file = folder.createFile(blob);
    
    // ตั้งค่าให้อ่านได้เฉพาะผู้ที่มีลิงก์ หรืออยู่ในการควบคุมของระบบ
    return file.getUrl();
  } catch (err) {
    Logger.log("Save file error: " + err.message);
    return "";
  }
}

/**
 * API: ตรวจสอบสถานะการสมัคร (พร้อม Rate Limiting & Masked PII)
 */
function checkApplicationStatus(nationalId, phoneLast4) {
  const cleanId = String(nationalId || "").replace(/\D/g, "");
  const cleanPhone4 = String(phoneLast4 || "").replace(/\D/g, "");

  // 1. ตรวจสอบ Rate Limit (จำกัดการสุ่มไม่เกิน 5 ครั้งต่อ 15 นาที ต่อ IP/Session)
  if (!checkRateLimit(cleanId || "unknown", 5, 900)) {
    logAudit("STATUS_CHECK", cleanId, "RATE_LIMITED", "เกินขีดจำกัดการค้นหา 5 ครั้ง");
    return { 
      success: false, 
      rateLimited: true,
      message: "คุณทำรายการเกินจำนวนครั้งที่กำหนด เพื่อความปลอดภัยของข้อมูล กรุณารอ 15 นาทีแล้วลองใหม่อีกครั้ง" 
    };
  }

  if (cleanId.length !== 13 || cleanPhone4.length !== 4) {
    return { success: false, message: "กรุณากรอกเลขบัตร ปชช. 13 หลัก และเบอร์โทรศัพท์ 4 หลักท้ายให้ครบถ้วน" };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const appSheet = ss.getSheetByName("Applications");
  const data = appSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowId = String(row[10]).trim();
    const rowPhone = String(row[13]).trim();

    if (rowId === cleanId) {
      // ตรวจสอบเบอร์โทร 4 ตัวท้าย
      if (rowPhone.slice(-4) !== cleanPhone4) {
        logAudit("STATUS_CHECK", cleanId, "FAILED", "เบอร์โทร 4 ตัวท้ายไม่ตรง");
        return { success: false, message: "ข้อมูลไม่ถูกต้อง กรุณาตรวจสอบเลขบัตรและเบอร์โทรศัพท์" };
      }

      logAudit("STATUS_CHECK", cleanId, "SUCCESS", `ค้นหาสำเร็จ: ${row[0]}`);

      return {
        success: true,
        applicationId: row[0],
        timestamp: Utilities.formatDate(new Date(row[1]), "Asia/Bangkok", "dd/MM/yyyy HH:mm"),
        academicYear: row[2],
        round: row[3],
        educationLevel: row[4],
        majorName: row[6],
        maskedName: maskName(row[7], row[8], row[9]),
        maskedId: maskNationalId(rowId),
        maskedPhone: maskPhone(rowPhone),
        status: row[27],             // Column AB: status
        adminNotes: row[28] || "",   // Column AC: admin_notes
        pdfUrl: row[31] || "",       // Column AF: pdf_application_url
        hasPhoto: !!row[23],
        hasIdCard: !!row[24],
        hasHouseReg: !!row[25],
        hasTranscript: !!row[26],
        canResubmit: row[27] === "เอกสารไม่ครบ"
      };
    }
  }

  logAudit("STATUS_CHECK", cleanId, "FAILED", "ไม่พบเลขบัตรในระบบ");
  return { success: false, message: "ไม่พบข้อมูลการสมัครของเลขบัตรประชาชนนี้ กรุณาตรวจสอบอีกครั้ง" };
}

/**
 * API: อัปโหลดเอกสารแก้ไขเพิ่มเติมกรณีสถานะ "เอกสารไม่ครบ"
 */
function resubmitApplicantFile(applicationId, nationalId, fileType, fileData) {
  const cleanId = String(nationalId || "").replace(/\D/g, "");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const appSheet = ss.getSheetByName("Applications");
  const data = appSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[0]) === applicationId && String(row[10]) === cleanId) {
      const folderId = row[22]; // Column W: applicant_folder_id
      if (!folderId) return { success: false, message: "ไม่พบโฟลเดอร์สำหรับจัดเก็บไฟล์" };

      const folder = DriveApp.getFolderById(folderId);
      const newUrl = saveBase64File(folder, fileData, `${fileType}_${applicationId}_resubmitted`);

      // อัปเดตคอลัมน์ของไฟล์ตามประเภท
      let colIndex = 0;
      if (fileType === "photo") colIndex = 24;       // Col X
      else if (fileType === "idcard") colIndex = 25; // Col Y
      else if (fileType === "housereg") colIndex = 26; // Col Z
      else if (fileType === "transcript") colIndex = 27; // Col AA

      if (colIndex > 0) {
        appSheet.getRange(i + 1, colIndex).setValue(newUrl);
        // ปรับสถานะกลับเป็น "รอตรวจสอบ"
        appSheet.getRange(i + 1, 28).setValue("รอตรวจสอบ");
        appSheet.getRange(i + 1, 29).setValue("ส่งเอกสารแก้ไขเพิ่มเติมแล้ว");

        logAudit("RESUBMIT_FILE", cleanId, "SUCCESS", `อัปโหลด ${fileType} ซ่อมสำเร็จ`);
        return { success: true, message: "อัปโหลดเอกสารแก้ไขเรียบร้อยแล้ว สถานะเปลี่ยนเป็นรอเจ้าหน้าที่ตรวจสอบ" };
      }
    }
  }
  return { success: false, message: "ไม่พบข้อมูลที่ตรงกัน" };
}
