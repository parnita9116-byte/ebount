/**
 * ==============================================================================
 * ระบบรับสมัครเรียนออนไลน์ครบวงจร (ปวช. / ปวส.) - วิทยาลัยอาชีวศึกษา
 * ไฟล์: InitSheet.gs
 * หน้าที่: สร้างโครงสร้างชีตฐานข้อมูลและค่าเริ่มต้นทั้งหมดอัตโนมัติ
 * ==============================================================================
 */

/**
 * ฟังก์ชันสร้างและตั้งค่าโครงสร้างชีตทั้งหมดอัตโนมัติ
 * กด Run ฟังก์ชันนี้ครั้งแรกเมื่อเปิด Google Sheets ใหม่
 */
function setupDatabaseSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. สร้างหรือดึงชีต Applications
  let appSheet = ss.getSheetByName("Applications");
  if (!appSheet) {
    appSheet = ss.insertSheet("Applications");
  }
  const appHeaders = [
    "application_id",       // A: รหัสการสมัคร (เช่น APP68-0001)
    "timestamp",            // B: วันเวลาที่สมัคร
    "academic_year",        // C: ปีการศึกษา (เช่น 2568)
    "admission_round",      // D: รอบการสมัคร
    "education_level",      // E: ระดับชั้น (ปวช./ปวส.)
    "major_code",           // F: รหัสสาขาวิชา
    "major_name",           // G: ชื่อสาขาวิชา
    "prefix",               // H: คำนำหน้า (นาย/นางสาว/เด็กชาย/เด็กหญิง)
    "first_name",           // I: ชื่อจริง
    "last_name",            // J: นามสกุล
    "national_id",          // K: เลขประจำตัวประชาชน 13 หลัก
    "birthdate",            // L: วันเกิด (YYYY-MM-DD หรือ วว/ดด/ปปปป)
    "gender",               // M: เพศ
    "phone",                // N: เบอร์โทรศัพท์ 10 หลัก
    "email",                // O: อีเมล
    "current_address",      // P: ที่อยู่ปัจจุบัน
    "prev_school",          // Q: โรงเรียนเดิม
    "prev_school_province", // R: จังหวัดโรงเรียนเดิม
    "gpa",                  // S: เกรดเฉลี่ยสะสม (GPAX)
    "referral_channel",     // T: ช่องทางที่รู้จักวิทยาลัย
    "consent_pdpa",         // U: ความยินยอม PDPA (TRUE/FALSE)
    "consent_timestamp",    // V: วันเวลาที่ยินยอม PDPA
    "applicant_folder_id",  // W: รหัสโฟลเดอร์ใน Google Drive
    "file_photo_url",       // X: ลิงก์รูปถ่าย
    "file_id_card_url",     // Y: ลิงก์สำเนาบัตรประชาชน
    "file_house_reg_url",   // Z: ลิงก์สำเนาทะเบียนบ้าน
    "file_transcript_url",  // AA: ลิงก์สำเนา รบ./ปพ.1
    "status",               // AB: สถานะ (รอตรวจสอบ / เอกสารไม่ครบ / ผ่านการคัดเลือก / มอบตัวแล้ว / สละสิทธิ์)
    "admin_notes",          // AC: หมายเหตุจากเจ้าหน้าที่ / เอกสารที่ต้องแก้ไข
    "reviewed_by",          // AD: เจ้าหน้าที่ผู้ตรวจ
    "reviewed_at",          // AE: วันเวลาที่ตรวจ
    "pdf_application_url"   // AF: ลิงก์ PDF ใบสมัคร
  ];
  setSheetHeaders(appSheet, appHeaders, "#1e3a8a"); // Navy Blue

  // 2. สร้างหรือดึงชีต Config
  let configSheet = ss.getSheetByName("Config");
  if (!configSheet) {
    configSheet = ss.insertSheet("Config");
  }
  configSheet.clear();
  
  // ส่วนที่ 1: ตารางระดับชั้นและสาขาวิชา
  const majorHeaders = ["level", "major_code", "major_name", "quota", "is_active"];
  configSheet.getRange(1, 1, 1, majorHeaders.length).setValues([majorHeaders])
    .setBackground("#047857").setFontColor("#ffffff").setFontWeight("bold");
    
  const defaultMajors = [
    ["ปวช.", "20101", "ช่างยนต์", 80, true],
    ["ปวช.", "20102", "ช่างกลโรงงาน", 40, true],
    ["ปวช.", "20103", "ช่างไฟฟ้ากำลัง", 80, true],
    ["ปวช.", "20104", "ช่างอิเล็กทรอนิกส์", 40, true],
    ["ปวช.", "20201", "การบัญชี", 80, true],
    ["ปวช.", "20202", "การตลาด", 40, true],
    ["ปวช.", "20204", "คอมพิวเตอร์ธุรกิจ", 70, true],
    ["ปวช.", "20401", "การโรงแรม", 40, true],
    ["ปวส.", "30101", "เทคนิคเครื่องกล", 40, true],
    ["ปวส.", "30102", "เทคนิคการผลิต", 30, true],
    ["ปวส.", "30103", "ไฟฟ้ากำลัง", 60, true],
    ["ปวส.", "30201", "การบัญชี", 60, true],
    ["ปวส.", "30204", "เทคโนโลยีธุรกิจดิจิทัล", 50, true],
    ["ปวส.", "30401", "การจัดการโรงแรมและบริการ", 30, true]
  ];
  configSheet.getRange(2, 1, defaultMajors.length, defaultMajors[0].length).setValues(defaultMajors);

  // ส่วนที่ 2: ตารางการตั้งค่าระบบ Key-Value (เริ่มคอลัมน์ G)
  const settingHeaders = ["setting_key", "setting_value", "description"];
  configSheet.getRange(1, 7, 1, settingHeaders.length).setValues([settingHeaders])
    .setBackground("#0369a1").setFontColor("#ffffff").setFontWeight("bold");
    
  const defaultSettings = [
    ["CURRENT_ACADEMIC_YEAR", "2568", "ปีการศึกษาปัจจุบันที่เปิดรับ"],
    ["CURRENT_ROUND", "รอบโควตาพิเศษ", "รอบการรับสมัครปัจจุบัน"],
    ["IS_REGISTRATION_OPEN", "TRUE", "สถานะเปิดรับสมัคร (TRUE / FALSE)"],
    ["LINE_CHANNEL_ACCESS_TOKEN", "", "LINE Messaging API Channel Access Token (Long-lived)"],
    ["LINE_ADMIN_GROUP_ID", "", "LINE Group ID หรือ User ID ที่รับแจ้งเตือนสำหรับเจ้าหน้าที่"],
    ["ADMIN_EMAILS", Session.getActiveUser().getEmail(), "อีเมลเจ้าหน้าที่ คั่นด้วยเครื่องหมายจุลภาค (,)"],
    ["ADMIN_PIN_CODE", "123456", "รหัส PIN 6 หลัก สำหรับล็อกอินแอดมิน (เปลี่ยนได้)"],
    ["DRIVE_ROOT_FOLDER_ID", "", "ID โฟลเดอร์ใน Google Drive สำหรับเก็บเอกสารทั้งหมด"],
    ["DOC_APPLICATION_TEMPLATE_ID", "", "ID Google Docs แม่แบบใบสมัคร"],
    ["DOC_ENROLLMENT_TEMPLATE_ID", "", "ID Google Docs แม่แบบสัญญามอบตัว"],
    ["STUDENT_ID_FORMAT", "{YEAR_SHORT}{LEVEL_CODE}{MAJOR_CODE}{SEQ:4}", "รูปแบบรหัส นศ. เช่น 68 2 0101 0001"],
    ["COLLEGE_NAME", "วิทยาลัยอาชีวศึกษาเทคโนโลยีสารสนเทศ", "ชื่อวิทยาลัยสำหรับแสดงในระบบและเอกสาร"]
  ];
  configSheet.getRange(2, 7, defaultSettings.length, defaultSettings[0].length).setValues(defaultSettings);

  // 3. สร้างหรือดึงชีต Enrollments (มอบตัว)
  let enrollSheet = ss.getSheetByName("Enrollments");
  if (!enrollSheet) {
    enrollSheet = ss.insertSheet("Enrollments");
  }
  const enrollHeaders = [
    "enrollment_id",      // A: รหัสการมอบตัว เช่น ENR68-0001
    "application_id",     // B: อ้างอิงรหัสใบสมัคร
    "student_id",         // C: เลขประจำตัวนักศึกษาที่ออกให้
    "national_id",        // D: เลขประจำตัวประชาชน
    "full_name",          // E: ชื่อ-นามสกุลนักศึกษา
    "education_level",    // F: ระดับชั้น
    "major_name",         // G: สาขาวิชา
    "guardian_name",      // H: ชื่อ-สกุล ผู้ปกครอง
    "guardian_relation",  // I: ความเกี่ยวข้อง
    "guardian_phone",     // J: เบอร์โทรผู้ปกครอง
    "guardian_address",   // K: ที่อยู่ผู้ปกครอง
    "witness_1",          // L: พยานคนที่ 1 (เจ้าหน้าที่)
    "witness_2",          // M: พยานคนที่ 2
    "enrollment_date",    // N: วันที่มอบตัว
    "payment_status",     // O: สถานะชำระเงิน/ค่าธรรมเนียม
    "pdf_enrollment_url"  // P: ลิงก์ PDF สัญญามอบตัว
  ];
  setSheetHeaders(enrollSheet, enrollHeaders, "#7c2d12"); // Dark Amber

  // 4. สร้างหรือดึงชีต AuditLogs (ความปลอดภัย & PDPA)
  let logSheet = ss.getSheetByName("AuditLogs");
  if (!logSheet) {
    logSheet = ss.insertSheet("AuditLogs");
  }
  const logHeaders = [
    "timestamp",        // A: วันเวลา
    "action_type",      // B: STATUS_CHECK, NEW_APPLICATION, STATUS_UPDATE, ENROLLMENT, ADMIN_LOGIN
    "target_id",        // C: เลขบัตร ปชช. (Masked เช่น 11002xxxxx345)
    "status",           // D: SUCCESS, FAILED, RATE_LIMITED
    "details",          // E: รายละเอียดเพิ่มเติม
    "user_agent"        // F: ข้อมูลผู้เข้าใช้เบื้องต้น
  ];
  setSheetHeaders(logSheet, logHeaders, "#374151"); // Charcoal Gray

  // ลบ Sheet1 เริ่มต้นทิ้งถ้ามี
  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  Logger.log("✅ สร้างฐานข้อมูลชีตครบทั้ง 4 ชีตเรียบร้อยแล้ว");
}

/**
 * ฟังก์ชันช่วยใส่หัวตาราง จัดกึ่งกลาง ตัวหนา และใส่สีพื้นหลัง
 */
function setSheetHeaders(sheet, headers, bgColor) {
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(bgColor)
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setFrozenRows(1);
}
