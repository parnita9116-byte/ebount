/**
 * ==============================================================================
 * ระบบรับสมัครเรียนออนไลน์ครบวงจร (ปวช. / ปวส.) - วิทยาลัยอาชีวศึกษา
 * ไฟล์: ExportStd02.gs
 * หน้าที่: ส่งออกข้อมูลผู้มอบตัวตามโครงสร้างรูปแบบที่นำเข้าระบบ ศธ.02 (สอศ.)
 * ==============================================================================
 */

/**
 * ดึงและแปลงข้อมูลนักศึกษาที่มอบตัวแล้วให้อยู่ในรูปแบบโครงสร้าง ศธ.02
 * พร้อมคืนค่าเป็น Base64 Data URL สำหรับดาวน์โหลดเป็นไฟล์ CSV (UTF-8 BOM สำหรับ Excel ภาษาไทย)
 */
function exportStd02Csv(filterYear, filterLevel, filterMajor) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const enrollSheet = ss.getSheetByName("Enrollments");
    const appSheet = ss.getSheetByName("Applications");

    const enrollData = enrollSheet.getDataRange().getValues();
    const appData = appSheet.getDataRange().getValues();

    if (enrollData.length <= 1) {
      return { success: false, message: "ยังไม่มีข้อมูลนักศึกษาที่มอบตัวในระบบ" };
    }

    // สร้าง Map ข้อมูลจาก Applications เพื่อ Join ตาม application_id
    const appMap = {};
    for (let i = 1; i < appData.length; i++) {
      const row = appData[i];
      appMap[String(row[0])] = {
        academicYear: String(row[2]),
        educationLevel: String(row[4]),
        majorCode: String(row[5]),
        majorName: String(row[6]),
        prefix: String(row[7]),
        firstName: String(row[8]),
        lastName: String(row[9]),
        nationalId: String(row[10]),
        birthdate: String(row[11]),
        gender: String(row[12]),
        phone: String(row[13]),
        email: String(row[14]),
        address: String(row[15]),
        prevSchool: String(row[16]),
        prevProvince: String(row[17]),
        gpa: String(row[18])
      };
    }

    // กำหนด Header ตามมาตรฐาน ศธ.02 ของสำนักงานคณะกรรมการการอาชีวศึกษา (สอศ.)
    const std02Headers = [
      "รหัสนักศึกษา",
      "เลขประจำตัวประชาชน",
      "คำนำหน้าชื่อ",
      "ชื่อ",
      "นามสกุล",
      "เพศ",
      "วันเดือนปีเกิด",
      "สัญชาติ",
      "ศาสนา",
      "ระดับการศึกษา",
      "รหัสสาขาวิชา",
      "ชื่อสาขาวิชา",
      "ปีการศึกษาที่เข้า",
      "โรงเรียนเดิม",
      "จังหวัดโรงเรียนเดิม",
      "เกรดเฉลี่ยเดิม",
      "ชื่อผู้ปกครอง",
      "ความสัมพันธ์ผู้ปกครอง",
      "เบอร์โทรผู้ปกครอง",
      "เบอร์โทรนักศึกษา",
      "ที่อยู่",
      "วันที่มอบตัว"
    ];

    const rows = [std02Headers];

    for (let j = 1; j < enrollData.length; j++) {
      const enr = enrollData[j];
      const appId = String(enr[1]);
      const app = appMap[appId] || {};

      const year = app.academicYear || "";
      const level = String(enr[5]) || app.educationLevel || "";
      const majorName = String(enr[6]) || app.majorName || "";

      // ตรวจสอบเงื่อนไขตัวกรอง
      if (filterYear && year !== String(filterYear)) continue;
      if (filterLevel && level !== String(filterLevel)) continue;
      if (filterMajor && majorName !== String(filterMajor) && app.majorCode !== String(filterMajor)) continue;

      const studentId = String(enr[2]);
      const nationalId = String(enr[3]);
      const prefix = app.prefix || "";
      const firstName = app.firstName || "";
      const lastName = app.lastName || "";
      const gender = app.gender || "";
      const birthdate = app.birthdate || "";
      const nationality = "ไทย";
      const religion = "พุทธ"; // ค่าเริ่มต้นตามปกติ หรืออ้างอิงจากฟิลด์เพิ่มเติม
      const majorCode = app.majorCode || "";
      const prevSchool = app.prevSchool || "";
      const prevProvince = app.prevProvince || "";
      const gpa = app.gpa || "";
      const guardianName = String(enr[7]);
      const guardianRelation = String(enr[8]);
      const guardianPhone = String(enr[9]);
      const studentPhone = app.phone || "";
      const address = String(enr[10]) || app.address || "";
      const enrollmentDate = String(enr[13]);

      rows.push([
        `="${studentId}"`,    // ใส่ ="..." เพื่อป้องกันโปรแกรม Excel ตัดเลข 0 ด้านหน้า
        `="${nationalId}"`,
        prefix,
        firstName,
        lastName,
        gender,
        birthdate,
        nationality,
        religion,
        level,
        majorCode,
        majorName,
        year,
        prevSchool,
        prevProvince,
        gpa,
        guardianName,
        guardianRelation,
        `="${guardianPhone}"`,
        `="${studentPhone}"`,
        `"${address.replace(/"/g, '""')}"`,
        enrollmentDate
      ]);
    }

    if (rows.length === 1) {
      return { success: false, message: "ไม่พบข้อมูลที่ตรงกับตัวกรองที่เลือก" };
    }

    // รวมข้อมูลเป็น CSV string
    const csvContent = rows.map(r => r.join(",")).join("\r\n");

    // เติม UTF-8 BOM (\uFEFF) เพื่อให้ Microsoft Excel บน Windows เปิดภาษาไทยได้ถูกต้อง 100% ไม่เป็นภาษาต่างดาว
    const bomCsv = "\uFEFF" + csvContent;
    const base64Data = Utilities.base64Encode(bomCsv, Utilities.Charset.UTF_8);

    const fileName = `std02_export_${filterYear || "all"}_${filterLevel || "all"}_${Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyyMMdd_HHmm")}.csv`;

    return {
      success: true,
      fileName: fileName,
      totalRecords: rows.length - 1,
      csvBase64: "data:text/csv;charset=utf-8;base64," + base64Data
    };

  } catch (err) {
    Logger.log("Export Std02 error: " + err.message);
    return { success: false, message: "เกิดข้อผิดพลาดในการส่งออก: " + err.message };
  }
}
