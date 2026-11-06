/**
 * ==============================================================================
 * ระบบรับสมัครเรียนออนไลน์ครบวงจร (ปวช. / ปวส.) - วิทยาลัยอาชีวศึกษา
 * ไฟล์: AdminService.gs
 * หน้าที่: การยืนยันสิทธิ์เจ้าหน้าที่, ดึงข้อมูลผู้สมัคร, ตรวจเอกสาร, มอบตัว, และสถิติ
 * ==============================================================================
 */

/**
 * ตรวจสอบสิทธิ์การเข้าใช้งานฝั่งเจ้าหน้าที่
 * รองรับทั้ง Google Workspace Domain Email และ PIN Code สำหรับเครื่องส่วนตัว
 */
function verifyAdminAuth(pin) {
  try {
    const settings = getSystemSettings();
    const activeEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
    const adminEmails = (settings["ADMIN_EMAILS"] || "")
      .toLowerCase()
      .split(",")
      .map(e => e.trim())
      .filter(e => e.length > 0);

    // 1. ตรวจสอบจาก Google Account ปัจจุบันถ้าเป็นเจ้าหน้าที่ในโดเมน
    if (activeEmail && adminEmails.includes(activeEmail)) {
      return {
        authenticated: true,
        userEmail: activeEmail,
        authMethod: "GOOGLE_ACCOUNT",
        role: "ADMIN"
      };
    }

    // 2. ตรวจสอบผ่านรหัส PIN 6 หลัก
    const adminPin = settings["ADMIN_PIN_CODE"] || "123456";
    if (pin && String(pin).trim() === adminPin) {
      logAudit("ADMIN_LOGIN", "ADMIN", "SUCCESS", "ล็อกอินผ่าน PIN สำเร็จ");
      return {
        authenticated: true,
        userEmail: activeEmail || "staff@college.local",
        authMethod: "PIN_CODE",
        role: "ADMIN"
      };
    }

    logAudit("ADMIN_LOGIN", "ADMIN", "FAILED", "รหัส PIN หรืออีเมลไม่ถูกต้อง");
    return {
      authenticated: false,
      message: "คุณไม่มีสิทธิ์เข้าถึงส่วนงานเจ้าหน้าที่ หรือรหัส PIN ไม่ถูกต้อง"
    };

  } catch (err) {
    return { authenticated: false, message: "เกิดข้อผิดพลาดในการตรวจสอบสิทธิ์: " + err.message };
  }
}

/**
 * ดึงรายการผู้สมัครตามตัวกรอง (ปีการศึกษา, ระดับชั้น, สาขาวิชา, สถานะ)
 */
function getAdminApplicants(filters) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appSheet = ss.getSheetByName("Applications");
    const data = appSheet.getDataRange().getValues();

    if (data.length <= 1) return [];

    const fYear = filters && filters.year ? String(filters.year) : "";
    const fLevel = filters && filters.level ? String(filters.level) : "";
    const fMajor = filters && filters.major ? String(filters.major) : "";
    const fStatus = filters && filters.status ? String(filters.status) : "";
    const fSearch = filters && filters.search ? String(filters.search).toLowerCase() : "";

    const results = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const appId = String(row[0]);
      const dateStr = Utilities.formatDate(new Date(row[1]), "Asia/Bangkok", "dd/MM/yyyy HH:mm");
      const year = String(row[2]);
      const round = String(row[3]);
      const level = String(row[4]);
      const majorCode = String(row[5]);
      const majorName = String(row[6]);
      const prefix = String(row[7]);
      const firstName = String(row[8]);
      const lastName = String(row[9]);
      const fullName = `${prefix} ${firstName} ${lastName}`;
      const nationalId = String(row[10]);
      const phone = String(row[13]);
      const email = String(row[14]);
      const prevSchool = String(row[16]);
      const gpa = row[18];
      const status = String(row[27] || "รอตรวจสอบ");
      const notes = String(row[28] || "");
      const pdfUrl = String(row[31] || "");

      // กรองเงื่อนไข
      if (fYear && year !== fYear) continue;
      if (fLevel && level !== fLevel) continue;
      if (fMajor && majorName !== fMajor && majorCode !== fMajor) continue;
      if (fStatus && status !== fStatus) continue;
      if (fSearch) {
        const matchesSearch = appId.toLowerCase().includes(fSearch) ||
                              fullName.toLowerCase().includes(fSearch) ||
                              nationalId.includes(fSearch) ||
                              phone.includes(fSearch);
        if (!matchesSearch) continue;
      }

      results.push({
        rowIndex: i + 1,
        applicationId: appId,
        timestamp: dateStr,
        academicYear: year,
        round: round,
        educationLevel: level,
        majorCode: majorCode,
        majorName: majorName,
        prefix: prefix,
        firstName: firstName,
        lastName: lastName,
        fullName: fullName,
        nationalId: nationalId,
        maskedId: maskNationalId(nationalId),
        birthdate: String(row[11]),
        gender: String(row[12]),
        phone: phone,
        email: email,
        currentAddress: String(row[15]),
        prevSchool: prevSchool,
        prevSchoolProvince: String(row[17]),
        gpa: gpa ? Number(gpa).toFixed(2) : "-",
        referralChannel: String(row[19]),
        filePhotoUrl: String(row[23]),
        fileIdCardUrl: String(row[24]),
        fileHouseRegUrl: String(row[25]),
        fileTranscriptUrl: String(row[26]),
        status: status,
        adminNotes: notes,
        reviewedBy: String(row[29]),
        reviewedAt: row[30] ? Utilities.formatDate(new Date(row[30]), "Asia/Bangkok", "dd/MM/yyyy HH:mm") : "",
        pdfUrl: pdfUrl
      });
    }

    // เรียงลำดับจากล่าสุดไปเก่าสุด
    return results.reverse();

  } catch (err) {
    Logger.log("Error getAdminApplicants: " + err.message);
    return [];
  }
}

/**
 * เจ้าหน้าที่อัปเดตสถานะและบันทึกหมายเหตุ
 */
function updateApplicantStatus(applicationId, newStatus, adminNotes, reviewerEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appSheet = ss.getSheetByName("Applications");
    const data = appSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === applicationId) {
        const rowNum = i + 1;
        const now = new Date();
        const nationalId = String(data[i][10]);
        const applicantEmail = String(data[i][14]);
        const fullName = `${data[i][7]} ${data[i][8]} ${data[i][9]}`;
        const majorName = String(data[i][6]);

        appSheet.getRange(rowNum, 28).setValue(newStatus);     // Col AB: status
        appSheet.getRange(rowNum, 29).setValue(adminNotes || ""); // Col AC: admin_notes
        appSheet.getRange(rowNum, 30).setValue(reviewerEmail || Session.getActiveUser().getEmail()); // Col AD
        appSheet.getRange(rowNum, 31).setValue(now);           // Col AE: reviewed_at

        logAudit("STATUS_UPDATE", nationalId, "SUCCESS", `เปลี่ยนเป็น [${newStatus}] เหตุผล: ${adminNotes}`);

        // ส่งการแจ้งเตือนผู้สมัครเมื่อสถานะเปลี่ยน (เฉพาะสถานะสำคัญ)
        try {
          if (typeof notifyStatusUpdate === "function") {
            notifyStatusUpdate({
              applicationId: applicationId,
              fullName: fullName,
              email: applicantEmail,
              majorName: majorName,
              newStatus: newStatus,
              adminNotes: adminNotes
            });
          }
        } catch (nErr) {
          Logger.log("Notify status error: " + nErr.message);
        }

        return { success: true, message: `อัปเดตสถานะผู้สมัคร ${applicationId} เป็น "${newStatus}" เรียบร้อยแล้ว` };
      }
    }

    return { success: false, message: "ไม่พบรหัสผู้สมัครในระบบ" };

  } catch (err) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + err.message };
  }
}

/**
 * บันทึกการมอบตัวและออกรหัสนักศึกษาตามรูปแบบที่กำหนด
 */
function enrollStudent(enrollData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const enrollSheet = ss.getSheetByName("Enrollments");
    const appSheet = ss.getSheetByName("Applications");
    const settings = getSystemSettings();

    const applicationId = enrollData.applicationId;
    const appData = appSheet.getDataRange().getValues();
    let targetRowIndex = -1;
    let targetApp = null;

    for (let i = 1; i < appData.length; i++) {
      if (String(appData[i][0]) === applicationId) {
        targetRowIndex = i + 1;
        targetApp = appData[i];
        break;
      }
    }

    if (!targetApp) {
      return { success: false, message: "ไม่พบข้อมูลใบสมัคร" };
    }

    const academicYear = String(targetApp[2]);
    const educationLevel = String(targetApp[4]);
    const majorCode = String(targetApp[5]);
    const majorName = String(targetApp[6]);
    const nationalId = String(targetApp[10]);
    const fullName = `${targetApp[7]} ${targetApp[8]} ${targetApp[9]}`;
    const applicantFolderId = String(targetApp[22]);

    // 1. ตรวจสอบว่าเคยมอบตัวแล้วหรือไม่
    const existingEnrollData = enrollSheet.getDataRange().getValues();
    for (let j = 1; j < existingEnrollData.length; j++) {
      if (String(existingEnrollData[j][1]) === applicationId) {
        return { 
          success: false, 
          message: `ผู้สมัครรายนี้ได้มอบตัวแล้ว ได้รับรหัสนักศึกษา ${existingEnrollData[j][2]}` 
        };
      }
    }

    // 2. คำนวณรหัสนักศึกษาตาม Format
    // Pattern เช่น 68 2 0101 0001
    const yearShort = academicYear.slice(-2);
    // รหัสระดับชั้นตามมาตรฐาน สอศ.: ปวช. = 2, ปวส. = 3
    const levelCode = educationLevel === "ปวช." ? "2" : "3";
    // ตัดรหัสสาขา 4 หลักท้าย (เช่น 20101 -> 0101 หรือใช้ majorCode ตรงๆ)
    const majorCleanCode = majorCode.length > 4 ? majorCode.slice(-4) : majorCode;

    // นับจำนวนคนที่เคยออกรหัสในสาขาและระดับเดียวกันเพื่อเรียงลำดับ (Sequence)
    let majorCount = 0;
    for (let k = 1; k < existingEnrollData.length; k++) {
      const existingStdId = String(existingEnrollData[k][2]);
      if (existingStdId.startsWith(`${yearShort}${levelCode}${majorCleanCode}`)) {
        majorCount++;
      }
    }
    const seq = ("0000" + (majorCount + 1)).slice(-4);
    const studentId = `${yearShort}${levelCode}${majorCleanCode}${seq}`;

    // 3. ออกรหัสการมอบตัว (ENR68-0001)
    const enrollId = `ENR${yearShort}-${("0000" + existingEnrollData.length).slice(-4)}`;
    const enrollmentDate = enrollData.enrollmentDate || Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy");

    // 4. สร้าง PDF สัญญามอบตัว
    let enrollmentPdfUrl = "";
    if (applicantFolderId && typeof generateEnrollmentPdf === "function") {
      try {
        const folder = DriveApp.getFolderById(applicantFolderId);
        enrollmentPdfUrl = generateEnrollmentPdf({
          studentId: studentId,
          applicationId: applicationId,
          fullName: fullName,
          nationalId: nationalId,
          educationLevel: educationLevel,
          majorName: majorName,
          guardianName: enrollData.guardianName,
          guardianRelation: enrollData.guardianRelation,
          guardianPhone: enrollData.guardianPhone,
          guardianAddress: enrollData.guardianAddress || targetApp[15],
          witness1: enrollData.witness1 || "เจ้าหน้าที่ทะเบียน",
          witness2: enrollData.witness2 || "หัวหน้างานรับสมัคร",
          enrollmentDate: enrollmentDate,
          collegeName: settings["COLLEGE_NAME"] || "วิทยาลัยอาชีวศึกษา"
        }, folder, settings["DOC_ENROLLMENT_TEMPLATE_ID"]);
      } catch (pdfErr) {
        Logger.log("Enrollment PDF error: " + pdfErr.message);
      }
    }

    // 5. บันทึกแถวใหม่ลงชีต Enrollments
    enrollSheet.appendRow([
      enrollId,
      applicationId,
      studentId,
      nationalId,
      fullName,
      educationLevel,
      majorName,
      enrollData.guardianName,
      enrollData.guardianRelation,
      enrollData.guardianPhone,
      enrollData.guardianAddress || targetApp[15],
      enrollData.witness1 || "เจ้าหน้าที่ทะเบียน",
      enrollData.witness2 || "หัวหน้างานรับสมัคร",
      enrollmentDate,
      enrollData.paymentStatus || "ชำระครบแล้ว",
      enrollmentPdfUrl
    ]);

    // 6. อัปเดตสถานะใน Applications เป็น "มอบตัวแล้ว"
    appSheet.getRange(targetRowIndex, 28).setValue("มอบตัวแล้ว");
    appSheet.getRange(targetRowIndex, 29).setValue(`มอบตัวสำเร็จ รหัสนักศึกษา: ${studentId}`);

    logAudit("ENROLLMENT", nationalId, "SUCCESS", `ออกรหัสนักศึกษา: ${studentId}`);

    return {
      success: true,
      studentId: studentId,
      enrollmentId: enrollId,
      pdfUrl: enrollmentPdfUrl,
      message: `มอบตัวสำเร็จ ออกรหัสนักศึกษา: ${studentId}`
    };

  } catch (err) {
    return { success: false, message: "เกิดข้อผิดพลาดในการมอบตัว: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ประมวลผลข้อมูลสถิติสำหรับแดชบอร์ด Chart.js
 */
function getDashboardStats(selectedYear) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appSheet = ss.getSheetByName("Applications");
    const enrollSheet = ss.getSheetByName("Enrollments");
    const settings = getSystemSettings();
    const year = selectedYear || settings["CURRENT_ACADEMIC_YEAR"] || "2568";

    const appData = appSheet.getDataRange().getValues();
    const enrollData = enrollSheet.getDataRange().getValues();

    let totalApplicants = 0;
    let pendingCount = 0;
    let incompleteCount = 0;
    let passedCount = 0;
    let enrolledCount = 0;

    const majorCounts = {};
    const levelCounts = { "ปวช.": 0, "ปวส.": 0 };
    const channelCounts = {};
    const schoolCounts = {};
    const provinceCounts = {};

    for (let i = 1; i < appData.length; i++) {
      const row = appData[i];
      const rowYear = String(row[2]);
      if (rowYear !== year) continue;

      totalApplicants++;
      const status = String(row[27]);
      const level = String(row[4]);
      const major = String(row[6]);
      const channel = String(row[19]) || "ไม่ระบุ";
      const school = String(row[16]) || "ไม่ระบุ";
      const province = String(row[17]) || "ไม่ระบุ";

      if (status === "รอตรวจสอบ") pendingCount++;
      else if (status === "เอกสารไม่ครบ") incompleteCount++;
      else if (status === "ผ่านการคัดเลือก") passedCount++;
      else if (status === "มอบตัวแล้ว") {
        passedCount++;
        enrolledCount++;
      }

      // นับแยกตามระดับ
      if (levelCounts[level] !== undefined) {
        levelCounts[level]++;
      }

      // นับแยกตามสาขา
      if (major) {
        majorCounts[major] = (majorCounts[major] || 0) + 1;
      }

      // ช่องทาง
      channelCounts[channel] = (channelCounts[channel] || 0) + 1;

      // โรงเรียนเดิม
      if (school !== "ไม่ระบุ") {
        schoolCounts[school] = (schoolCounts[school] || 0) + 1;
      }

      // จังหวัด
      if (province !== "ไม่ระบุ") {
        provinceCounts[province] = (provinceCounts[province] || 0) + 1;
      }
    }

    // Top 5 โรงเรียนเดิม
    const topSchools = Object.keys(schoolCounts)
      .map(k => ({ name: k, count: schoolCounts[k] }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 5);

    return {
      academicYear: year,
      summary: {
        total: totalApplicants,
        pending: pendingCount,
        incomplete: incompleteCount,
        passed: passedCount,
        enrolled: enrolledCount
      },
      majors: majorCounts,
      levels: levelCounts,
      channels: channelCounts,
      provinces: provinceCounts,
      topSchools: topSchools
    };

  } catch (err) {
    Logger.log("Error getDashboardStats: " + err.message);
    return null;
  }
}
