/**
 * ==============================================================================
 * ระบบรับสมัครเรียนออนไลน์ครบวงจร (ปวช. / ปวส.) - วิทยาลัยอาชีวศึกษา
 * ไฟล์: PdfEngine.gs
 * หน้าที่: สร้างไฟล์ PDF ใบสมัครและสัญญามอบตัวจาก Google Docs Template (แทน AutoCrat)
 * ==============================================================================
 */

/**
 * สร้างไฟล์ PDF ใบสมัคร
 * @param {Object} data ข้อมูลผู้สมัคร
 * @param {Folder} folder โฟลเดอร์ใน Drive ของผู้สมัคร
 * @param {String} templateId รหัส Google Docs Template (ถ้าไม่มีจะใช้ตัวสร้างมาตรฐาน)
 * @return {String} URL ของไฟล์ PDF ใน Google Drive
 */
function generateApplicationPdf(data, folder, templateId) {
  try {
    let docId = "";
    let isTemporaryDoc = false;

    if (templateId && templateId.trim() !== "") {
      // 1. กรณีมี Template Doc ให้ Copy Template มาใช้งาน
      const templateFile = DriveApp.getFileById(templateId.trim());
      const copy = templateFile.makeCopy(`ใบสมัคร_${data.applicationId}_temp`, folder);
      docId = copy.getId();
      isTemporaryDoc = true;

      const doc = DocumentApp.openById(docId);
      const body = doc.getBody();

      // แทนที่ตัวแปร Placeholder ข้อความ
      body.replaceText("{{applicationId}}", data.applicationId || "");
      body.replaceText("{{timestamp}}", data.timestamp || "");
      body.replaceText("{{academicYear}}", data.academicYear || "");
      body.replaceText("{{round}}", data.admissionRound || "");
      body.replaceText("{{educationLevel}}", data.educationLevel || "");
      body.replaceText("{{majorName}}", data.majorName || "");
      body.replaceText("{{majorCode}}", data.majorCode || "");
      body.replaceText("{{fullName}}", data.fullName || "");
      body.replaceText("{{nationalId}}", data.nationalId || "");
      body.replaceText("{{birthdate}}", data.birthdate || "");
      body.replaceText("{{gender}}", data.gender || "");
      body.replaceText("{{phone}}", data.phone || "");
      body.replaceText("{{email}}", data.email || "-");
      body.replaceText("{{currentAddress}}", data.currentAddress || "-");
      body.replaceText("{{prevSchool}}", data.prevSchool || "-");
      body.replaceText("{{prevSchoolProvince}}", data.prevSchoolProvince || "-");
      body.replaceText("{{gpa}}", data.gpa || "-");
      body.replaceText("{{collegeName}}", data.collegeName || "วิทยาลัยอาชีวศึกษา");

      // แทรกรูปถ่ายลงตรงจุด {{photo}} ถ้ามีรูปแนบมา
      insertPhotoPlaceholder(body, data.photoFile);

      doc.saveAndClose();

    } else {
      // 2. Fallback: สร้างเอกสารใบสมัครมาตรฐานให้อัตโนมัติในกรณีที่ยังไม่ได้สร้าง Template Doc
      docId = createDefaultApplicationDoc(data, folder);
      isTemporaryDoc = true;
    }

    // แปลง Google Doc เป็นไฟล์ PDF Blob
    const docFile = DriveApp.getFileById(docId);
    const pdfBlob = docFile.getAs(MimeType.PDF).setName(`ใบสมัคร_${data.applicationId}.pdf`);
    const pdfFile = folder.createFile(pdfBlob);

    // ลบไฟล์ Google Doc ชั่วคราวทิ้ง เพื่อไม่ให้เปลืองพื้นที่
    if (isTemporaryDoc) {
      docFile.setTrashed(true);
    }

    return pdfFile.getUrl();

  } catch (error) {
    Logger.log("Error generating application PDF: " + error.toString());
    return "";
  }
}

/**
 * สร้างไฟล์ PDF สัญญามอบตัวนักศึกษาใหม่
 */
function generateEnrollmentPdf(data, folder, templateId) {
  try {
    let docId = "";
    let isTemporaryDoc = false;

    if (templateId && templateId.trim() !== "") {
      const templateFile = DriveApp.getFileById(templateId.trim());
      const copy = templateFile.makeCopy(`สัญญามอบตัว_${data.studentId}_temp`, folder);
      docId = copy.getId();
      isTemporaryDoc = true;

      const doc = DocumentApp.openById(docId);
      const body = doc.getBody();

      body.replaceText("{{studentId}}", data.studentId || "");
      body.replaceText("{{applicationId}}", data.applicationId || "");
      body.replaceText("{{fullName}}", data.fullName || "");
      body.replaceText("{{nationalId}}", data.nationalId || "");
      body.replaceText("{{educationLevel}}", data.educationLevel || "");
      body.replaceText("{{majorName}}", data.majorName || "");
      body.replaceText("{{guardianName}}", data.guardianName || "");
      body.replaceText("{{guardianRelation}}", data.guardianRelation || "");
      body.replaceText("{{guardianPhone}}", data.guardianPhone || "");
      body.replaceText("{{guardianAddress}}", data.guardianAddress || "-");
      body.replaceText("{{witness1}}", data.witness1 || "");
      body.replaceText("{{witness2}}", data.witness2 || "");
      body.replaceText("{{enrollmentDate}}", data.enrollmentDate || "");
      body.replaceText("{{collegeName}}", data.collegeName || "วิทยาลัยอาชีวศึกษา");

      doc.saveAndClose();

    } else {
      docId = createDefaultEnrollmentDoc(data, folder);
      isTemporaryDoc = true;
    }

    const docFile = DriveApp.getFileById(docId);
    const pdfBlob = docFile.getAs(MimeType.PDF).setName(`สัญญามอบตัว_${data.studentId}.pdf`);
    const pdfFile = folder.createFile(pdfBlob);

    if (isTemporaryDoc) {
      docFile.setTrashed(true);
    }

    return pdfFile.getUrl();

  } catch (error) {
    Logger.log("Error generating enrollment PDF: " + error.toString());
    return "";
  }
}

/**
 * ฟังก์ชันช่วยแทรกรูปถ่ายแทนที่ข้อความ {{photo}}
 */
function insertPhotoPlaceholder(body, photoFile) {
  try {
    if (!photoFile || !photoFile.data) return;
    const found = body.findText("{{photo}}");
    if (!found) return;

    const element = found.getElement();
    const parent = element.getParent();

    const parts = photoFile.data.split(",");
    const base64Data = parts.length > 1 ? parts[1] : parts[0];
    const decoded = Utilities.base64Decode(base64Data);
    const imgBlob = Utilities.newBlob(decoded, photoFile.type || "image/jpeg", "applicant_photo.jpg");

    if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const paragraph = parent.asParagraph();
      const inlineImg = paragraph.appendInlineImage(imgBlob);
      // กำหนดขนาดสัดส่วน 3x4 ซม. (ประมาณ 110 x 140 pt)
      inlineImg.setWidth(110);
      inlineImg.setHeight(140);
      element.removeFromParent();
    }
  } catch (err) {
    Logger.log("Error inserting photo into Doc: " + err.message);
  }
}

/**
 * สร้างเอกสารใบสมัครมาตรฐานอัตโนมัติ (ไม่ต้องพึ่ง Template เริ่มต้น)
 */
function createDefaultApplicationDoc(data, folder) {
  const doc = DocumentApp.create(`ใบสมัคร_${data.applicationId}_temp`);
  const file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  const body = doc.getBody();
  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(40);
  body.setMarginRight(40);

  // หัวเอกสาร
  const header = body.appendParagraph(data.collegeName || "วิทยาลัยอาชีวศึกษา");
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  header.setAlignment(DocumentApp.TextAlignment.CENTER);
  header.setFontSize(16);
  header.setBold(true);

  const subHeader = body.appendParagraph(`ใบสมัครเข้าศึกษาต่อระดับ ${data.educationLevel} ประจำปีการศึกษา ${data.academicYear}`);
  subHeader.setAlignment(DocumentApp.TextAlignment.CENTER);
  subHeader.setFontSize(14);
  subHeader.setBold(true);

  const roundPara = body.appendParagraph(`รอบการสมัคร: ${data.admissionRound} | เลขที่ใบสมัคร: ${data.applicationId} | วันที่สมัคร: ${data.timestamp}`);
  roundPara.setAlignment(DocumentApp.TextAlignment.CENTER);
  roundPara.setFontSize(11);
  roundPara.setItalic(true);

  body.appendHorizontalRule();

  // ตารางข้อมูลผู้สมัคร
  const tableData = [
    ["รหัสประจำตัวประชาชน", data.nationalId, "สาขาวิชาที่สมัคร", `${data.majorName} (${data.majorCode})`],
    ["ชื่อ - นามสกุล", data.fullName, "เพศ", data.gender],
    ["วันเดือนปีเกิด", data.birthdate, "เบอร์โทรศัพท์", data.phone],
    ["อีเมล", data.email, "เกรดเฉลี่ย (GPAX)", data.gpa],
    ["โรงเรียนเดิม", data.prevSchool, "จังหวัดโรงเรียนเดิม", data.prevSchoolProvince],
    ["ที่อยู่ปัจจุบัน", data.currentAddress, "", ""]
  ];

  const table = body.appendTable(tableData);
  table.setBorderWidth(0.5);
  table.setBorderColor("#9ca3af");

  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.setPaddingTop(5);
      cell.setPaddingBottom(5);
      cell.setPaddingLeft(6);
      cell.setPaddingRight(6);
      if (c % 2 === 0) {
        cell.setBackgroundColor("#f3f4f6");
        cell.getChild(0).asParagraph().setBold(true);
      }
    }
  }

  // รวม Cell แถวที่อยู่
  table.getRow(5).getCell(1).setColSpan(3);

  body.appendParagraph("");
  const certPara = body.appendParagraph("ข้าพเจ้าขอรับรองว่าข้อความดังกล่าวข้างต้นเป็นความจริงทุกประการ หากปรากฏในภายหลังว่าข้อความไม่เป็นความจริง ข้าพเจ้ายินยอมให้วิทยาลัยตัดสิทธิ์การเข้าศึกษาต่อทันที");
  certPara.setFontSize(10);
  certPara.setItalic(true);

  body.appendParagraph("");
  body.appendParagraph("");
  const signPara = body.appendParagraph(`ลงชื่อ............................................................ผู้สมัคร\n( ${data.fullName} )\nวันที่........เดือน...................พ.ศ...........`);
  signPara.setAlignment(DocumentApp.TextAlignment.RIGHT);
  signPara.setFontSize(11);

  doc.saveAndClose();
  return doc.getId();
}

/**
 * สร้างเอกสารสัญญามอบตัวมาตรฐานอัตโนมัติ
 */
function createDefaultEnrollmentDoc(data, folder) {
  const doc = DocumentApp.create(`สัญญามอบตัว_${data.studentId}_temp`);
  const file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  const body = doc.getBody();
  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(40);
  body.setMarginRight(40);

  const header = body.appendParagraph(data.collegeName || "วิทยาลัยอาชีวศึกษา");
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  header.setAlignment(DocumentApp.TextAlignment.CENTER);
  header.setFontSize(16);
  header.setBold(true);

  const subHeader = body.appendParagraph(`สัญญามอบตัวนักศึกษาใหม่ ระดับ ${data.educationLevel}`);
  subHeader.setAlignment(DocumentApp.TextAlignment.CENTER);
  subHeader.setFontSize(14);
  subHeader.setBold(true);

  const metaPara = body.appendParagraph(`เลขประจำตัวนักศึกษา: ${data.studentId} | สาขาวิชา: ${data.majorName}`);
  metaPara.setAlignment(DocumentApp.TextAlignment.CENTER);
  metaPara.setFontSize(12);

  body.appendHorizontalRule();

  const contractText = `ทำขึ้น ณ ${data.collegeName} วันที่ ${data.enrollmentDate}\n\n` +
    `ข้าพเจ้า ${data.guardianName} เกี่ยวข้องเป็น ${data.guardianRelation} ของ ${data.fullName} ` +
    `เลขประจำตัวประชาชน ${data.nationalId} อยู่บ้านเลขที่ ${data.guardianAddress || "-"} เบอร์โทรศัพท์ ${data.guardianPhone}\n\n` +
    `ขอทำสัญญามอบตัวนักศึกษาไว้ต่อวิทยาลัย โดยมีข้อตกลงว่าจะกำกับดูแลให้นักศึกษาปฏิบัติตามกฎระเบียบ ข้อบังคับ และวินัยของวิทยาลัยอย่างเคร่งครัด หากนักศึกษาประพฤติผิดระเบียบ ข้าพเจ้ายินยอมให้ทางวิทยาลัยลงโทษตามระเบียบทุกประการ`;

  const p = body.appendParagraph(contractText);
  p.setFontSize(11);
  p.setLineSpacing(1.3);

  body.appendParagraph("");
  body.appendParagraph("");

  const signTableData = [
    [`ลงชื่อ................................................ผู้ปกครอง\n( ${data.guardianName} )`, `ลงชื่อ................................................นักศึกษา\n( ${data.fullName} )`],
    [`ลงชื่อ................................................พยาน\n( ${data.witness1 || "................................................"} )`, `ลงชื่อ................................................พยาน/เจ้าหน้าที่\n( ${data.witness2 || "................................................"} )`]
  ];

  const signTable = body.appendTable(signTableData);
  signTable.setBorderWidth(0);
  for (let r = 0; r < signTable.getNumRows(); r++) {
    for (let c = 0; c < 2; c++) {
      const cell = signTable.getRow(r).getCell(c);
      cell.getChild(0).asParagraph().setAlignment(DocumentApp.TextAlignment.CENTER).setFontSize(10);
      cell.setPaddingBottom(20);
    }
  }

  doc.saveAndClose();
  return doc.getId();
}
