/**
 * ==============================================================================
 * ระบบรับสมัครเรียนออนไลน์ครบวงจร (ปวช. / ปวส.) - วิทยาลัยอาชีวศึกษา
 * ไฟล์: Notification.gs
 * หน้าที่: ระบบแจ้งเตือนอัตโนมัติผ่าน LINE Official Account (Messaging API) และ Gmail
 * หมายเหตุ: LINE Notify ปิดให้บริการถาวรแล้ว ระบบนี้จึงใช้ LINE Messaging API 100%
 * ==============================================================================
 */

/**
 * แจ้งเตือนเมื่อมีผู้สมัครรายใหม่เข้ามาในระบบ
 * 1. ส่ง LINE Flex Message ไปยัง LINE Admin Group/User
 * 2. ส่งอีเมลยืนยันการรับสมัครไปยังผู้สมัคร พร้อมแนบลิงก์ดาวน์โหลดใบสมัคร PDF
 */
function notifyNewApplication(app, settings) {
  const lineToken = settings["LINE_CHANNEL_ACCESS_TOKEN"];
  const adminTargetId = settings["LINE_ADMIN_GROUP_ID"]; // User ID หรือ Group ID ใน LINE
  const collegeName = settings["COLLEGE_NAME"] || "วิทยาลัยอาชีวศึกษา";

  // 1. ส่ง LINE Flex Message เข้ากลุ่มเจ้าหน้าที่
  if (lineToken && adminTargetId) {
    const flexPayload = createNewAppFlexMessage(app, collegeName);
    sendLinePushMessage(lineToken, adminTargetId, flexPayload);
  }

  // 2. ส่งอีเมลยืนยันไปยังผู้สมัคร
  if (app.email && app.email.includes("@")) {
    sendApplicantConfirmationEmail(app, collegeName);
  }

  // 3. ส่งอีเมลแจ้งเตือนเจ้าหน้าที่
  const adminEmails = settings["ADMIN_EMAILS"];
  if (adminEmails) {
    sendAdminEmailAlert(app, adminEmails, collegeName);
  }
}

/**
 * แจ้งเตือนผู้สมัครเมื่อเจ้าหน้าที่ตรวจเอกสารเสร็จ / เปลี่ยนสถานะ
 */
function notifyStatusUpdate(data) {
  if (!data.email || !data.email.includes("@")) return;

  const subject = `[${data.newStatus}] ผลการตรวจสอบเอกสารการสมัครเรียน - ${data.applicationId}`;
  let statusColor = "#2563eb"; // Blue
  let statusDesc = "";

  if (data.newStatus === "ผ่านการคัดเลือก") {
    statusColor = "#16a34a"; // Green
    statusDesc = "ยินดีด้วย! คุณผ่านการคัดเลือกเข้าศึกษาต่อ กรุณาติดตามกำหนดการมอบตัวและเตรียมเอกสารตัวจริง";
  } else if (data.newStatus === "เอกสารไม่ครบ") {
    statusColor = "#dc2626"; // Red
    statusDesc = `เอกสารของท่านยังไม่สมบูรณ์ กรุณาเข้าสู่ระบบ 'ตรวจสอบสถานะ' เพื่ออัปโหลดเอกสารแก้ไขเพิ่มเติม\n\nสิ่งที่ต้องแก้ไข: ${data.adminNotes || "กรุณาดูรายละเอียดในระบบ"}`;
  } else if (data.newStatus === "มอบตัวแล้ว") {
    statusColor = "#059669";
    statusDesc = "ท่านได้ดำเนินการมอบตัวเป็นนักศึกษาเรียบร้อยแล้ว";
  } else {
    statusDesc = `สถานะปัจจุบัน: ${data.newStatus} ${data.adminNotes ? " (" + data.adminNotes + ")" : ""}`;
  }

  const htmlBody = `
    <div style="font-family: 'Sarabun', Tahoma, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden;">
      <div style="background-color: #1e3a8a; color: #ffffff; padding: 20px; text-align: center;">
        <h2 style="margin: 0; font-size: 20px;">แจ้งผลการตรวจสอบเอกสาร</h2>
        <p style="margin: 5px 0 0; font-size: 14px; opacity: 0.9;">ระบบรับสมัครเรียนออนไลน์</p>
      </div>
      <div style="padding: 24px; color: #374151; font-size: 15px; line-height: 1.6;">
        <p>เรียนคุณ <strong>${data.fullName}</strong>,</p>
        <p>เจ้าหน้าที่ได้ตรวจสอบเอกสารการสมัครเรียนของท่านสำหรับรหัสใบสมัคร <strong>${data.applicationId}</strong> (สาขาวิชา: ${data.majorName}) เรียบร้อยแล้ว</p>
        
        <div style="background-color: #f8fafc; border-left: 4px solid ${statusColor}; padding: 15px; margin: 20px 0; border-radius: 4px;">
          <div style="font-weight: bold; color: ${statusColor}; font-size: 16px; margin-bottom: 5px;">
            สถานะ: ${data.newStatus}
          </div>
          <div style="white-space: pre-line; color: #4b5563;">
            ${statusDesc}
          </div>
        </div>

        <p style="font-size: 13px; color: #6b7280;">ท่านสามารถเข้าสู่ระบบเพื่อตรวจสอบสถานะหรืออัปโหลดเอกสารแก้ไขได้ตลอด 24 ชั่วโมง โดยใช้เลขประจำตัวประชาชน 13 หลัก</p>
      </div>
      <div style="background-color: #f3f4f6; padding: 12px; text-align: center; font-size: 12px; color: #9ca3af;">
        อีเมลนี้เป็นข้อความอัตโนมัติ กรุณาอย่าตอบกลับ
      </div>
    </div>
  `;

  try {
    GmailApp.sendEmail(data.email, subject, "", {
      htmlBody: htmlBody,
      name: "งานรับสมัครนักศึกษา"
    });
  } catch (err) {
    Logger.log("Send status email error: " + err.message);
  }
}

/**
 * ส่งข้อความผ่าน LINE Messaging API (Push Message)
 */
function sendLinePushMessage(channelAccessToken, toId, messages) {
  try {
    const url = "https://api.line.me/v2/bot/message/push";
    const payload = {
      to: toId,
      messages: Array.isArray(messages) ? messages : [messages]
    };

    const options = {
      method: "post",
      contentType: "application/json",
      headers: {
        "Authorization": "Bearer " + channelAccessToken
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    if (code !== 200) {
      Logger.log("LINE Push Failed (" + code + "): " + response.getContentText());
    } else {
      Logger.log("LINE Push Success");
    }
  } catch (err) {
    Logger.log("LINE Push Exception: " + err.message);
  }
}

/**
 * สร้าง Flex Message การ์ดแจ้งเตือนผู้สมัครใหม่สำหรับส่งเข้า LINE
 */
function createNewAppFlexMessage(app, collegeName) {
  return {
    type: "flex",
    altText: `มีผู้สมัครใหม่: ${app.applicationId} (${app.fullName})`,
    contents: {
      type: "bubble",
      size: "mega",
      header: {
        type: "box",
        layout: "vertical",
        backgroundColor: "#1e3a8a",
        paddingAll: "15px",
        contents: [
          {
            type: "text",
            text: "🔔 แจ้งเตือน: มีผู้สมัครใหม่",
            weight: "bold",
            color: "#ffffff",
            size: "md"
          },
          {
            type: "text",
            text: collegeName,
            color: "#93c5fd",
            size: "xs",
            margin: "xs"
          }
        ]
      },
      body: {
        type: "box",
        layout: "vertical",
        spacing: "md",
        contents: [
          {
            type: "box",
            layout: "horizontal",
            contents: [
              { type: "text", text: "เลขที่สมัคร:", size: "sm", color: "#6b7280", flex: 3 },
              { type: "text", text: app.applicationId, size: "sm", weight: "bold", color: "#111827", flex: 5 }
            ]
          },
          {
            type: "box",
            layout: "horizontal",
            contents: [
              { type: "text", text: "ชื่อ-สกุล:", size: "sm", color: "#6b7280", flex: 3 },
              { type: "text", text: app.fullName, size: "sm", weight: "bold", color: "#111827", flex: 5 }
            ]
          },
          {
            type: "box",
            layout: "horizontal",
            contents: [
              { type: "text", text: "ระดับชั้น:", size: "sm", color: "#6b7280", flex: 3 },
              { type: "text", text: app.educationLevel, size: "sm", color: "#111827", flex: 5 }
            ]
          },
          {
            type: "box",
            layout: "horizontal",
            contents: [
              { type: "text", text: "สาขาวิชา:", size: "sm", color: "#6b7280", flex: 3 },
              { type: "text", text: app.majorName, size: "sm", color: "#111827", flex: 5, wrap: true }
            ]
          },
          {
            type: "box",
            layout: "horizontal",
            contents: [
              { type: "text", text: "เบอร์โทร:", size: "sm", color: "#6b7280", flex: 3 },
              { type: "text", text: app.phone, size: "sm", color: "#2563eb", flex: 5 }
            ]
          }
        ]
      },
      footer: {
        type: "box",
        layout: "vertical",
        spacing: "sm",
        contents: [
          {
            type: "text",
            text: "กรุณาเข้าสู่ระบบแอดมินเพื่อตรวจสอบเอกสาร",
            size: "xs",
            color: "#9ca3af",
            align: "center"
          }
        ]
      }
    }
  };
}

/**
 * ส่งอีเมลยืนยันการรับสมัครไปยังผู้สมัคร
 */
function sendApplicantConfirmationEmail(app, collegeName) {
  const subject = `ยืนยันการสมัครเรียนออนไลน์ - รหัส ${app.applicationId} (${collegeName})`;
  const htmlBody = `
    <div style="font-family: 'Sarabun', Tahoma, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden;">
      <div style="background-color: #047857; color: #ffffff; padding: 20px; text-align: center;">
        <h2 style="margin: 0; font-size: 20px;">ลงทะเบียนสมัครเรียนสำเร็จ</h2>
        <p style="margin: 5px 0 0; font-size: 14px; opacity: 0.9;">${collegeName}</p>
      </div>
      <div style="padding: 24px; color: #374151; font-size: 15px; line-height: 1.6;">
        <p>เรียนคุณ <strong>${app.fullName}</strong>,</p>
        <p>วิทยาลัยได้รับข้อมูลและเอกสารการสมัครเรียนของท่านเรียบร้อยแล้ว รายละเอียดการสมัครดังนี้:</p>
        
        <table style="width: 100%; border-collapse: collapse; margin: 15px 0;">
          <tr>
            <td style="padding: 8px; border-bottom: 1px solid #e5e7eb; color: #6b7280; width: 35%;">เลขที่ใบสมัคร</td>
            <td style="padding: 8px; border-bottom: 1px solid #e5e7eb; font-weight: bold; color: #1e3a8a;">${app.applicationId}</td>
          </tr>
          <tr>
            <td style="padding: 8px; border-bottom: 1px solid #e5e7eb; color: #6b7280;">ระดับชั้น</td>
            <td style="padding: 8px; border-bottom: 1px solid #e5e7eb;">${app.educationLevel}</td>
          </tr>
          <tr>
            <td style="padding: 8px; border-bottom: 1px solid #e5e7eb; color: #6b7280;">สาขาวิชา</td>
            <td style="padding: 8px; border-bottom: 1px solid #e5e7eb; font-weight: bold;">${app.majorName}</td>
          </tr>
          <tr>
            <td style="padding: 8px; border-bottom: 1px solid #e5e7eb; color: #6b7280;">สถานะเอกสาร</td>
            <td style="padding: 8px; border-bottom: 1px solid #e5e7eb; color: #d97706; font-weight: bold;">รอการตรวจสอบ</td>
          </tr>
        </table>

        ${app.pdfUrl ? `<p style="text-align: center; margin: 25px 0;"><a href="${app.pdfUrl}" target="_blank" style="background-color: #1e3a8a; color: #ffffff; padding: 10px 20px; text-decoration: none; border-radius: 6px; font-weight: bold; display: inline-block;">📄 ดาวน์โหลดใบสมัคร (PDF)</a></p>` : ""}

        <p style="font-size: 13px; color: #6b7280;">ท่านสามารถติดตามผลการตรวจสอบเอกสารได้ที่หน้าเว็บ 'ตรวจสอบสถานะ' โดยใช้เลขประจำตัวประชาชน 13 หลัก</p>
      </div>
      <div style="background-color: #f3f4f6; padding: 12px; text-align: center; font-size: 12px; color: #9ca3af;">
        งานรับสมัครและแนะแนวการศึกษา | ${collegeName}
      </div>
    </div>
  `;

  try {
    GmailApp.sendEmail(app.email, subject, "", {
      htmlBody: htmlBody,
      name: collegeName
    });
  } catch (err) {
    Logger.log("Send applicant email error: " + err.message);
  }
}

/**
 * ส่งอีเมลแจ้งเตือนเจ้าหน้าที่
 */
function sendAdminEmailAlert(app, adminEmails, collegeName) {
  const subject = `[มีผู้สมัครใหม่] ${app.applicationId} - ${app.fullName} (${app.majorName})`;
  const body = `มีผู้สมัครเรียนรายใหม่เข้ามาในระบบ:\n\n` +
    `เลขที่สมัคร: ${app.applicationId}\n` +
    `ชื่อ-นามสกุล: ${app.fullName}\n` +
    `ระดับชั้น: ${app.educationLevel}\n` +
    `สาขาวิชา: ${app.majorName}\n` +
    `เบอร์โทร: ${app.phone}\n` +
    `อีเมล: ${app.email || "-"}\n\n` +
    `กรุณาเข้าสู่ระบบเจ้าหน้าที่เพื่อตรวจสอบเอกสารและอนุมัติผล`;

  try {
    GmailApp.sendEmail(adminEmails, subject, body, {
      name: "ระบบรับสมัครออนไลน์"
    });
  } catch (err) {
    Logger.log("Send admin alert email error: " + err.message);
  }
}
