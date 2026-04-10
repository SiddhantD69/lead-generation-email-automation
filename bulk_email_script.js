function sendBulkEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");

  if (!sheet) {
    throw new Error("Sheet 'Sheet1' not found. Please rename your sheet.");
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const company = data[i][0];
    const email = data[i][1] ? String(data[i][1]).trim() : "";
    const cc = data[i][2] ? String(data[i][2]).trim() : "";
    const status = data[i][3];

    // Skip if already sent
    if (status === "Email Sent") continue;

    // Validate email
    if (!email.includes("@") || !email.includes(".")) {
      sheet.getRange(i + 1, 4).setValue("Invalid Email ❌");
      continue;
    }

    // Subject (Personalized)
    const subject = "Company Profile - " + company;

    // HTML Email Body
    const htmlBody = `
      <p>Dear Sir,</p>

      <p>We would like to introduce <strong>M/S R K Power, Pune</strong> as a manufacturer of 
      <strong>High Voltage Rectifier Transformers</strong>.</p>

      <p>We have developed <strong>High-Frequency Rectifier Technology</strong> which improves 
      performance and increases power output by <strong>25% to 30%</strong>.</p>

      <p>This leads to improved ESP efficiency, reduced emissions, and better energy savings.</p>

      <p>We request you to kindly register us in your vendor list and share your valuable inquiries.</p>

      <br>

      <p>Thanks & Regards,<br>
      <strong>Siddhant Deshpande</strong><br>
      Sales & Marketing<br>
      R.K Power</p>
    `;

    try {
      // Attachments from Google Drive
      const attachments = [
        DriveApp.getFileById("1M-1rBk8647gtoQUoINEATFd4JtnpqY9d"),
        DriveApp.getFileById("1bU9Hb2mpGCJB_UVrYEkBaB9MPb1uRtWE"),
        DriveApp.getFileById("1eOYic61EzYRbYfmAhFd50cbDIZsk94Im")
      ];

      GmailApp.sendEmail(email, subject, "", {
        htmlBody: htmlBody,
        attachments: attachments,
        cc: cc
      });

      // Update status
      sheet.getRange(i + 1, 4).setValue("Email Sent ✅");

      // Timestamp
      sheet.getRange(i + 1, 5).setValue(new Date());

    } catch (error) {
      sheet.getRange(i + 1, 4).setValue("Failed ❌");
      sheet.getRange(i + 1, 5).setValue(error.message);
    }
  }
}
