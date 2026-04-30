function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('✉️ CDRF Automail')
      .addItem('Send to Unchecked Rows', 'sendPendingEmails')
      .addToUi();
}

function sendPendingEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  
  const headers = data[0];
  const firstNameIndex = headers.indexOf("Firstname");
  const lastNameIndex = headers.indexOf("Lastname");
  const emailIndex = headers.indexOf("email");
  const linkSentIndex = headers.indexOf("Link Sent");
  
  // Stop if headers are missing
  if (firstNameIndex === -1 || lastNameIndex === -1 || emailIndex === -1 || linkSentIndex === -1) {
    SpreadsheetApp.getUi().alert("Could not find required columns. Please ensure 'Firstname', 'Lastname', 'email', and 'Link Sent' headers exist.");
    return;
  }
  
  let emailsSent = 0;
  let skippedRows = 0;
  const rowsToCheck = []; // collect rows to update in one batch

  // Loop through all rows (starting at 1 to skip the header row)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const isSent = row[linkSentIndex];
    
    // If the checkbox is NOT checked (false or empty), evaluate the row
    if (isSent === false || isSent === "" || isSent === "FALSE") {
      
      const firstName = row[firstNameIndex] ? row[firstNameIndex].toString().trim() : "";
      const lastName = row[lastNameIndex] ? row[lastNameIndex].toString().trim() : "";
      const recipientEmail = row[emailIndex] ? row[emailIndex].toString().trim() : ""; 
      
      if (firstName !== "" && lastName !== "" && recipientEmail !== "") {
        
        const subject = "CageDroneRF (CDRF) Benchmark Access";

        const plainText = `Dear ${firstName} ${lastName},\n\n` +
          `Thanks for reaching out! We are thrilled to see your interest in the CageDroneRF (CDRF) benchmark.\n\n` +
          `Please use the links below to access the materials:\n` +
          `Code and Tooling: https://github.com/DroneGoHome/U-RAPTOR-PUB\n` +
          `Data Access: REDACTED\n\n` +
          `Citation:\n@article{rostami2026cagedronerf, title={CageDroneRF: A Large-Scale RF Benchmark and Toolkit for Drone Perception}, author={Rostami, Mohammad and Faysal, Atik and Xia, Hongtao and Kasasbeh, Hadi and Gao, Ziang and Wang, Huaxia}, journal={arXiv preprint arXiv:2601.03302}, year={2026}}\n\n` +
          `All the Best,\nMohammad Rostami\nPh.D. Candidate, Rowan University`;

        const htmlBody = `
          <div style="font-family: Arial, sans-serif; max-width: 640px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;">
            <!-- Header -->
            <div style="background-color: #1a237e; padding: 28px 32px;">
              <h1 style="margin: 0; color: #ffffff; font-size: 22px; letter-spacing: 0.5px;">CageDroneRF (CDRF)</h1>
              <p style="margin: 4px 0 0; color: #90caf9; font-size: 14px;">Benchmark Access</p>
            </div>

            <!-- Body -->
            <div style="padding: 32px; background-color: #ffffff; color: #212121; line-height: 1.7; font-size: 15px;">
              <p>Dear <strong>${firstName} ${lastName}</strong>,</p>
              <p>
                Thanks for reaching out! We are thrilled to see your interest in the <strong>CageDroneRF (CDRF)</strong> benchmark.
                To accelerate progress toward robust RF perception models, we have made the dataset, code, and trained models publicly available.
                You will also find our open-source tools for data generation, preprocessing, augmentation, and evaluation.
              </p>

              <h3 style="color: #1a237e; border-bottom: 2px solid #e8eaf6; padding-bottom: 6px;">Access the Materials</h3>
              <table style="width: 100%; border-collapse: collapse;">
                <tr>
                  <td style="padding: 10px 0; border-bottom: 1px solid #f0f0f0; width: 140px; color: #555; font-size: 14px;">Code &amp; Tooling</td>
                  <td style="padding: 10px 0; border-bottom: 1px solid #f0f0f0;">
                    <a href="https://github.com/DroneGoHome/U-RAPTOR-PUB" style="color: #1565c0; text-decoration: none; font-weight: bold;">github.com/DroneGoHome/U-RAPTOR-PUB</a>
                  </td>
                </tr>
                <tr>
                  <td style="padding: 10px 0; color: #555; font-size: 14px;">Data Access</td>
                  <td style="padding: 10px 0;">
                    <a href="REDACTED" style="color: #1565c0; text-decoration: none; font-weight: bold;">Google Drive Dataset</a>
                  </td>
                </tr>
              </table>

              <h3 style="color: #1a237e; border-bottom: 2px solid #e8eaf6; padding-bottom: 6px; margin-top: 28px;">Citation</h3>
              <!-- Citation block -->
              <div style="border-radius: 6px; overflow: hidden; border: 1px solid #3e4451;">
                <!-- Top bar -->
                <div style="background-color: #3e4451; padding: 7px 14px;">
                  <span style="color: #abb2bf; font-size: 12px; font-family: monospace; font-weight: bold;">BibTeX</span>
                </div>
                <!-- Syntax-highlighted code -->
                <pre style="margin: 0; padding: 16px; background-color: #282c34; font-size: 13px; font-family: 'Courier New', monospace; line-height: 1.8; overflow-x: auto; white-space: pre-wrap; word-wrap: break-word;"><span style="color:#c678dd;">@article</span><span style="color:#abb2bf;">{</span><span style="color:#e06c75;">rostami2026cagedronerf</span><span style="color:#abb2bf;">,
  </span><span style="color:#61afef;">title</span><span style="color:#abb2bf;">={</span><span style="color:#98c379;">CageDroneRF: A Large-Scale RF Benchmark and Toolkit for Drone Perception</span><span style="color:#abb2bf;">},
  </span><span style="color:#61afef;">author</span><span style="color:#abb2bf;">={</span><span style="color:#98c379;">Rostami, Mohammad and Faysal, Atik and Xia, Hongtao and Kasasbeh, Hadi and Gao, Ziang and Wang, Huaxia</span><span style="color:#abb2bf;">},
  </span><span style="color:#61afef;">journal</span><span style="color:#abb2bf;">={</span><span style="color:#98c379;">arXiv preprint arXiv:2601.03302</span><span style="color:#abb2bf;">},
  </span><span style="color:#61afef;">year</span><span style="color:#abb2bf;">={</span><span style="color:#e5c07b;">2026</span><span style="color:#abb2bf;">}
}</span></pre>
              </div>

              <p style="margin-top: 28px;">Thank you for your time and attention.</p>
              <p style="margin: 0;">All the Best,</p>
            </div>

            <!-- Footer / Signature -->
            <div style="background-color: #1e1e2e; padding: 24px 32px; border-top: 1px solid #333;">
              <table style="border-collapse: collapse; width: 100%;">
                <tr>
                  <td style="vertical-align: middle; padding-right: 20px; width: 80px;">
                    <img src="https://lh7-us.googleusercontent.com/8Vg-tFcjJe83DThpK8qTDUMyTHJVCV0ZbilGGqhyhGBObOrM5FwfijMsM_gSk4pKXRuJB-Q44TzuuRU7L8j8EgfnNQJgPUmgLASbv05r7_Fsa0_wio1WpZQpnQbmMEVTKGUTtUBuMeS6VQhgtY6lWdg"
                      width="72" height="72"
                      style="border-radius: 50%; border: 2px solid #4fc3f7; object-fit: cover; display: block;" />
                  </td>
                  <td style="vertical-align: middle;">
                    <p style="margin: 0; font-weight: bold; color: #4fc3f7; font-size: 16px;">Mohammad Rostami</p>
                    <p style="margin: 2px 0 0; color: #e0e0e0; font-size: 13px; font-weight: bold;">Ph.D. Candidate</p>
                    <p style="margin: 2px 0 0; color: #aaa; font-size: 12px; font-style: italic;">Rowan University</p>
                    <p style="margin: 2px 0 6px; color: #aaa; font-size: 12px; font-style: italic;">Computer and Electrical Engineering</p>
                    <!-- Social Icons -->
                    <a href="https://www.linkedin.com/in/woreom/" style="display: inline-block; margin-right: 6px; background-color: #24292e; padding: 5px 8px; border-radius: 4px; text-decoration: none; vertical-align: middle;">
                      <img src="https://content.linkedin.com/content/dam/me/business/en-us/amp/xbu/linkedin-revised-brand-guidelines/in-logo/fg/brand-inlogo-download-fg-dsk-v01.png.original.png" width="14" height="14" alt="LinkedIn" style="display: block;" />
                    </a>
                    <a href="https://github.com/woreom" style="display: inline-block; margin-right: 6px; background-color: #24292e; padding: 5px 8px; border-radius: 4px; text-decoration: none; vertical-align: middle;">
                      <img src="https://cdn.simpleicons.org/github/ffffff" width="14" height="14" alt="GitHub" style="display: block;" />
                    </a>
                    <a href="https://sites.google.com/view/woreom" style="display: inline-block; background-color: #24292e; padding: 5px 8px; border-radius: 4px; text-decoration: none; vertical-align: middle;">
                      <img src="https://lh3.googleusercontent.com/sitesv/AA5AbUDJuet5ULM2cqEXV9sa6k7Uiw4GbM8qd96ElE6MIsKkqv1REU908RUYN6AhAVctJukWQJMOg3veefLJbbSxnf3t0lwlKI8n7ceHwzIRa7WhZCWpgOhC-H7TOjIjRphn5J6AFDixMY9y0YvWblM3vPKQVuK9xW5jUe-U4Lf2MfcnmNi2osNnVo1gox8=w16383" width="14" height="14" alt="Website" style="display: block;" />
                    </a>
                  </td>
                </tr>
              </table>
            </div>
          </div>
        `;

        MailApp.sendEmail(recipientEmail, subject, plainText, { htmlBody: htmlBody });
        
        rowsToCheck.push(i + 1); // remember this row index, update later
        emailsSent++;
      } else {
        skippedRows++;
      }
    }
  }
  
  // Batch update all checkboxes in one operation instead of one per row
  rowsToCheck.forEach(rowNum => {
    sheet.getRange(rowNum, linkSentIndex + 1).setValue(true);
  });

  if (skippedRows > 0) {
    SpreadsheetApp.getUi().alert(`Done! Sent ${emailsSent} emails successfully.`);
  } else {
    SpreadsheetApp.getUi().alert(`Done! Sent ${emailsSent} emails successfully.`);
  }
}