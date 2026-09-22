const CONFIG = {
  DRAFT_SUBJECT: "Your Draft Subject Line Here",
  SENDER_NAME: "Your Sender Name Here",

  // Optional. Leave [] to send only what's already attached to the draft.
  EXTRA_ATTACHMENT_IDS: [],

  SEND_DELAY_MS: 300,
};

function createTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("onChangeSendEmail")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onChange()
    .create();

  Logger.log("✅ Trigger installed! Paste rows and emails auto-send.");
}

function onChangeSendEmail(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    Logger.log("⏳ Another execution is running — skipping this run.");
    return;
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    setupHeaders(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const draftToSend = findDraft(CONFIG.DRAFT_SUBJECT);
    if (!draftToSend) {
      Logger.log("❌ Draft not found — nothing sent this run.");
      return;
    }
    const message = draftToSend.getMessage();

    // CC/BCC: pulled live from the draft itself. If the draft has none,
    // these come back as "" and are simply omitted below — no config needed.
    const { cc: draftCc, bcc: draftBcc } = getDraftCcBcc(draftToSend.getId());
    Logger.log(`ℹ️ Draft Cc="${draftCc}" Bcc="${draftBcc}"`);

    // Attachments: whatever is attached to the draft, plus any optional
    // extras from CONFIG. If the draft has none and CONFIG has none, this
    // is just an empty array — no error.
    const extraAttachments = CONFIG.EXTRA_ATTACHMENT_IDS.map(id => {
      try {
        return DriveApp.getFileById(id).getBlob();
      } catch (err) {
        Logger.log(`⚠️ Could not load extra attachment ${id}: ${err.message}`);
        return null;
      }
    }).filter(Boolean);

    const alreadySent = getAlreadySentEmails(sheet, lastRow);

    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

    for (let idx = 0; idx < data.length; idx++) {
      const rowNum = idx + 2;
      const email = (data[idx][0] || "").toString().trim();
      const name = (data[idx][1] || "").toString().trim() || "Participant";
      const status = (data[idx][2] || "").toString().trim();

      if (!email || status !== "") continue;

      const emailList = email.split(",")
        .map(addr => addr.trim())
        .filter(addr => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(addr));

      if (emailList.length === 0) {
        sheet.getRange(rowNum, 3).setValue("Failed: no valid email");
        continue;
      }

      if (MailApp.getRemainingDailyQuota() <= 0) {
        Logger.log("⚠️ Daily email quota exhausted — stopping.");
        break;
      }

      sendRowEmails(sheet, rowNum, emailList, name, message, extraAttachments, alreadySent, draftCc, draftBcc);
      Utilities.sleep(CONFIG.SEND_DELAY_MS);
    }
  } finally {
    lock.releaseLock();
  }
}

function findDraft(subject) {
  const target = subject.trim().toLowerCase();
  const drafts = GmailApp.getDrafts();
  for (let i = 0; i < drafts.length; i++) {
    if (drafts[i].getMessage().getSubject().trim().toLowerCase() === target) {
      return drafts[i];
    }
  }
  return null;
}

function getDraftCcBcc(draftId) {
  try {
    const draft = Gmail.Users.Drafts.get("me", draftId, { format: "full" });
    const headers = (draft.message && draft.message.payload && draft.message.payload.headers) || [];
    const findHeader = name => {
      const h = headers.find(h => h.name.toLowerCase() === name);
      return h ? h.value : "";
    };
    return { cc: findHeader("cc"), bcc: findHeader("bcc") };
  } catch (err) {
    Logger.log("⚠️ Advanced Gmail API not available/enabled (" + err.message +
      ") — falling back to basic GmailApp Cc/Bcc reading, which may be unreliable for drafts. " +
      "Enable it under Services → Gmail API in the Apps Script editor if you rely on draft Cc/Bcc.");
    return { cc: "", bcc: "" };
  }
}

function getAlreadySentEmails(sheet, lastRow) {
  const sent = new Set();
  if (lastRow < 2) return sent;

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  data.forEach(row => {
    const status = (row[2] || "").toString();

    if (status === "Sent") {
      (row[0] || "").toString().split(",")
        .map(e => e.trim().toLowerCase())
        .filter(Boolean)
        .forEach(e => sent.add(e));
    } else if (status.startsWith("Partial")) {
      const match = status.match(/sent to (.*?)(?:;\s*failed:|$)/i);
      if (match) {
        match[1].split(",")
          .map(e => e.trim().toLowerCase())
          .filter(Boolean)
          .forEach(e => sent.add(e));
      }
    }
  });

  return sent;
}

function sendRowEmails(sheet, row, emailList, name, message, extraAttachments, alreadySent, draftCc, draftBcc) {
  const attachments = message.getAttachments().concat(extraAttachments || []);

  const succeeded = [];
  const failed = [];
  const duplicates = [];

  emailList.forEach(recipient => {
    const key = recipient.toLowerCase();

    if (alreadySent.has(key)) {
      duplicates.push(recipient);
      return;
    }

    try {
      const subject = message.getSubject().replace(/{{name}}/g, name);
      const plainBody = message.getPlainBody().replace(/{{name}}/g, name);
      const htmlBody = message.getBody().replace(/{{name}}/g, name);

      const mailOptions = {
        htmlBody: htmlBody,
        attachments: attachments,
        name: CONFIG.SENDER_NAME,
      };
      if (draftCc) mailOptions.cc = draftCc;
      if (draftBcc) mailOptions.bcc = draftBcc;

      GmailApp.sendEmail(recipient, subject, plainBody, mailOptions);
      succeeded.push(recipient);
      alreadySent.add(key);
      Logger.log(`✅ Sent to ${name} <${recipient}>`);
    } catch (err) {
      failed.push(recipient + " (" + err.message + ")");
      Logger.log(`❌ Failed for ${name} <${recipient}>: ` + err.message);
    }
  });

  const parts = [];
  if (succeeded.length) parts.push("sent to " + succeeded.join(", "));
  if (duplicates.length) parts.push("duplicate (already sent): " + duplicates.join(", "));
  if (failed.length) parts.push("failed: " + failed.join("; "));

  let statusValue;
  if (succeeded.length === 0 && failed.length === 0 && duplicates.length > 0) {
    statusValue = "Duplicate — already sent: " + duplicates.join(", ");
  } else if (failed.length === 0 && duplicates.length === 0) {
    statusValue = "Sent";
  } else {
    statusValue = "Partial — " + parts.join("; ");
  }

  sheet.getRange(row, 3).setValue(statusValue);
}

function setupHeaders(sheet) {
  if (sheet.getLastRow() === 0 || sheet.getRange("A1").getValue() === "") {
    sheet.getRange("A1:C1").setValues([["Email", "Name", "Status"]]);
    sheet.getRange("A1:C1").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
}
