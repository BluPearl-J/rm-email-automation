/**
 * SMART CRM EMAIL FOLLOW-UP AUTOMATION
 * =====================================
 * Built by Joy O. Reuben-Atohor
 * 
 * What it does:
 * - Tracks leads/contacts in Google Sheets (your CRM)
 * - Automatically sends personalised follow-up emails at configured intervals
 * - Logs every email sent with timestamp
 * - Flags overdue contacts in red
 * - Creates Google Calendar reminders for high-priority leads
 * - Sends you a daily digest of all pending follow-ups
 * 
 * This is the EXACT automation a VA role like Sadecia Tutors needs.
 */

// ============================================================
// CONFIGURATION — edit these to match your setup
// ============================================================
const CONFIG = {
  SHEET_NAME: 'CRM',
  FOLLOW_UP_DAYS: 3,          // follow up after 3 days of no response
  URGENT_DAYS: 7,             // flag as urgent after 7 days
  DIGEST_EMAIL: 'atohorokeno@gmail.com', // sends daily digest to you
  EMAIL_SUBJECT_PREFIX: '[Sadecia Tutors] ',
};

// ============================================================
// COLUMN MAPPING — matches your Google Sheet headers
// ============================================================
const COL = {
  NAME: 1,
  EMAIL: 2,
  PHONE: 3,
  STATUS: 4,        // New / Contacted / Follow-Up / Enrolled / Closed
  LAST_CONTACT: 5,
  NEXT_FOLLOW_UP: 6,
  NOTES: 7,
  EMAIL_SENT: 8,    // tracks how many emails sent
  PRIORITY: 9,      // High / Medium / Low
};

// ============================================================
// EMAIL TEMPLATES
// ============================================================
const TEMPLATES = {
  first_follow_up: (name) => ({
    subject: CONFIG.EMAIL_SUBJECT_PREFIX + `Following up — ${name}`,
    body: `Hi ${name},

I hope you're doing well! I wanted to follow up on your interest in Sadecia Tutors.

We have some great tutoring options available and I'd love to help find the right fit for you. Do you have 10 minutes this week for a quick chat?

Feel free to reply to this email or book a time directly using the link below.

Warm regards,
Joy
Sadecia Tutors Team`
  }),

  second_follow_up: (name) => ({
    subject: CONFIG.EMAIL_SUBJECT_PREFIX + `Still here for you, ${name}`,
    body: `Hi ${name},

Just checking in one more time! I know things get busy.

If you're still interested in tutoring support, we're here and happy to help. If your plans have changed, no worries at all — just let me know and I'll update your record.

Best,
Joy
Sadecia Tutors Team`
  }),

  urgent: (name) => ({
    subject: CONFIG.EMAIL_SUBJECT_PREFIX + `Last check-in — ${name}`,
    body: `Hi ${name},

I wanted to reach out one final time regarding tutoring support with Sadecia Tutors.

If this is still something you're interested in, I'd love to connect. Otherwise, I'll mark this as closed on our end and won't send any further emails.

Either way, we wish you all the best!

Warm regards,
Joy
Sadecia Tutors Team`
  }),
};

// ============================================================
// MAIN FUNCTION — run this on a daily trigger
// ============================================================
function runDailyFollowUp() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const results = {
    sent: [],
    overdue: [],
    skipped: 0,
  };

  // Start from row 2 (skip header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowNum = i + 1;

    const name = row[COL.NAME - 1];
    const email = row[COL.EMAIL - 1];
    const status = row[COL.STATUS - 1];
    const lastContact = row[COL.LAST_CONTACT - 1];
    const emailsSent = row[COL.EMAIL_SENT - 1] || 0;

    // Skip empty rows or closed/enrolled contacts
    if (!name || !email || status === 'Enrolled' || status === 'Closed') {
      results.skipped++;
      continue;
    }

    // Calculate days since last contact
    let lastContactDate = null;
    if (lastContact) {
      const parsed = new Date(lastContact);
      lastContactDate = isNaN(parsed.getTime()) ? null : parsed;
    }
    const daysSince = lastContactDate
      ? Math.floor((today - lastContactDate) / (1000 * 60 * 60 * 24))
      : 999;

    // Determine if follow-up is needed
    if (daysSince >= CONFIG.FOLLOW_UP_DAYS) {
      const template = getTemplate(emailsSent);
      if (!template) {
        // Max follow-ups reached — flag and skip
        highlightRow(sheet, rowNum, '#FFE0E0'); // red
        results.overdue.push({ name, email, daysSince });
        continue;
      }

      // Send the email
      const { subject, body } = template(name);
      sendEmail(email, subject, body);

      // Update the sheet
      const now = new Date();
      sheet.getRange(rowNum, COL.LAST_CONTACT).setValue(now);
      sheet.getRange(rowNum, COL.NEXT_FOLLOW_UP).setValue(addDays(now, CONFIG.FOLLOW_UP_DAYS));
      sheet.getRange(rowNum, COL.STATUS).setValue('Follow-Up');
      sheet.getRange(rowNum, COL.EMAIL_SENT).setValue(emailsSent + 1);

      // Highlight based on urgency
      if (daysSince >= CONFIG.URGENT_DAYS) {
        highlightRow(sheet, rowNum, '#FFF3CD'); // amber — urgent
      } else {
        highlightRow(sheet, rowNum, '#D4EDDA'); // green — followed up
      }

      results.sent.push({ name, email, template: getTemplateName(emailsSent) });

      // Add calendar reminder for high priority
      if (row[COL.PRIORITY - 1] === 'High') {
        createCalendarReminder(name, email, addDays(now, CONFIG.FOLLOW_UP_DAYS));
      }

      // Throttle to avoid Gmail rate limits
      Utilities.sleep(1000);
    }
  }

  // Send daily digest to yourself
  sendDailyDigest(results);

  Logger.log(`Follow-up run complete. Sent: ${results.sent.length}, Overdue: ${results.overdue.length}, Skipped: ${results.skipped}`);
}

// ============================================================
// HELPER FUNCTIONS
// ============================================================

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = createCRMSheet(ss);
  }
  return sheet;
}

function createCRMSheet(ss) {
  const sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  const headers = [
    'Name', 'Email', 'Phone', 'Status',
    'Last Contact', 'Next Follow-Up', 'Notes',
    'Emails Sent', 'Priority'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1B2A4A')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Force Last Contact (col E) and Next Follow-Up (col F) to Date format
  sheet.getRange(2, COL.LAST_CONTACT, 1000)
    .setNumberFormat('MM/dd/yyyy');
  sheet.getRange(2, COL.NEXT_FOLLOW_UP, 1000)
    .setNumberFormat('MM/dd/yyyy');

  // STRICT date validation — rejects ALL non-date input including text
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText('Must be a valid date. Click the cell and use the date picker, or type MM/DD/YYYY')
    .build();
  sheet.getRange(2, COL.LAST_CONTACT, 1000).setDataValidation(dateRule);
  sheet.getRange(2, COL.NEXT_FOLLOW_UP, 1000).setDataValidation(dateRule);

  // Add data validation for Status column
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['New', 'Contacted', 'Follow-Up', 'Enrolled', 'Closed'])
    .build();
  sheet.getRange(2, COL.STATUS, 1000).setDataValidation(statusRule);

  // Add data validation for Priority column
  const priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['High', 'Medium', 'Low'])
    .build();
  sheet.getRange(2, COL.PRIORITY, 1000).setDataValidation(priorityRule);

  Logger.log('CRM sheet created successfully.');
  return sheet;
}

function getTemplate(emailsSent) {
  if (emailsSent === 0) return TEMPLATES.first_follow_up;
  if (emailsSent === 1) return TEMPLATES.second_follow_up;
  if (emailsSent === 2) return TEMPLATES.urgent;
  return null; // max follow-ups reached
}

function getTemplateName(emailsSent) {
  const names = ['First Follow-Up', 'Second Follow-Up', 'Final Follow-Up'];
  return names[emailsSent] || 'Unknown';
}

function sendEmail(to, subject, body) {
  try {
    GmailApp.sendEmail(to, subject, body);
    Logger.log(`Email sent to ${to}: ${subject}`);
  } catch (e) {
    Logger.log(`Failed to send email to ${to}: ${e.message}`);
  }
}

function highlightRow(sheet, rowNum, colour) {
  sheet.getRange(rowNum, 1, 1, Object.keys(COL).length).setBackground(colour);
}

function addDays(date, days) {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

function createCalendarReminder(name, email, date) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    calendar.createEvent(
      `Follow up with ${name}`,
      date,
      addDays(date, 0),
      {
        description: `Email: ${email}\nHigh priority lead — follow up today.`,
        reminders: [{ method: 'email', minutesBefore: 60 }]
      }
    );
  } catch (e) {
    Logger.log(`Calendar reminder failed for ${name}: ${e.message}`);
  }
}

function sendDailyDigest(results) {
  if (results.sent.length === 0 && results.overdue.length === 0) return;

  let body = `DAILY CRM DIGEST — ${new Date().toDateString()}\n\n`;

  if (results.sent.length > 0) {
    body += `✅ EMAILS SENT (${results.sent.length})\n`;
    results.sent.forEach(r => {
      body += `  • ${r.name} (${r.email}) — ${r.template}\n`;
    });
    body += '\n';
  }

  if (results.overdue.length > 0) {
    body += `⚠️ MAX FOLLOW-UPS REACHED — NEEDS MANUAL REVIEW (${results.overdue.length})\n`;
    results.overdue.forEach(r => {
      body += `  • ${r.name} (${r.email}) — ${r.daysSince} days since last contact\n`;
    });
    body += '\n';
  }

  body += `Skipped (enrolled/closed/empty): ${results.skipped}\n`;
  body += `\nLog into your CRM sheet to review: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;

  GmailApp.sendEmail(
    CONFIG.DIGEST_EMAIL,
    `[CRM Digest] ${results.sent.length} emails sent — ${new Date().toDateString()}`,
    body
  );
}

// ============================================================
// SETUP FUNCTION — run once to configure triggers
// ============================================================
function setupTriggers() {
  // Delete existing triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Run daily at 8am
  ScriptApp.newTrigger('runDailyFollowUp')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log('Daily trigger set for 8am. CRM automation is live.');
  SpreadsheetApp.getUi().alert('✅ Setup complete! CRM automation will run daily at 8am.');
}

// ============================================================
// MANUAL TRIGGER — add to Google Sheets menu
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🤖 CRM Automation')
    .addItem('Run Follow-Up Now', 'runDailyFollowUp')
    .addItem('Setup Daily Trigger (8am)', 'setupTriggers')
    .addItem('Create CRM Sheet', 'createCRMSheetManual')
    .addToUi();
}

function createCRMSheetManual() {
  createCRMSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('✅ CRM sheet created! Add your contacts and run the automation.');
}