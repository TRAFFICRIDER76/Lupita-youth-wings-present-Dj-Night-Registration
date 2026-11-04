 /**
 * Lupita YOUTH WINGS presents DJ NIGHT — Admin Approval, Ticket ID & QR Pass Automation
 * Add this file to the SAME Apps Script project as Code.gs (same Google Sheet).
 * Do NOT redeclare SHEET_NAME / COL_* constants here — they already live in Code.gs.
 *
 * WHAT THIS DOES
 * ----------------
 * 1. Adds a custom menu "🎧 DJ Night Admin" to your Google Sheet.
 * 2. "Setup Sheet Template" — builds Responses + Config tabs with headers,
 *    dropdowns (Status, Parish) and color-coded conditional formatting.
 *    SAFE TO RE-RUN: if the sheet already has registration data, it will
 *    only sync the header row / add missing columns — it will NOT clear
 *    existing rows. Config values you've already customized are preserved
 *    too (only missing keys get added).
 * 3. "Install Approval Trigger" — wires up an onEdit watcher (one-time setup).
 * 4. Whenever Admin changes a row's Status to APPROVED (or uses "Approve
 *    Entire Group"):
 *      - Auto-generates a unique Ticket ID (e.g. DJ-2026-101), protected by
 *        LockService so concurrent approvals never produce duplicate IDs.
 *      - Generates a QR Pass image (free QR Server API, no key needed)
 *      - Creates a clickable "📲 Send Pass to Participant" WhatsApp link
 *      - Creates a clickable "📲 Alert Youth Leader" WhatsApp link
 *      - Emails the participant a backup pass with the QR code inline
 *
 * ONE-TIME SETUP
 * ----------------
 * a) Paste this file into Apps Script (Extensions > Apps Script > + > Script).
 * b) Reload the Google Sheet (refresh browser tab) — menu "🎧 DJ Night Admin" appears.
 * c) Click menu > "1️⃣ Setup Sheet Template (run once)".
 * d) Click menu > "2️⃣ Install Approval Trigger (run once)" and approve permissions.
 * e) Open the "Config" tab and edit Event name/year/parish list etc. as needed.
 * f) (Optional, for OCR screening) Project Settings > Script Properties >
 *    add a property named OCR_API_KEY with your free OCR.space key.
 *
 * DAILY USE
 * ----------------
 * - Admin checks bank app for UTR match, then sets Status = APPROVED in the Sheet
 *   (or uses "✅ Approve Entire Group (Selected Row)" to approve everyone who
 *   paid together in one shot).
 * - Everything else (Ticket ID, QR, WhatsApp links, Email) happens automatically.
 * - Admin clicks the "📲 Send Pass to Participant" / "📲 Alert Youth Leader" cells
 *   to actually open WhatsApp and hit send (per the SRS's 1-click design).
 *
 * SECURITY NOTE: OCR Status / OCR Extracted UTR / UTR Match Result / Duplicate
 * UTR Result / OCR Extracted Amount columns (filled in by Code.gs at
 * submission time) are SCREENING AIDS ONLY. They never set Status to
 * APPROVED automatically — that decision always stays with the admin.
 */

/* ---------------- MENU ---------------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎧 DJ Night Admin')
    .addItem('1️⃣ Setup Sheet Template (run once)', 'setupSheetTemplate')
    .addItem('2️⃣ Install Approval Trigger (run once)', 'installApprovalTrigger')
    .addSeparator()
    .addItem('✅ Approve Entire Group (Selected Row)', 'approveEntireGroupForSelectedRow')
    .addItem('🔁 Resend Pass for Selected Row', 'resendPassForSelectedRow')
    .addItem('👁 Preview QR Pass for Selected Row', 'previewQrPassForSelectedRow')
    .addItem('↩️ Reset Check-in for Selected Row', 'resetCheckInForSelectedRow')
    .addSeparator()
    .addItem('📊 Build / Refresh Dashboard Tab', 'buildDashboardSheet')
    .addToUi();
}

/* ---------------- 1. READY-MADE SHEET TEMPLATE (NON-DESTRUCTIVE ON RE-RUN) ---------------- */

function setupSheetTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const headers = [
    'Timestamp', 'Group ID', 'Participant No', 'Full Name', 'Age', 'Phone', 'Email',
    'Parish', 'Leader Name', 'Leader Phone', 'Guardian Name', 'Guardian Phone', 'Talents',
    'Quantity', 'Total Amount', 'UTR',
    'Screenshot URL', 'Conduct Agreed', 'Status', 'Ticket ID',
    'WhatsApp Pass Link', 'Leader Alert Link', 'Entry Status', 'Check-in Time',
    'OCR Status', 'OCR Extracted UTR', 'UTR Match Result', 'Duplicate UTR Result',
    'OCR Extracted Amount', 'Admin Notes'
  ];

  // --- Responses sheet ---
  let sheet = ss.getSheetByName(SHEET_NAME);
  const isFreshSheet = !sheet || sheet.getLastRow() < 1;
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  if (isFreshSheet) {
    sheet.clear();
    sheet.clearConditionalFormatRules();
  }
  // Header row is always (re)written — this only touches row 1, never rows 2+,
  // so existing registrations are preserved when the sheet already has data.
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#8a2be2').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, headers.length, 150);

  // Status dropdown (PENDING / APPROVED / REJECTED) — re-applying validation is
  // safe/non-destructive, it does not touch existing cell values.
  const statusRange = sheet.getRange(2, COL_STATUS, 999, 1);
  statusRange.setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['PENDING', 'APPROVED', 'REJECTED'], true)
      .setAllowInvalid(false)
      .build()
  );

  // Entry Status range — filled in automatically by the Gate Check-in scan
  // (handleGateCheckIn in Code.gs), not manually editable like Status.
  const entryStatusRange = sheet.getRange(2, COL_ENTRYSTATUS, 999, 1);

  // Color-coded conditional formatting (replacing the rule list is safe —
  // it only changes formatting, never cell values).
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('PENDING').setBackground('#fff3cd').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('APPROVED').setBackground('#d4edda').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('REJECTED').setBackground('#f8d7da').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Checked-In').setBackground('#cfe2ff').setFontColor('#084298').setBold(true)
      .setRanges([entryStatusRange]).build()
  ]);

  // --- Config sheet: created fresh on first run; on re-run, only ADD missing
  // default keys — never overwrite values the admin has already customized. ---
  let config = ss.getSheetByName(CONFIG_SHEET_NAME);
  const isFreshConfig = !config || config.getLastRow() < 1;
  if (!config) config = ss.insertSheet(CONFIG_SHEET_NAME);

  const defaults = [
    ['EventName', 'Lupita YOUTH WINGS presents DJ NIGHT'],
    ['EventYear', 2026],
    ['TicketPrefix', 'DJ'],
    ['TicketStartNumber', 101],
    ['CountryCode', '91'],
    ['ParishList', "OUR LADY OF GUADALUPE CHURCH, TELCO; ST. JOSEPH'S CATHEDRAL, GOLMURI; ST. MARY'S CHURCH, BISTUPUR; ST. ROBERT'S CHURCH, PARSUDIH; ST. ANTHONY'S CHURCH, MANGO; HOLY FAMILY CHURCH, RAHRGORA (SUB-STATION); INFANT JESUS CHURCH, SONARI; INFANT JESUS CHURCH, JADUGORA; ST. ANTHONY'S CHURCH, MAUBHANDAR; ST. BARBARA CATHOLIC CHURCH, MOSABONI; OUR LADY OF NATIVITY CHURCH, SINI; JISU MARSAL CHURCH, CHANDIL"],
    ['AdminEmail', 'nikhilminz04@gmail.com']
  ];

  if (isFreshConfig) {
    config.clear();
    config.getRange(1, 1, 1, 2).setValues([['Key', 'Value']])
      .setFontWeight('bold').setBackground('#27e0ff');
    config.getRange(2, 1, defaults.length, 2).setValues(defaults);
    config.setColumnWidths(1, 2, 220);
  } else {
    // Merge: only append keys that don't already exist, leave everything else untouched.
    const existingLastRow = config.getLastRow();
    const existingPairs = existingLastRow >= 2
      ? config.getRange(2, 1, existingLastRow - 1, 2).getValues()
      : [];
    const existingKeys = existingPairs.map(function (r) { return r[0]; });
    const toAppend = defaults.filter(function (d) { return existingKeys.indexOf(d[0]) === -1; });
    if (toAppend.length > 0) {
      config.getRange(existingLastRow + 1, 1, toAppend.length, 2).setValues(toAppend);
    }
  }

  // Parish column on Responses: allow dropdown values OR any custom typed value
  // (form sends either a preset parish name, or whatever the user typed under "Other").
  const cfg = getConfig();
  const parishList = String(cfg.ParishList || '')
    .split(';').map(function (s) { return s.trim(); }).filter(Boolean).concat('Other');
  sheet.getRange(2, COL_PARISH, 999, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(parishList, true).setAllowInvalid(true).build()
  );

  if (isFreshSheet) {
    ui.alert('✅ Fresh setup done! "Responses" + "Config" tabs ready with dropdowns & color formatting.\n\nAb "Config" tab mein apna Event Name, Year, Parish list waghera edit kar lo.');
  } else {
    ui.alert('✅ Template synced — headers/columns updated, aapka existing registration data safe hai (koi row delete nahi hui).');
  }
}

/* ---------------- 2. INSTALL APPROVAL TRIGGER ---------------- */

function installApprovalTrigger() {
  const ui = SpreadsheetApp.getUi();
  const already = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'onEditApproval');
  if (already) {
    ui.alert('Trigger already installed ✅ (dobara install karne ki zaroorat nahi)');
    return;
  }
  ScriptApp.newTrigger('onEditApproval')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  ui.alert('✅ Approval trigger installed! Ab jab bhi Status ko "APPROVED" karoge, Ticket ID + WhatsApp links + Email auto-generate ho jayenge.');
}

/* ---------------- 3. APPROVAL WATCHER (fires on every manual edit) ---------------- */

function onEditApproval(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_NAME) return;
    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row === 1 || col !== COL_STATUS) return; // only react to Status column edits

    processApprovalForRow(sheet, row);
  } catch (err) {
    Logger.log('onEditApproval error: ' + err);
  }
}

/**
 * Core "approve this row" logic — shared by the onEdit trigger AND the manual
 * "Approve Entire Group" menu action, so there is exactly ONE code path that
 * can generate a Ticket ID / send a pass. The entire check-generate-send
 * sequence runs inside a script lock so that:
 *  - Two near-simultaneous triggers for the same row can never both generate
 *    a ticket (the "ticket already exists" check is re-verified safely).
 *  - Two different rows being approved at nearly the same time can never be
 *    assigned the same Ticket ID (NEXT_TICKET_NUMBER read+increment is
 *    effectively serialized by this lock).
 * IMPORTANT: generateTicketId() must only ever be called from inside this
 * lock — it has no locking of its own.
 */
function processApprovalForRow(sheet, row) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (lockErr) {
    Logger.log('processApprovalForRow: could not acquire lock for row ' + row + ': ' + lockErr);
    return; // fail safe — skip rather than risk a duplicate ticket
  }

  try {
    const status = sheet.getRange(row, COL_STATUS).getDisplayValue().trim().toUpperCase();
    const existingTicket = sheet.getRange(row, COL_TICKETID).getValue();

    if (status === 'APPROVED' && !existingTicket) {
      const cfg = getConfig();
      const ticketId = generateTicketId(cfg);
      sheet.getRange(row, COL_TICKETID).setValue(ticketId);
      buildAndSendPass(sheet, row, ticketId, cfg);
    }
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- 3b. GROUP-LEVEL APPROVAL (manual menu action) ----------------
 * One payment can cover multiple participants (same Group ID). This lets the
 * admin review the payment ONCE at the group level, then approve every
 * participant row in that group in a single click. Each participant still
 * keeps their OWN independent Status / Ticket ID / Entry Status — approving
 * the group just sets Status = APPROVED on every row sharing that Group ID
 * and runs the normal ticket/pass pipeline for each one. Rows already
 * APPROVED are left untouched (idempotent, safe to click more than once).
 *
 * NOTE: because this sets Status via script (not a manual Sheet edit), the
 * onEdit trigger does NOT fire for these cells — that's why this function
 * calls processApprovalForRow() itself for every row, instead of relying on
 * onEditApproval(). -------------------------------------------------------- */

function approveEntireGroupForSelectedRow() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveCell().getRow();

  if (sheet.getName() !== SHEET_NAME || row === 1) {
    ui.alert('Responses sheet mein pehle ek participant ki row select karo.');
    return;
  }

  const groupId = sheet.getRange(row, COL_GROUPID).getValue();
  if (!groupId) {
    ui.alert('Is row ka Group ID nahi mila — yeh purana/manually-added row ho sakta hai.');
    return;
  }

  const result = approveGroupById(groupId);
  if (!result.found) {
    ui.alert('Group ID "' + groupId + '" ke liye koi row nahi mili.');
    return;
  }

  ui.alert('✅ Group "' + groupId + '": ' + result.newlyApproved + ' participant(s) approved abhi, ' +
    result.alreadyApproved + ' pehle se approved the. Total ' + result.affectedRows.length +
    ' participant(s) is group mein. Tickets/passes ban gaye jinke bhi pending the.');
}

/**
 * Shared core: approves every row sharing the given Group ID (Status ->
 * APPROVED) and runs the normal ticket/pass pipeline for each one. Used by
 * BOTH the "Approve Entire Group" Sheet menu action AND the password-gated
 * Admin Panel on dashboard.html (via Code.gs's doPost admin action) — this
 * is the single code path that can approve a group, so behavior is always
 * identical no matter which UI triggered it. Idempotent: rows already
 * APPROVED are left untouched (safe to call more than once).
 */
function approveGroupById(groupId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { found: false, newlyApproved: 0, alreadyApproved: 0, affectedRows: [] };

  const groupIdValues = sheet.getRange(2, COL_GROUPID, lastRow - 1, 1).getValues();

  let newlyApproved = 0;
  let alreadyApproved = 0;
  let affectedRows = [];

  for (let i = 0; i < groupIdValues.length; i++) {
    if (String(groupIdValues[i][0]) === String(groupId)) {
      const sheetRow = i + 2;
      affectedRows.push(sheetRow);
      const currentStatus = sheet.getRange(sheetRow, COL_STATUS).getDisplayValue().trim().toUpperCase();
      if (currentStatus === 'APPROVED') {
        alreadyApproved++;
      } else {
        sheet.getRange(sheetRow, COL_STATUS).setValue('APPROVED');
        newlyApproved++;
      }
    }
  }

  // Run the ticket/pass pipeline for every row in the group (covers rows we
  // just approved, and any row that was already APPROVED earlier but for
  // some reason never got a ticket/pass yet).
  affectedRows.forEach(function (sheetRow) {
    processApprovalForRow(sheet, sheetRow);
  });

  return { found: affectedRows.length > 0, newlyApproved: newlyApproved, alreadyApproved: alreadyApproved, affectedRows: affectedRows };
}

/**
 * Shared core: rejects every row sharing the given Group ID (Status ->
 * REJECTED), but SKIPS any row that is already APPROVED (so the Admin Panel
 * can never accidentally retract a ticket that was already issued/sent —
 * admin must use the Sheet directly for that edge case). Used by the
 * password-gated Admin Panel on dashboard.html.
 */
function rejectGroupById(groupId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { found: false, newlyRejected: 0, skippedApproved: 0, affectedRows: [] };

  const groupIdValues = sheet.getRange(2, COL_GROUPID, lastRow - 1, 1).getValues();

  let newlyRejected = 0;
  let skippedApproved = 0;
  let affectedRows = [];

  for (let i = 0; i < groupIdValues.length; i++) {
    if (String(groupIdValues[i][0]) === String(groupId)) {
      const sheetRow = i + 2;
      affectedRows.push(sheetRow);
      const currentStatus = sheet.getRange(sheetRow, COL_STATUS).getDisplayValue().trim().toUpperCase();
      if (currentStatus === 'APPROVED') {
        skippedApproved++;
      } else {
        sheet.getRange(sheetRow, COL_STATUS).setValue('REJECTED');
        newlyRejected++;
      }
    }
  }

  return { found: affectedRows.length > 0, newlyRejected: newlyRejected, skippedApproved: skippedApproved, affectedRows: affectedRows };
}

/**
 * Reads back the WhatsApp 1-click links + Ticket ID already generated (by
 * buildAndSendPass, via the HYPERLINK formulas it writes into the sheet) for
 * a list of row numbers. Used right after an Admin Panel approval so the
 * panel can immediately show "📲 Send to Participant" / "📲 Alert Leader"
 * buttons without the admin needing to open the Sheet at all.
 */
function getApprovalLinksForGroup(affectedRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  function extractUrl(formula) {
    const match = String(formula || '').match(/HYPERLINK\("([^"]+)"/);
    return match ? match[1] : '';
  }

  return affectedRows.map(function (sheetRow) {
    const rowData = sheet.getRange(sheetRow, 1, 1, LAST_COL).getValues()[0];
    return {
      fullName: rowData[COL_FULLNAME - 1],
      ticketId: rowData[COL_TICKETID - 1],
      status: rowData[COL_STATUS - 1],
      waParticipantUrl: extractUrl(sheet.getRange(sheetRow, COL_WA_PASS_LINK).getFormula()),
      waLeaderUrl: extractUrl(sheet.getRange(sheetRow, COL_WA_LEADER_LINK).getFormula())
    };
  });
}

/**
 * Admin Panel "search any registration" helper — works for a solo/individual
 * submission OR any single member of a bigger group, APPROVED or PENDING or
 * REJECTED, no time limit (unlike the Approve-button flash message, which
 * only shows WA links for a few seconds right after approving). Matches by
 * Ticket ID, Full Name, Phone or Group ID (case-insensitive substring).
 * Returns the SAME WhatsApp 1-click links + Ticket ID that already live in
 * the Sheet's "WhatsApp Pass Link"/"Leader Alert Link" columns (via
 * buildAndSendPass's HYPERLINK formulas) — this is just a read-only lookup,
 * it never changes Status or re-sends anything by itself.
 */
function searchRegistrationsForAdminPanel(query) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || !query) return [];

  const q = String(query).trim().toLowerCase();
  if (!q) return [];

  function extractUrl(formula) {
    const match = String(formula || '').match(/HYPERLINK\("([^"]+)"/);
    return match ? match[1] : '';
  }

  const data = sheet.getRange(2, 1, lastRow - 1, LAST_COL).getValues();
  const results = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const sheetRow = i + 2;
    const ticketId = String(row[COL_TICKETID - 1] || '');
    const fullName = String(row[COL_FULLNAME - 1] || '');
    const phone    = String(row[COL_PHONE - 1] || '');
    const groupId  = String(row[COL_GROUPID - 1] || '');

    const haystack = (ticketId + ' ' + fullName + ' ' + phone + ' ' + groupId).toLowerCase();
    if (haystack.indexOf(q) === -1) continue;

    results.push({
      rowNumber: sheetRow,
      groupId: groupId,
      fullName: fullName,
      phone: phone,
      email: row[COL_EMAIL - 1],
      parish: row[COL_PARISH - 1],
      leaderName: row[COL_LEADERNAME - 1],
      leaderPhone: row[COL_LEADERPHONE - 1],
      status: String(row[COL_STATUS - 1] || ''),
      ticketId: ticketId,
      waParticipantUrl: extractUrl(sheet.getRange(sheetRow, COL_WA_PASS_LINK).getFormula()),
      waLeaderUrl: extractUrl(sheet.getRange(sheetRow, COL_WA_LEADER_LINK).getFormula())
    });

    if (results.length >= 25) break; // keep the Admin Panel light even on a big sheet
  }

  return results;
}

/**
 * Admin Panel "Resend Email" button (mirrors the Sheet menu's "🔁 Resend
 * Pass for Selected Row", just addressed by row number instead of the
 * active cell). Only works on a row that is already APPROVED with a Ticket
 * ID — it re-runs buildAndSendPass, which re-writes the same WhatsApp
 * HYPERLINK formulas (idempotent) and re-sends the backup email. Never
 * changes Status and never generates a new/second Ticket ID.
 */
function resendPassForRowNumber(rowNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const row = parseInt(rowNumber, 10);
  if (!row || row < 2) return { success: false, message: 'Invalid row number.' };

  const status = sheet.getRange(row, COL_STATUS).getDisplayValue().trim().toUpperCase();
  const ticketId = sheet.getRange(row, COL_TICKETID).getValue();
  if (status !== 'APPROVED' || !ticketId) {
    return { success: false, message: 'Yeh registration abhi APPROVED nahi hai ya Ticket ID nahi bana — pehle Pending list se Approve karo.' };
  }

  const cfg = getConfig();
  buildAndSendPass(sheet, row, ticketId, cfg);
  return { success: true, message: 'Email dobara bhej diya aur WhatsApp links refresh kar diye.' };
}

/**
 * Scans the Responses sheet for all PENDING rows, grouped by Group ID, and
 * returns full details (name/phone/UTR/screenshot/OCR flags etc.) for the
 * password-gated Admin Panel on dashboard.html. Key-gating happens in
 * Code.gs's handleAdminPanelData() BEFORE this is ever called — this
 * function itself has no access control of its own.
 */
function getPendingGroupsForAdminPanel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, LAST_COL).getValues();
  const groups = {};
  const order = [];

  data.forEach(function (row, idx) {
    const status = String(row[COL_STATUS - 1] || '').trim().toUpperCase();
    if (status !== 'PENDING') return;

    const groupId = row[COL_GROUPID - 1];
    if (!groups[groupId]) {
      groups[groupId] = {
        groupId: groupId,
        timestamp: row[COL_TIMESTAMP - 1],
        quantity: row[COL_QUANTITY - 1],
        totalAmount: row[COL_TOTALAMOUNT - 1],
        utr: row[COL_UTR - 1],
        screenshotUrl: row[COL_SCREENSHOT - 1],
        talents: row[COL_TALENTS - 1],
        ocrStatus: row[COL_OCR_STATUS - 1],
        ocrExtractedUtr: row[COL_OCR_UTR - 1],
        utrMatch: row[COL_UTR_MATCH - 1],
        duplicateUtrResult: row[COL_DUP_UTR - 1],
        ocrExtractedAmount: row[COL_OCR_AMOUNT - 1],
        adminNotes: row[COL_ADMIN_NOTES - 1],
        participants: []
      };
      order.push(groupId);
    }

    groups[groupId].participants.push({
      rowNumber: idx + 2,
      fullName: row[COL_FULLNAME - 1],
      age: row[COL_AGE - 1],
      phone: row[COL_PHONE - 1],
      email: row[COL_EMAIL - 1],
      parish: row[COL_PARISH - 1],
      leaderName: row[COL_LEADERNAME - 1],
      leaderPhone: row[COL_LEADERPHONE - 1],
      guardianName: row[COL_GUARDIANNAME - 1],
      guardianPhone: row[COL_GUARDIANPHONE - 1]
    });
  });

  return order.map(function (gid) { return groups[gid]; });
}

/* ---------------- 4. CONFIG READER ---------------- */

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  const cfg = {
    EventName: 'Lupita YOUTH WINGS presents DJ NIGHT', EventYear: new Date().getFullYear(),
    TicketPrefix: 'DJ', TicketStartNumber: 101, CountryCode: '91'
  };
  if (!configSheet) return cfg;
  const lastRow = configSheet.getLastRow();
  if (lastRow < 2) return cfg;
  const values = configSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  values.forEach(([key, value]) => { if (key) cfg[key] = value; });
  return cfg;
}

/* ---------------- 5. TICKET ID GENERATOR ----------------
 * SECURITY NOTE: this function reads+increments a shared counter in
 * PropertiesService with NO locking of its own. It must only ever be called
 * from inside processApprovalForRow(), which already holds a script lock for
 * its entire duration — that is what actually prevents two approvals from
 * ever receiving the same Ticket ID. Do not call this from anywhere else
 * without wrapping it in the same kind of lock. ------------------------- */

function generateTicketId(cfg) {
  const props = PropertiesService.getScriptProperties();
  let next = props.getProperty('NEXT_TICKET_NUMBER');
  next = next ? parseInt(next, 10) : (parseInt(cfg.TicketStartNumber, 10) || 101);
  const ticketId = `${cfg.TicketPrefix}-${cfg.EventYear}-${next}`;
  props.setProperty('NEXT_TICKET_NUMBER', String(next + 1));
  return ticketId;
}

/* ---------------- 6. QR PASS GENERATOR (reusable, free API) ---------------- */

// The QR code now encodes a CLICKABLE GATE CHECK-IN LINK (not just plain text).
// Scanning it with any phone camera/QR app opens this link in a browser, which
// hits doGet() in Code.gs, marks the ticket "Checked-In" the first time (✅ Entry
// Granted), and shows a warning (⚠️ Already Checked In, with original time) if
// the same QR is scanned again — exactly the "blue tick" re-entry check the
// admin wanted. The resulting page shows Full Name, Parish, Ticket ID, Youth
// Leader Name+Phone and Guardian Name+Phone (all safely HTML-escaped on the
// Code.gs side) so gate staff can verify a participant on the spot.
function generateQrPassUrl(ticketId, fullName, parish, guardianPhone, leaderName, leaderPhone, size) {
  size = size || 300;
  const checkInLink = WEB_APP_URL + '?ticket=' + encodeURIComponent(ticketId);
  return `https://api.qrserver.com/v1/create-qr-code/?size=${size}x${size}&data=${encodeURIComponent(checkInLink)}`;
}

/* ---------------- 7. BUILD WHATSAPP LINKS + SEND BACKUP EMAIL ---------------- */

function buildAndSendPass(sheet, row, ticketId, cfg) {
  const rowData = sheet.getRange(row, 1, 1, LAST_COL).getValues()[0];
  const fullName      = rowData[COL_FULLNAME - 1];
  const phone         = String(rowData[COL_PHONE - 1]).replace(/\D/g, '');
  const email         = rowData[COL_EMAIL - 1];
  const parish        = rowData[COL_PARISH - 1];
  const leaderName    = rowData[COL_LEADERNAME - 1];
  const leaderPhone   = String(rowData[COL_LEADERPHONE - 1]).replace(/\D/g, '');
  const guardianPhone = String(rowData[COL_GUARDIANPHONE - 1]).replace(/\D/g, '');
  const quantity       = rowData[COL_QUANTITY - 1] || 1;
  const totalAmount    = rowData[COL_TOTALAMOUNT - 1] || (quantity * 100);

  const qrUrl = generateQrPassUrl(ticketId, fullName, parish, guardianPhone, leaderName, leaderPhone);

  // --- Participant WhatsApp pass link (encodeURIComponent already protects
  // this — WhatsApp deep links are just text, not rendered as HTML). ---
  const passMessage =
`🎧 ${cfg.EventName} — Entry Pass
Name: ${fullName}
Ticket: ${ticketId}
Parish: ${parish}
Tickets Paid: ${quantity} (₹${totalAmount})

Your registration is APPROVED ✅
View/Download your QR Pass here:
${qrUrl}

Please show this QR + your name at the gate. Pass is personal & non-transferable.`;

  const waParticipantUrl = `https://wa.me/${cfg.CountryCode}${phone}?text=${encodeURIComponent(passMessage)}`;
  sheet.getRange(row, COL_WA_PASS_LINK).setFormula(
    `=HYPERLINK("${waParticipantUrl}","📲 Send Pass to Participant")`
  );

  // --- Youth Leader WhatsApp alert link ---
  const leaderMessage =
`📋 Parish Audit Alert — ${cfg.EventName}
Hi ${leaderName}, ${fullName} from your parish (${parish}) has been APPROVED.
Ticket: ${ticketId}
Tickets Paid: ${quantity} (₹${totalAmount})
Guardian contact: ${guardianPhone}`;

  const waLeaderUrl = `https://wa.me/${cfg.CountryCode}${leaderPhone}?text=${encodeURIComponent(leaderMessage)}`;
  sheet.getRange(row, COL_WA_LEADER_LINK).setFormula(
    `=HYPERLINK("${waLeaderUrl}","📲 Alert Youth Leader")`
  );

  // --- Backup email with inline QR pass ---
  sendPassEmail(email, fullName, ticketId, parish, qrUrl, cfg);
}

function sendPassEmail(email, fullName, ticketId, parish, qrUrl, cfg) {
  if (!email) return;
  try {
    const qrBlob = UrlFetchApp.fetch(qrUrl).getBlob().setName('qr_pass.png');
    // fullName/parish are participant-controlled text — escape before
    // embedding in the email's HTML body to avoid markup injection.
    const safeName   = escapeHtml(fullName);
    const safeParish = escapeHtml(parish);
    const safeTicket = escapeHtml(ticketId);
    const htmlBody = `
      <div style="font-family:Arial,sans-serif;max-width:480px;margin:auto;">
        <h2 style="color:#8a2be2;">🎧 ${escapeHtml(cfg.EventName)}</h2>
        <p>Hi <b>${safeName}</b>,</p>
        <p>Your registration is <b style="color:green;">APPROVED</b> ✅</p>
        <p><b>Ticket ID:</b> ${safeTicket}<br/><b>Parish:</b> ${safeParish}</p>
        <p>Your Entry QR Pass is below. Show this + your name at the gate.</p>
        <img src="cid:qrPassImage" style="width:220px;height:220px;" />
        <p style="color:#777;font-size:12px;">This pass is personal & non-transferable.</p>
      </div>`;
    MailApp.sendEmail({
      to: email,
      subject: `✅ Your ${cfg.EventName} Entry Pass — ${ticketId}`,
      htmlBody: htmlBody,
      inlineImages: { qrPassImage: qrBlob }
    });
  } catch (err) {
    Logger.log('Email send failed for ' + email + ': ' + err);
  }
}

/* ---------------- 8. MANUAL HELPERS (from the Admin menu) ---------------- */

function resendPassForSelectedRow() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveCell().getRow();

  if (sheet.getName() !== SHEET_NAME || row === 1) {
    ui.alert('Responses sheet mein pehle ek participant ki row select karo.');
    return;
  }
  const ticketId = sheet.getRange(row, COL_TICKETID).getValue();
  if (!ticketId) {
    ui.alert('Is row ka Ticket ID abhi generate nahi hua. Pehle Status ko APPROVED karo.');
    return;
  }
  const cfg = getConfig();
  buildAndSendPass(sheet, row, ticketId, cfg);
  ui.alert('✅ Pass links & email dobara bana diye row ' + row + ' ke liye.');
}

function previewQrPassForSelectedRow() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveCell().getRow();

  if (sheet.getName() !== SHEET_NAME || row === 1) {
    ui.alert('Responses sheet mein pehle ek participant ki row select karo.');
    return;
  }
  const rowData = sheet.getRange(row, 1, 1, LAST_COL).getValues()[0];
  const ticketId = rowData[COL_TICKETID - 1] || 'PREVIEW-0000';
  const fullName = rowData[COL_FULLNAME - 1];
  const parish = rowData[COL_PARISH - 1];
  const guardianPhone = rowData[COL_GUARDIANPHONE - 1];
  const leaderName = rowData[COL_LEADERNAME - 1];
  const leaderPhone = rowData[COL_LEADERPHONE - 1];

  const qrUrl = generateQrPassUrl(ticketId, fullName, parish, guardianPhone, leaderName, leaderPhone);
  const html = HtmlService.createHtmlOutput(
    `<div style="text-align:center;font-family:Arial;">
       <p><b>${escapeHtml(fullName)}</b> — ${escapeHtml(parish)}</p>
       <img src="${qrUrl}" width="260" height="260"/>
       <p style="font-size:12px;color:#666;">${escapeHtml(ticketId)}</p>
     </div>`
  ).setWidth(320).setHeight(380);
  ui.showModalDialog(html, 'QR Pass Preview');
}

function resetCheckInForSelectedRow() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveCell().getRow();

  if (sheet.getName() !== SHEET_NAME || row === 1) {
    ui.alert('Responses sheet mein pehle ek participant ki row select karo.');
    return;
  }

  const entryStatus = sheet.getRange(row, COL_ENTRYSTATUS).getValue();
  if (!entryStatus) {
    ui.alert('Is row ka abhi tak koi gate check-in record hi nahi hai — reset karne ki zaroorat nahi.');
    return;
  }

  const fullName = sheet.getRange(row, COL_FULLNAME).getValue();
  const confirmResp = ui.alert(
    'Reset Check-in?',
    'Pakka? "' + fullName + '" (row ' + row + ') ka check-in record clear ho jaayega, aur agli scan par phir se "Entry Granted" dikhega.',
    ui.ButtonSet.YES_NO
  );
  if (confirmResp !== ui.Button.YES) return;

  sheet.getRange(row, COL_ENTRYSTATUS).clearContent();
  sheet.getRange(row, COL_CHECKIN_TIME).clearContent();

  // Lightweight audit trail — appended to Admin Notes rather than a whole new column.
  const existingNotes = sheet.getRange(row, COL_ADMIN_NOTES).getValue();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy, hh:mm a');
  const resetNote = '[Check-in manually reset on ' + stamp + ']';
  sheet.getRange(row, COL_ADMIN_NOTES).setValue(existingNotes ? (existingNotes + ' | ' + resetNote) : resetNote);

  ui.alert('✅ Check-in reset ho gaya row ' + row + ' ke liye. Ab agli baar QR scan karne par "Entry Granted" dikhega.');
}

/* ---------------- 9. LIVE DASHBOARD TAB ----------------
 * Builds a "Dashboard" sheet made entirely of native Google Sheets formulas
 * (COUNTIF / SUMPRODUCT / QUERY) that read the Responses sheet. Because
 * these are live formulas (not a one-time snapshot), the numbers update
 * INSTANTLY whenever Responses changes — no manual refresh needed, and no
 * extra code/trigger required. This tab only ever reads Responses, it never
 * writes to it, so it is 100% safe to rebuild any time.
 *
 * This can also be used as a clean, pre-aggregated data source for a
 * prettier Google Looker Studio report later (Looker Studio can connect to
 * either this "Dashboard" tab or directly to "Responses").
 *
 * NOTE: formulas below assume registrations stay under ~1000 rows. If your
 * event grows bigger than that, just bump the "1000" in the ranges below. ---- */

function buildDashboardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  if (!ss.getSheetByName(SHEET_NAME)) {
    ui.alert('"Responses" sheet nahi mila. Pehle "Setup Sheet Template" chalao.');
    return;
  }

  let dash = ss.getSheetByName('Dashboard');
  if (dash) {
    dash.clear(); // Safe — this sheet only ever contains formulas, never source data.
  } else {
    dash = ss.insertSheet('Dashboard');
  }

  const R = "'" + SHEET_NAME + "'"; // sheet name reference for formulas, quoted for safety

  dash.getRange('A1:B1').merge().setValue('📊 ' + getConfig().EventName + ' — Live Dashboard')
    .setFontWeight('bold').setFontSize(16).setBackground('#8a2be2').setFontColor('#ffffff');
  dash.getRange('A2:B2').merge()
    .setValue('Yeh numbers LIVE hain — Responses sheet mein koi edit hote hi yahan khud update ho jaate hain. Manual refresh ki zaroorat nahi.')
    .setFontStyle('italic').setFontColor('#666666').setWrap(true);

  // --- Overview scorecards ---
  const overview = [
    ['Total Registrations (Groups)', `=COUNTIF(${R}!C2:C1000,"1 of *")`],
    ['Total Participants',           `=COUNTA(${R}!D2:D1000)`],
    ['Pending',                      `=COUNTIF(${R}!S2:S1000,"PENDING")`],
    ['Approved',                     `=COUNTIF(${R}!S2:S1000,"APPROVED")`],
    ['Rejected',                     `=COUNTIF(${R}!S2:S1000,"REJECTED")`],
    ['Total Revenue — Approved Groups (₹)', `=SUMPRODUCT((LEFT(${R}!C2:C1000,4)="1 of")*(${R}!S2:S1000="APPROVED")*(${R}!O2:O1000))`],
    ['Checked-In at Gate',           `=COUNTIF(${R}!W2:W1000,"Checked-In")`],
    ['Approved but Still Outside',   `=COUNTIF(${R}!S2:S1000,"APPROVED")-COUNTIF(${R}!W2:W1000,"Checked-In")`],
    ['⚠️ Duplicate UTR Flags',        `=COUNTIF(${R}!AB2:AB1000,"DUPLICATE*")`],
    ['⚠️ OCR UTR Mismatch Flags',     `=COUNTIF(${R}!AA2:AA1000,"Mismatch*")`]
  ];
  dash.getRange(4, 1, overview.length, 2).setValues(overview);
  dash.getRange(4, 1, overview.length, 1).setFontWeight('bold');
  dash.getRange(4, 2, overview.length, 1).setBackground('#f3eefc').setHorizontalAlignment('center');
  dash.setColumnWidth(1, 280);
  dash.setColumnWidth(2, 160);

  // --- Parish-wise breakdown ---
  const parishRow = overview.length + 6;
  dash.getRange(parishRow, 1).setValue('⛪ Parish-wise Participants').setFontWeight('bold').setFontSize(13);
  dash.getRange(parishRow + 1, 1).setFormula(
    `=QUERY(${R}!H2:H1000,"select H, count(H) where H is not null group by H order by count(H) desc label count(H) 'Participants'",0)`
  );

  // --- Needs-review list: anything flagged by duplicate-UTR or OCR mismatch screening ---
  const reviewRow = parishRow + 20; // leave room for the parish table above
  dash.getRange(reviewRow, 1).setValue('🔎 Needs Manual Review (duplicate UTR / OCR mismatch)').setFontWeight('bold').setFontSize(13);
  dash.getRange(reviewRow + 1, 1).setFormula(
    `=QUERY(${R}!A2:AD1000,"select D, H, P, AB, AA, AD where AB contains 'DUPLICATE' or AA contains 'Mismatch' label D 'Full Name', H 'Parish', P 'UTR', AB 'Duplicate Check', AA 'OCR UTR Match', AD 'Admin Notes'",0)`
  );

  ui.alert('✅ Dashboard tab ready! "Dashboard" sheet tab kholo — yeh Responses change hote hi khud update hota rahega.\n\nAgar aage Looker Studio se aur pretty charts banane hain, to Looker Studio mein "Google Sheets" data source se isi "Dashboard" ya "Responses" tab ko connect kar sakte ho.');
}