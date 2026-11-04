 /**
 * Lupita YOUTH WINGS presents DJ NIGHT — Custom Registration Backend
 * Replaces the Google Form front-end. Keeps the SAME Google Sheet
 * structure/workflow described in the SRS (PENDING -> APPROVED,
 * UTR verification, WhatsApp 1-click links, QR pass issuance).
 *
 * NOTE: This file lives in the SAME Apps Script project as AdminApproval.gs.
 * Column constants (COL_*, SHEET_NAME) used by both files are declared
 * once below — do not redeclare them in AdminApproval.gs.
 *
 * GROUP REGISTRATION SUPPORT:
 * One form submission can cover multiple tickets/participants. Each
 * participant gets their OWN row in the Sheet (own Status, Ticket ID,
 * QR pass, WhatsApp links) so admin can approve/track each person
 * individually, but all rows from the same submission share the same
 * Group ID, Quantity, Total Amount, UTR, Screenshot and OCR verification
 * results (one payment covers the whole group).
 *
 * SECURITY / RELIABILITY NOTES (Phase 2 audit fixes — read before editing):
 * - Ticket ID generation and gate check-in writes are both wrapped in
 *   LockService to prevent duplicate-ticket / double-grant race conditions
 *   under concurrent admin approvals or concurrent gate scans.
 * - All participant-controlled data (name, parish, etc.) is HTML-escaped
 *   before being rendered on the public gate check-in page or in emails.
 * - The gate check-in page shows Full Name, Parish, Ticket ID, Youth Leader
 *   Name+Phone and Guardian Name+Phone (per the admin's explicit requirement
 *   so gate staff can cross-check a participant/leader/guardian on the
 *   spot). All of these values are passed through escapeHtml() before being
 *   rendered, so this does not reopen the XSS issue found in the audit.
 * - OCR (OCR.space) is a SCREENING aid only. It never auto-approves or
 *   auto-rejects a payment — Status is always set manually by the admin.
 *
 * SETUP STEPS:
 * 1. Create / open your Google Sheet (same one your admin uses for verification).
 * 2. Run menu "🎧 DJ Night Admin > 1️⃣ Setup Sheet Template" (from AdminApproval.gs)
 *    to auto-create the "Responses" + "Config" tabs with the right headers.
 * 3. Create a Google Drive folder for payment screenshots, copy its folder ID
 *    (from the URL: drive.google.com/drive/folders/<THIS_IS_THE_ID>) into DRIVE_FOLDER_ID below.
 * 4. Paste this file into Extensions > Apps Script.
 * 5. Add your OCR.space API key via Project Settings > Script Properties
 *    (Key: OCR_API_KEY) — NEVER paste it directly into this source file.
 * 6. Deploy > New deployment > type: "Web app" > Execute as: Me > Who has access: Anyone.
 * 7. Copy the generated Web App URL into index.html's SCRIPT_URL variable AND into WEB_APP_URL below.
 * 8. Re-deploy (Manage deployments > Edit > New version) any time you change this code.
 */

const SHEET_NAME = 'Responses';
const DRIVE_FOLDER_ID = '1v2SBCWPSDwNwGzYq2k4CrX9UR9eKUa0G';
const TICKET_PRICE = 100; // ₹ per ticket — Total Amount must always be quantity x this

// The deployed Web App URL (Deploy > Manage deployments > Web app).
// Used to build the Gate Check-in link embedded in each QR pass.
// IMPORTANT: only update this if you ever create a brand-new deployment —
// if you always use "Manage deployments > Edit > New version", this URL
// never changes and you never need to touch this constant again.
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbx75pAL6AhfcWhMTBzNo694EKzCSM-PQaCPKHMJ6AoMaHgdyySsBpBkt1XZJazx-DOt/exec';

// Column numbers (1-indexed) — shared with AdminApproval.gs
const COL_TIMESTAMP      = 1;
const COL_GROUPID        = 2;
const COL_PARTICIPANTNO  = 3;
const COL_FULLNAME       = 4;
const COL_AGE            = 5;
const COL_PHONE          = 6;
const COL_EMAIL          = 7;
const COL_PARISH         = 8;
const COL_LEADERNAME     = 9;
const COL_LEADERPHONE    = 10;
const COL_GUARDIANNAME   = 11;
const COL_GUARDIANPHONE  = 12;
const COL_TALENTS        = 13;
const COL_QUANTITY       = 14;
const COL_TOTALAMOUNT    = 15; // Also doubles as "Expected Amount" for OCR amount comparison — no separate column needed.
const COL_UTR            = 16;
const COL_SCREENSHOT     = 17;
const COL_CONDUCT        = 18;
const COL_STATUS         = 19; // Also doubles as "Payment Approval Status" — admin-controlled, OCR never writes here.
const COL_TICKETID       = 20;
const COL_WA_PASS_LINK   = 21;
const COL_WA_LEADER_LINK = 22;
const COL_ENTRYSTATUS    = 23;
const COL_CHECKIN_TIME   = 24;

// --- OCR / payment-screening columns (added in Phase 2 audit — always APPENDED
// at the end so none of the existing column numbers above ever shift). ---
const COL_OCR_STATUS     = 25; // Processed / Unreadable / Error / Skipped (no API key / rate-limited)
const COL_OCR_UTR        = 26; // UTR text OCR could read out of the screenshot (if any)
const COL_UTR_MATCH      = 27; // Match / Mismatch / Not Found in Screenshot / N/A
const COL_DUP_UTR        = 28; // Unique / DUPLICATE - also used in another row
const COL_OCR_AMOUNT     = 29; // Amount OCR could read, plus a Match/Mismatch note vs Total Amount
const COL_ADMIN_NOTES    = 30; // Free-text, admin-editable; auto-prefilled with warning flags when relevant

const LAST_COL           = 30;

const CONFIG_SHEET_NAME  = 'Config';

// --- Screenshot upload limits (server-side enforced; index.html also checks client-side) ---
const ALLOWED_SCREENSHOT_MIME = ['image/jpeg', 'image/jpg', 'image/png', 'image/webp'];
const MAX_SCREENSHOT_BYTES    = 8 * 1024 * 1024; // 8 MB

// --- OCR.space config ---
const OCR_SPACE_ENDPOINT = 'https://api.ocr.space/parse/image';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // Admin Panel approve/reject/search/resend actions are routed separately
    // from the registration-submission flow below — they never touch the
    // participants/quantity/UTR/screenshot validation that follows.
    if (data && (data.action === 'approveGroup' || data.action === 'rejectGroup' ||
                 data.action === 'searchRegistrations' || data.action === 'resendPass')) {
      return handleAdminAction(data);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      throw new Error('Sheet tab "' + SHEET_NAME + '" not found. Run "Setup Sheet Template" from the Admin menu first.');
    }

    const participants = data.participants;
    if (!Array.isArray(participants) || participants.length < 1) {
      throw new Error('No participant details received.');
    }

    // --- Quantity & Total Amount: NEVER trust the amount sent by the browser,
    // always recompute server-side as quantity x TICKET_PRICE (₹100/ticket, minimum 1) ---
    const quantity = parseInt(data.quantity, 10);
    if (!quantity || quantity < 1) {
      throw new Error('Invalid ticket quantity. Minimum 1 ticket (₹' + TICKET_PRICE + ') required.');
    }
    if (participants.length !== quantity) {
      throw new Error('Number of participant detail sets (' + participants.length +
        ') does not match the selected ticket quantity (' + quantity + ').');
    }
    const totalAmount = quantity * TICKET_PRICE;

    if (!/^[0-9]{12}$/.test(data.utr)) throw new Error('UTR must be exactly 12 digits');
    if (!data.conduct) throw new Error('Code of Conduct must be accepted');

    // --- Validate every participant's details (never trust the browser alone) ---
    const requiredFields = ['fullName', 'age', 'phone', 'email', 'parish',
      'leaderName', 'leaderPhone', 'guardianName', 'guardianPhone'];

    participants.forEach(function (p, idx) {
      const label = 'Participant ' + (idx + 1);
      requiredFields.forEach(function (field) {
        if (!p[field] || String(p[field]).trim() === '') {
          throw new Error(label + ': missing ' + field);
        }
      });
      if (!/^[0-9]{10}$/.test(p.phone)) throw new Error(label + ': invalid phone number');
      if (!/^[0-9]{10}$/.test(p.leaderPhone)) throw new Error(label + ': invalid leader phone number');
      if (!/^[0-9]{10}$/.test(p.guardianPhone)) throw new Error(label + ': invalid guardian phone number');
    });

    // --- Server-side screenshot validation (MIME allowlist + size cap) ---
    if (!data.screenshotBase64) {
      throw new Error('Payment screenshot is required.');
    }
    const mimeType = (data.screenshotMime || 'image/jpeg').toLowerCase();
    if (ALLOWED_SCREENSHOT_MIME.indexOf(mimeType) === -1) {
      throw new Error('Invalid screenshot file type. Please upload a JPG, PNG or WEBP image.');
    }
    const base64Data = data.screenshotBase64.split(',')[1] || data.screenshotBase64;
    const approxBytes = Math.floor(base64Data.length * 0.75);
    if (approxBytes > MAX_SCREENSHOT_BYTES) {
      throw new Error('Screenshot file is too large (max 8 MB). Please upload a smaller or compressed image.');
    }

    // --- Duplicate UTR check (across ALL existing rows, any status) — screening only,
    // never blocks submission. Admin sees the flag and decides. ---
    const duplicateUtrResult = checkDuplicateUtr(sheet, data.utr);

    // --- Save payment screenshot to Drive (ONE screenshot covers the whole group).
    // Kept PRIVATE by default (not "anyone with the link") — only the script owner
    // and the configured AdminEmail(s) can open it, to avoid leaking payment
    // screenshots to anyone who gets hold of the Sheet/email link. ---
    const cfg = getConfig();
    let screenshotUrl = '';
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const fileName = (participants[0].fullName || 'group').replace(/\s+/g, '_') + '_' + data.utr;
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
    const file = folder.createFile(blob);
    try {
      if (cfg.AdminEmail) file.addViewer(cfg.AdminEmail);
    } catch (shareErr) {
      Logger.log('Could not add viewer for screenshot: ' + shareErr);
    }
    screenshotUrl = file.getUrl();

    // --- OCR screening (best-effort, never blocks registration) ---
    const ocrResult = runOcrVerification(base64Data, mimeType, data.utr, totalAmount);

    // --- Build an auto-suggested Admin Notes flag for anything suspicious ---
    const noteFlags = [];
    if (duplicateUtrResult.indexOf('DUPLICATE') === 0) noteFlags.push('⚠️ Duplicate UTR — verify before approving');
    if (ocrResult.utrMatch && ocrResult.utrMatch.indexOf('Mismatch') === 0) noteFlags.push('⚠️ OCR UTR mismatch — verify screenshot');
    if (ocrResult.amountNote && ocrResult.amountNote.indexOf('Mismatch') !== -1) noteFlags.push('⚠️ OCR amount mismatch — verify screenshot');
    if (ocrResult.ocrStatus.indexOf('Error') === 0 || ocrResult.ocrStatus.indexOf('Unreadable') === 0) noteFlags.push('ℹ️ OCR could not read screenshot clearly — verify manually');
    const adminNotes = noteFlags.join(' | ');

    const groupId = 'GRP-' + new Date().getTime();
    const timestamp = new Date();
    const talents = data.talents || 'None';

    // --- Append ONE row PER participant — same groupId/quantity/amount/UTR/screenshot/
    // OCR results, but each participant gets their own Status, Ticket ID & QR pass later. ---
    participants.forEach(function (p, idx) {
      const row = new Array(LAST_COL).fill('');
      row[COL_TIMESTAMP - 1]      = timestamp;
      row[COL_GROUPID - 1]        = groupId;
      row[COL_PARTICIPANTNO - 1]  = (idx + 1) + ' of ' + quantity;
      row[COL_FULLNAME - 1]       = p.fullName;
      row[COL_AGE - 1]            = p.age;
      row[COL_PHONE - 1]          = p.phone;
      row[COL_EMAIL - 1]          = p.email;
      row[COL_PARISH - 1]         = p.parish;
      row[COL_LEADERNAME - 1]     = p.leaderName;
      row[COL_LEADERPHONE - 1]    = p.leaderPhone;
      row[COL_GUARDIANNAME - 1]   = p.guardianName;
      row[COL_GUARDIANPHONE - 1]  = p.guardianPhone;
      row[COL_TALENTS - 1]        = talents;
      row[COL_QUANTITY - 1]       = quantity;
      row[COL_TOTALAMOUNT - 1]    = totalAmount;
      row[COL_UTR - 1]            = data.utr;
      row[COL_SCREENSHOT - 1]     = screenshotUrl;
      row[COL_CONDUCT - 1]        = 'Yes';
      row[COL_STATUS - 1]         = 'PENDING';   // Admin flips this to APPROVED (per participant) — NEVER auto-set by OCR
      // COL_TICKETID, COL_WA_PASS_LINK, COL_WA_LEADER_LINK stay blank until approval
      row[COL_ENTRYSTATUS - 1]    = '';
      row[COL_CHECKIN_TIME - 1]   = '';
      row[COL_OCR_STATUS - 1]     = ocrResult.ocrStatus;
      row[COL_OCR_UTR - 1]        = ocrResult.extractedUtr;
      row[COL_UTR_MATCH - 1]      = ocrResult.utrMatch;
      row[COL_DUP_UTR - 1]        = duplicateUtrResult;
      row[COL_OCR_AMOUNT - 1]     = ocrResult.extractedAmount ? (ocrResult.extractedAmount + (ocrResult.amountNote ? ' (' + ocrResult.amountNote + ')' : '')) : '';
      row[COL_ADMIN_NOTES - 1]    = adminNotes;

      sheet.appendRow(row);
    });

    return jsonResponse({
      status: 'success',
      message: 'Registration saved as PENDING for ' + quantity + ' ticket(s)/participant(s), ₹' + totalAmount + ' total.'
    });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message || String(err) });
  }
}

/* ---------------- DUPLICATE UTR SCREENING ---------------- */

function checkDuplicateUtr(sheet, utr) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'Unique';
  const utrValues = sheet.getRange(2, COL_UTR, lastRow - 1, 1).getValues();
  for (let i = 0; i < utrValues.length; i++) {
    if (String(utrValues[i][0]).trim() === String(utr).trim()) {
      return 'DUPLICATE — this UTR already exists in row ' + (i + 2) + ' (verify before approving)';
    }
  }
  return 'Unique';
}

/* ---------------- OCR.SPACE PAYMENT SCREENSHOT SCREENING ----------------
 * This is a SCREENING AID ONLY. It never approves or rejects a payment —
 * Status is always set manually by the admin after checking the bank
 * notification/transaction record. OCR failures/rate-limits are handled
 * gracefully and never block registration. ---------------------------- */

function runOcrVerification(base64Data, mimeType, enteredUtr, expectedAmount) {
  const result = {
    ocrStatus: 'Skipped',
    extractedUtr: '',
    utrMatch: 'N/A',
    extractedAmount: '',
    amountNote: ''
  };

  const apiKey = PropertiesService.getScriptProperties().getProperty('OCR_API_KEY');
  if (!apiKey) {
    result.ocrStatus = 'Skipped (no OCR_API_KEY configured in Script Properties)';
    return result;
  }

  try {
    const payload = {
      apikey: apiKey,
      base64Image: 'data:' + mimeType + ';base64,' + base64Data,
      OCREngine: '2',
      scale: 'true',
      isTable: 'false'
    };
    const options = {
      method: 'post',
      payload: payload,
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(OCR_SPACE_ENDPOINT, options);
    const code = response.getResponseCode();

    if (code !== 200) {
      result.ocrStatus = 'Error (HTTP ' + code + ')';
      return result;
    }

    const json = JSON.parse(response.getContentText());

    if (json.IsErroredOnProcessing) {
      result.ocrStatus = 'Error: ' + (Array.isArray(json.ErrorMessage) ? json.ErrorMessage.join(', ') : (json.ErrorMessage || 'Unknown OCR error'));
      return result;
    }

    const parsedText = (json.ParsedResults && json.ParsedResults[0] && json.ParsedResults[0].ParsedText) || '';
    if (!parsedText.trim()) {
      result.ocrStatus = 'Unreadable (no text detected in screenshot)';
      return result;
    }

    result.ocrStatus = 'Processed';

    // --- UTR comparison ---
    if (parsedText.indexOf(enteredUtr) !== -1) {
      result.extractedUtr = enteredUtr;
      result.utrMatch = 'Match';
    } else {
      const numMatch = parsedText.match(/\b\d{12}\b/);
      if (numMatch) {
        result.extractedUtr = numMatch[0];
        result.utrMatch = 'Mismatch (screenshot shows ' + numMatch[0] + ')';
      } else {
        result.utrMatch = 'Not Found in Screenshot';
      }
    }

    // --- Amount comparison (best-effort; only flagged when confidently read) ---
    const amtMatch = parsedText.match(/(?:₹|Rs\.?|INR)\s*([0-9][0-9,]*(?:\.[0-9]{1,2})?)/i);
    if (amtMatch) {
      const extractedNum = parseFloat(amtMatch[1].replace(/,/g, ''));
      result.extractedAmount = amtMatch[1].replace(/,/g, '');
      if (!isNaN(extractedNum)) {
        result.amountNote = (Math.round(extractedNum) === Math.round(expectedAmount))
          ? 'Match'
          : ('Mismatch — expected ₹' + expectedAmount);
      }
    }

    return result;

  } catch (err) {
    result.ocrStatus = 'Error: ' + (err.message || String(err));
    return result;
  }
}

function doGet(e) {
  // QR passes encode a link like: WEB_APP_URL?ticket=DJ-2026-101
  // Scanning the QR with ANY camera/QR app opens this link in a browser,
  // which runs this exact check and marks the ticket as checked-in.
  if (e && e.parameter && e.parameter.ticket) {
    return handleGateCheckIn(e.parameter.ticket);
  }
  // Live dashboard data feed: WEB_APP_URL?dashboard=1[&key=ADMIN_KEY]
  if (e && e.parameter && e.parameter.dashboard === '1') {
    return handleDashboardData(e.parameter.key || '');
  }
  // Admin Panel data feed (full PII — key REQUIRED): WEB_APP_URL?admin=1&key=ADMIN_KEY
  if (e && e.parameter && e.parameter.admin === '1') {
    return handleAdminPanelData(e.parameter.key || '');
  }
  return ContentService.createTextOutput('DJ Night Registration API is live ✅');
}

/* ---------------- ADMIN PANEL: FETCH PENDING REGISTRATIONS (full PII, key-gated) ----------------
 * Unlike handleDashboardData() (aggregate-only, safe without a key), this
 * returns full participant detail (name/phone/UTR/screenshot/OCR flags) for
 * every PENDING group — so it is ALWAYS key-gated, with no partial/public
 * response. Reuses the same DASHBOARD_ADMIN_KEY shared secret as the
 * dashboard's "Needs Review" list for a single consistent admin password. */
function handleAdminPanelData(key) {
  const adminKey = PropertiesService.getScriptProperties().getProperty('DASHBOARD_ADMIN_KEY');
  const isAuthorized = !!adminKey && key === adminKey;
  if (!isAuthorized) {
    return jsonResponse({ status: 'error', message: 'Unauthorized — check the admin key.' });
  }
  try {
    const groups = getPendingGroupsForAdminPanel();
    return jsonResponse({ status: 'success', groups: groups });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message || String(err) });
  }
}

/* ---------------- ADMIN PANEL: APPROVE / REJECT / SEARCH / RESEND (key-gated write/read actions) ----------------
 * Called from dashboard.html's Admin Panel via doPost with
 * { action: 'approveGroup' | 'rejectGroup' | 'searchRegistrations' | 'resendPass', ... , key }.
 * approveGroup/rejectGroup reuse the EXACT same approveGroupById()/rejectGroupById()
 * core used by the Sheet's "Approve Entire Group" menu action (AdminApproval.gs) —
 * so there is still only one code path that can actually flip a Status to
 * APPROVED/REJECTED, no matter which UI triggered it. */
function handleAdminAction(data) {
  const adminKey = PropertiesService.getScriptProperties().getProperty('DASHBOARD_ADMIN_KEY');
  const isAuthorized = !!adminKey && data.key === adminKey;
  if (!isAuthorized) {
    return jsonResponse({ status: 'error', message: 'Unauthorized — check the admin key.' });
  }

  try {
    if (data.action === 'approveGroup') {
      if (!data.groupId) return jsonResponse({ status: 'error', message: 'Missing groupId.' });
      const result = approveGroupById(data.groupId);
      if (!result.found) {
        return jsonResponse({ status: 'error', message: 'Group ID "' + data.groupId + '" ke liye koi row nahi mili.' });
      }
      const links = getApprovalLinksForGroup(result.affectedRows);
      return jsonResponse({
        status: 'success',
        message: result.newlyApproved + ' participant(s) approved, ' + result.alreadyApproved + ' pehle se approved the.',
        participants: links
      });
    }

    if (data.action === 'rejectGroup') {
      if (!data.groupId) return jsonResponse({ status: 'error', message: 'Missing groupId.' });
      const result = rejectGroupById(data.groupId);
      if (!result.found) {
        return jsonResponse({ status: 'error', message: 'Group ID "' + data.groupId + '" ke liye koi row nahi mili.' });
      }
      let message = result.newlyRejected + ' participant(s) rejected.';
      if (result.skippedApproved > 0) {
        message += ' ' + result.skippedApproved + ' already-APPROVED participant(s) is group mein the — unhe skip kiya gaya (Sheet se manually handle karo agar zaroorat ho).';
      }
      return jsonResponse({ status: 'success', message: message });
    }

    // --- Search any registration (individual OR group member, any status) by
    // Ticket ID / Name / Phone / Group ID — read-only lookup used by the
    // Admin Panel's "🔍 Find / Resend" box so WhatsApp links & Ticket ID are
    // never "lost" once the Approve-click flash message disappears. ---
    if (data.action === 'searchRegistrations') {
      if (!data.query || !String(data.query).trim()) {
        return jsonResponse({ status: 'error', message: 'Search query missing.' });
      }
      const results = searchRegistrationsForAdminPanel(data.query);
      return jsonResponse({ status: 'success', results: results });
    }

    // --- Re-send the pass email + refresh the WhatsApp links for one
    // already-APPROVED row (mirrors the Sheet menu's "Resend Pass"). ---
    if (data.action === 'resendPass') {
      if (!data.rowNumber) return jsonResponse({ status: 'error', message: 'Missing rowNumber.' });
      const result = resendPassForRowNumber(data.rowNumber);
      return jsonResponse({ status: result.success ? 'success' : 'error', message: result.message });
    }

    return jsonResponse({ status: 'error', message: 'Unknown action: ' + data.action });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message || String(err) });
  }
}

/* ---------------- LIVE DASHBOARD DATA FEED ----------------
 * Returns aggregate JSON stats (counts, revenue, parish breakdown) for the
 * custom dashboard.html page to poll and render. Aggregate numbers contain
 * NO personal data (no names/phones/UTRs) and are safe to show to anyone
 * with the link. The "needsReview" detail list (which does include names)
 * is only included when the caller supplies the correct DASHBOARD_ADMIN_KEY
 * (set via Script Properties) — this is a simple shared-secret, not real
 * authentication, so treat the dashboard URL + key as "keep it private",
 * not as a fully secured admin page. --------------------------------- */

function handleDashboardData(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResponse({ status: 'error', message: 'Responses sheet not found.' });

  const stats = {
    totalGroups: 0,
    totalParticipants: 0,
    pending: 0,
    approved: 0,
    rejected: 0,
    totalRevenue: 0,
    checkedIn: 0,
    duplicateFlags: 0,
    ocrMismatchFlags: 0,
    parishCounts: {},
    needsReview: []
  };

  const adminKey = PropertiesService.getScriptProperties().getProperty('DASHBOARD_ADMIN_KEY');
  const isAuthorized = !!adminKey && key === adminKey;

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, LAST_COL).getValues();

    data.forEach(function (row) {
      const participantNo = String(row[COL_PARTICIPANTNO - 1] || '');
      const status         = String(row[COL_STATUS - 1] || '').trim().toUpperCase();
      const parish         = String(row[COL_PARISH - 1] || '').trim();
      const entryStatus    = String(row[COL_ENTRYSTATUS - 1] || '').trim();
      const dupUtr         = String(row[COL_DUP_UTR - 1] || '');
      const utrMatch       = String(row[COL_UTR_MATCH - 1] || '');
      const totalAmount    = Number(row[COL_TOTALAMOUNT - 1]) || 0;

      stats.totalParticipants++;
      if (participantNo.indexOf('1 of') === 0) {
        stats.totalGroups++;
        if (status === 'APPROVED') stats.totalRevenue += totalAmount;
      }
      if (status === 'PENDING') stats.pending++;
      else if (status === 'APPROVED') stats.approved++;
      else if (status === 'REJECTED') stats.rejected++;
      if (entryStatus === 'Checked-In') stats.checkedIn++;
      if (dupUtr.indexOf('DUPLICATE') === 0) stats.duplicateFlags++;
      if (utrMatch.indexOf('Mismatch') === 0) stats.ocrMismatchFlags++;
      if (parish) stats.parishCounts[parish] = (stats.parishCounts[parish] || 0) + 1;

      if (isAuthorized && (dupUtr.indexOf('DUPLICATE') === 0 || utrMatch.indexOf('Mismatch') === 0)) {
        stats.needsReview.push({
          fullName: row[COL_FULLNAME - 1],
          parish: parish,
          utr: row[COL_UTR - 1],
          duplicateCheck: dupUtr,
          ocrMatch: utrMatch,
          adminNotes: row[COL_ADMIN_NOTES - 1]
        });
      }
    });
  }

  stats.authorized = isAuthorized;
  stats.lastUpdated = new Date().toISOString();
  return jsonResponse({ status: 'success', data: stats });
}

/* ---------------- GATE CHECK-IN (triggered by scanning the QR pass) ---------------- */

function handleGateCheckIn(rawTicketId) {
  // Validate the ticket ID shape BEFORE doing anything else — rejects obviously
  // malformed/malicious input early (e.g. script injection attempts in the URL).
  const ticketId = String(rawTicketId || '').trim();
  if (!/^[A-Za-z0-9\-]{3,40}$/.test(ticketId)) {
    return gateHtml('❌', '#ff5c6c', 'Invalid Ticket', '<p>This QR code / link is not a valid ticket.</p>');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return gateHtml('❌', '#ff5c6c', 'System Error', '<p>Responses sheet not found.</p>');

  // The whole "find row -> check status -> read/write Entry Status" sequence is
  // wrapped in a script lock so that two guards scanning the SAME ticket at the
  // same moment can never both get a successful "first scan" result.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (lockErr) {
    return gateHtml('❌', '#ff5c6c', 'Busy', '<p>System is busy verifying another ticket, please scan again in a moment.</p>');
  }

  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return gateHtml('❌', '#ff5c6c', 'Ticket Not Found', '<p>Ticket ID: ' + escapeHtml(ticketId) + '</p>');

    const data = sheet.getRange(2, 1, lastRow - 1, LAST_COL).getValues();

    for (let i = 0; i < data.length; i++) {
      const rowTicketId = String(data[i][COL_TICKETID - 1]).trim();
      if (rowTicketId && rowTicketId === ticketId) {
        const sheetRow       = i + 2;
        const fullName       = data[i][COL_FULLNAME - 1];
        const parish         = data[i][COL_PARISH - 1];
        const leaderName     = data[i][COL_LEADERNAME - 1];
        const leaderPhone    = data[i][COL_LEADERPHONE - 1];
        const guardianName   = data[i][COL_GUARDIANNAME - 1];
        const guardianPhone  = data[i][COL_GUARDIANPHONE - 1];
        const status         = String(data[i][COL_STATUS - 1]).trim().toUpperCase();
        const entryStatus    = String(data[i][COL_ENTRYSTATUS - 1] || '').trim();

        if (status !== 'APPROVED') {
          return gateHtml('❌', '#ff5c6c', 'Entry Denied',
            '<p>This registration is not APPROVED yet (Status: ' + escapeHtml(status || 'PENDING') + ').</p>' +
            buildGateDetails(fullName, parish, ticketId, leaderName, leaderPhone, guardianName, guardianPhone));
        }

        if (entryStatus === 'Checked-In') {
          const checkinTimeRaw = data[i][COL_CHECKIN_TIME - 1];
          const timeStr = checkinTimeRaw
            ? Utilities.formatDate(new Date(checkinTimeRaw), Session.getScriptTimeZone(), 'dd MMM, hh:mm a')
            : 'earlier';
          return gateHtml('⚠️', '#ffc107', 'Already Checked In',
            buildGateDetails(fullName, parish, ticketId, leaderName, leaderPhone, guardianName, guardianPhone) +
            '<p style="margin-top:14px;color:#b45309;font-weight:700;">This ticket was already scanned at ' +
            escapeHtml(timeStr) + '. Possible duplicate / re-entry — verify manually before allowing in.</p>');
        }

        // First valid scan — mark as checked-in right now (still inside the lock).
        const now = new Date();
        sheet.getRange(sheetRow, COL_ENTRYSTATUS).setValue('Checked-In');
        sheet.getRange(sheetRow, COL_CHECKIN_TIME).setValue(now);

        return gateHtml('✅', '#28a745', 'Entry Granted',
          buildGateDetails(fullName, parish, ticketId, leaderName, leaderPhone, guardianName, guardianPhone));
      }
    }

    return gateHtml('❌', '#ff5c6c', 'Ticket Not Found', '<p>No registration found for Ticket ID: ' + escapeHtml(ticketId) + '</p>');

  } finally {
    lock.releaseLock();
  }
}

// Info shown on the gate check-in page when a QR is scanned. Per the admin's
// explicit requirement, this includes Full Name, Parish, Ticket ID, Youth
// Leader Name+Phone, and Guardian Name+Phone — every participant-controlled
// value is passed through escapeHtml() before being inserted into the HTML,
// so this stays safe from stored/reflected XSS even though it shows more
// data than a minimal "name-only" gate page would.
function buildGateDetails(fullName, parish, ticketId, leaderName, leaderPhone, guardianName, guardianPhone) {
  return `
    <table style="width:100%;margin-top:16px;font-size:15px;text-align:left;border-collapse:collapse;">
      <tr><td style="padding:5px 0;color:#888;">Name</td><td style="padding:5px 0;font-weight:700;">${escapeHtml(fullName)}</td></tr>
      <tr><td style="padding:5px 0;color:#888;">Ticket ID</td><td style="padding:5px 0;font-weight:700;">${escapeHtml(ticketId)}</td></tr>
      <tr><td style="padding:5px 0;color:#888;">Parish</td><td style="padding:5px 0;">${escapeHtml(parish)}</td></tr>
      <tr><td style="padding:5px 0;color:#888;">Youth Leader</td><td style="padding:5px 0;">${escapeHtml(leaderName)} (${escapeHtml(leaderPhone)})</td></tr>
      <tr><td style="padding:5px 0;color:#888;">Guardian</td><td style="padding:5px 0;">${escapeHtml(guardianName)} (${escapeHtml(guardianPhone)})</td></tr>
    </table>`;
}

function gateHtml(icon, color, title, bodyHtml) {
  const html = `
    <!DOCTYPE html>
    <html><head><meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
      body{font-family:Arial,sans-serif;background:#0b0014;color:#f1e9ff;margin:0;padding:24px;}
      .card{max-width:420px;margin:30px auto;background:#160a26;border-radius:16px;padding:28px 22px;text-align:center;box-shadow:0 0 30px rgba(138,43,226,0.3);}
      .icon{font-size:56px;}
      h1{font-size:22px;margin:10px 0;color:${color};}
    </style></head>
    <body>
      <div class="card">
        <div class="icon">${icon}</div>
        <h1>${escapeHtml(title)}</h1>
        ${bodyHtml}
      </div>
    </body></html>`;
  return HtmlService.createHtmlOutput(html).setTitle('Gate Check-in — ' + title);
}

/* ---------------- SHARED HELPERS ---------------- */

// Escapes a value for safe insertion into HTML (prevents stored/reflected XSS
// from participant-controlled fields like Full Name / Parish, and from the
// raw `ticket` URL parameter on the public gate check-in page).
function escapeHtml(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}