/**
 * Google Apps Script — Expert Poll Result Collector
 *
 * SETUP:
 * 1. Go to https://script.google.com
 * 2. Create a new project
 * 3. Paste this code
 * 4. Replace SHEET_ID below with your Google Sheet ID
 *    (from the URL: docs.google.com/spreadsheets/d/{SHEET_ID}/edit)
 * 5. Deploy as Web App:
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Copy the deployment URL and set it as BACKEND_URL in index.html
 *
 * NOTE: The frontend sends Content-Type: text/plain to avoid CORS preflight.
 * The body is still JSON — we parse it from e.postData.contents.
 */

const SHEET_ID = '1DNdHVzLtI57Z9WmvdT--tGZa8rX70j3DEoZnRzZ0Y0s';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const isPartial = data._isPartial === true;

    // Save raw JSON to "Raw" sheet
    let rawSheet = ss.getSheetByName('Raw');
    if (!rawSheet) {
      rawSheet = ss.insertSheet('Raw');
      rawSheet.appendRow(['Timestamp', 'Reviewer', 'Type', 'JSON']);
    }
    rawSheet.appendRow([
      new Date().toISOString(),
      data.reviewer?.name || 'Anonymous',
      isPartial ? 'partial' : 'final',
      JSON.stringify(data)
    ]);

    // Only write to Summary on final submission
    if (!isPartial) {
      let summarySheet = ss.getSheetByName('Summary');
      if (!summarySheet) {
        summarySheet = ss.insertSheet('Summary');
        summarySheet.appendRow([
          'Timestamp', 'Reviewer', 'Expertise',
          'Q1_rank_X', 'Q1_rank_Y', 'Q1_rank_Z', 'Q1_confidence',
          'Q2_rank_X', 'Q2_rank_Y', 'Q2_rank_Z', 'Q2_confidence',
          'Q3_rank_X', 'Q3_rank_Y', 'Q3_rank_Z', 'Q3_confidence',
          'Q4_rank_X', 'Q4_rank_Y', 'Q4_rank_Z', 'Q4_confidence',
          'Q5_rank_X', 'Q5_rank_Y', 'Q5_rank_Z', 'Q5_confidence',
          'Q6_rank_X', 'Q6_rank_Y', 'Q6_rank_Z', 'Q6_confidence',
          'Q7_rank_X', 'Q7_rank_Y', 'Q7_rank_Z', 'Q7_confidence',
          'Q8_rank_X', 'Q8_rank_Y', 'Q8_rank_Z', 'Q8_confidence',
          'Q9_rank_X', 'Q9_rank_Y', 'Q9_rank_Z', 'Q9_confidence',
          'Q11_rank_X', 'Q11_rank_Y', 'Q11_rank_Z', 'Q11_confidence',
          'Q12_rank_X', 'Q12_rank_Y', 'Q12_rank_Z', 'Q12_confidence',
          'Final_directions', 'Final_feedback'
        ]);
      }

      const row = [
        new Date().toISOString(),
        data.reviewer?.name || '',
        data.reviewer?.expertise || ''
      ];

      const qids = ['Q1','Q2','Q3','Q4','Q5','Q6','Q7','Q8','Q9','Q11','Q12'];
      qids.forEach(qid => {
        const q = data.questions?.[qid] || {};
        row.push(q.rankings?.X || '', q.rankings?.Y || '', q.rankings?.Z || '', q.confidence || '');
      });

      row.push(data.final?.directions || '', data.final?.feedback || '');
      summarySheet.appendRow(row);
    }

    return ContentService
      .createTextOutput(JSON.stringify({status: 'ok', type: isPartial ? 'partial' : 'final'}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  const email = (e && e.parameter && e.parameter.email) || '';

  if (action === 'load' && email) {
    return loadProgress(email);
  }

  return ContentService
    .createTextOutput(JSON.stringify({status: 'ok', message: 'Expert Poll backend is running'}))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Load the most recent saved state for a given email.
 * Searches the "Raw" sheet in reverse order for the latest entry
 * matching the email (case-insensitive), then returns its JSON payload.
 */
function loadProgress(email) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const rawSheet = ss.getSheetByName('Raw');
    if (!rawSheet) {
      return ContentService
        .createTextOutput(JSON.stringify({status: 'not_found', message: 'No submissions yet'}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data = rawSheet.getDataRange().getValues();
    const emailLower = email.toLowerCase().trim();

    // Walk backwards to find the most recent matching entry
    for (let i = data.length - 1; i >= 1; i--) {
      const jsonStr = data[i][3]; // Column D = JSON payload
      if (!jsonStr) continue;
      try {
        const parsed = JSON.parse(jsonStr);
        const savedEmail = (parsed.reviewer && parsed.reviewer.email) || '';
        if (savedEmail.toLowerCase().trim() === emailLower) {
          return ContentService
            .createTextOutput(JSON.stringify({status: 'ok', data: parsed}))
            .setMimeType(ContentService.MimeType.JSON);
        }
      } catch (parseErr) {
        continue;
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({status: 'not_found', message: 'No saved progress for this email'}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
