/**
 * Google Apps Script — Expert Poll v2 Result Collector
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
 * v2 CHANGES:
 * - 2-arm preference (X better / Y better / Equal) instead of 3-arm ranking
 * - Q1-Q9 comparison questions with preference + confidence + reasoning
 * - Q10 hypothesis with per-proposal scoring (novelty, groundedness, feasibility, worth, known, confidence)
 * - Frontend sends {reviewer_name, reviewer_email, type, payload: {field_confidence, answers}}
 */

const SHEET_ID = '1DNdHVzLtI57Z9WmvdT--tGZa8rX70j3DEoZnRzZ0Y0s';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const isPartial = data.type === 'partial';

    // Save raw JSON to "Raw_v2" sheet
    let rawSheet = ss.getSheetByName('Raw_v2');
    if (!rawSheet) {
      rawSheet = ss.insertSheet('Raw_v2');
      rawSheet.appendRow(['Timestamp', 'Reviewer', 'Email', 'Type', 'JSON']);
    }
    rawSheet.appendRow([
      new Date().toISOString(),
      data.reviewer_name || 'Anonymous',
      data.reviewer_email || '',
      isPartial ? 'partial' : 'final',
      JSON.stringify(data)
    ]);

    // Only write to Summary_v2 on final submission
    if (!isPartial) {
      let summarySheet = ss.getSheetByName('Summary_v2');
      if (!summarySheet) {
        summarySheet = ss.insertSheet('Summary_v2');
        const headers = [
          'Timestamp', 'Reviewer', 'Email', 'Field_Confidence',
          // Q1-Q9: preference + confidence + reasoning
          'Q1_preference', 'Q1_confidence', 'Q1_reasoning',
          'Q2_preference', 'Q2_confidence', 'Q2_reasoning',
          'Q3_preference', 'Q3_confidence', 'Q3_reasoning',
          'Q4_preference', 'Q4_confidence', 'Q4_reasoning',
          'Q5_preference', 'Q5_confidence', 'Q5_reasoning',
          'Q6_preference', 'Q6_confidence', 'Q6_reasoning',
          'Q7_preference', 'Q7_confidence', 'Q7_reasoning',
          'Q8_preference', 'Q8_confidence', 'Q8_reasoning',
          'Q9_preference', 'Q9_confidence', 'Q9_reasoning',
        ];
        // Q10 hypothesis: 5 proposals from X, 7 from Y
        for (let i = 1; i <= 5; i++) {
          headers.push(
            `Q10_X_P${i}_novelty`, `Q10_X_P${i}_ground`, `Q10_X_P${i}_feas`,
            `Q10_X_P${i}_worth`, `Q10_X_P${i}_known`, `Q10_X_P${i}_conf`
          );
        }
        for (let i = 1; i <= 7; i++) {
          headers.push(
            `Q10_Y_P${i}_novelty`, `Q10_Y_P${i}_ground`, `Q10_Y_P${i}_feas`,
            `Q10_Y_P${i}_worth`, `Q10_Y_P${i}_known`, `Q10_Y_P${i}_conf`
          );
        }
        // Final questions
        headers.push('Final_directions', 'Final_ai_questions', 'Final_feedback', 'Final_other');
        summarySheet.appendRow(headers);
      }

      const answers = data.payload?.answers || {};
      const row = [
        new Date().toISOString(),
        data.reviewer_name || '',
        data.reviewer_email || '',
        data.payload?.field_confidence || ''
      ];

      // Q1-Q9
      for (let i = 1; i <= 9; i++) {
        const qid = 'Q' + i;
        const q = answers[qid] || {};
        row.push(q.preference || '', q.confidence || '', q.reasoning || '');
      }

      // Q10 hypothesis proposals
      const q10 = answers['Q10'] || {};
      const xProposals = q10.proposals_x || {};
      const yProposals = q10.proposals_y || {};

      for (let i = 1; i <= 5; i++) {
        const key = 'x_p' + i;
        const p = xProposals[key] || {};
        row.push(p.novelty || '', p.groundedness || '', p.feasibility || '',
                 p.worth_investigating || '', p.already_known || '', p.confidence || '');
      }
      for (let i = 1; i <= 7; i++) {
        const key = 'y_p' + i;
        const p = yProposals[key] || {};
        row.push(p.novelty || '', p.groundedness || '', p.feasibility || '',
                 p.worth_investigating || '', p.already_known || '', p.confidence || '');
      }

      // Final questions
      const fq = data.payload?.final_questions || {};
      row.push(fq.directions || '', fq.ai_questions || '', fq.feedback || '', fq.other || '');

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
    .createTextOutput(JSON.stringify({status: 'ok', message: 'Expert Poll v2 backend is running'}))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Load the most recent saved state for a given email.
 * Searches the "Raw_v2" sheet in reverse order for the latest entry
 * matching the email (case-insensitive), then returns its JSON payload.
 */
function loadProgress(email) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const rawSheet = ss.getSheetByName('Raw_v2');
    if (!rawSheet) {
      return ContentService
        .createTextOutput(JSON.stringify({status: 'not_found', message: 'No submissions yet'}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data = rawSheet.getDataRange().getValues();
    const emailLower = email.toLowerCase().trim();

    // Walk backwards to find the most recent matching entry
    for (let i = data.length - 1; i >= 1; i--) {
      const jsonStr = data[i][4]; // Column E = JSON payload (shifted by new Email column)
      if (!jsonStr) continue;
      try {
        const parsed = JSON.parse(jsonStr);
        const savedEmail = (parsed.reviewer_email || '').toLowerCase().trim();
        if (savedEmail === emailLower) {
          return ContentService
            .createTextOutput(JSON.stringify({status: 'ok', payload: parsed.payload}))
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
