/**
 * Google Apps Script — Expert Poll Result Collector
 *
 * SETUP:
 * 1. Go to https://script.google.com
 * 2. Create a new project
 * 3. Paste this code
 * 4. Create a Google Sheet and copy its ID from the URL
 * 5. Replace SHEET_ID below with your sheet ID
 * 6. Deploy as Web App:
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 7. Copy the deployment URL and set it as BACKEND_URL in index.html
 *
 * The script saves each submission as a new row in the sheet,
 * plus stores the full JSON in a separate "Raw" sheet tab.
 */

const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID_HERE';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);

    // Save raw JSON to "Raw" sheet
    let rawSheet = ss.getSheetByName('Raw');
    if (!rawSheet) {
      rawSheet = ss.insertSheet('Raw');
      rawSheet.appendRow(['Timestamp', 'Reviewer', 'JSON']);
    }
    rawSheet.appendRow([
      new Date().toISOString(),
      data.reviewer?.name || 'Anonymous',
      JSON.stringify(data)
    ]);

    // Save summary to "Summary" sheet
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

    return ContentService
      .createTextOutput(JSON.stringify({status: 'ok'}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({status: 'ok', message: 'Expert Poll backend is running'}))
    .setMimeType(ContentService.MimeType.JSON);
}
