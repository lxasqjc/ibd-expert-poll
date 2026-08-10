/**
 * Run this ONCE in Google Apps Script to recreate the Summary sheet
 * with corrected schema (Q10 included, Q12 removed, Q11 per-proposal columns).
 * Also backfills Kexin's Q11 per-proposal data from the Raw sheet.
 *
 * Steps:
 * 1. Open Google Apps Script for this spreadsheet
 * 2. Paste this function
 * 3. Run rebuildSummary()
 */

function rebuildSummary() {
  const ss = SpreadsheetApp.openById('1DNdHVzLtI57Z9WmvdT--tGZa8rX70j3DEoZnRzZ0Y0s');

  // Delete old Summary sheet
  const oldSummary = ss.getSheetByName('Summary');
  if (oldSummary) ss.deleteSheet(oldSummary);

  // Create new Summary sheet with corrected headers
  const summarySheet = ss.insertSheet('Summary');
  const headers = [
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
    'Q10_rank_X', 'Q10_rank_Y', 'Q10_rank_Z', 'Q10_confidence'
  ];

  // Q11 per-proposal columns: X has 5 proposals, Y has 7
  const q11Counts = {X: 5, Y: 7};
  ['X', 'Y'].forEach(label => {
    for (let i = 1; i <= q11Counts[label]; i++) {
      headers.push(
        `Q11_${label}_P${i}_novelty`,
        `Q11_${label}_P${i}_ground`,
        `Q11_${label}_P${i}_feas`,
        `Q11_${label}_P${i}_worth`,
        `Q11_${label}_P${i}_invest`,
        `Q11_${label}_P${i}_conf`
      );
    }
  });
  headers.push('Q11_confidence', 'Final_directions', 'Final_ai_questions', 'Final_feedback', 'Final_other');
  summarySheet.appendRow(headers);

  // Find the last final submission for each reviewer in Raw
  const rawSheet = ss.getSheetByName('Raw');
  if (!rawSheet) return;

  const rawData = rawSheet.getDataRange().getValues();
  const latestByEmail = {};

  for (let i = 1; i < rawData.length; i++) {
    const jsonStr = rawData[i][3];
    if (!jsonStr) continue;
    try {
      const parsed = JSON.parse(jsonStr);
      const email = (parsed.reviewer && parsed.reviewer.email) || parsed._syncEmail || '';
      if (!email) continue;
      // Keep the latest entry per email (overwrite)
      latestByEmail[email] = {timestamp: rawData[i][0], data: parsed};
    } catch(e) { continue; }
  }

  // Write one summary row per reviewer
  Object.values(latestByEmail).forEach(({timestamp, data}) => {
    // Skip test accounts
    if (data.reviewer && data.reviewer.name && data.reviewer.name.toLowerCase().includes('chrome test')) return;

    const row = [
      timestamp,
      data.reviewer?.name || '',
      data.reviewer?.field_confidence || ''
    ];

    // Q1-Q10 rankings
    const rankQids = ['Q1','Q2','Q3','Q4','Q5','Q6','Q7','Q8','Q9','Q10'];
    rankQids.forEach(qid => {
      const q = data.questions?.[qid] || {};
      row.push(q.rankings?.X || '', q.rankings?.Y || '', q.rankings?.Z || '', q.confidence || '');
    });

    // Q11 per-proposal
    const q11 = data.questions?.Q11 || {};
    const q11Proposals = q11.hypothesis_proposals || {};
    ['X', 'Y'].forEach(label => {
      const proposals = q11Proposals[label] || [];
      const count = q11Counts[label];
      for (let i = 0; i < count; i++) {
        const p = proposals[i] || {};
        row.push(
          p.novelty || '',
          p.groundedness || '',
          p.feasibility || '',
          p.worth_investigating ? 'Y' : '',
          p.already_investigating ? 'Y' : '',
          p.proposal_confidence || ''
        );
      }
    });
    row.push(q11.confidence || '');

    row.push(
      data.final?.directions || '',
      data.final?.ai_questions || '',
      data.final?.feedback || '',
      data.final?.other || ''
    );

    summarySheet.appendRow(row);
  });

  Logger.log('Summary sheet rebuilt with ' + (Object.keys(latestByEmail).length) + ' reviewer(s)');
}
