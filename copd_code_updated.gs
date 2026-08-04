// ============================================================
// PIM AUTOMATION SCRIPT
// ============================================================

const SPREADSHEET_ID = '1O9riaBS2q0QWrlL3EsNDeYcXFO3DCPiCjq_S2SMrDls';
const CLAUDE_API_KEY = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');

// Sheet names
const SHEETS = {
  DICTIONARY: 'DICTIONARY',
  HYBRIS: 'HYBRIS',
  PRODUCT_LIST: 'PRODUCT_LIST',
  CATEGORY_LIST: 'CATEGORY_LIST',
  BATCH: 'BATCH',
  SCRIPT_STATE: 'SCRIPT_STATE',
  PHASE_1AB_WORKING: 'PHASE_1AB_WORKING',
  PHASE_2_WORKING: 'PHASE_2_WORKING',
  PHASE_3_WORKING: 'PHASE_3_WORKING',
  NEWADDITIONS_STAGING: 'NEWADDITIONS_STAGING',
  NEWADDITIONS_APPROVED: 'NEWADDITIONS_APPROVED',
  DICTIONARY_SUBSET: 'DICTIONARY_SUBSET',
  PL_ATTRIBUTELIST: 'PL_ATTRIBUTELIST',
  REVIEW: 'REVIEW',
  PROMPTS: 'PROMPTS',
  MASTER_OUTPUT: 'MASTER_OUTPUT',
  DICT_READY: 'DICT_READY',
  DICT_READY_DROPDOWNS: 'DICT_READY_DROPDOWNS',
  LOG: 'LOG',
  INTENDED_FOR: 'INTENDED_FOR',
  REVIEW_NOTES: 'REVIEW_NOTES',
  PHASE2_DICT_ENTRIES: 'PHASE2_DICT_ENTRIES',
  PHASE2_DICT_DROPDOWNS: 'PHASE2_DICT_DROPDOWNS',
  JUDGE_NOTEBOOKLM: 'JUDGE_NOTEBOOKLM',
  JUDGE_GEMINI: 'JUDGE_GEMINI',
  DEBATE_REPROMPT: 'DEBATE_REPROMPT'
};

// ============================================================
// UTILITY FUNCTIONS
// ============================================================

function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found`);
  return sheet;
}

function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
}

function getSheetData(sheetName) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { headers: data[0] || [], rows: [] };
  const headers = data[0];
  const rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
  return { headers, rows };
}

function writeToSheet(sheetName, headers, rows) {
  const sheet = getSheet(sheetName);
  sheet.clearContents();
  if (rows.length === 0) return;
  const output = [headers, ...rows.map(r => headers.map(h => r[h] || ''))];
  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}

function appendToSheet(sheetName, headers, rows) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  if (rows.length === 0) return;
  const output = rows.map(r => headers.map(h => r[h] || ''));
  sheet.getRange(lastRow + 1, 1, output.length, output[0].length).setValues(output);
}

function log(message) {
  Logger.log(message)
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('LOG');
    if (!sheet) sheet = ss.insertSheet('LOG');
    const timestamp = new Date().toLocaleTimeString();
    sheet.insertRowBefore(2);
    sheet.getRange(2, 1).setValue(`${timestamp}  ${message}`);
    // Keep only last 100 log entries
    if (sheet.getLastRow() > 101) {
      sheet.deleteRows(102, sheet.getLastRow() - 101);
    }
  } catch(e) {
    // Silent fail — logging should never break the main script
  }
}

function clearLog() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('LOG');
  if (sheet) {
    sheet.clearContents();
    sheet.getRange(1, 1).setValue('Log').setFontWeight('bold');
  }
}

function saveState(state) {
  const sheet = getSheet(SHEETS.SCRIPT_STATE);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
  Object.entries(state).forEach(([key, value], i) => {
    sheet.getRange(i + 2, 1, 1, 2).setValues([[key, JSON.stringify(value)]]);
  });
}

function loadState() {
  const { rows } = getSheetData(SHEETS.SCRIPT_STATE);
  const state = {};
  rows.forEach(r => {
    try { state[r['Key']] = JSON.parse(r['Value']); }
    catch (e) { state[r['Key']] = r['Value']; }
  });
  return state;
}

function clearState() {
  getSheet(SHEETS.SCRIPT_STATE).clearContents();
}

function sendNotification(subject, body) {
  //GmailApp.sendEmail(Session.getActiveUser().getEmail(), subject, body);
  log(`NOTIFICATION: ${subject}`);
}

function extractStepContent(output, stepTag) {
  const startTag = `[${stepTag}]`;
  const endTag = `[/${stepTag}]`;
  const startIndex = output.indexOf(startTag);
  if (startIndex === -1) return null;
  const endIndex = output.indexOf(endTag, startIndex);
  if (endIndex === -1) return output.substring(startIndex + startTag.length).trim();
  return output.substring(startIndex + startTag.length, endIndex).trim();
}

function formatDictionaryAsCSV(dictRows) {
  const headers = ['Attribute Name', 'Description'];
  const lines = [headers.join(',')];
  dictRows.forEach(r => {
    const line = headers.map(h => {
      const val = (r[h] || '').toString().replace(/"/g, '""');
      return `"${val}"`;
    });
    lines.push(line.join(','));
  });
  return lines.join('\n');
}

function formatProductListAsCSV(productRows) {
  const headers = ['L4 Category', 'Product Name'];
  const lines = [headers.join(',')];
  productRows.forEach(r => {
    const line = [
      `"${(r['Retail_Category4'] || r['L4 Category'] || '').toString().replace(/"/g, '""')}"`,
      `"${(r['Product Name'] || '').toString().replace(/"/g, '""')}"`
    ];
    lines.push(line.join(','));
  });
  return lines.join('\n');
}

function doGet(e) {
  const action = e.parameter.action;

  if (action === 'getReview') {
    try {
      const { rows } = getSheetData(SHEETS.REVIEW);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, data: rows }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === 'getBatch') {
    try {
      const { rows } = getSheetData(SHEETS.BATCH);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, data: rows }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === 'runPhase') {
    const phase = e.parameter.phase;
    try {
      clearLog();
      log('Log cleared');
      //log(`doGet runPhase called with phase: ${phase}`); 
      // Save which phase to run
      PropertiesService.getScriptProperties().setProperty('PENDING_PHASE', phase);
      log(`PENDING_PHASE set to: ${phase}`); 
      // Create a trigger to fire in 1 minute
      ScriptApp.newTrigger('runPendingPhase')
        .timeBased()
        .after(10000) // 1 second
        .create();
      return ContentService
        .createTextOutput(JSON.stringify({ 
          success: true, 
          message: `Phase ${phase} queued — check your Sheet in a few minutes` 
        }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === 'getLog') {
    try {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheet = ss.getSheetByName('LOG');
      if (!sheet) return ContentService
        .createTextOutput(JSON.stringify({ success: true, entries: [] }))
        .setMimeType(ContentService.MimeType.JSON);
      const data = sheet.getDataRange().getValues();
      const entries = data.slice(1).map(r => r[0]).filter(Boolean);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, entries }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === 'clearLog') {
    try {
      clearLog();
      return ContentService
        .createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === 'getLastMasterRow') {
    try {
      const sheet = getSheet(SHEETS.MASTER_OUTPUT);
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: true, data: null }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const lastRowData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
      const row = {};
      headers.forEach((h, i) => { if (h) row[h] = lastRowData[i]; });
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, data: row, totalRows: lastRow - 1 }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === 'getPhase2Suggestions') {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.PHASE2_DICT_ENTRIES);
    if (!sheet || sheet.getLastRow() < 2) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, suggestions: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nameCol = headers.indexOf('Attribute Name');
    const catCol = headers.indexOf('L4 Category');
    const suggestions = data.slice(1)
      .filter(r => r[nameCol])
      .map(r => catCol > -1 
        ? `${r[nameCol]} (${r[catCol]})` 
        : r[nameCol]
      ).reverse();
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, suggestions }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

  if (action === 'getPrompts') {
    try {
      const sheet = getSheet(SHEETS.PROMPTS);
      const data = sheet.getDataRange().getValues();
      const entries = data.map(r => r[0]).filter(Boolean);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, data: entries }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === 'getJudgeStatus') {
    try {
      const status = {};
      [['notebooklm', SHEETS.JUDGE_NOTEBOOKLM], ['gemini', SHEETS.JUDGE_GEMINI]].forEach(pair => {
        try {
          const text = readJudgeTab(pair[1]);
          status[pair[0]] = { chars: text.length, cards: (text.match(/🔸/g) || []).length };
        } catch(e) { status[pair[0]] = { chars: 0, cards: 0 }; }
      });
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, status }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === 'getDebateReprompt') {
    try {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheet = ss.getSheetByName(SHEETS.DEBATE_REPROMPT);
      if (!sheet || sheet.getLastRow() < 2) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: true, items: [] }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      const data = sheet.getDataRange().getValues();
      const items = data.slice(1).filter(r => r[0]).map(r => ({ attribute: r[0], prompt: r[1] }));
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, items }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ success: false, error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    // Existing review save
    if (payload.type === 'saveReview') {
      const { rowIndex, decision, updatedName, updatedDropdownName, notes } = payload;
      const sheet = getSheet(SHEETS.REVIEW);
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const yourDecisionCol = headers.indexOf('Your Decision') + 1;
      const updatedNameCol = headers.indexOf('Updated Attribute Name') + 1;
      const updatedDropdownCol = headers.indexOf('Updated Dropdown Name') + 1;
      const notesCol = headers.indexOf('Notes') + 1;
      const sheetRow = rowIndex + 1;
      if (yourDecisionCol > 0) sheet.getRange(sheetRow, yourDecisionCol).setValue(decision);
      if (updatedNameCol > 0) sheet.getRange(sheetRow, updatedNameCol).setValue(updatedName);
      if (updatedDropdownCol > 0) sheet.getRange(sheetRow, updatedDropdownCol).setValue(updatedDropdownName);
      if (notesCol > 0) sheet.getRange(sheetRow, notesCol).setValue(notes);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Batch status update
    if (payload.type === 'updateBatch') {
      const { rowIndex, status } = payload;
      const sheet = getSheet(SHEETS.BATCH);
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const statusCol = headers.indexOf('Status') + 1;
      if (statusCol > 0) sheet.getRange(rowIndex + 1, statusCol).setValue(status);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (payload.type === 'saveJudgeOutput') {
      const { judge, text } = payload;
      const tabName = judge === 'gemini' ? SHEETS.JUDGE_GEMINI : SHEETS.JUDGE_NOTEBOOKLM;
      const sheet = getSheet(tabName);
      sheet.clearContents();
      const maxChars = 45000;
      let remaining = text || '';
      let row = 1;
      while (remaining.length > 0) {
        sheet.getRange(row, 1).setValue(remaining.substring(0, maxChars));
        remaining = remaining.substring(maxChars);
        row++;
      }
      const cardCount = ((text || '').match(/🔸/g) || []).length;
      log(`📥 ${tabName} saved: ${(text || '').length.toLocaleString()} chars, ${cardCount} cards detected`);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, tab: tabName, chars: (text || '').length, cards: cardCount }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: 'Unknown post type' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function runPendingPhase() {
  const phase = PropertiesService.getScriptProperties().getProperty('PENDING_PHASE');
  log(`runPendingPhase fired with PENDING_PHASE value: ${phase}`);
  if (!phase) {
    log('No pending phase found — exiting');
    return;
  }
  
  // Clear the pending phase immediately
  PropertiesService.getScriptProperties().deleteProperty('PENDING_PHASE');
  
  // Delete this trigger
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runPendingPhase')
    .forEach(t => ScriptApp.deleteTrigger(t));
    log('runPendingPhase: Triggers cleaned up');
  
  // Run the phase
  try {
    log(`Starting ${phase}...`);
    if (phase === '1AB') runPhase1AB();
    else if (phase === 'postReview') runPostReviewPipeline();
    else if (phase === 'promptOnly') generatePromptOnly();
    else if (phase === 'phase2SubBatch') runPendingPhase2SubBatch();
    else if (phase === 'arbitrate') runJudgeArbitration();
    else if (phase === 'finishWrite') finishPhase1ABWrite();
    // 1AB and finishWrite log their own completion (1AB hands the write to a
    // second trigger, so it isn't "complete" when runPhase1AB returns).
    if (phase !== 'postReview' && phase !== 'finishWrite' && phase !== '1AB') {
      log(`✅ Phase ${phase} completed via trigger`);
    }
  } catch(e) {
    log(`❌ Phase ${phase} trigger failed: ${e.message}`);
  }
}

function generatePromptOnly() {
  try {
    log('Generating review prompts...');

    const batch = getReadyBatch();
    log(`Batch: ${batch.map(r => r['L4 Category']).join(', ')}`);
    const categoryRows = getCategoryInfo(batch);

    const { rows: stagingRows } = getSheetData(SHEETS.NEWADDITIONS_STAGING);
    if (stagingRows.length === 0) {
      log('⚠️ NEWADDITIONS_STAGING is empty — run Phase 1AB first');
      return;
    }

    // Build ph1bContent directly from staging rows
    const lines = ['| Attribute Name | Description |'];
    stagingRows.forEach(r => {
      const name = (r['Attribute Name'] || '').toString().trim();
      const desc = (r['Description'] || '').toString().trim();
      if (name && name !== 'Attribute Name') {
        lines.push(`| ${name} | ${desc} |`);
      }
    });
    const ph1bContent = lines.join('\n');
    log(`Built ph1bContent from staging: ${stagingRows.length} rows`);

    // Build namesOnly
    const namesOnly = stagingRows
      .map(r => (r['Attribute Name'] || '').toString().trim())
      .filter(n => n && n !== 'Attribute Name')
      .join('\n');

    // Get category info for prompt context
    const { l2, l3List, l4List, paths } = formatCategoryInfo(batch, categoryRows);

    // Build prompts manually using the same template as generateReviewPrompts
    const promptTemplate = `Act as my practical, no-nonsense taxonomy reviewer for our Akeneo PIM. Review the following proposed attributes against our existing dictionary. [Attached CSV].
CRITICAL RULES FOR MERGING (The Akeneo Reality Check):
You must consider the data types (Simple Select, Multi-Select, Numeric, Boolean, Text).
1. Dropdown Pollution: Do not merge Simple Select or Multi-Select attributes unless the option lists are practically identical. Keep domain-specific lists (like "Defence Spray Type") isolated.
2. Facet Collision: Do not merge attributes if they would cause bizarre cross-shopping filter results on the website (e.g., merging "Compatible With" for smart home alarms and power tools).
3. Metric Free-Pass: Merge Metric (number + unit) and Boolean (Yes/No) attributes aggressively, as long as it doesn't break Rule 4.
4. Meaning Shift: Do not merge if the word's actual physical definition changes depending on the product family.
5. Component-Specific Fitment vs. Whole-Product Identity (The Core/Component Rule):
Do not merge an attribute into a generic global field (like Length, Width, Capacity, or Surface) if it represents a critical component dimension or a specific fitment variable of a larger, complex assembly. These must remain domain-specific to prevent frontend facet collision and preserve exact compatibility mapping.
6. Logistics vs. Engineering Boundary: Never merge functional fitment dimensions into generic fields (Length, Width) if those fields act as ERP/WMS shipping bounding boxes. Keep engineering specs isolated from logistics. You may merge into a generic attribute only if the attribute defines the identity or physical bound of the entire product as a whole (e.g., a luggage strap's length, a bucket's capacity, or a storage box's material).
7. Global Search & Data Type Test: If merging creates a chaotic global filter of unrelated L2 categories (e.g., plumbing and screws under 'Thread Type'), add a prefix to isolate it. Never merge faceted filters (Dropdowns) into descriptive Text fields.
8. Naming Convention: When an attribute is kept domain-specific to protect the shopping experience, always prefix or suffix the name to make its specific scope/domain clear to both the system and the catalog team to prevent crossover with other categories.
9. The Akeneo 1:1 Label Constraint: Akeneo forces a strict 1:1 relationship between the backend attribute and the frontend display label per locale. You must never merge attributes across different product families if it forces a generic, compromised, or incorrect label onto the e-commerce storefront. 
10. The Customer Vocabulary Test: Always validate a proposed merge against the target customer's actual shopping vocabulary. If a tradesperson, hobbyist, or professional expects a highly specific industry term (e.g., searching for a saw's "Arbor" instead of a generic "Mounting Hole", or a jigsaw's "Shoe" instead of a "Base Plate"), you must reject the merge to protect the UX. 
11. Acceptable Controlled Bloat: It is entirely acceptable to create a new, slightly redundant attribute to preserve the frontend shopping experience. When you do this, you must apply Rule 8 (Naming Convention) and prefix the attribute with its L2 overarching category (e.g., "Power Tool") so this "bloat" remains perfectly grouped alphabetically in the backend and does not scatter the taxonomy.

OUTPUT FORMAT:
Do not use tables. Sort the attributes in alphabetical order. Use the following vertical "card" structure for each attribute so it is easy to read on a mobile phone screen:
🔸 [Proposed Attribute Name]
Dictionary Matches: [List of semantic matches from Attribute Dictionary]
Action: [Keep / Merge / Rename] ➡️ [Target PIM Name]
Data Type & Rule Fired: [Assumed Data Type] | [Which of the 4 rules applies?]
Customer View: [Importance: Critical / Important / Technical / Low] | [Customer-Friendliness Assessment]
The Rationale: [A direct, fluff-free explanation of why we are doing this]
Categories: [Typical Categories]

Category context:
L2: ${l2}
L3 Categories: ${l3List}
L4 Categories: ${l4List}`;

    const fullPrompt = `${promptTemplate}\n\nHere are the proposed new attributes to review:\n\n${ph1bContent}`;
    const condensedPrompt = `${promptTemplate}\n\nHere are the proposed new attribute names to review:\n\n${namesOnly}`;

    const prompts = {
      fullPrompt,
      condensedPrompt,
      fullCharCount: fullPrompt.length,
      condensedCharCount: condensedPrompt.length,
      namesOnly,
      ph1bContent
    };

    writePromptsTab(prompts, batch);
    log(`✅ PROMPTS tab written — ${stagingRows.length} attributes from staging`);

  } catch(e) {
    log(`❌ Error: ${e.message}`);
    throw e;
  }
}

function validateReview() {
  const { rows } = getSheetData(SHEETS.REVIEW);
  const errors = [];
  const warnings = [];

  // Filter to actual attribute rows
  const attrRows = rows.filter(r => {
    const n = (r['Attribute Name']||'').toString().trim();
    return n && n !== 'Attribute Name' && !n.startsWith('---');
  }).map(r => {
    // Coerce all values to strings
    const coerced = {};
    Object.keys(r).forEach(k => {
      coerced[k] = r[k] !== null && r[k] !== undefined ? String(r[k]) : '';
    });
    return coerced;
  });

  if (attrRows.length === 0) {
    errors.push('REVIEW tab is empty. Run Phase 1AB first');
    return { valid: false, errors, warnings, rows: attrRows };
  }

  attrRows.forEach((r, i) => {
    const name = (r['Attribute Name'] || '').trim();
    const decision = (r['Your Decision'] || '').trim().toLowerCase();
    const updatedName = (r['Updated Attribute Name'] || '').trim();

    // Every row must have a decision
    if (!decision) {
      errors.push(`Row ${i + 2}: "${name}" has no decision`);
      return;
    }

    // Decision must be one of the valid options
    if (!['keep', 'rename', 'merge', 'delete'].includes(decision)) {
      errors.push(`Row ${i + 2}: "${name}" has invalid decision "${decision}" — must be Keep, Rename, Merge or Delete`);
    }

    // Rename requires an updated name
    if (decision === 'rename' && !updatedName) {
      errors.push(`Row ${i + 2}: "${name}" is marked Rename but has no Updated Attribute Name`);
    }

    // Merge requires an updated name (the target dictionary attribute)
    if (decision === 'merge' && !updatedName) {
      errors.push(`Row ${i + 2}: "${name}" is marked Merge but has no Updated Attribute Name (target dictionary attribute required)`);
    }

    // Rename — warn if updated name is same as original
    if (decision === 'rename' && updatedName.toLowerCase() === name.toLowerCase()) {
      warnings.push(`Row ${i + 2}: "${name}" is marked Rename but Updated Attribute Name is identical — consider changing to Keep`);
    }
  });

  // Check total decisions add up
  const decided = attrRows.filter(r => (r['Your Decision'] || '').trim() !== '').length;
  if (decided < attrRows.length) {
    errors.push(`${attrRows.length - decided} attribute(s) still have no decision — complete all reviews before running`);
  }

  return {
    valid: errors.length === 0,
    errors,
    warnings,
    rows: attrRows,
    counts: {
      keep: attrRows.filter(r => (r['Your Decision']||'').toLowerCase() === 'keep').length,
      rename: attrRows.filter(r => (r['Your Decision']||'').toLowerCase() === 'rename').length,
      merge: attrRows.filter(r => (r['Your Decision']||'').toLowerCase() === 'merge').length,
      delete: attrRows.filter(r => (r['Your Decision']||'').toLowerCase() === 'delete').length,
      total: attrRows.length
    }
  };
}




// ============================================================
// CUSTOM MENU ON SHEETS
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PIM Automation')
    .addItem('▶ Run Phase 1AB', 'runPhase1AB')
    .addItem('▶ Run Post-Review Pipeline', 'runPostReviewPipeline')
    .addItem('▶ Generate Prompt Only', 'generatePromptOnly')
    .addItem('⚖ Run Judge Arbitration', 'runJudgeArbitration')
    .addSeparator()
    .addItem('🗑 Clear DICT_READY + PHASE2_DICT_ENTRIES + MASTER_OUTPUT + PL_ATTRIBUTELIST', 'clearDeliverableTabs')
    .addSeparator()
    .addItem('📋 Authorise All Permissions', 'authoriseAllPermissions')
    .addToUi();
}

function clearDeliverableTabs() {
  const tabs = [
    SHEETS.DICT_READY,
    SHEETS.DICT_READY_DROPDOWNS,
    SHEETS.MASTER_OUTPUT,
    SHEETS.PL_ATTRIBUTELIST,
    SHEETS.PHASE2_DICT_ENTRIES,
    SHEETS.PHASE2_DICT_DROPDOWNS
  ];
  tabs.forEach(tabName => {
    try {
      const sheet = getSheet(tabName);
      sheet.clearContents();
      log(`✅ Cleared: ${tabName}`);
    } catch(e) {
      log(`⚠️ Could not clear ${tabName}: ${e.message}`);
    }
  });
  SpreadsheetApp.getUi().alert('Done — DICT_READY, DICT_READY_DROPDOWNS, MASTER_OUTPUT, PL_ATTRIBUTELIST, PHASE2_DICT_ENTRIES, PHASE2_DICT_DROPDOWNS cleared.');
}

// ============================================================
// SHARED WRITE HELPERS
// ============================================================

function writeTextBlockToSheet(sheet, currentRow, text) {
  const maxChars = 40000;
  if (text.length <= maxChars) {
    sheet.getRange(currentRow, 1).setValue(text);
    sheet.getRange(currentRow, 1).setFontSize(9).setFontWeight('normal').setWrap(true);
    currentRow += 3;
  } else {
    let remaining = text;
    while (remaining.length > 0) {
      const chunk = remaining.substring(0, maxChars);
      remaining = remaining.substring(maxChars);
      sheet.getRange(currentRow, 1).setValue(chunk);
      sheet.getRange(currentRow, 1).setFontSize(9).setFontWeight('normal').setWrap(true);
      currentRow += 2;
    }
    currentRow += 1;
  }
  return currentRow;
}

// ============================================================
// READ BATCH
// ============================================================

function getReadyBatch() {
  const { rows } = getSheetData(SHEETS.BATCH);
  const ready = rows.filter(r => r['Status'] === 'Ready');
  if (ready.length === 0) throw new Error('No categories with Status = Ready found in BATCH tab');
  return ready;
}

function markBatchDone(batch) {
  const sheet = getSheet(SHEETS.BATCH);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const statusCol = headers.indexOf('Status');
  const l4Col = headers.indexOf('L4 Category');
  const doneCats = batch.map(r => r['L4 Category']);
  for (let i = 1; i < data.length; i++) {
    if (doneCats.includes(data[i][l4Col])) {
      sheet.getRange(i + 1, statusCol + 1).setValue('Done');
    }
  }
}

// ============================================================
// FILTER INPUTS FOR BATCH
// ============================================================

function getFilteredProducts(batch) {
  const l4Categories = batch.map(r => r['L4 Category']);
  const { rows } = getSheetData(SHEETS.PRODUCT_LIST);
  const filtered = rows.filter(r => l4Categories.includes(r['Retail_Category4']));
  // Deduplicate by product name per category
  const seen = new Set();
  return filtered.filter(r => {
    const key = `${r['Retail_Category4']}||${r['Product Name']}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function getFilteredHybris(batch) {
  const l4Categories = batch.map(r => r['L4 Category']);
  const { rows } = getSheetData(SHEETS.HYBRIS);
  return rows.filter(r => 
    l4Categories.includes(r['L4 Category']) || 
    r['L4 Category'].startsWith('L3:') || 
    r['L4 Category'].startsWith('L2:')
  );
}

function checkBatchComplexity(batch, hybris, products) {
  const hybrisCount = hybris.filter(r => 
    !r['L4 Category'].startsWith('L3:') && 
    !r['L4 Category'].startsWith('L2:')
  ).length;
  const productCount = products.length;
  const categoryCount = batch.length;

  let risk = 'LOW';
  let reason = '';

  if (hybrisCount > 80) {
    risk = 'HIGH';
    reason = `${hybrisCount} Hybris rows — too large for a single API call. Reduce batch size.`;
  } else if (hybrisCount > 60) {
    risk = 'MEDIUM';
    reason = `${hybrisCount} Hybris rows — Phase 2 will split into sub-batches.`;
  } else {
    reason = `${hybrisCount} Hybris rows — looks good.`;
  }

  log(`🔍 Batch complexity: ${risk} — ${reason}`);
  log(`📦 Products: ${productCount} | Hybris rows: ${hybrisCount} | Categories: ${categoryCount}`);

  return { risk, hybrisCount, productCount, categoryCount, reason };
}

function getDictionary() {
  const { rows: dictRows } = getSheetData(SHEETS.DICTIONARY);
  
  // Append DICT_READY rows so latest approved attributes are included
  // even if not yet copied to master dictionary
  try {
    const { rows: dictReadyRows } = getSheetData(SHEETS.DICT_READY);
    if (dictReadyRows.length > 0) {
      // Map DICT_READY columns to DICTIONARY columns
      const mapped = dictReadyRows.map(r => ({
        'Attribute Name': r['Attribute Name'] || '',
        'Description': r['Description'] || '',
        'Datatype': r['Datatype'] || '',
        'Dropdown Name': r['Dropdown Name'] || '',
        'Unit Of Measure Metric': r['Unit'] || '',
        'Mandate_Source': r['Mandate Source'] || '',
        'Needs Localisation': r['Needs Localisation'] || '',
        'Sample Values': r['Sample Values'] || ''
      }));
      // Deduplicate — DICTIONARY takes priority, skip DICT_READY entries already in DICTIONARY
      const existingNames = new Set(dictRows.map(r => (r['Attribute Name'] || '').toLowerCase()));
      const newRows = mapped.filter(r => !existingNames.has((r['Attribute Name'] || '').toLowerCase()));
      log(`📚 Dictionary: ${dictRows.length} master + ${newRows.length} from DICT_READY (${dictReadyRows.length - newRows.length} already in master)`);
      return [...dictRows, ...newRows];
    }
  } catch(e) {
    log(`⚠️ Could not load DICT_READY for dictionary merge: ${e.message}`);
  }
  
  return dictRows;
}

function getCategoryInfo(batch) {
  const l4Categories = batch.map(r => r['L4 Category']);
  const { rows } = getSheetData(SHEETS.CATEGORY_LIST);
  return rows.filter(r => l4Categories.includes(r['Retail_Category4']));
}

// ============================================================
// TEST CONNECTION
// ============================================================

function testSetup() {
  try {
    // Test sheet access
    const batch = getReadyBatch();
    log(`✅ BATCH tab read successfully. Found ${batch.length} Ready categories:`);
    batch.forEach(r => log(`   - ${r['L3 Category']} > ${r['L4 Category']}`));

    // Test product filter
    const products = getFilteredProducts(batch);
    log(`✅ PRODUCT_LIST filtered: ${products.length} products for this batch`);

    // Test Hybris filter
    const hybris = getFilteredHybris(batch);
    log(`✅ HYBRIS filtered: ${hybris.length} attribute rows for this batch`);

    // Test dictionary
    const dict = getDictionary();
    log(`✅ DICTIONARY loaded: ${dict.length} attributes`);

    // Test Claude API key exists
    if (!CLAUDE_API_KEY) throw new Error('CLAUDE_API_KEY not found in Script Properties');
    log(`✅ CLAUDE_API_KEY found`);

    log(`\n✅ All systems ready. Safe to proceed to Phase 1AB.`);

  } catch(e) {
    log(`❌ Setup error: ${e.message}`);
  }
}

// ============================================================
// FORMAT DATA FOR PROMPT INJECTION
// ============================================================

function formatDictionaryForPrompt(dictRows) {
  const headers = ['Attribute Name', 'Description', 'Sample Values'];
  const lines = [headers.join('\t')];
  dictRows.forEach(r => {
    lines.push(headers.map(h => r[h] || '').join('\t'));
  });
  return lines.join('\n');
}

function formatProductListForPrompt(productRows) {
  const lines = ['L4 Category\tProduct Name'];
  productRows.forEach(r => {
    lines.push(`${r['Retail_Category4'] || r['L4 Category']}\t${r['Product Name']}`);
  });
  return lines.join('\n');
}

function formatHybrisForPrompt(hybrisRows) {
  const lines = ['L4 Category\tHybris Swedish\tHybris English\tDictionary Attributes'];
  hybrisRows.forEach(r => {
    lines.push([
      r['L4 Category'],
      r['Hybris Swedish'],
      r['Hybris English'],
      r['Dictionary Attributes']
    ].join('\t'));
  });
  return lines.join('\n');
}

function formatCategoryInfo(batch, categoryRows) {
  // Build L4 list
  const l4List = [...new Set(batch.map(r => r['L4 Category']))].join('\n');
  
  // Build L3 list
  const l3List = [...new Set(batch.map(r => r['L3 Category']))].join('\n');
  
  // Build L2
  const l2 = categoryRows.length > 0 ? 
    (categoryRows[0]['Retail_Category2'] || batch[0]['L3 Category']) : 
    'Unknown';

  // Build category paths
  const paths = batch.map(r => {
    const catRow = categoryRows.find(c => c['Retail_Category4'] === r['L4 Category']);
    if (catRow && catRow['L2_L4_Path']) return catRow['L2_L4_Path'];
    return `${l2} -> ${r['L3 Category']} -> ${r['L4 Category']}`;
  }).join('\n');

  return { l2, l3List, l4List, paths };
}

// ============================================================
// ASSEMBLE PHASE 1AB PROMPT
// ============================================================

function assemblePhase1ABPrompt(batch, products, hybris, dictionary, categoryRows) {
  const { l2, l3List, l4List, paths } = formatCategoryInfo(batch, categoryRows);
  const productsFormatted = formatProductListForPrompt(products);
  const hybrisFormatted = formatHybrisForPrompt(hybris);

  const prompt = `Phase 1AB: Attribute Generation
Category L2: ${l2}

1. Inputs
[ATTRIBUTE_DICTIONARY]: (Attached CSV) Authority for semantics and naming.
[L3_LIST] / [L4_LIST]: Sub-category hierarchy.
[CATEGORY_PATH]: Hierarchy mapping.
[PRODUCT_LIST]: (Attached CSV) Product names (English/Swedish).
[HYBRIS]: Legacy mapping for modernization.

2. Global Execution Rules
Follow steps strictly in order. Do not skip or merge.
Once an attribute is defined, do not duplicate or rename.
Prefer reuse and semantic matching from [ATTRIBUTE_DICTIONARY] over creating new names.
Swedish names are for context only; final output must be in English.
Use "-na-" as a marker for an empty cell unless stated otherwise.
Exclude these attributes unless they exist in [HYBRIS] for this category: Brand, Colour, Item Weight, Dimensions Breadth, Dimensions Height, Dimensions Depth, Warranty Period. If they exist in Hybris they must be mapped and included. If they do not exist in Hybris do not generate them even if they seem relevant — they exist on the PDP by default and are out of scope for this project.

3. Machine-Readable Formatting Rules (CRITICAL)
To ensure system compatibility, you must wrap each step's output in the following tags:
[STEP_X] ... [/STEP_X]
Tables must use the | pipe format.
Character Constraint: No single table cell may exceed 1,000 characters. If a description is longer, summarize it.
No Preamble: Do not include conversational filler (e.g., "Here is Step 1..."). Start and end only with the tags.

4. Step-by-Step Instructions
[STEP_0]
Create a subset of [PRODUCT_LIST] for categories in [L4_LIST], called [PRODUCT_LIST_L4].
Output: A table showing each L4 category and the number of products within it.
No need to print [PRODUCT_LIST_L4]
[/STEP_0]


[STEP_1]
For the given product list in [PRODUCT_LIST_L4] associated with each category in [L4_LIST], explain the general functions of the products in each category and who it is for (leaning consumer or leaning technical). Based on who it is for, what are the attributes they (both general customers and technical customers) most often look for when purchasing products within this category that helps them make an informed purchase decision without having to call customer support. 
Output: A table called [PL_ATTRIBUTELIST_V1] showing Category, General product types/functions, Who it is for, Attribute list
No need to print [PL_ATTRIBUTELIST_V1]
[/STEP_1]


[STEP_2]
Use the attributes in [HYBRIS] for each category, find the most appropriate match in [ATTRIBUTE_DICTIONARY] if there are no strong matches in the Dictionary Attributes column. Use these matches to add to the attribute set in [PL_ATTRIBUTELIST_V1]. Dedupe attributes in this list prioritizing [ATTRIBUTE_DICTIONARY] names over newly generated names. For any unmatched attributes, use a modern customer-friendly name to add to the [PL_ATTRIBUTELIST_V1] if it doesn't already exist. For each category then list this updated attribute set for each category in a single table called [PL_ATTRIBUTELIST_V2] : category name, number of products within the category, general product types, who it is for, updated attribute set. 
Output: A single table called [PL_ATTRIBUTELIST_V2] containing: Category, Product Count, Product Types, Target Customer, What is important for each target customer, Updated Attribute Set.
Print [PL_ATTRIBUTELIST_V2]
[/STEP_2]

[STEP_2a]
Using the attribute list in [PL_ATTRIBUTELIST_V2] create an internal logic table called [NEW_ATTRIBUTES_PRODUCTLIST] with format: L3 Category|L4 category | attribute name (in Title Case) | description of attribute (what it is for) | customer importance  | Exists in Hybris (Yes/No/Renamed (old -> new name)) | Attribute Dictionary Match (Yes/No)  | Related matches/attributes  |. List attributes in alphabetical order. Call this table [NEW_ATTRIBUTES_PRODUCTLIST]. Exclude these attributes: Brand, Colour, Dimensions, Item Weight, Material (unless they are critical purchase decision drivers in this category). 
Customer importance: critical would be when they're the absolute must-have information for this product, important would be when they're important to know for most customers, good-to-have but most regular customers don't need it, low importance is when it is irrelevant for most customrs expect the super enthusiast crowd, irrelevant when the attribute is irrelevant for this category
No need to print [NEW_ATTRIBUTES_PRODUCTLIST] table
[/STEP_2a]


[STEP_3]
For attributes under Hybris, give me a table [HYBRIS_PASS1] in this format: L4 category | Hybris English attribute name (in Title Case)| Dictionary Attribute | description of attribute (what it is for) | number of products that need this attribute in the category  | Is this attribute addressed by [NEW_ATTRIBUTES_PRODUCTLIST] or [ATTRIBUTE_DICTIONARY] and name the matched attribute | Should this be a standalone attribute, be consolidated into another attribute or deleted due to irrelevance for the product set for this category | Reasoning
If the L4 category in [HYBRIS] contains 'L3: ' or 'L2: ' use these attributes within their L3 or L4 categories only if appropriate for the individual L4 category.
Print [HYBRIS_PASS1]
[/STEP_3]

[STEP_4]
Step 4: List all attributes from [NEW_ATTRIBUTES_PRODUCTLIST] that did have a match in [ATTRIBUTE_DICTIONARY] with this structure: Attribute name, Sample Values (at least 5 unless Boolean), Description. Call this table [DICTIONARY_SUBSET].
Print [DICTIONARY_SUBSET]
[/STEP_4]

[STEP_5]
For attributes in [NEW_ATTRIBUTES_PRODUCTLIST] that do not have a direct match in [ATTRIBUTE_DICTIONARY], give me a table called [ADD_TO_DICT_PL] in this structure:
Attribute Name | Description | Datatype| Dropdown Name | Unit of measure | Mandate source | Language Localization | Sample values |  Matches from [ATTRIBUTE_DICTIONARY] | Customer Friendly Attribute Name  | Categories . 

Column definitions and rule for the [ADD_TO_DICT_PL]:
Description: a description of the attribute that helps category managers understand what it is
Datatype: one of these types listed: Numeric, Text, Boolean, Dropdown, or Dropdown (Multi)
Dropdown Name: In Title Case. Only populate if the datatype is Dropdown or Dropdown (Multi)
Unit of measure: Most commonly used in the Nordics
Mandate source: Product Feature or Compliance)
Needs language localization: Yes if datatype is of the dropdown type and values have english words, leave blank if not
Sample values: at least 5 unless Boolean
Match in [ATTRIBUTE_DICTIONARY]: Attributes from [ATTRIBUTE_DICTIONARY] that are semantically similar to the Attribute Name
Customer friendly attribute name: What does this attribute typically appear as on retail product pages
Categories: What categories this attribute applies to in the current analysis. 
No need to Print [ADD_TO_DICT_PL]
[/STEP_5]

[STEP_6]
Please give me in a table format, dropdown names from [ADD_TO_DICT_PL] and an exhaustive comma separated list of values for each dropdown (both single and multi) in the list. Call this table [ADD_TO_DICT_PL_PICKLIST].
No need to Print [ADD_TO_DICT_PL_PICKLIST]
[/STEP_6]

[STEP_7]
For attributes in [HYBRIS_PASS1] that are not addressed by [NEW_ATTRIBUTES_PRODUCTLIST], check for attribute matches in [ATTRIBUTE_DICTIONARY]. give me a table [ADD_TO_DICT_HYBRIS] with the same structure and rules as [ADD_TO_DICT_PL]: 
| Attribute Name | Description | Datatype| Dropdown Name | Unit of measure | Mandate source | Language Localization | Sample values |  Matches from [ATTRIBUTE_DICTIONARY] | Customer Friendly Attribute Name  | Categories |. 
 
No need to print [ADD_TO_DICT_HYBRIS].
[/STEP_7]

[STEP_8]
Step 8: Please give me in a table format, dropdown names from [ADD_TO_DICT_HYBRIS] and an exhaustive comma separated list of values for each dropdown (both single and multi) in the list. Call this table [ADD_TO_DICT_HYBRIS_PICKLIST].
No need to Print [ADD_TO_DICT_HYBRIS_PICKLIST].
[/STEP_8]

[STEP_9]
Combine [ADD_TO_DICT_PL] and [ADD_TO_DICT_HYBRIS] into [ADD_ME_TO_THE_DICTIONARY_PH1B]. Combine their dropdown value tables [ADD_TO_DICT_PL_PICKLIST] and [ADD_TO_DICT_HYBRIS_PICKLIST] into one and call it [DROPDOWN_PH1B] and print it in alphabetical order.
Print [ADD_ME_TO_THE_DICTIONARY_PH1B], with an additional column for source ([ADD_TO_DICT_PL] or [ADD_TO_DICT_HYBRIS]).
Print [DROPDOWN_PH1B] in alphabetical order.
[/STEP_9]

[STEP_10]
Consolidated Attribute Review Table

Act as my practical, no-nonsense taxonomy reviewer for our Akeneo PIM. Review the following proposed attributes against our existing dictionary.

For each attribute in [ADD_ME_TO_THE_DICTIONARY_PH1B], produce a single consolidated review table called [ATTRIBUTE_REVIEW] that combines the naming analysis with the Hybris context and attribute metadata. This table is the primary output for human review.

Use this exact table structure:
| Row # | Attribute Name | Source | Hybris English Original | Hybris Reasoning | Existing Dictionary Match | Recommendation | Recommended Name | Rationale | Description | Sample Values | Datatype | Applicable Categories | Sample Products | Your Decision | Updated Attribute Name | Notes |

Column definitions:
Row #: Sequential row number starting from 1.
Attribute Name: The attribute name from [ADD_ME_TO_THE_DICTIONARY_PH1B].
Source: Either ADD_TO_DICT_PL (came from product analysis) or ADD_TO_DICT_HYBRIS (came from unmatched Hybris attribute).
Hybris English Original: The original Hybris English attribute name from [HYBRIS_PASS1] if Source is ADD_TO_DICT_HYBRIS. If Source is ADD_TO_DICT_PL write "New".
Hybris Reasoning: The keep/consolidate/delete reasoning from [HYBRIS_PASS1] for this attribute if Source is ADD_TO_DICT_HYBRIS. If Source is ADD_TO_DICT_PL write "New".
Existing Dictionary Match: Any semantically similar attributes already in [ATTRIBUTE_DICTIONARY]. List all matches. If none write "None".
Recommendation: One of — Keep / Rename / Merge / Delete. Evaluate from both customer language and industry terminology perspectives along with input from column 'Customer Friendly Attribute Name' in [ADD_ME_TO_THE_DICTIONARY_PH1B] to find the best representation of this attribute.
Recommended Name: If Recommendation is Keep — use the current Attribute Name. If Recommendation is Rename or Merge — suggest the best name, preferring an existing [ATTRIBUTE_DICTIONARY] name if a strong match exists. If Recommendation is Delete — write "N/A".
Rationale: One concise sentence explaining the recommendation. Explicitly call out if something should not be merged even if similar, or if a business may choose to keep both despite overlap.
Description: The attribute description from [ADD_ME_TO_THE_DICTIONARY_PH1B].
Sample Values: At least 5 sample values unless Boolean. From [ADD_ME_TO_THE_DICTIONARY_PH1B].
Datatype: From [ADD_ME_TO_THE_DICTIONARY_PH1B].
Applicable Categories: From [ADD_ME_TO_THE_DICTIONARY_PH1B].
Sample Products: 2-3 example product names from [PRODUCT_LIST] where this attribute applies.
Your Decision: Leave blank — for human reviewer to fill in (Keep / Rename / Merge / Delete). Merge means this attribute already exists in [ATTRIBUTE_DICTIONARY] under the name in Updated Attribute Name.
Updated Attribute Name: Leave blank — for human reviewer to fill in. For Merge decisions, write the existing [ATTRIBUTE_DICTIONARY] attribute name it merges into.
Updated Dropdown Name: Leave blank — for human reviewer to fill in. Only needed if Datatype is Dropdown or Dropdown (Multi) and the dropdown name changes. If blank, defaults to Updated Attribute Name.
Notes: Leave blank — for human reviewer to fill in.

Rules:.
Do not invent new attributes.
Do not remove attributes silently.
When in doubt prefer flagging for review over making a definitive call.
Treat [ATTRIBUTE_DICTIONARY] as authoritative but not infallible.
No single cell may exceed 1000 characters.
You must consider the data types:
1. Dropdown Pollution: Do not merge Simple Select or Multi-Select attributes unless the option lists are practically identical. Keep domain-specific lists (like "Defence Spray Type") isolated.
2. Facet Collision: Do not merge attributes if they would cause bizarre cross-shopping filter results on the website (e.g., merging "Compatible With" for smart home alarms and power tools).
3. Metric Free-Pass: Merge Metric (number + unit) and Boolean (Yes/No) attributes aggressively, as long as it doesn't break Rule 4.
4. Meaning Shift: Do not merge if the word's actual physical definition changes depending on the product family.
5. Component-Specific Fitment vs. Whole-Product Identity (The Core/Component Rule):
Do not merge an attribute into a generic global field (like Length, Width, Capacity, or Surface) if it represents a critical component dimension or a specific fitment variable of a larger, complex assembly. These must remain domain-specific to prevent frontend facet collision and preserve exact compatibility mapping.
6. Logistics vs. Engineering Boundary: Never merge functional fitment dimensions into generic fields (Length, Width) if those fields act as ERP/WMS shipping bounding boxes. Keep engineering specs isolated from logistics. You may merge into a generic attribute only if the attribute defines the identity or physical bound of the entire product as a whole (e.g., a luggage strap's length, a bucket's capacity, or a storage box's material).
7. Global Search & Data Type Test: If merging creates a chaotic global filter of unrelated L2 categories (e.g., plumbing and screws under 'Thread Type'), add a prefix to isolate it. Never merge faceted filters (Dropdowns) into descriptive Text fields.
8. Naming Convention: When an attribute is kept domain-specific to protect the shopping experience, always prefix or suffix the name to make its specific scope/domain clear to both the system and the catalog team to prevent crossover with other categories.
Print [ATTRIBUTE_REVIEW] in alphabetical order.

[/STEP_10]

[L4_CATLIST]
${l4List}

[L3_CATLIST]
${l3List}

[CATEGORY_PATH]
${paths}

[HYBRIS]
${hybrisFormatted}

[PRODUCT_LIST]
${productsFormatted}`;

  return prompt;
}

// ============================================================
// EXPORT DICTIONARY TO GOOGLE DRIVE
// ============================================================

function exportDictionaryToDrive() {
  try {
    //const DRIVE_FOLDER_NAME = 'Project CO_PD'; // Change to your preferred folder name
    const FILE_NAME = 'CO_PD_Dictionary.csv';

    // Get dictionary data
    const { headers, rows } = getSheetData(SHEETS.DICTIONARY);
    
    // Build CSV content — only 3 columns needed for Gemini
    const exportHeaders = ['Attribute Name', 'Description', 'Sample Values'];
    const csvLines = [exportHeaders.join(',')];
    rows.forEach(r => {
      const line = exportHeaders.map(h => {
        const val = (r[h] || '').toString().replace(/"/g, '""');
        return `"${val}"`;
      });
      csvLines.push(line.join(','));
    });
    const csvContent = csvLines.join('\n');
    const csvBlob = Utilities.newBlob(csvContent, 'text/csv', FILE_NAME);

// Save to root of My Drive — overwrite if exists
    const fileSearch = DriveApp.getFilesByName(FILE_NAME);
    if (fileSearch.hasNext()) {
      const existingFile = fileSearch.next();
      existingFile.setContent(csvContent);
      log(`✅ Dictionary CSV updated in Drive root: ${FILE_NAME} (${rows.length} attributes)`);
    } else {
      DriveApp.createFile(csvBlob);
      log(`✅ Dictionary CSV created in Drive root: ${FILE_NAME} (${rows.length} attributes)`);
    }

  } catch(e) {
    log(`⚠️ Could not export dictionary to Drive: ${e.message}`);
    // Non-fatal — don't throw, just warn
  }
}

// ============================================================
// CALL CLAUDE API
// ============================================================

function callClaudeAPI(prompt, documents) {
  const url = 'https://api.anthropic.com/v1/messages';
  
  const content = [];
  
  // Add documents as labelled text blocks
  if (documents && documents.length > 0) {
    documents.forEach(doc => {
      content.push({
        type: 'text',
        text: `[${doc.title}]\n${doc.content}\n[/${doc.title}]`
      });
    });
  }
  
  // Add prompt text
  content.push({
    type: 'text',
    text: prompt
  });

  const payload = {
    model: 'claude-sonnet-4-6',
    max_tokens: 50000,
    messages: [
      {
        role: 'user',
        content: content
      }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': CLAUDE_API_KEY,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    throw new Error(`Claude API error ${responseCode}: ${responseText}`);
  }

  const parsed = JSON.parse(responseText);
  return parsed.content[0].text;
}

// ============================================================
// MARKDOWN TABLE PARSER
// ============================================================

function parseMarkdownTable(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(l => l.startsWith('|'));
  if (lines.length < 2) return null;
  
  // Parse headers
  const headers = lines[0].split('|').map(c => c.trim()).filter(Boolean);
  
  // Skip separator line (|---|---|)
  const dataLines = lines.slice(1).filter(l => !l.match(/^\|[-\s|]+\|$/));
  
  // Parse rows
  const rows = dataLines.map(line => {
    const cols = line.split('|').map(c => c.trim()).filter(Boolean);
    const row = {};
    headers.forEach((h, i) => {
      row[h] = cols[i] || '';
    });
    return row;
  }).filter(row => Object.values(row).some(v => v !== ''));
  
  return { headers, rows };
}

function extractSection(output, startMarker, endMarkers) {
  const startIndex = output.indexOf(startMarker);
  if (startIndex === -1) return null;
  
  let endIndex = output.length;
  endMarkers.forEach(marker => {
    const idx = output.indexOf(marker, startIndex + startMarker.length);
    if (idx !== -1 && idx < endIndex) endIndex = idx;
  });
  
  return output.substring(startIndex + startMarker.length, endIndex).trim();
}

function reprocessPhase2Output() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const debugSheet = ss.getSheetByName('RAW_OUTPUT_DEBUG');
    const data = debugSheet.getDataRange().getValues();
    const output = data.map(r => r[0]).filter(Boolean).join('');
    log(`Output: ${output.length} chars`);

    const step1Start = output.indexOf('[STEP_1]');
    const step1End = output.indexOf('[/STEP_1]');
    const step1Content = output.substring(
      step1Start + '[STEP_1]'.length,
      step1End > -1 ? step1End : output.length
    );
    log(`STEP_1 content length: ${step1Content.length}`);

    const lines = step1Content.split('\n');
    const tableStart = lines.findIndex(l => l.trim().startsWith('|'));
    log(`Table starts at line: ${tableStart}`);

    if (tableStart === -1) { log('❌ No table found'); return; }

    // Find where table ends — skip blank lines, --- separators, and **bold** headers
    let tableEnd = lines.length;
    for (let i = tableStart + 2; i < lines.length; i++) {
      const trimmed = lines[i].trim();
      if (trimmed.length === 0 || trimmed.startsWith('---') || trimmed.startsWith('**')) continue;
      if (!trimmed.startsWith('|')) {
        tableEnd = i;
        break;
      }
    }

    log(`Table ends at line: ${tableEnd}, total table lines: ${tableEnd - tableStart}`);
    const tableContent = lines.slice(tableStart, tableEnd).join('\n');
    log(`First line: ${lines[tableStart].substring(0, 100)}`);

    const parsed = parseMarkdownTable(tableContent);
    if (!parsed || !parsed.rows || parsed.rows.length === 0) {
      log('❌ Could not parse table');
      return;
    }

    log(`Parsed ${parsed.rows.length} rows`);

    const masterSheet = getSheet(SHEETS.MASTER_OUTPUT);
    if (masterSheet.getLastRow() === 0) {
      masterSheet.getRange(1, 1, 1, parsed.headers.length).setValues([parsed.headers]);
    }
    const rows = parsed.rows.map(r => parsed.headers.map(h => r[h] || ''));
    masterSheet.getRange(masterSheet.getLastRow() + 1, 1, rows.length, parsed.headers.length).setValues(rows);
    log(`✅ Written ${parsed.rows.length} rows to MASTER_OUTPUT`);

  } catch(e) {
    log(`❌ Error: ${e.message}`);
  }
}

function diagnoseOutput() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const debugSheet = ss.getSheetByName('RAW_OUTPUT_DEBUG');
  const output = debugSheet.getRange(1, 1).getValue();
  
  log(`Total output length: ${output.length}`);
  
  // Check each step tag content length
  const steps = ['STEP_0', 'STEP_1', 'STEP_2', 'STEP_2a', 'STEP_3', 'STEP_4', 
                 'STEP_5', 'STEP_6', 'STEP_7', 'STEP_8', 'STEP_9', 'STEP_10'];
  steps.forEach(step => {
    const content = extractStepContent(output, step);
    if (content) {
      log(`${step}: ${content.length} chars`);
      if (content.length > 45000) {
        log(`⚠️ ${step} exceeds 45k — will cause cell limit error`);
      }
    } else {
      log(`${step}: not found`);
    }
  });
}

// ============================================================
// WRITE PARSED TABLES TO PHASE_1AB_WORKING
// ============================================================

function writePhase1ABOutput(output, batch) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Nuclear reset — delete and recreate tab entirely
  const existingSheet = ss.getSheetByName(SHEETS.PHASE_1AB_WORKING);
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }
  const sheet = ss.insertSheet(SHEETS.PHASE_1AB_WORKING);
  SpreadsheetApp.flush();
  log(`Phase 1AB: Tab recreated. Max rows: ${sheet.getMaxRows()}`);

  const rawBatchInfo = batch.map(r => `${r['L3 Category']} > ${r['L4 Category']}`).join(', ');
  const batchInfo = rawBatchInfo.substring(0, 1000);
  let currentRow = 1;

  log(`\n Phase 1AB: Batch Info ${batchInfo}`);
  SpreadsheetApp.flush();
  
function writeSectionHeader(title) {
    const safeTitle = title.length > 1000 ? title.substring(0, 1000) : title;
    try {
  if (title.toString().length > 50000) {
    console.error("FOUND IT: Variable length is " + title.toString().length);
  }
} catch(e) {}
    log(`\n Phase 1AB: Writing section header. ${safeTitle}`);
    sheet.getRange(currentRow, 1).setValue(safeTitle);
    sheet.getRange(currentRow, 1).setFontWeight('bold').setFontSize(10);
    currentRow += 2;
  }
  
function writeTextBlock(text) {
   try {
  if (text.toString().length > 50000) {
    console.error("FOUND IT: Variable length is " + text.toString().length);
  }
} catch(e) {}
    const safeText = text.length > 49000 ? text.substring(0, 49000) : text;
    log(`\n Phase 1AB: Writing TextBlock. ${safeText}`);
    currentRow = writeTextBlockToSheet(sheet, currentRow, safeText);
  }
  
function writeTable(headers, rows, tabName) {
  
  log(`\n Phase 1AB: Writing Table. ${tabName}`);
    if (!headers || !rows || headers.length === 0) {
      sheet.getRange(currentRow, 1).setValue('⚠️ No data found for this section');
      sheet.getRange(currentRow, 1).setFontSize(9).setFontWeight('normal');
      currentRow += 3;
      return;
    }

    // OPTIMIZATION: Ensure headers themselves aren't massive
  const safeHeaders = headers.map(h => String(h).substring(0, 1000));

    // Header row — bold
    sheet.getRange(currentRow, 1, 1, safeHeaders.length).setValues([safeHeaders]);
    sheet.getRange(currentRow, 1, 1, safeHeaders.length)
      .setFontWeight('bold')
      .setFontSize(9);
    currentRow++;
    // Data rows — normal weight
    if (rows && rows.length > 0) {
      const maxCellChars = 45000;
      const data = rows.map(r => safeHeaders.map(h => {
        const val = (r[h] || '').toString();
        return val.length > maxCellChars ? val.substring(0, maxCellChars) + '...[truncated]' : val;
      }));
      // Write in chunks of 50 rows to avoid timeout on large tables
      const chunkSize = 50;
      for (let i = 0; i < data.length; i += chunkSize) {
        const chunk = data.slice(i, i + chunkSize);
        sheet.getRange(currentRow + i, 1, chunk.length, safeHeaders.length).setValues(chunk);
      }
      sheet.getRange(currentRow, 1, data.length, safeHeaders.length)
        .setFontWeight('normal')
        .setFontSize(9);
      currentRow += data.length;
    }
    if (tabName) {
      try {
        const destSheet = getSheet(tabName);
        destSheet.clearContents();
        const allData = [safeHeaders, ...rows.map(r => safeHeaders.map(h => r[h] || ''))];
        destSheet.getRange(1, 1, allData.length, safeHeaders.length).setValues(allData);
        safeHeaders.forEach((_, i) => destSheet.autoResizeColumn(i + 1));
      } catch(e) {
        log(`Could not write to tab ${tabName}: ${e.message}`);
      }
    }
    currentRow += 3;
  }

  // ── HEADER ──────────────────────────────────────────────
  writeSectionHeader(`PHASE 1AB OUTPUT — ${batchInfo} — Generated: ${new Date().toLocaleString()}`);

  // ── STEP 0 ───────────────────────────────────────────────
  writeSectionHeader('STEP 0 — Product Count per Category');
  const step0Content = extractStepContent(output, 'STEP_0');
  if (step0Content) {
    const parsed = parseMarkdownTable(step0Content);
    if (parsed && parsed.rows.length > 0) writeTable(parsed.headers, parsed.rows, null);
    else writeTextBlock(step0Content.substring(0, 2000));
  }

  writeSectionHeader('STEP 2 — PL_ATTRIBUTELIST_V2 (Attribute List per Category)');
  const step2Content = extractStepContent(output, 'STEP_2');
  if (step2Content) {
    const parsed = parseMarkdownTable(step2Content);
    if (parsed && parsed.rows.length > 0) {
      log(`Step 2: writing ${parsed.rows.length} rows to PL_ATTRIBUTELIST`);
      writeTable(parsed.headers, parsed.rows, SHEETS.PL_ATTRIBUTELIST);
    } else {
      log(`Step 2: no parseable table found in step 2 content`); 
      writeTextBlock(step2Content.substring(0, 2000));
    }
 } else {
    log(`Step 2: no content extracted from STEP_2 tag`);
  }

  // ── STEP 3: HYBRIS_PASS1 ─────────────────────────────────
  writeSectionHeader('STEP 3 — HYBRIS_PASS1 (Legacy Attribute Mapping)');
  const step3Content = extractStepContent(output, 'STEP_3');
  if (step3Content) {
    const parsed = parseMarkdownTable(step3Content);
    if (parsed && parsed.rows.length > 0) writeTable(parsed.headers, parsed.rows, null);
    else writeTextBlock(step3Content.substring(0, 2000));
  }

  // ── STEP 4: DICTIONARY_SUBSET ────────────────────────────
  writeSectionHeader('STEP 4 — DICTIONARY_SUBSET (Matched Existing Attributes)');
  const step4Content = extractStepContent(output, 'STEP_4');
  if (step4Content) {
    const parsed = parseMarkdownTable(step4Content);
    if (parsed && parsed.rows.length > 0) writeTable(parsed.headers, parsed.rows, SHEETS.DICTIONARY_SUBSET);
    else writeTextBlock(step4Content.substring(0, 2000));
  }

  // ── STEP 5: ADD_TO_DICT_PL ──────────────────────────────
  writeSectionHeader('STEP 5 — ADD_TO_DICT_PL (New Attributes for Dictionary)');
  const step5Content = extractStepContent(output, 'STEP_5');
  if (step5Content) {
    const parsed = parseMarkdownTable(step5Content);
    if (parsed && parsed.rows.length > 0) writeTable(parsed.headers, parsed.rows, SHEETS.NEWADDITIONS_STAGING);
    else writeTextBlock(step5Content.substring(0, 2000));
  }

  // ── STEP 6: Dropdowns from ADD_TO_DICT_PL ───────────────
  writeSectionHeader('STEP 6 — Dropdown Values (from ADD_TO_DICT_PL)');
  const step6Content = extractStepContent(output, 'STEP_6');
  if (step6Content) {
    const parsed = parseMarkdownTable(step6Content);
    if (parsed && parsed.rows.length > 0) writeTable(parsed.headers, parsed.rows, null);
    else writeTextBlock(step6Content.substring(0, 2000));
  }

  // ── STEP 7: ADD_TO_DICT_HYBRIS ───────────────────────────
  writeSectionHeader('STEP 7 — ADD_TO_DICT_HYBRIS (Unmatched Hybris Attributes)');
  const step7Content = extractStepContent(output, 'STEP_7');
  if (step7Content) {
    const parsed = parseMarkdownTable(step7Content);
    if (parsed && parsed.rows.length > 0) writeTable(parsed.headers, parsed.rows, null);
    else writeTextBlock(step7Content.substring(0, 2000));
  }

  // ── STEP 8: Dropdowns from HYBRIS ───────────────────────
  writeSectionHeader('STEP 8 — Dropdown Values (from ADD_TO_DICT_HYBRIS)');
  try {
    const step8Content = extractStepContent(output, 'STEP_8');
    if (step8Content) {
      const parsed = parseMarkdownTable(step8Content);
      if (parsed && parsed.rows.length > 0) writeTable(parsed.headers, parsed.rows, null);
      else writeTextBlock(step8Content.substring(0, 2000));
    }
  } catch(e) {
    log(`⚠️ STEP 8 write skipped: ${e.message}`);
    writeTextBlock(`⚠️ STEP 8 could not be written: ${e.message}`);
  }

  // ── STEP 9: PH1B COMBINED + DROPDOWNS ───────────────────
  writeSectionHeader('STEP 9 — ADD_ME_TO_THE_DICTIONARY_PH1B (Combined New Attributes)');
  const step9Content = extractStepContent(output, 'STEP_9');
  if (step9Content) {
    // Extract PH1B table
    const ph1bSection = extractSection(step9Content, '[ADD_ME_TO_THE_DICTIONARY_PH1B]',
      ['[DROPDOWN_PH1B]']);
    if (ph1bSection) {
      const parsed = parseMarkdownTable(ph1bSection);
      if (parsed && parsed.rows.length > 0) {
        writeTable(parsed.headers, parsed.rows, SHEETS.NEWADDITIONS_STAGING);
      } else {
        writeTextBlock('✅ No new attributes to add — all attributes matched existing dictionary entries.');
      }
    }

    currentRow += 2;

    // Extract DROPDOWN_PH1B table
    writeSectionHeader('STEP 9 — DROPDOWN_PH1B (All Dropdown Values)');
    const dropdownSection = extractSection(step9Content, '[DROPDOWN_PH1B]', []);
    if (dropdownSection) {
      const parsed = parseMarkdownTable(dropdownSection);
      if (parsed && parsed.rows.length > 0) {
        writeTable(parsed.headers, parsed.rows, null);
      } else {
        writeTextBlock('✅ No dropdown values to add for this batch.');
      }
    }
  }

  // ── STEP 10 → appended to NEWADDITIONS_STAGING ──────────
  // ── STEP 10 → written to REVIEW tab ─────────────────────
  writeSectionHeader('STEP 10 — ATTRIBUTE_REVIEW (Consolidated Review Table)');
  const step10Content = extractStepContent(output, 'STEP_10');
  if (step10Content) {
    const reviewSection = extractSection(step10Content, '[ATTRIBUTE_REVIEW]', []);
    const parsed = reviewSection ? parseMarkdownTable(reviewSection) : parseMarkdownTable(step10Content);
    if (parsed && parsed.rows.length > 0) {
      writeTable(parsed.headers, parsed.rows, SHEETS.REVIEW);
      log(`✅ ATTRIBUTE_REVIEW written to REVIEW tab — ${parsed.rows.length} attributes`);
    } else {
      writeTextBlock('✅ No new attributes to review for this batch.');
    }
  }

  for (let i = 1; i <= 10; i++) {
    sheet.autoResizeColumn(i);
  }
  sheet.setColumnWidth(1, 200);

  log(`✅ PHASE_1AB_WORKING written with structured sections at row ${currentRow}`);
}

// ============================================================
// GENERATE NOTEBOOKLM + GEMINI PROMPTS
// ============================================================

function generateReviewPrompts(output, batch, categoryRows) {
  let ph1bContent = '';
  let rawContent = '';
  const { l2, l3List, l4List } = formatCategoryInfo(batch, categoryRows);

  // Build recently approved attributes section
  let recentlyApprovedSection = '';
  try {
    const { rows: dictReadyRows } = getSheetData(SHEETS.DICT_READY);
    if (dictReadyRows.length > 0) {
      const lines = ['\n\n━━━ RECENTLY APPROVED ATTRIBUTES IN DICTIONARY ━━━\n'];
      lines.push('Attribute Name\tDescription\tSample Values');
      dictReadyRows.forEach(r => {
        const name = (r['Attribute Name'] || '').toString().trim();
        const desc = (r['Description'] || '').toString().trim();
        const samples = (r['Sample Values'] || '').toString().trim();
        if (name) lines.push(`${name}\t${desc}\t${samples}`);
      });
      recentlyApprovedSection = lines.join('\n');
    }
  } catch(e) {
    // silent fail
  }

  // Use tag-based extraction first
  const step9Content = extractStepContent(output, 'STEP_9');
  if (step9Content) {
    const ph1bSection = extractSection(step9Content, '[ADD_ME_TO_THE_DICTIONARY_PH1B]', ['[DROPDOWN_PH1B]']);
    if (ph1bSection && ph1bSection.trim().length > 100) {
      rawContent = ph1bSection.trim();
    }
  }

  // Fall back to individual steps if PH1B is empty
  if (!rawContent) {
    const step5Content = extractStepContent(output, 'STEP_5') || '';
    const step7Content = extractStepContent(output, 'STEP_7') || '';
    rawContent = [step5Content, step7Content].filter(Boolean).join('\n\n');
  }

  if (rawContent) {
    const lines = rawContent.split('\n').filter(l => l.trim());
    const filtered = lines.map(l => {
      if (!l.includes('|')) return l;
      if (l.match(/^\|[-\s|]+\|$/)) return null;
      const cols = l.split('|').map(c => c.trim()).filter(Boolean);
      if (cols[0].toLowerCase().includes('attribute name')) {
        return '| Attribute Name | Description |';
      }
      return `| ${cols[0]} | ${cols[1] || ''} |`;
    }).filter(Boolean);
    ph1bContent = filtered.join('\n');
  } else {
    ph1bContent = '⚠️ No new attributes to review for this batch — all Hybris attributes matched existing dictionary entries. Check PHASE_1AB_WORKING tab to confirm.';
  }

  // Extract just attribute names for condensed version
  const lines = ph1bContent.split('\n').filter(l => l.trim());
  const attributeLines = lines.filter(l =>
    l.includes('|') &&
    !l.startsWith('|--') &&
    !l.startsWith('| ---') &&
    !l.toLowerCase().includes('attribute name')
  );

  const namesOnly = attributeLines
    .map(l => {
      const cols = l.split('|').map(c => c.trim()).filter(Boolean);
      return cols[0] || '';
    })
    .filter(Boolean)
    .join('\n');

  const promptTemplate = `Act as my practical, no-nonsense taxonomy reviewer for our Akeneo PIM. Review the following proposed attributes against our existing dictionary. 
CRITICAL RULES FOR MERGING (The Akeneo Reality Check):
You must consider the data types (Simple Select, Multi-Select, Numeric, Boolean, Text).
1. Dropdown Pollution: Do not merge Simple Select or Multi-Select attributes unless the option lists are practically identical. Keep domain-specific lists (like "Defence Spray Type") isolated.
2. Facet Collision: Do not merge attributes if they would cause bizarre cross-shopping filter results on the website (e.g., merging "Compatible With" for smart home alarms and power tools).
3. Metric Free-Pass: Merge Metric (number + unit) and Boolean (Yes/No) attributes aggressively, as long as it doesn't break Rule 4.
4. Meaning Shift: Do not merge if the word's actual physical definition changes depending on the product family.
5. Component-Specific Fitment vs. Whole-Product Identity (The Core/Component Rule):
Do not merge an attribute into a generic global field (like Length, Width, Capacity, or Surface) if it represents a critical component dimension or a specific fitment variable of a larger, complex assembly. These must remain domain-specific to prevent frontend facet collision and preserve exact compatibility mapping.
6. Logistics vs. Engineering Boundary: Never merge functional fitment dimensions into generic fields (Length, Width) if those fields act as ERP/WMS shipping bounding boxes. Keep engineering specs isolated from logistics. You may merge into a generic attribute only if the attribute defines the identity or physical bound of the entire product as a whole (e.g., a luggage strap's length, a bucket's capacity, or a storage box's material).
7. Global Search & Data Type Test: If merging creates a chaotic global filter of unrelated L2 categories (e.g., plumbing and screws under 'Thread Type'), add a prefix to isolate it. Never merge faceted filters (Dropdowns) into descriptive Text fields.
8. Naming Convention: When an attribute is kept domain-specific to protect the shopping experience, always prefix or suffix the name to make its specific scope/domain clear to both the system and the catalog team to prevent crossover with other categories.
9. The Akeneo 1:1 Label Constraint: Akeneo forces a strict 1:1 relationship between the backend attribute and the frontend display label per locale. You must never merge attributes across different product families if it forces a generic, compromised, or incorrect label onto the e-commerce storefront. 
10. The Customer Vocabulary Test: Always validate a proposed merge against the target customer's actual shopping vocabulary. If a tradesperson, hobbyist, or professional expects a highly specific industry term (e.g., searching for a saw's "Arbor" instead of a generic "Mounting Hole", or a jigsaw's "Shoe" instead of a "Base Plate"), you must reject the merge to protect the UX. 
11. Acceptable Controlled Bloat: It is entirely acceptable to create a new, slightly redundant attribute to preserve the frontend shopping experience. When you do this, you must apply Rule 8 (Naming Convention) and prefix the attribute with its L2 overarching category (e.g., "Power Tool") so this "bloat" remains perfectly grouped alphabetically in the backend and does not scatter the taxonomy.

OUTPUT FORMAT:
Sort in alphabetical order. Do not use tables. Use the following vertical "card" structure for each attribute so it is easy to read on a mobile phone screen:
🔸 [Proposed Attribute Name]
Dictionary Matches: [List of semantic matches from Attribute Dictionary]
Action: [Keep / Merge / Rename] ➡️ [Target PIM Name]
Data Type & Rule Fired: [Assumed Data Type] | [Which of the 4 rules applies?]
Customer View: [Importance: Critical / Important / Technical / Low] | [Customer-Friendliness Assessment]
The Rationale: [A direct, fluff-free explanation of why we are doing this, considering the existing semantic matches, the customer and the system]
Categories: [Typical Categories]

Category context:
L2: ${l2}
L4 Categories: ${l4List}`;

  const fullPrompt = `${promptTemplate}

Here are the proposed new attributes to review:

${ph1bContent}`;

  const condensedPrompt = `${promptTemplate}

Here are the proposed new attribute names to review:

${namesOnly || ph1bContent}`;

  const fullCharCount = fullPrompt.length;
  const condensedCharCount = condensedPrompt.length;

  return {
    fullPrompt,
    condensedPrompt,
    fullCharCount,
    condensedCharCount,
    namesOnly,
    ph1bContent
  };
}

function writePromptsTab(prompts, batch) {
  const sheet = getSheet(SHEETS.PROMPTS);
  sheet.clear();
  sheet.clearFormats();

  const batchInfo = batch.map(r => r['L4 Category']).join(', ');
  const { l2, l3List, l4List, paths } = formatCategoryInfo(batch, getCategoryInfo(batch));

  let row = 1;

  function writeLabel(text) {
    sheet.getRange(row, 1).setValue(text);
    sheet.getRange(row, 1).setFontWeight('bold').setFontSize(10);
    row += 2;
  }

  function writeContent(text) {
    sheet.getRange(row, 1).setValue(text);
    sheet.getRange(row, 1).setFontSize(9).setFontWeight('normal').setWrap(true);
    row += 2;
  }

  

  // ── HEADER ───────────────────────────────────────────────
  writeLabel(`REVIEW PROMPTS — Batch: ${batchInfo}`);

  // ── CHARACTER COUNTS ─────────────────────────────────────
  writeLabel('CHARACTER COUNTS');
  sheet.getRange(row, 1).setValue(`Full version (names + descriptions): ${prompts.fullCharCount.toLocaleString()} characters`);
  sheet.getRange(row, 1).setFontSize(9);
  row++;
  sheet.getRange(row, 1).setValue(`Condensed version (names only): ${prompts.condensedCharCount.toLocaleString()} characters`);
  sheet.getRange(row, 1).setFontSize(9);
  row++;
  sheet.getRange(row, 1).setValue(`NotebookLM: Use condensed version (${prompts.condensedCharCount.toLocaleString()} characters)`);
  sheet.getRange(row, 1).setFontWeight('bold').setFontSize(9);
  row += 3;

  // ── QUICK REFERENCE ──────────────────────────────────────
  writeLabel('CATEGORY CONTEXT & NEW ATTRIBUTES — quick reference');
  const quickRef = [`L2: ${l2}`, `L3: ${l3List}`, `L4: ${l4List}`, '', 'New Attributes:', prompts.ph1bContent].join('\n');
  writeContent(quickRef);

  // ── GEMINI PROMPT ────────────────────────────────────────
  writeLabel('GEMINI PROMPT — Full context including dictionary (attach dictionary CSV)');

  const geminiPrompt = `Phase 1B: Attribute Validation and Customer Friendliness Check

Inputs:
Input 1 (attached csv file): [ATTRIBUTE_DICTIONARY] An attribute dictionary containing attributes, their generalized description and sample values. Use these descriptions and sample values to arrive at the semantics of the attribute name to be used in subsequent steps where we need semantic match.
Input 2: [ADD_TO_DICT_PL]: List of attributes that are potentially not existing in [ATTRIBUTE_DICTIONARY]
Input 3: [L3_LIST] List of sub categories of the category L2 above
Input 4: [L4_LIST] List of sub categories of categories in [L3_LIST]
Input 5: [CATEGORY_PATH] This helps you determine the hierarchy of L2, L3 and L4 categories

Act as my practical, no-nonsense taxonomy reviewer for our Akeneo PIM. Review the following proposed attributes against our existing dictionary. Existing dictionary includes the attached attribute dictionary and the attributes under recently approved attribute list.

CRITICAL RULES FOR MERGING (The Akeneo Reality Check):
You must consider the data types (Simple Select, Multi-Select, Numeric, Boolean, Text).
1. Dropdown Pollution: Do not merge Simple Select or Multi-Select attributes unless the option lists are practically identical. Keep domain-specific lists (like "Defence Spray Type") isolated.
2. Facet Collision: Do not merge attributes if they would cause bizarre cross-shopping filter results on the website (e.g., merging "Compatible With" for smart home alarms and power tools).
3. Metric Free-Pass: Merge Metric (number + unit) and Boolean (Yes/No) attributes aggressively, as long as it doesn't break Rule 4.
4. Meaning Shift: Do not merge if the word's actual physical definition changes depending on the product family.
5. Component-Specific Fitment vs. Whole-Product Identity (The Core/Component Rule):
Do not merge an attribute into a generic global field (like Length, Width, Capacity, or Surface) if it represents a critical component dimension or a specific fitment variable of a larger, complex assembly. These must remain domain-specific to prevent frontend facet collision and preserve exact compatibility mapping.
6. Logistics vs. Engineering Boundary: Never merge functional fitment dimensions into generic fields (Length, Width) if those fields act as ERP/WMS shipping bounding boxes. Keep engineering specs isolated from logistics. You may merge into a generic attribute only if the attribute defines the identity or physical bound of the entire product as a whole (e.g., a luggage strap's length, a bucket's capacity, or a storage box's material).
7. Global Search & Data Type Test: If merging creates a chaotic global filter of unrelated L2 categories (e.g., plumbing and screws under 'Thread Type'), add a prefix to isolate it. Never merge faceted filters (Dropdowns) into descriptive Text fields.
8. Naming Convention: When an attribute is kept domain-specific to protect the shopping experience, always prefix or suffix the name to make its specific scope/domain clear to both the system and the catalog team to prevent crossover with other categories.
9. The Akeneo 1:1 Label Constraint: Akeneo forces a strict 1:1 relationship between the backend attribute and the frontend display label per locale. You must never merge attributes across different product families if it forces a generic, compromised, or incorrect label onto the e-commerce storefront. 
10. The Customer Vocabulary Test: Always validate a proposed merge against the target customer's actual shopping vocabulary. If a tradesperson, hobbyist, or professional expects a highly specific industry term (e.g., searching for a saw's "Arbor" instead of a generic "Mounting Hole", or a jigsaw's "Shoe" instead of a "Base Plate"), you must reject the merge to protect the UX. 
11. Acceptable Controlled Bloat: It is entirely acceptable to create a new, slightly redundant attribute to preserve the frontend shopping experience. When you do this, you must apply Rule 8 (Naming Convention) and prefix the attribute with its L2 overarching category (e.g., "Power Tool") so this "bloat" remains perfectly grouped alphabetically in the backend and does not scatter the taxonomy.

OUTPUT FORMAT:
Sort in alphabetical order. Do not use tables. Use the following vertical "card" structure for each attribute so it is easy to read on a mobile phone screen:
🔸 [Proposed Attribute Name]
Dictionary Matches: [List of semantic matches from Attribute Dictionary]
Action: [Keep / Merge / Rename] ➡️ [Target PIM Name]
Data Type & Rule Fired: [Assumed Data Type] | [Which of the 4 rules applies?]
Customer View: [Importance: Critical / Important / Technical / Low] | [Customer-Friendliness Assessment]
The Rationale: [A direct, fluff-free explanation of why we are doing this, considering the existing semantic matches, the customer and the system]
Categories: [Typical Categories]

Treat [ATTRIBUTE_DICTIONARY] as authoritative but not infallible.

[ADD_TO_DICT_PL]
Attribute name\tDescription
${prompts.ph1bContent}

Category L2: ${l2}

[L4_CATLIST]
${l4List}

[L3_CATLIST]
${l3List}

[CATEGORY_PATH]
${paths}`;

  writeContent(geminiPrompt);

    // ── NOTEBOOKLM PROMPT — FULL ─────────────────────────────
  writeLabel('NOTEBOOKLM PROMPT — FULL VERSION (names + descriptions)');
  writeContent(prompts.fullPrompt);

  // ── NOTEBOOKLM PROMPT — CONDENSED ────────────────────────
  writeLabel('NOTEBOOKLM PROMPT — CONDENSED VERSION (names only)');
  writeContent(prompts.condensedPrompt);

  // ── RECENTLY APPROVED ATTRIBUTES — reference only ────────
  try {
    const { rows: dictReadyRows } = getSheetData(SHEETS.DICT_READY);
    if (dictReadyRows.length > 0) {
      sheet.getRange(row, 1).setValue('━━━ RECENTLY APPROVED ATTRIBUTES — reference only ━━━');
      sheet.getRange(row, 1).setFontWeight('bold').setFontSize(10);
      row += 2;
      dictReadyRows.forEach(r => {
        const name = (r['Attribute Name'] || '').toString().trim();
        const desc = (r['Description'] || '').toString().trim();
        const samples = (r['Sample Values'] || '').toString().trim();
        if (name) {
          sheet.getRange(row, 1).setValue(`${name}\n${desc ? '↳ ' + desc : ''}\n${samples ? '⟐ ' + samples : ''}`);
          sheet.getRange(row, 1).setFontSize(9).setFontWeight('normal').setWrap(true);
          row += 2;
        }
      });
    }
  } catch(e) {
    log(`⚠️ Could not load DICT_READY for reference section: ${e.message}`);
  }


  sheet.setColumnWidth(1, 800);
  
}

// ============================================================
// MAIN PHASE 1AB RUNNER
// ============================================================

function runPhase1AB() {
  try {
    //clearLog();
    //log('Log cleared');
    
    // Save initial state
    saveState({ 
      phase: 'Phase1AB', 
      status: 'running', 
      startTime: new Date().toISOString() 
    });

    // Clear all output tabs fresh
    log('Clearing output tabs...');
    clearOutputTabs();

    // Read inputs
    log('Reading batch...');
    const batch = getReadyBatch();
    log(`Batch: ${batch.map(r => r['L4 Category']).join(', ')}`);

    log('Filtering products...');
    const products = getFilteredProducts(batch);
    log(`${products.length} products found`);
  
    // Abort if product count is too high for a single execution
    if (products.length > 350) {
      log(`⛔ AUTO-ABORTED — ${products.length} products exceeds limit of 350.`);
      clearState();
      return;
    }

    log('Filtering Hybris...');
    const hybris = getFilteredHybris(batch);
    log(`${hybris.length} Hybris rows found`);

    log(`📦 Products: ${products.length} | Hybris rows: ${hybris.length} | Categories: ${batch.length}`);

    // ── COMPLEXITY CHECK ─────────────────────────────────
    const complexity = checkBatchComplexity(batch, hybris, products);
    if (complexity.risk === 'HIGH') {
      log(`⛔ AUTO-ABORTED — ${complexity.reason}`);
      log(`💡 Reduce batch size and rerun`);
      clearState();
      return;
    }
    // ─────────────────────────────────────────────────────

    log('Loading dictionary...');
    const dictionary = getDictionary();
    log(`${dictionary.length} dictionary attributes loaded`);

    log('Loading category info...');
    const categoryRows = getCategoryInfo(batch);

    // Export dictionary to Drive for Gemini attachment
    //log('Exporting dictionary to Drive...');
    //exportDictionaryToDrive();

    // Assemble prompt
    log('Assembling prompt...');
    const prompt = assemblePhase1ABPrompt(batch, products, hybris, dictionary, categoryRows);
    log(`Prompt assembled: ${prompt.length.toLocaleString()} characters`);
    const dictCSV = formatDictionaryAsCSV(dictionary);
    const productCSV = formatProductListAsCSV(products);
    savePromptForManualRun(prompt, 'Phase 1AB', dictCSV, productCSV);

    // Build dictionary attachment
    const documents = [
      { title: 'ATTRIBUTE_DICTIONARY', content: dictCSV }
    ];
    log(`Dictionary attachment: ${dictCSV.length.toLocaleString()} characters`);


    // Call Claude API with dictionary as attachment
    log('Calling Claude API — this may take several minutes...');
    const output = callClaudeAPI(prompt, documents);
    log(`Response received: ${output.length.toLocaleString()} characters`);
    saveRawOutput(output); // full response persisted before any writing

    // ── Hand the write to a FRESH execution so it gets its own ~6-min budget ──
    // The Claude call above can consume most of this run's time; writing all the
    // tabs in the same execution risks Google killing it mid-write (that's the
    // "stopped at STEP 8, no error" symptom). The raw output is safe in
    // RAW_OUTPUT_DEBUG, so a second trigger rebuilds from it and finishes the
    // write + prompts with a clean clock. See finishPhase1ABWrite().
    saveState({ phase: 'Phase1AB', status: 'writing' });
    log('Raw output saved. Scheduling a fresh execution to write the tabs...');
    scheduleFinishWrite();
    log('✅ Phase 1AB API step done — write + prompts queued in a new session (watch the log).');
    return;

  } catch(e) {
    // Save error state
    saveState({ 
      phase: 'Phase1AB', 
      status: 'error', 
      error: e.message,
      time: new Date().toISOString()
    });
    
    // Send error notification
    sendNotification(
      `❌ PIM Phase 1AB Error`,
      `Phase 1AB encountered an error:\n\n${e.message}\n\nCheck SCRIPT_STATE tab for details.`
    );
    
    log(`❌ Error: ${e.message}`);
    throw e;
  }
}

// ── RECOVERY: finish the write from already-saved output ─────
// Phase 1AB saves the full Claude response to RAW_OUTPUT_DEBUG *before* writing.
// If the run stopped mid-write (e.g. hit the ~6-min execution limit after the
// slow API call), run THIS from the Apps Script editor to finish the write and
// prompts from that saved output — no Claude call, so it completes fast.
function finishPhase1ABWrite() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const debugSheet = ss.getSheetByName('RAW_OUTPUT_DEBUG');
    if (!debugSheet || debugSheet.getLastRow() < 1) {
      log('⛔ RAW_OUTPUT_DEBUG is empty — nothing to resume from. Re-run Phase 1AB.');
      return;
    }
    // Rejoin the 45k-char cells with newlines (saveRawOutput split on line boundaries)
    const output = debugSheet.getRange(1, 1, debugSheet.getLastRow(), 1)
      .getValues().map(r => r[0]).filter(Boolean).join('\n');
    log(`Resuming Phase 1AB write from saved output: ${output.length.toLocaleString()} chars`);

    const batch = getReadyBatch() || [];
    if (batch.length === 0) log('⚠️ No Ready batch found — writing anyway; batch is only used for headers/prompts.');

    log('Writing output to PHASE_1AB_WORKING tab...');
    writePhase1ABOutput(output, batch);

    log('Generating review prompts...');
    const categoryRows = getCategoryInfo(batch);
    const prompts = generateReviewPrompts(output, batch, categoryRows);
    writePromptsTab(prompts, batch);
    log(`Full prompt: ${prompts.fullCharCount.toLocaleString()} chars | Condensed: ${prompts.condensedCharCount.toLocaleString()} chars`);

    const reviewSheet = getSheet(SHEETS.REVIEW);
    const reviewCount = Math.max(0, reviewSheet.getLastRow() - 1);

    clearState();

    if (batch.length > 0) {
      sendNotification(
        `✅ PIM Phase 1AB Complete — ${batch.map(r => r['L4 Category']).join(', ')}`,
        `Phase 1AB has completed successfully.\n\nBatch: ${batch.map(r => r['L3 Category'] + ' > ' + r['L4 Category']).join('\n')}\n\nNext steps:\n1. Review PHASE_1AB_WORKING tab\n2. Go to PROMPTS tab — copy the recommended prompt version into NotebookLM and Gemini\n3. Paste results into REVIEW tab\n4. Approve edits into NEWADDITIONS_APPROVED tab\n5. Run Phase 2 when ready`
      );
    }
    log(`✅ Phase 1AB complete — ${reviewCount} attributes ready for review (write finished in a fresh session).`);
  } catch(e) {
    log(`❌ finishPhase1ABWrite error: ${e.message}`);
    throw e;
  }
}

// Schedule the Phase 1AB write to run in a brand-new execution with a fresh
// ~6-min budget. Reuses the runPendingPhase dispatcher (which deletes its own
// trigger on each fire, so triggers don't pile up).
function scheduleFinishWrite() {
  PropertiesService.getScriptProperties().setProperty('PENDING_PHASE', 'finishWrite');
  ScriptApp.newTrigger('runPendingPhase').timeBased().after(10000).create();
  log('PENDING_PHASE set to: finishWrite — write step will run in a fresh execution (~10s).');
}

function testPromptAssembly() {
  const batch = getReadyBatch();
  const products = getFilteredProducts(batch);
  const hybris = getFilteredHybris(batch);
  const dictionary = getDictionary();
  const categoryRows = getCategoryInfo(batch);
  
  const prompt = assemblePhase1ABPrompt(batch, products, hybris, dictionary, categoryRows);
  
  log(`Prompt length: ${prompt.length.toLocaleString()} characters`);
  log(`Estimated tokens: ~${Math.round(prompt.length / 4).toLocaleString()}`);
  log(`\nFirst 500 characters of prompt:\n${prompt.substring(0, 500)}`);
  log(`\nLast 200 characters of prompt:\n${prompt.substring(prompt.length - 200)}`);
}

function testPromptCompleteness() {
  const batch = getReadyBatch();
  const products = getFilteredProducts(batch);
  const hybris = getFilteredHybris(batch);
  const dictionary = getDictionary();
  const categoryRows = getCategoryInfo(batch);
  
  const prompt = assemblePhase1ABPrompt(batch, products, hybris, dictionary, categoryRows);
  
  // Check all sections are present
  const checks = [
    ['[L4_CATLIST]', prompt.includes('[L4_CATLIST]')],
    ['[L3_CATLIST]', prompt.includes('[L3_CATLIST]')],
    ['[CATEGORY_PATH]', prompt.includes('[CATEGORY_PATH]')],
    ['[ATTRIBUTE_DICTIONARY]', prompt.includes('[ATTRIBUTE_DICTIONARY]')],
    ['[PRODUCT_LIST]', prompt.includes('[PRODUCT_LIST]')],
    ['[HYBRIS]', prompt.includes('[HYBRIS]')],
    ['Blenders in Hybris', prompt.includes('Blenders')],
    ['Juicers in Hybris', prompt.includes('Juicers')],
    ['Rice cookers in Hybris', prompt.includes('Rice cookers')],
    ['Blenders in products', (prompt.match(/Blenders/g) || []).length > 1],
  ];
  
  checks.forEach(([name, result]) => {
    log(`${result ? '✅' : '❌'} ${name}`);
  });

  // Count Hybris rows injected
  const hybrisRowCount = (prompt.match(/\n/g) || []).length;
  log(`\nTotal lines in prompt: ${hybrisRowCount}`);
  log(`Hybris rows injected: ${hybris.length}`);
  log(`Products injected: ${products.length}`);
  log(`Dictionary attributes injected: ${dictionary.length}`);
}

function testPromptExtraction() {
  const sheet = getSheet(SHEETS.PHASE_1AB_WORKING);
  const output = sheet.getRange(4, 1).getValue();
  
  log(`Total output length: ${output.length.toLocaleString()} characters`);
  
  // Find all section headers Claude used
  const headers = output.match(/\[.*?\]/g) || [];
  const unique = [...new Set(headers)];
  log(`\nSection headers found in output:`);
  unique.forEach(h => log(`  ${h}`));
  
  // Show output around the PH1B section
  const ph1bIndex = output.indexOf('ADD_ME_TO_THE_DICTIONARY_PH1B');
  log(`\nADD_ME_TO_THE_DICTIONARY_PH1B found at index: ${ph1bIndex}`);
  if (ph1bIndex > -1) {
    log(`Context: ...${output.substring(ph1bIndex - 20, ph1bIndex + 100)}...`);
  }
}

function testTableFormatting() {
  const sheet = getSheet(SHEETS.PHASE_1AB_WORKING);
  const output = sheet.getRange(4, 1).getValue();
  
  // Find HYBRIS_PASS1 section and show 500 chars of it
  const hybrisIndex = output.indexOf('[HYBRIS_PASS1]');
  if (hybrisIndex > -1) {
    log('HYBRIS_PASS1 sample:');
    log(output.substring(hybrisIndex, hybrisIndex + 500));
  }
  
  // Find DICTIONARY_SUBSET and show 500 chars
  const dictIndex = output.indexOf('[DICTIONARY_SUBSET]');
  if (dictIndex > -1) {
    log('\nDICTIONARY_SUBSET sample:');
    log(output.substring(dictIndex, dictIndex + 500));
  }

  // Find ADD_TO_DICT_PL and show 500 chars
  const addDictIndex = output.indexOf('[ADD_TO_DICT_PL]');
  if (addDictIndex > -1) {
    log('\nADD_TO_DICT_PL sample:');
    log(output.substring(addDictIndex, addDictIndex + 500));
  }

  // Find STEP 10 and show 500 chars
  const step10Index = output.indexOf('Step 10');
  if (step10Index > -1) {
    log('\nSTEP 10 sample:');
    log(output.substring(step10Index, step10Index + 500));
  }
}

function testStep10() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const debugSheet = ss.getSheetByName('RAW_OUTPUT_DEBUG');
  if (!debugSheet) {
    log('RAW_OUTPUT_DEBUG tab not found — run Phase 1AB first');
    return;
  }
  
  const rawOutput = debugSheet.getRange(1, 1).getValue();
  if (!rawOutput) {
    log('No raw output found in RAW_OUTPUT_DEBUG');
    return;
  }
  
  log(`Total output length: ${rawOutput.length.toLocaleString()} characters`);
  
  const idx = rawOutput.indexOf('Step 10');
  log(`\nStep 10 found at index: ${idx}`);
  if (idx > -1) {
    log(`Content around Step 10:\n${rawOutput.substring(idx, idx + 800)}`);
  } else {
    log('Step 10 not found in output');
    log(`\nLast 500 chars:\n${rawOutput.substring(rawOutput.length - 500)}`);
  }
  
  // Check where --- appears relative to Step 10
  const dashes = [];
  let pos = 0;
  while ((pos = rawOutput.indexOf('---', pos)) !== -1) {
    dashes.push(pos);
    pos += 3;
  }
  log(`\n'---' appears at ${dashes.length} positions: ${dashes.join(', ')}`);
  if (idx > -1) {
    const dashesAfterStep10 = dashes.filter(d => d > idx);
    log(`'---' appearances after Step 10: ${dashesAfterStep10.join(', ')}`);
  }
}

function saveRawOutput(output) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let debugSheet = ss.getSheetByName('RAW_OUTPUT_DEBUG');
  log('Saving RawOutput to RAW_OUTPUT_DEBUG');
  if (!debugSheet) {
    debugSheet = ss.insertSheet('RAW_OUTPUT_DEBUG');
  }
  debugSheet.clearContents();
  
  // Split into 45k chunks across multiple cells
  const maxChars = 45000;
  const lines = output.split('\n');
  let currentChunk = '';
  let row = 1;
  for (const line of lines) {
    if (currentChunk.length + line.length + 1 > maxChars && currentChunk.length > 0) {
      debugSheet.getRange(row, 1).setValue(currentChunk);
      row++;
      currentChunk = '';
    }
    currentChunk += (currentChunk.length > 0 ? '\n' : '') + line;
  }
  if (currentChunk.length > 0) {
    debugSheet.getRange(row, 1).setValue(currentChunk);
    row++;
  }
  log(`✅ RawOutput saved across ${row - 1} cell(s)`);
}

function savePromptForManualRun(prompt, phase, dictCSV, productCSV) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tabName = 'LAST_PROMPT';
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
  }
  sheet.clearContents();
  sheet.setColumnWidth(1, 800);

  let row = 1;

  // Header info
  sheet.getRange(row, 1).setValue(`Phase: ${phase}`);
  sheet.getRange(row, 1).setFontWeight('bold').setFontSize(10);
  row++;
  sheet.getRange(row, 1).setValue(`Generated: ${new Date().toLocaleString()}`);
  sheet.getRange(row, 1).setFontSize(9);
  row++;
  sheet.getRange(row, 1).setValue(`Prompt: ${prompt.length.toLocaleString()} chars | Dictionary: ${dictCSV ? dictCSV.length.toLocaleString() : 'N/A'} chars | Product list: ${productCSV ? productCSV.length.toLocaleString() : 'N/A'} chars`);
  sheet.getRange(row, 1).setFontSize(9);
  row += 2;

  // Prompt section
  sheet.getRange(row, 1).setValue('━━━ PROMPT — paste this into Claude/ChatGPT ━━━');
  sheet.getRange(row, 1).setFontWeight('bold').setFontSize(10);
  row++;
  const maxChars = 45000;
  let remaining = prompt;
  while (remaining.length > 0) {
    const chunk = remaining.substring(0, maxChars);
    remaining = remaining.substring(maxChars);
    sheet.getRange(row, 1).setValue(chunk);
    sheet.getRange(row, 1).setFontSize(9).setWrap(true);
    row++;
  }
  row++;

  // Dictionary section
  if (dictCSV) {
    sheet.getRange(row, 1).setValue('━━━ ATTRIBUTE DICTIONARY — attach as CSV ━━━');
    sheet.getRange(row, 1).setFontWeight('bold').setFontSize(10);
    row++;
    let remainingDict = dictCSV;
    while (remainingDict.length > 0) {
      const chunk = remainingDict.substring(0, maxChars);
      remainingDict = remainingDict.substring(maxChars);
      sheet.getRange(row, 1).setValue(chunk);
      sheet.getRange(row, 1).setFontSize(9).setWrap(true);
      row++;
    }
    row++;
  }

  // Product list section
  if (productCSV) {
    sheet.getRange(row, 1).setValue('━━━ PRODUCT LIST — attach as CSV ━━━');
    sheet.getRange(row, 1).setFontWeight('bold').setFontSize(10);
    row++;
    let remainingProducts = productCSV;
    while (remainingProducts.length > 0) {
      const chunk = remainingProducts.substring(0, maxChars);
      remainingProducts = remainingProducts.substring(maxChars);
      sheet.getRange(row, 1).setValue(chunk);
      sheet.getRange(row, 1).setFontSize(9).setWrap(true);
      row++;
    }
  }

  log(`✅ LAST_PROMPT tab written — prompt + dictionary + product list`);
}

function clearOutputTabs() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Delete and recreate PHASE_1AB_WORKING
  const workingSheet = ss.getSheetByName(SHEETS.PHASE_1AB_WORKING);
  if (workingSheet) ss.deleteSheet(workingSheet);
  ss.insertSheet(SHEETS.PHASE_1AB_WORKING);
  log(`✅ Recreated: ${SHEETS.PHASE_1AB_WORKING}`);

  // Clear other tabs normally
  const tabsToClear = [
    SHEETS.DICTIONARY_SUBSET,
    SHEETS.NEWADDITIONS_STAGING,
    SHEETS.PROMPTS,
    SHEETS.REVIEW
  ];
  tabsToClear.forEach(tabName => {
    try {
      const sheet = getSheet(tabName);
      sheet.clear();
      sheet.clearFormats();
      log(`✅ Cleared: ${tabName}`);
    } catch(e) {
      log(`⚠️ Could not clear ${tabName}: ${e.message}`);
    }
  });
}

// ============================================================
// PHASE 1C — RECONCILIATION
// ============================================================

function runPostReviewPipeline() {

    try {
    
    saveState({ phase: 'PostReview', status: 'running', startTime: new Date().toISOString() });
    log('Post Review State Saved');

    // ── READ BATCH UPFRONT ────────────────────────────────
    const batch = getReadyBatch();
    log(`Batch: ${batch.map(r => r['L4 Category']).join(', ')}`);

    // ── STEP 1: VALIDATE ──────────────────────────────────
    log('Validating REVIEW tab...');
    const validation = validateReview();

    if (validation.warnings.length > 0) {
      validation.warnings.forEach(w => log(`⚠️ Warning: ${w}`));
    }

    if (!validation.valid) {
      const errorMsg = `REVIEW tab validation failed:\n\n${validation.errors.join('\n')}`;
      log(`❌ ${errorMsg}`);
      saveState({ phase: 'PostReview', status: 'error', error: errorMsg });
      throw new Error(errorMsg);
    }

    const { counts } = validation;
    log(`✅ Validation passed — Keep: ${counts.keep} | Rename: ${counts.rename} | Merge: ${counts.merge} | Delete: ${counts.delete} | Total: ${counts.total}`);

    // ── SAVE NOTES ────────────────────────────────────────
    const notesRows = validation.rows
      .filter(r => (r['Notes'] || '').trim() !== '')
      .map(r => ({
        'Batch Date': new Date().toLocaleDateString(),
        'L4 Category': (r['Applicable Categories'] || '').trim(),
        'Attribute Name': (r['Attribute Name'] || '').trim(),
        'Decision': (r['Your Decision'] || '').trim(),
        'Updated Name': (r['Updated Attribute Name'] || '').trim(),
        'Notes': (r['Notes'] || '').trim()
      }));

    if (notesRows.length > 0) {
      const notesHeaders = ['Batch Date', 'L4 Category', 'Attribute Name', 'Decision', 'Updated Name', 'Notes'];
      const notesSheet = getSheet(SHEETS.REVIEW_NOTES);
      if (notesSheet.getLastRow() === 0) {
        notesSheet.getRange(1, 1, 1, notesHeaders.length).setValues([notesHeaders]);
      }
      const data = notesRows.map(r => notesHeaders.map(h => r[h] || ''));
      notesSheet.getRange(notesSheet.getLastRow() + 1, 1, data.length, notesHeaders.length).setValues(data);
      log(`📝 REVIEW_NOTES: ${notesRows.length} notes saved`);
    }

    // ── STEP 2: PROCESS REVIEW ────────────────────────────
    log('Processing review decisions...');

    const dictReadyRows = [];
    const dictReadyDropdownRows = [];
    const newAdditionsApprovedRows = [];
    const dictSubsetAppendRows = [];

    // Read staging for full metadata
    const { rows: stagingRows } = getSheetData(SHEETS.NEWADDITIONS_STAGING);
    const stagingLookup = {};
    stagingRows.forEach(r => {
      const name = (r['Attribute Name'] || r['attribute name'] || '').trim();
      if (name) stagingLookup[name.toLowerCase()] = r;
    });

    // Find dropdown section in staging
    const dropdownStartIndex = stagingRows.findIndex(r =>
      Object.values(r).some(v => (v||'').toString().includes('DROPDOWN_PH1B'))
    );
    const dropdownRows = dropdownStartIndex > -1 ? stagingRows.slice(dropdownStartIndex + 1) : [];
    const dropdownLookup = {};
    dropdownRows.forEach(r => {
      const name = (r['Dropdown Name'] || r['dropdown name'] || '').trim();
      if (name) dropdownLookup[name.toLowerCase()] = r;
    });

    // Build rename map simultaneously
    const renameMap = {};

    log(`Processing ${validation.rows.length} rows...`);
    validation.rows.forEach(r => {
      // Coerce all values to strings — Sheets can auto-convert numbers/booleans
      Object.keys(r).forEach(k => { 
        if (r[k] !== null && r[k] !== undefined) r[k] = String(r[k]); 
      });   
      log(`Processing: ${String(r['Attribute Name'])}`); // temporary   
      const originalName = (r['Attribute Name'] !== null && r['Attribute Name'] !== undefined ? String(r['Attribute Name']) : '').trim();
      const decision = (r['Your Decision'] || '').trim().toLowerCase();
      const updatedName = (r['Updated Attribute Name'] || '').trim() || originalName;
      const updatedDropdownName = (r['Updated Dropdown Name'] || '').trim() || updatedName;
      const description = (r['Description'] || '').trim();
      const sampleValues = (r['Sample Values'] || '').trim();
      const datatype = (r['Datatype'] || '').trim();
      const categories = (r['Applicable Categories'] || '').trim();
      const stagingMeta = stagingLookup[originalName.toLowerCase()] || {};
      const unit = stagingMeta['Unit of Measure'] || stagingMeta['unit of measure'] || '';
      const mandateSource = stagingMeta['Mandate Source'] || stagingMeta['mandate source'] || '';
      const needsLocalisation = stagingMeta['Needs Language Localization'] || stagingMeta['needs language localization'] || '';

      if (decision === 'delete') {
        return;
      }

if (decision === 'merge') {
  dictSubsetAppendRows.push({
    'Attribute Name': updatedName,
    'Sample Values': '',
    'Description': ''
  });
  // Add to renameMap so PL_ATTRIBUTELIST gets updated with target name
  if (updatedName !== originalName) {
    renameMap[originalName.toLowerCase()] = {
      original: originalName,
      updatedName,
      action: 'Rename'
    };
  }
  // Save proposed dropdown values to REVIEW_NOTES for reference
  if (datatype.toLowerCase().includes('dropdown')) {
    const origDropdown = stagingMeta['Dropdown Name'] || stagingMeta['dropdown name'] || originalName;
    const dropData = dropdownLookup[origDropdown.toLowerCase()] || dropdownLookup[originalName.toLowerCase()] || {};
    const proposedValues = dropData['Values'] || dropData['Exhaustive list of allowed values'] || sampleValues;
    if (proposedValues) {
      try {
        const notesHeaders = ['Batch Date', 'L4 Category', 'Attribute Name', 'Decision', 'Updated Name', 'Notes'];
        const notesSheet = getSheet(SHEETS.REVIEW_NOTES);
        if (notesSheet.getLastRow() === 0) {
          notesSheet.getRange(1, 1, 1, notesHeaders.length).setValues([notesHeaders]);
        }
        const row = [
          new Date().toLocaleDateString(),
          categories,
          originalName,
          `Merge → ${updatedName}`,
          updatedName,
          `Proposed dropdown values: ${proposedValues}`
        ];
        notesSheet.getRange(notesSheet.getLastRow() + 1, 1, 1, notesHeaders.length).setValues([row]);
      } catch(e) {
        log(`⚠️ Could not save merge dropdown values to REVIEW_NOTES: ${e.message}`);
      }
    }
  }
  return;
}

      // Keep or Rename
      const action = (decision === 'rename' && updatedName !== originalName) ? 'Rename' : 'Keep';
      renameMap[originalName.toLowerCase()] = { original: originalName, updatedName, action };

      dictReadyRows.push({
        'Attribute Name': updatedName,
        'Description': description,
        'Datatype': datatype,
        'Dropdown Name': updatedDropdownName,
        'Unit': unit,
        'Mandate Source': mandateSource,
        'Needs Localisation': needsLocalisation,
        'Sample Values': sampleValues,
        'Applicable Categories': categories
      });

      if (datatype.toLowerCase().includes('dropdown')) {
        const origDropdown = stagingMeta['Dropdown Name'] || stagingMeta['dropdown name'] || originalName;
        const dropData = dropdownLookup[origDropdown.toLowerCase()] || dropdownLookup[originalName.toLowerCase()] || {};
        const values = dropData['Values'] || dropData['Exhaustive list of allowed values'] || sampleValues;
        dictReadyDropdownRows.push({ 'Dropdown Name': updatedDropdownName, 'Values': values });
      }

      newAdditionsApprovedRows.push({
        'Updated Attribute Name': updatedName,
        'Description': description,
        'Sample Values': sampleValues
      });

    });

// Write DICT_READY
    // Append DICT_READY
const dictReadyHeaders = ['Attribute Name','Description','Datatype','Dropdown Name','Unit','Mandate Source','Needs Localisation','Sample Values','Applicable Categories'];
const dictReadySheet = getSheet(SHEETS.DICT_READY);
if (dictReadySheet.getLastRow() === 0) {
  dictReadySheet.getRange(1, 1, 1, dictReadyHeaders.length).setValues([dictReadyHeaders]);
}
if (dictReadyRows.length > 0) {
  const data = dictReadyRows.map(r => dictReadyHeaders.map(h => r[h] || ''));
  dictReadySheet.getRange(dictReadySheet.getLastRow() + 1, 1, data.length, dictReadyHeaders.length).setValues(data);
}
log(`✅ DICT_READY: ${dictReadyRows.length} attributes appended`);

// Append DICT_READY_DROPDOWNS
const dropdownHeaders = ['Dropdown Name','Values'];
const dropdownSheet = getSheet(SHEETS.DICT_READY_DROPDOWNS);
if (dropdownSheet.getLastRow() === 0) {
  dropdownSheet.getRange(1, 1, 1, dropdownHeaders.length).setValues([dropdownHeaders]);
}
if (dictReadyDropdownRows.length > 0) {
  const dropData = dictReadyDropdownRows.map(r => dropdownHeaders.map(h => r[h] || ''));
  dropdownSheet.getRange(dropdownSheet.getLastRow() + 1, 1, dropData.length, dropdownHeaders.length).setValues(dropData);
}
log(`✅ DICT_READY_DROPDOWNS: ${dictReadyDropdownRows.length} dropdowns appended`);

    // Write NEWADDITIONS_APPROVED
    writeToSheet(SHEETS.NEWADDITIONS_APPROVED, [
      'Updated Attribute Name','Description','Sample Values'
    ], newAdditionsApprovedRows);
    log(`✅ NEWADDITIONS_APPROVED: ${newAdditionsApprovedRows.length} attributes`);
    log(`About to read products/hybris for Phase 2 setup...`);

    // Append merges to DICTIONARY_SUBSET
    if (dictSubsetAppendRows.length > 0) {
      const { headers: dictSubHeaders, rows: existingDictSubRows } = getSheetData(SHEETS.DICTIONARY_SUBSET);
      const toStr = v => (v !== null && v !== undefined ? String(v) : '');
      const existingNames = new Set(existingDictSubRows.map(r => toStr(r['Attribute Name']).toLowerCase()));
      const newRows = dictSubsetAppendRows.filter(r => !existingNames.has(toStr(r['Attribute Name']).toLowerCase()));
      if (newRows.length > 0) {
        const headers = dictSubHeaders.length > 0 ? dictSubHeaders : ['Attribute Name','Sample Values','Description'];
        appendToSheet(SHEETS.DICTIONARY_SUBSET, headers, newRows);
        log(`✅ DICTIONARY_SUBSET: ${newRows.length} merged attributes appended`);
      }
    }

    // ── STEP 3: PHASE 1C — UPDATE PL_ATTRIBUTELIST ────────
    log('Running Phase 1C — updating PL_ATTRIBUTELIST to V3...');

    const { headers: plHeaders, rows: allPlRows } = getSheetData(SHEETS.PL_ATTRIBUTELIST);
    if (allPlRows.length === 0) {
  log(`⚠️ PL_ATTRIBUTELIST is empty — skipping Phase 1C, Phase 2 will run without completeness check`);
    } else {

    // Only process rows from the current batch — not previous accumulated rows
    const batchL4Categories = batch.map(r => r['L4 Category']);
    let plRows = allPlRows.filter(r => 
      batchL4Categories.includes((r['Category'] || r['L4 Category'] || r['category'] || '').trim())
    );
    log(`Batch L4 categories: ${JSON.stringify(batchL4Categories)}`);
    log(`PL_ATTRIBUTELIST Category values: ${JSON.stringify(allPlRows.map(r => r['Category']))}`);

    if (plRows.length === 0) {
      // Fallback — if category column name doesn't match, take the last N rows
      // where N = number of rows added in the current Phase 1AB run
      log(`⚠️ Could not filter PL_ATTRIBUTELIST by category — using last ${batchL4Categories.length} rows`);
      plRows = allPlRows.slice(-batchL4Categories.length);
    }
    
    log(`PL_ATTRIBUTELIST: ${allPlRows.length} total rows, ${plRows.length} rows for current batch`);

    const updatedPlRows = plRows.map(row => {
      const newRow = { ...row };
      Object.keys(newRow).forEach(key => {
        const cellValue = (newRow[key] || '').toString().trim();
      // Strip parenthetical suffixes like (Boolean), (Dropdown) before matching
      const normalizedValue = cellValue.replace(/\s*\([^)]*\)\s*$/, '').trim();
      const match = renameMap[normalizedValue.toLowerCase()] || renameMap[cellValue.toLowerCase()];
        if (match) {
          newRow[key] = match.action === 'Rename' ? match.updatedName : cellValue;
        }
      });
      return newRow;
    });

    const deduplicatedPlRows = updatedPlRows.map(row => {
      const newRow = { ...row };
      Object.keys(newRow).forEach(key => {
        const val = (newRow[key] || '').toString();
        if (val.includes(',')) {
          const parts = val.split(',').map(v => v.trim()).filter(Boolean);
          const unique = [...new Set(parts.map(p => p.toLowerCase()))].map(lower =>
            parts.find(p => p.toLowerCase() === lower)
          );
          newRow[key] = unique.join(', ');
        }
      });
      return newRow;
    });

    // Remove existing rows for this batch, then append updated versions
    const plSheet = getSheet(SHEETS.PL_ATTRIBUTELIST);
    const allData = plSheet.getDataRange().getValues();
    const headers2 = allData[0];
    const categoryColIndex = headers2.indexOf('Category') > -1 
      ? headers2.indexOf('Category') 
      : headers2.indexOf('L4 Category');
    
    // Delete rows belonging to current batch (from bottom up to avoid index shifting)
    for (let i = allData.length - 1; i >= 1; i--) {
      if (batchL4Categories.includes((allData[i][categoryColIndex] || '').trim())) {
        plSheet.deleteRow(i + 1);
      }
    }

    // Append updated versions
    if (deduplicatedPlRows.length > 0) {
      const data = deduplicatedPlRows.map(r => plHeaders.map(h => r[h] || ''));
      plSheet.getRange(plSheet.getLastRow() + 1, 1, data.length, plHeaders.length).setValues(data);
    }
    log(`✅ PL_ATTRIBUTELIST: replaced ${plRows.length} rows for current batch with updated V3 versions`);

    }

    // ── STEP 4: PHASE 2 ───────────────────────────────────
   // ── STEP 4: PHASE 2 ───────────────────────────────────
    log('Running Phase 2...');


    const products = getFilteredProducts(batch);
    const hybris = getFilteredHybris(batch);
    const dictionary = getDictionary();
    const categoryRows = getCategoryInfo(batch);
    const { rows: newAdditions } = getSheetData(SHEETS.NEWADDITIONS_APPROVED);
    if (newAdditions.length === 0) {
      log(`ℹ️ NEWADDITIONS_APPROVED is empty — all attributes matched existing dictionary entries`);
    }
    const { rows: dictSubset } = getSheetData(SHEETS.DICTIONARY_SUBSET);
    const { rows: plV3 } = getSheetData(SHEETS.PL_ATTRIBUTELIST);

    const sharedData = {
      counts,
      dictReadyCount: dictReadyRows.length,
      dictReadyDropdownCount: dictReadyDropdownRows.length,
      newAdditionsCount: newAdditionsApprovedRows.length,
      dictSubsetCount: dictSubsetAppendRows.length
    };

    // Split batch if needed
    const phase2Complexity = checkBatchComplexity(batch, hybris, products);
    //const subBatches = phase2Complexity.risk === 'HIGH' || batch.length > 4
      const subBatches = phase2Complexity.risk === 'HIGH' || phase2Complexity.risk === 'MEDIUM'
    //const subBatches = phase2Complexity.risk === 'HIGH'
      ? [batch.slice(0, Math.ceil(batch.length / 2)), batch.slice(Math.ceil(batch.length / 2))]
      : [batch];

    log(`Phase 2 will run in ${subBatches.length} sub-batch${subBatches.length > 1 ? 'es' : ''}`);

    // Queue Phase 2 as separate trigger — fresh 6 minute window
    PropertiesService.getScriptProperties().setProperty('PENDING_PHASE2_STATE', JSON.stringify({
      stage: 'generate',
      subBatches: JSON.stringify(subBatches),
      currentIndex: 0,
      sharedData: JSON.stringify(sharedData)
    }));
    ScriptApp.newTrigger('runPendingPhase2SubBatch')
      .timeBased()
      .after(10000)
      .create();
    log(`✅ Phase 1C complete — Phase 2 queued, starting in 10 seconds`);
    log(`📋 Phase 2 will process ${subBatches.length} sub-batch${subBatches.length > 1 ? 'es' : ''}`);
    clearState();

  } catch(e) {
    saveState({ phase: 'PostReview', status: 'error', error: e.message, time: new Date().toISOString() });
    log(`❌ Error: ${e.message}`);
    throw e;
  }
}

function runPhase2SubBatch(subBatches, currentIndex, sharedData) {
  if (currentIndex === 0) {
    try {
      const s = getSheet(SHEETS.PHASE_2_WORKING);
      s.clear(); s.clearFormats();
      log(`✅ Cleared: PHASE_2_WORKING`);
    } catch(e) { log(`⚠️ Could not clear PHASE_2_WORKING: ${e.message}`); }
  }
  
  if (currentIndex >= subBatches.length) {
    // All sub-batches done
    log(`✅ All Phase 2 sub-batches complete`);
    finalisePostReview(sharedData);
    return;
  }

  const subBatch = subBatches[currentIndex];
  log(`Running Phase 2 sub-batch ${currentIndex + 1}/${subBatches.length}: ${subBatch.map(r => r['L4 Category']).join(', ')}`);

  const subProducts = getFilteredProducts(subBatch);
  const subHybris = getFilteredHybris(subBatch);
  const subCategoryRows = getCategoryInfo(subBatch);
  const dictionary = getDictionary();
  const dictCSV = formatDictionaryAsCSV(dictionary);
  const documents = [{ title: 'ATTRIBUTE_DICTIONARY', content: dictCSV }];

  const { rows: newAdditions } = getSheetData(SHEETS.NEWADDITIONS_APPROVED);
  const { rows: dictSubset } = getSheetData(SHEETS.DICTIONARY_SUBSET);
  const { rows: plV3 } = getSheetData(SHEETS.PL_ATTRIBUTELIST);

  const prompt = assemblePhase2Prompt(subBatch, subProducts, subHybris, dictionary, subCategoryRows, newAdditions, dictSubset, plV3);
  log(`Phase 2 sub-batch ${currentIndex + 1} prompt: ${prompt.length.toLocaleString()} characters`);

  savePromptForManualRun(prompt, `Phase 2 sub-batch ${currentIndex + 1}`, null, formatProductListAsCSV(subProducts)); // ← add this line

  log(`Calling Claude API for Phase 2 sub-batch ${currentIndex + 1}...`);
  const output = callClaudeAPI(prompt, documents);
  log(`Phase 2 sub-batch ${currentIndex + 1} response: ${output.length.toLocaleString()} characters`);
  saveRawOutput(output);
  log(`✅ Phase 2 sub-batch ${currentIndex + 1} generated`);

  // Hand the write to a FRESH execution so it gets its own ~6-min budget.
  // The raw output is safe in RAW_OUTPUT_DEBUG; writePhase2SubBatch() rebuilds
  // from it, writes PHASE_2_WORKING, then advances to the next sub-batch.
  const state = {
    stage: 'write',
    subBatches: JSON.stringify(subBatches),
    currentIndex: currentIndex,
    sharedData: JSON.stringify(sharedData)
  };
  PropertiesService.getScriptProperties().setProperty('PENDING_PHASE2_STATE', JSON.stringify(state));
  ScriptApp.newTrigger('runPendingPhase2SubBatch')
    .timeBased()
    .after(10000)
    .create();
  log(`Phase 2 sub-batch ${currentIndex + 1} — write queued in a fresh session`);
}

// STAGE B: write one Phase 2 sub-batch from the saved raw output, then advance.
function writePhase2SubBatch(subBatches, currentIndex, sharedData) {
  const subBatch = subBatches[currentIndex];
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const debugSheet = ss.getSheetByName('RAW_OUTPUT_DEBUG');
  if (!debugSheet || debugSheet.getLastRow() < 1) {
    log(`⛔ RAW_OUTPUT_DEBUG empty — cannot write Phase 2 sub-batch ${currentIndex + 1}`);
    return;
  }
  const output = debugSheet.getRange(1, 1, debugSheet.getLastRow(), 1)
    .getValues().map(r => r[0]).filter(Boolean).join('\n');
  log(`Writing Phase 2 sub-batch ${currentIndex + 1}/${subBatches.length} from saved output: ${output.length.toLocaleString()} chars`);
  writePhase2Output(output, subBatch);
  log(`✅ Phase 2 sub-batch ${currentIndex + 1} written`);

  // Advance: queue the NEXT sub-batch's generate stage, or finalise.
  if (currentIndex + 1 < subBatches.length) {
    const state = {
      stage: 'generate',
      subBatches: JSON.stringify(subBatches),
      currentIndex: currentIndex + 1,
      sharedData: JSON.stringify(sharedData)
    };
    PropertiesService.getScriptProperties().setProperty('PENDING_PHASE2_STATE', JSON.stringify(state));
    ScriptApp.newTrigger('runPendingPhase2SubBatch')
      .timeBased()
      .after(10000)
      .create();
    log(`Phase 2 sub-batch ${currentIndex + 2} queued via trigger`);
  } else {
    finalisePostReview(sharedData);
  }
}

function runPendingPhase2SubBatch() {
  const raw = PropertiesService.getScriptProperties().getProperty('PENDING_PHASE2_STATE');
  if (!raw) { log('No pending Phase 2 sub-batch state found'); return; }

  PropertiesService.getScriptProperties().deleteProperty('PENDING_PHASE2_STATE');
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runPendingPhase2SubBatch')
    .forEach(t => ScriptApp.deleteTrigger(t));

  try {
    const state = JSON.parse(raw);
    const subBatches = JSON.parse(state.subBatches);
    const sharedData = JSON.parse(state.sharedData);
    // stage 'write' finishes the current sub-batch's write in a fresh session;
    // anything else (or missing, for the initial queue) runs the generate stage.
    if (state.stage === 'write') {
      writePhase2SubBatch(subBatches, state.currentIndex, sharedData);
    } else {
      runPhase2SubBatch(subBatches, state.currentIndex, sharedData);
    }
  } catch(e) {
    log(`❌ Phase 2 sub-batch failed: ${e.message}`);
  }
}

function finalisePostReview(sharedData) {
  const { counts, dictReadyCount, dictReadyDropdownCount, newAdditionsCount, dictSubsetCount, deduplicatedPlCount } = sharedData;
  const masterTotal = getSheet(SHEETS.MASTER_OUTPUT).getLastRow() - 1;
  const dictReadyTotal = getSheet(SHEETS.DICT_READY).getLastRow() - 1;
  const plTotal = getSheet(SHEETS.PL_ATTRIBUTELIST).getLastRow() - 1;

  log(`✅ Post-review pipeline complete`);
  log(`📊 MASTER_OUTPUT: ${masterTotal} total rows`);
  log(`📚 DICT_READY: ${dictReadyTotal} total rows (${dictReadyCount} added this run)`);
  log(`📋 PL_ATTRIBUTELIST: ${plTotal} total rows`);
  log(`🔤 DICTIONARY_SUBSET: merged ${dictSubsetCount} attributes appended`);
  clearState();
}

// ============================================================
// PHASE 2 — LEGACY MAPPING & FINAL ATTRIBUTE ASSEMBLY
// ============================================================

function formatNewAdditionsForPrompt(rows) {
  if (!rows || rows.length === 0) return 'No new additions approved yet.';
  const lines = [];
  rows.forEach(r => {
    const name = (r['Updated Attribute Name'] || '').trim();
    const desc = (r['Description'] || '').trim();
    const samples = (r['Sample Values'] || '').trim();
    if (!name) return;
    lines.push(`[${name}] [${desc}] [${samples}]`);
  });
  return lines.join('\n');
}

function formatDictionarySubsetForPrompt(rows) {
  if (!rows || rows.length === 0) return 'No dictionary subset available.';
  return rows.map(r => r['Attribute Name'] || r['Attribute name'] || '').filter(Boolean).join('\n');
}

function formatPLAttributeListV2ForPrompt(rows) {
  if (!rows || rows.length === 0) return 'No attribute list available.';
  const headers = Object.keys(rows[0]);
  const lines = [headers.join('\t')];
  rows.forEach(r => lines.push(headers.map(h => r[h] || '').join('\t')));
  return lines.join('\n');
}

function assemblePhase2Prompt(batch, products, hybris, dictionary, categoryRows, newAdditions, dictSubset, plAttributeListV2) {
  const { l2, l3List, l4List, paths } = formatCategoryInfo(batch, categoryRows);
  const productsFormatted = formatProductListForPrompt(products);
  const hybrisFormatted = formatHybrisForPrompt(hybris);
  const newAdditionsFormatted = formatNewAdditionsForPrompt(newAdditions);
  const dictSubsetFormatted = formatDictionarySubsetForPrompt(dictSubset);
  const plV2Formatted = formatPLAttributeListV2ForPrompt(plAttributeListV2);

  const prompt = `Phase 2: Legacy Mapping & Final Attribute Assembly

Role
You are a Product Data Architect assisting with legacy-to-modern attribute migration for a Nordic home goods retailer.
Your task is to generate final, customer-complete L4 attribute sets using:
the existing attribute dictionary
approved new attribute additions
legacy (Hybris) attributes as reference only
Do not invent alternative naming unless explicitly required.
Do not use any other source for attributes

Inputs
[DICTIONARY_SUBSET] Existing approved subset of attribute names from the attribute dictionary
[NEWADDITIONS_DICTIONARY] Newly approved modern attributes that may be used in addition to the dictionary.
[HYBRIS] Legacy attributes mapped to L4 categories that require modernization. If under [HYBRIS] there are L4 categories written in the format 'All L4 categories under L3 category x' where x is the name of the L3 category, then all L4 categories under the L3 category will have the given set of attributes in addition to attributes that already exist for those L4 categories.
[L4_CATLIST_SET] Set of L4 categories that need to be processed
[PL_ATTRIBUTELIST_V3] for attribute completeness check
[PRODUCT_LIST] A file containing L4 category and the products within them
[ATTRIBUTE_DICTIONARY] The master attribute set containing all attributes including attributes in [DICTIONARY_SUBSET]

Before starting, verify that all required input blocks are present.
If [PRODUCT_LIST] and [ATTRIBUTE_DICTIONARY] are missing, use the latest of these files available.
If [L4_CATLIST_SET] is absent use the most recent [L4_CATLIST]

Machine-Readable Formatting Rules (CRITICAL)
Wrap each step's output in tags: [STEP_X] ... [/STEP_X]
Tables must use the | pipe format.
No single table cell may exceed 1,000 characters. If content is longer, summarize it.
No preamble or conversational filler. Start and end only with the tags.
Use "-na-" as a marker for an empty cell unless stated otherwise.

[STEP_1]
Build the modern base attribute set per L4 category.
For each category in [L4_CATLIST_SET], processed one category at a time and grouped by the parent L3 category in [L3_CATLIST] determined by [CATEGORY_PATH].
Use product names from [PRODUCT_LIST] for each category in [L4_CATLIST_SET] to choose relevant attributes from [NEWADDITIONS_DICTIONARY] and [DICTIONARY_SUBSET].
Start from the Hybris attributes listed for the category (if any).
For each Hybris attribute:
Map it to the closest semantic match in: [NEWADDITIONS_DICTIONARY] (first priority), [DICTIONARY_SUBSET] (second priority) and [ATTRIBUTE_DICTIONARY] (third priority if no matches are found in [NEWADDITIONS_DICTIONARY] or [DICTIONARY_SUBSET])
If no suitable modern attribute exists, flag it as "No modern equivalent".
After mapping Hybris attributes:
Add any relevant attributes from [NEWADDITIONS_DICTIONARY] that are important for this category but were not triggered by Hybris.
Then add any relevant attributes from [DICTIONARY_SUBSET] not already included.
If a category has no Hybris attributes, generate a complete attribute set using [NEWADDITIONS_DICTIONARY] and [DICTIONARY_SUBSET].
Flag clearly if any new attribute would still be required.
Do not remove any Hybris attribute. Use relevant Attribute_Status to mark them appropriately.
If the L4 category in [HYBRIS] is 'L3' or 'L2' use these attributes at their respective levels where necessary.
Do not summarize any categories. Do Step 1 in full for all categories listed under [L4_CATLIST_SET].
Exclude these attributes unless they already exist in [HYBRIS] for this category: Colour, Item Weight, Dimensions Breadth, Dimensions Height, Dimensions Depth, Warranty Period. If present in Hybris they must be mapped with the appropriate Attribute Change Status. If not present in Hybris do not add them as new attributes — they exist on the PDP by default and are out of scope.

For each category provide number of products from [PRODUCT_LIST] then produce a single combined table [L4_SET_2] with one row per attribute covering all categories.

Table structure:
| L3 Category | L4 Category | Hybris Attribute Swedish | Hybris Attribute English | Final Modern Attribute | Final Modern Attribute translated to Swedish | Attribute Change Status | Attribute Importance | Data_Exists | Sample_Values | Customer_Reason | Customer_Importance | Attribute_Source | Products |

Rules for [L4_SET_2]:
Final Modern Attribute must come from [NEWADDITIONS_DICTIONARY], [DICTIONARY_SUBSET] or [ATTRIBUTE_DICTIONARY]. If none exist, use [NEW_ATTRIBUTE: suggested name].
Do not dedupe.
At the end of the table report: number of duplicates and row numbers where they occur.
Output all categories as a single continuous [L4_SET_2] table. Do not insert any text, headers, separators, or blank rows between category groups. The table must start immediately after the [L4_SET_2] marker with no preamble.

Attribute Change Status values:
Keep: hybris attribute is a case-insensitive exact match of the modern attribute
Replace: hybris attribute is replaced by a newer attribute
Remove: irrelevant for this category or does not apply to modern products released less than 8 years ago
Add New: newly added attribute not from Hybris

Attribute Importance values:
Mandatory: customer importance >= important
Optional: customer importance = low priority
Remove: if Attribute Change Status is Remove

Customer_Importance values:
critical - feature
critical - compliance
important for customer
low priority
irrelevant

Data_Exists values:
Hybris: attribute source is Hybris
Unknown: newly added attribute

Products: list 2-5 product names from [PRODUCT_LIST] where this exact attribute applies.
Customer_Reason: one detailed line explaining why this attribute matters for purchase decisions.
Sample_Values: at least 5 values unless Boolean.

Print [L4_SET_2].
[/STEP_1]

[STEP_2]
Validate completeness and mark attribute intent.
For each category produce a table comparing:
- The attribute list from [PL_ATTRIBUTELIST_V3]
- The final modern attributes from [L4_SET_2]
- What attributes in [PL_ATTRIBUTELIST_V3] are missing from [L4_SET_2] and why
- What attributes were added in [L4_SET_2] vs [PL_ATTRIBUTELIST_V3] and why

Table structure:
| L4 Category | Attribute | In PL_ATTRIBUTELIST_V3 | In L4_SET_2 | Status | Reason |

Status values: Present in both / Added in Phase 2 / Missing from Phase 2

Print [STEP_2] output table.
[/STEP_2]

[STEP_3]
Dictionary gap report.
If [L4_SET_2] contains any Final Modern Attribute marked as [NEW_ATTRIBUTE: ...], create table [DICT_ADD_SUGGESTIONS].

Table structure:
| L4 Category | Attribute Name | Description | Datatype | Dropdown Name | Unit | Mandate Source | Localization Needed | Sample Values |

If no new attributes exist output exactly: No dictionary gaps found.
Print [DICT_ADD_SUGGESTIONS].
[/STEP_3]

[STEP_3_1]
Dropdown governance.
For attributes in [DICT_ADD_SUGGESTIONS] with Datatype = Dropdown or Dropdown (Multi), provide:

Table structure:
| Dropdown Name | Exhaustive list of allowed values |

If no dropdowns exist output exactly: No dropdown governance needed.
Print [STEP_3_1] output table.
[/STEP_3_1]

[STEP_4]
If there is an attribute named Intended For within the categories in [L4_CATLIST_SET], generate:

Table structure:
| Category | List of Target Audiences/Uses |

If Intended For does not apply output exactly: Intended For not applicable for this batch.
Print [STEP_4] output table.
[/STEP_4]

Hard constraints
Do not rename attributes from [NEWADDITIONS_DICTIONARY] or [DICTIONARY_SUBSET]
Do not invent alternatives unless flagged as NEW_ATTRIBUTE_REQUIRED
Prefer reuse over creation
Legacy (Hybris) attributes do not dictate final naming

[L4_CATLIST_SET]
${l4List}

[L3_CATLIST]
${l3List}

[CATEGORY_PATH]
${paths}

[HYBRIS]
${hybrisFormatted}

[DICTIONARY_SUBSET]
${dictSubsetFormatted}

[NEWADDITIONS_DICTIONARY]
${newAdditionsFormatted}

[PL_ATTRIBUTELIST_V3]
${plV2Formatted}

[PRODUCT_LIST]
${productsFormatted}`;

  return prompt;
}

// ============================================================
// WRITE PHASE 2 OUTPUT
// ============================================================

function writePhase2Output(output, batch) {
  const sheet = getSheet(SHEETS.PHASE_2_WORKING);
  sheet.clear();
  sheet.clearFormats();

  // Diagnostic — log all section markers found
  const markers = ['[L4_SET_2]', 'Step 2', '[DICT_ADD_SUGGESTIONS]', 'Step 3.1', 'Step 4'];
  markers.forEach(m => {
    //log(`Marker "${m}" found: ${output.indexOf(m) > -1} at index ${output.indexOf(m)}`);
  });

  const batchInfo = batch.map(r => `${r['L3 Category']} > ${r['L4 Category']}`).join(', ');
  let currentRow = 1;

function writeSectionHeader(title) {
    //log(`Writing section header at row ${currentRow}: ${title.substring(0, 50)}`);
    sheet.getRange(currentRow, 1).setValue(title);
    sheet.getRange(currentRow, 1).setFontWeight('bold').setFontSize(10);
    currentRow += 2;
  }

function writeTextBlock(text) {
    currentRow = writeTextBlockToSheet(sheet, currentRow, text);
  }

function writeTable(headers, rows, tabName) {
    if (!headers || !rows || headers.length === 0) {
      sheet.getRange(currentRow, 1).setValue('⚠️ No data found for this section');
      sheet.getRange(currentRow, 1).setFontSize(9).setFontWeight('normal');
      currentRow += 3;
      return;
    }
    const safeHeaders = headers.map(h => String(h).substring(0, 1000));
    sheet.getRange(currentRow, 1, 1, safeHeaders.length).setValues([safeHeaders]);
    sheet.getRange(currentRow, 1, 1, safeHeaders.length)
      .setFontWeight('bold')
      .setFontSize(9);
    currentRow++;
    if (rows && rows.length > 0) {
      const maxCellChars = 45000;
      const data = rows.map(r => safeHeaders.map(h => {
        const val = (r[h] || '').toString();
        return val.length > maxCellChars ? val.substring(0, maxCellChars) + '...[truncated]' : val;
      }));
      sheet.getRange(currentRow, 1, data.length, safeHeaders.length).setValues(data);
      sheet.getRange(currentRow, 1, data.length, safeHeaders.length)
        .setFontWeight('normal')
        .setFontSize(9);
      currentRow += data.length;
    }
if (tabName) {
  try {
    const destSheet = getSheet(tabName);
    if (tabName === SHEETS.PL_ATTRIBUTELIST || tabName === SHEETS.MASTER_OUTPUT || tabName === SHEETS.INTENDED_FOR || tabName === SHEETS.PHASE2_DICT_ENTRIES || tabName === SHEETS.PHASE2_DICT_DROPDOWNS) {
      // Append — preserve data from previous sub-batches
      if (destSheet.getLastRow() === 0) {
        destSheet.getRange(1, 1, 1, safeHeaders.length).setValues([safeHeaders]);
      }
      const data = rows.map(r => safeHeaders.map(h => r[h] || ''));
      destSheet.getRange(destSheet.getLastRow() + 1, 1, data.length, safeHeaders.length).setValues(data);
    } else {
      // Overwrite for working tabs
      destSheet.clearContents();
      const allData = [safeHeaders, ...rows.map(r => safeHeaders.map(h => r[h] || ''))];
      destSheet.getRange(1, 1, allData.length, safeHeaders.length).setValues(allData);
    }
    safeHeaders.forEach((_, i) => destSheet.autoResizeColumn(i + 1));
  } catch(e) {
    log(`Could not write to tab ${tabName}: ${e.message}`);
  }
}
    currentRow += 3;
  }

  // ── HEADER ───────────────────────────────────────────────
  writeSectionHeader(`PHASE 2 OUTPUT — ${batchInfo} — Generated: ${new Date().toLocaleString()}`);

  // ── STEP 1: L4_SET_2 ─────────────────────────────────────
  writeSectionHeader('STEP 1 — L4_SET_2 (Final Attribute Sets per Category)');
  const step1Content = extractStepContent(output, 'STEP_1');
  if (step1Content) {
    const l4Section = extractSection(step1Content, '[L4_SET_2]', []);
    
    // If no [L4_SET_2] marker, find the first pipe table in the content
    let parsed = null;
    if (l4Section) {
      parsed = parseMarkdownTable(l4Section);
    } else {
      // Extract just the table portion — find first | line and stop at first non-pipe line
      const lines = step1Content.split('\n');
      const tableStart = lines.findIndex(l => l.trim().startsWith('|'));
      if (tableStart > -1) {
        let tableEnd = lines.length;
        for (let i = tableStart + 2; i < lines.length; i++) {
          const trimmed = lines[i].trim();
          // Skip blank lines, --- separators, and **bold** L3 headers between sub-tables
          if (trimmed.length === 0 || trimmed.startsWith('---') || trimmed.startsWith('**')) continue;
          if (!trimmed.startsWith('|')) {
            tableEnd = i;
            break;
          }
        }
        const tableContent = lines.slice(tableStart, tableEnd).join('\n');
        parsed = parseMarkdownTable(tableContent);
      }
    }
    
    if (parsed && parsed.headers && parsed.rows && parsed.rows.length > 0) {
      writeTable(parsed.headers, parsed.rows, SHEETS.MASTER_OUTPUT);
      log(`✅ Final Attribute Sets L4_SET_2 written to MASTER_OUTPUT — ${parsed.rows.length} rows`);
    } else {
      writeTextBlock(step1Content.substring(0, 2000));
      log(`⚠️ Could not parse L4_SET_2 table from Step 1`);
    }
  }
  

  // ── STEP 2: Completeness validation ──────────────────────
  writeSectionHeader('STEP 2 — Completeness Validation');
  const step2Content = extractStepContent(output, 'STEP_2');
  if (step2Content) {
    const parsed = parseMarkdownTable(step2Content);
    if (parsed && parsed.headers && parsed.rows && parsed.rows.length > 0) {
      writeTable(parsed.headers, parsed.rows, null);
    } else {
      writeTextBlock(step2Content.substring(0, 2000));
    }
  }

  // ── STEP 3: Dictionary gap report ────────────────────────
  writeSectionHeader('STEP 3 — DICT_ADD_SUGGESTIONS (Dictionary Gaps)');
  const step3Content = extractStepContent(output, 'STEP_3');
  if (step3Content) {
    const dictSection = extractSection(step3Content, '[DICT_ADD_SUGGESTIONS]', []);
    const parsed = dictSection ? parseMarkdownTable(dictSection) : parseMarkdownTable(step3Content);
    if (parsed && parsed.headers && parsed.rows && parsed.rows.length > 0) {
      writeTable(parsed.headers, parsed.rows, SHEETS.PHASE2_DICT_ENTRIES);
      log(`✅ DICT_ADD_SUGGESTIONS appended — ${parsed.rows.length} new attributes`);
    } else {
      writeTextBlock('✅ No dictionary gaps found for this batch.');
    }
  }

  // ── STEP 3.1: Dropdown governance ────────────────────────
  writeSectionHeader('STEP 3.1 — Dropdown Governance');
  const step31Content = extractStepContent(output, 'STEP_3_1');
  if (step31Content) {
    const parsed = parseMarkdownTable(step31Content);
    if (parsed && parsed.headers && parsed.rows && parsed.rows.length > 0) {
      writeTable(parsed.headers, parsed.rows, SHEETS.PHASE2_DICT_DROPDOWNS);
      log(`✅ Dropdown governance appended — ${parsed.rows.length} dropdowns`);
    } else {
      writeTextBlock('✅ No dropdown governance needed for this batch.');
    }
  }

  // ── STEP 4: Intended For ──────────────────────────────────
  writeSectionHeader('STEP 4 — Intended For (Target Audiences)');
  const step4Content = extractStepContent(output, 'STEP_4');
  if (step4Content) {
    const parsed = parseMarkdownTable(step4Content);
    if (parsed && parsed.headers && parsed.rows && parsed.rows.length > 0) {
      writeTable(parsed.headers, parsed.rows, SHEETS.INTENDED_FOR);
    } else {
      writeTextBlock(step4Content.substring(0, 1000));
    }
  }

  for (let i = 1; i <= 14; i++) {
    sheet.autoResizeColumn(i);
  }

  log(`✅ PHASE_2_WORKING written — ${currentRow} rows used`);
}


// ============================================================
// MAIN PHASE 2 RUNNER
// ============================================================

function runPhase2() {
  try {
    log('Starting Phase 2...');
    saveState({
      phase: 'Phase2',
      status: 'running',
      startTime: new Date().toISOString()
    });

    // Clear output tabs
    [SHEETS.PHASE_2_WORKING].forEach(tabName => {
      try {
        const sheet = getSheet(tabName);
        sheet.clear();
        sheet.clearFormats();
        log(`✅ Cleared tab: ${tabName}`);
      } catch(e) {
        log(`⚠️ Could not clear tab ${tabName}: ${e.message}`);
      }
    });

    // Read inputs
    log('Reading batch...');
    const batch = getReadyBatch();
    log(`Batch: ${batch.map(r => r['L4 Category']).join(', ')}`);

    log('Filtering products...');
    const products = getFilteredProducts(batch);
    log(`${products.length} products found`);

    log('Filtering Hybris...');
    const hybris = getFilteredHybris(batch);
    log(`${hybris.length} Hybris rows found`);

    log('Loading dictionary...');
    const dictionary = getDictionary();
    log(`${dictionary.length} dictionary attributes loaded`);

    log('Loading category info...');
    const categoryRows = getCategoryInfo(batch);

    // Read NEWADDITIONS_APPROVED — 3 columns used for prompt
    //log('Loading NEWADDITIONS_APPROVED...');
    const { rows: newAdditions } = getSheetData(SHEETS.NEWADDITIONS_APPROVED);
    if (newAdditions.length === 0) {
      throw new Error('NEWADDITIONS_APPROVED is empty — complete your review and paste approved rows first');
    }
    log(`✅ NEWADDITIONS_APPROVED: ${newAdditions.length} rows`);

    // Read DICTIONARY_SUBSET — manually updated by user
    //log('Loading DICTIONARY_SUBSET...');
    const { rows: dictSubset } = getSheetData(SHEETS.DICTIONARY_SUBSET);
    log(`✅ DICTIONARY_SUBSET: ${dictSubset.length} rows`);

    // Read PL_ATTRIBUTELIST (V2 after Phase 1C)
    //log('Loading PL_ATTRIBUTELIST_V2...');
    const { rows: plV2 } = getSheetData(SHEETS.PL_ATTRIBUTELIST);
    log(`✅ PL_ATTRIBUTELIST_V2: ${plV2.length} rows`);

    // Assemble prompt
    log('Assembling Phase 2 prompt...');
    const prompt = assemblePhase2Prompt(
      batch, products, hybris, dictionary,
      categoryRows, newAdditions, dictSubset, plV2
    );
    log(`Prompt assembled: ${prompt.length.toLocaleString()} characters`);
    const productCSV = formatProductListAsCSV(products);
    savePromptForManualRun(prompt, 'Phase 2', null, productCSV);

    // Call Claude API
    const dictCSV = formatDictionaryAsCSV(dictionary);
    const documents = [
      { title: 'ATTRIBUTE_DICTIONARY', content: dictCSV }
    ];
    log(`Dictionary attachment: ${dictCSV.length.toLocaleString()} characters`);
    const output = callClaudeAPI(prompt, documents);
    log(`Response received: ${output.length.toLocaleString()} characters`);

    if (!output) {
  throw new Error('Claude API returned empty output — check RAW_OUTPUT_DEBUG tab');
}
log(`Output length before writing: ${output.length}`);

    // Save raw output for debugging
    saveRawOutput(output);

    // Write output
    log('Writing Phase 2 output...');
    
    writePhase2Output(output, batch);

    clearState();

    sendNotification(
      `✅ PIM Phase 2 Complete — ${batch.map(r => r['L4 Category']).join(', ')}`,
      `Phase 2 has completed successfully.\n\nBatch: ${batch.map(r => r['L3 Category'] + ' > ' + r['L4 Category']).join('\n')}\n\nNext steps:\n1. Review PHASE_2_WORKING tab\n2. Check MASTER_OUTPUT for final L4_SET_2 attribute sets\n3. Run Phase 3 when ready`
    );

    log('✅ Phase 2 complete. Check PHASE_2_WORKING and MASTER_OUTPUT tabs.');

  } catch(e) {
    saveState({
      phase: 'Phase2',
      status: 'error',
      error: e.message,
      time: new Date().toISOString()
    });
    sendNotification(
      `❌ PIM Phase 2 Error`,
      `Phase 2 encountered an error:\n\n${e.message}`
    );
    log(`❌ Error: ${e.message}`);
    throw e;
  }
}

// ============================================================
// SHEET UI — MENU + CONTROLS
// ============================================================


function createControlsTab() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Delete existing CONTROLS tab if present
  const existing = ss.getSheetByName('CONTROLS');
  if (existing) ss.deleteSheet(existing);
  
  // Create fresh CONTROLS tab and move to front
  const sheet = ss.insertSheet('CONTROLS', 0);
  
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 280);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 50);

  // Title
  sheet.getRange('B1').setValue('PIM Automation Controls');
  sheet.getRange('B1').setFontSize(16).setFontWeight('bold');
  sheet.getRange('B2').setValue('Tap a button to run a phase. Check execution log for progress.');
  sheet.getRange('B2').setFontSize(11).setFontColor('#666666');

  const buttons = [
    { row: 4, label: '▶  Run Phase 1AB', fn: 'runPhase1AB', desc: 'Generate attributes for current BATCH', color: '#1a73e8', textColor: '#ffffff' },
    { row: 6, label: '▶  Run Post-Review Pipeline', fn: 'runPostReviewPipeline', desc: 'Validate → Process review → Phase 1C → Phase 2', color: '#188038', textColor: '#ffffff' },
    { row: 8, label: '▶  Generate Prompt Only', fn: 'generatePromptOnly', desc: 'Generate Review Prompt', color: '#e37400', textColor: '#ffffff' },
    { row: 10, label: '🗑  Clear Deliverable Tabs', fn: 'clearDeliverableTabs', desc: 'Clear DICT_READY + MASTER_OUTPUT', color: '#c5221f', textColor: '#ffffff' },
  ];

  buttons.forEach(b => {
    // Button cell
    const btnCell = sheet.getRange(b.row, 2);
    btnCell.setValue(b.label);
    btnCell.setBackground(b.color);
    btnCell.setFontColor(b.textColor);
    btnCell.setFontSize(13);
    btnCell.setFontWeight('bold');
    btnCell.setHorizontalAlignment('center');
    btnCell.setVerticalAlignment('middle');
    sheet.setRowHeight(b.row, 44);

    // Description cell
    const descCell = sheet.getRange(b.row, 3);
    descCell.setValue(b.desc);
    descCell.setFontSize(11);
    descCell.setFontColor('#444444');
    descCell.setVerticalAlignment('middle');

    // Assign script to button via drawing — done separately below
  });

  // Status row
  sheet.getRange('B12').setValue('Status');
  sheet.getRange('B12').setFontWeight('bold').setFontSize(11);
  sheet.getRange('B13').setValue('Check Apps Script execution log for detailed progress.');
  sheet.getRange('B13').setFontSize(11).setFontColor('#888888');

  SpreadsheetApp.getUi().alert('CONTROLS tab created. Use the PIM Automation menu to run phases, or assign scripts to drawings manually for tap-to-run buttons.');
}

/*
function diagnosSheet() {
  const sheet = getSheet(SHEETS.PHASE_1AB_WORKING);
  log(`Max rows: ${sheet.getMaxRows()}`);
  log(`Max cols: ${sheet.getMaxColumns()}`);
  log(`Last row: ${sheet.getLastRow()}`);
  log(`Last col: ${sheet.getLastColumn()}`);
  
  // Check first cell
  const cell = sheet.getRange(1, 1);
  log(`Cell A1 value length: ${(cell.getValue() || '').toString().length}`);
  log(`Cell A1 value: ${(cell.getValue() || '').toString().substring(0, 100)}`);
  
  // Check if any cells have very long content
  const data = sheet.getDataRange().getValues();
  data.forEach((row, ri) => {
    row.forEach((cell, ci) => {
      if ((cell || '').toString().length > 1000) {
        log(`Long cell at row ${ri+1}, col ${ci+1}: ${(cell || '').toString().length} chars`);
      }
    });
  });
}
*/

// ============================================================
// JUDGE ARBITRATION MODULE
// Parses NotebookLM + Gemini card outputs, auto-fills REVIEW
// where panel is unanimous, arbitrates disagreements via one
// Claude API call. Replaces the manual debate loop.
// ============================================================

// ── STEP A: READ + PARSE JUDGE TABS ─────────────────────────

function readJudgeTab(sheetName) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return '';
  const data = sheet.getRange(1, 1, lastRow, 1).getValues();
  return data.map(r => (r[0] || '').toString()).filter(Boolean).join('\n');
}

function normaliseAction(raw) {
  const a = (raw || '').toLowerCase();
  if (a.includes('merge')) return 'merge';
  if (a.includes('rename')) return 'rename';
  if (a.includes('delete') || a.includes('remove')) return 'delete';
  if (a.includes('keep')) return 'keep';
  return '';
}

function parseJudgeCards(text) {
  // Splits judge output on the 🔸 card marker your prompts enforce.
  // Returns { attrNameLower: { name, action, target, rationale } }
  const cards = {};
  if (!text) return cards;
  const blocks = text.split('🔸').slice(1);
  blocks.forEach(block => {
    const lines = block.split('\n').map(l => l.trim()).filter(Boolean);
    if (lines.length === 0) return;
    // First line = attribute name; strip brackets, bold markers, trailing colons
    const name = lines[0].replace(/[\[\]*_:]/g, '').trim();
    if (!name) return;

    let action = '', target = '', rationale = '';
    lines.forEach(l => {
      const lower = l.toLowerCase();
      if (lower.startsWith('action')) {
        // Format: Action: Keep / Merge / Rename ➡️ Target Name
        const afterColon = l.substring(l.indexOf(':') + 1);
        const parts = afterColon.split('➡️');
        action = normaliseAction(parts[0]);
        if (parts.length > 1) {
          target = parts[1].replace(/[\[\]*_]/g, '').trim();
        }
      }
      if (lower.startsWith('the rationale') || lower.startsWith('rationale')) {
        rationale = l.substring(l.indexOf(':') + 1).trim();
      }
    });

    // Keep with no explicit target = keep own name
    if (action === 'keep' && !target) target = name;

    cards[name.toLowerCase()] = { name, action, target, rationale };
  });
  return cards;
}

// ── STEP B: MAIN RUNNER ─────────────────────────────────────

function runJudgeArbitration() {
  try {
    saveState({ phase: 'Arbitration', status: 'running', startTime: new Date().toISOString() });
    log('Starting judge arbitration...');

    // Read judges
    const nlmText = readJudgeTab(SHEETS.JUDGE_NOTEBOOKLM);
    const gemText = readJudgeTab(SHEETS.JUDGE_GEMINI);
    if (!nlmText || !gemText) {
      log('⛔ One or both judge tabs are empty — paste NotebookLM output into JUDGE_NOTEBOOKLM and Gemini output into JUDGE_GEMINI first');
      clearState();
      return;
    }
    const nlmCards = parseJudgeCards(nlmText);
    const gemCards = parseJudgeCards(gemText);
    log(`Parsed judges — NotebookLM: ${Object.keys(nlmCards).length} cards | Gemini: ${Object.keys(gemCards).length} cards`);

    // Read REVIEW tab directly (don't use validateReview — decisions aren't made yet)
    const reviewSheet = getSheet(SHEETS.REVIEW);
    const data = reviewSheet.getDataRange().getValues();
    if (data.length < 2) {
      log('⛔ REVIEW tab is empty — run Phase 1AB first');
      clearState();
      return;
    }
    // Make sure the three explanation columns exist so we can store the full
    // NotebookLM / Gemini / Claude reasoning the reviewer needs to see.
    const headers = ensureReviewColumns(data, ['Panel Verdict', 'NLM Verdict', 'Gemini Verdict', 'Debate Log']);
    const col = name => headers.indexOf(name); // -1 if missing
    const cName = col('Attribute Name');
    const cRec = col('Recommendation');
    const cRecName = col('Recommended Name');
    const cDictMatch = col('Existing Dictionary Match');
    const cDecision = col('Your Decision');
    const cUpdated = col('Updated Attribute Name');
    const cVerdict = col('Panel Verdict');   // arbitration writes here — the Notes column is left for you
    const cNLM = col('NLM Verdict');
    const cGem = col('Gemini Verdict');
    const cDebate = col('Debate Log');
    if (cName === -1 || cDecision === -1) {
      throw new Error('REVIEW tab missing required columns (Attribute Name / Your Decision)');
    }

    let autoFilled = 0, skippedExisting = 0, missingJudge = 0;
    const disagreements = [];

    // Load the live dictionary once so each debated case can carry the real
    // target record(s) — datatype, option list, sample values — that the
    // merge-VALIDITY debate needs. (Existence of a match is NotebookLM's call.)
    const dictIndex = buildDictIndex();

    for (let i = 1; i < data.length; i++) {
      const name = (data[i][cName] || '').toString().trim();
      if (!name || name === 'Attribute Name' || name.startsWith('---')) continue;

      // Whether a decision is already committed for this row (by you or a prior
      // arbitration run). We never overwrite it — but we still attach the panel
      // trail below so the judging notes show regardless of save status.
      const alreadyDecided = (data[i][cDecision] || '').toString().trim() !== '';

      const key = name.toLowerCase();
      const j1 = nlmCards[key];
      const j2 = gemCards[key];
      const phase1Rec = normaliseAction(data[i][cRec]);
      const phase1Target = (cRecName > -1 ? data[i][cRecName] : '').toString().trim();

      if (!j1 || !j2) {
        missingJudge++;
        if (cVerdict > -1) data[i][cVerdict] = alreadyDecided
          ? '⚠️ Not found in one or both judge outputs'
          : '⚠️ Not found in one or both judge outputs — review manually';
        if (cNLM > -1) data[i][cNLM] = j1 ? fmtJudge(j1) : '(missing from NotebookLM)';
        if (cGem > -1) data[i][cGem] = j2 ? fmtJudge(j2) : '(missing from Gemini)';
        continue;
      }

      // Both judges weighed in — record their initial positions as the trail,
      // whether or not the row is already decided.
      if (cNLM > -1) data[i][cNLM] = fmtJudge(j1);
      if (cGem > -1) data[i][cGem] = fmtJudge(j2);

      // Unanimity test: both judges + Phase 1AB recommendation agree on action,
      // and for merge/rename the target names also agree (case-insensitive)
      const actionsAgree = j1.action && j1.action === j2.action && j1.action === phase1Rec;
      let targetsAgree = true;
      let finalTarget = '';
      if (actionsAgree && (j1.action === 'merge' || j1.action === 'rename')) {
        const t1 = j1.target.toLowerCase(), t2 = j2.target.toLowerCase();
        targetsAgree = t1 && t1 === t2;
        finalTarget = j1.target;
        // If Phase 1AB suggested a different target, flag it but trust the judges
        if (targetsAgree && phase1Target && phase1Target.toLowerCase() !== t1) {
          targetsAgree = false; // three-way target split → arbitrate
        }
      }

      // Already decided → leave YOUR decision untouched, but attach a Panel
      // Verdict comparing the panel's leaning to what you chose. This is the
      // cheap path: no Claude debate is run for rows you've already settled.
      if (alreadyDecided) {
        skippedExisting++;
        const yourDec = normaliseAction(data[i][cDecision]);
        if (cVerdict > -1) {
          if (actionsAgree && targetsAgree) {
            data[i][cVerdict] = yourDec === j1.action
              ? `✅ Panel unanimous (${cap(j1.action)}) — matches your decision`
              : `⚠️ Panel unanimous: ${cap(j1.action)}${finalTarget ? ' → ' + finalTarget : ''} — you chose ${cap(yourDec)}`;
          } else {
            data[i][cVerdict] = `⚖ Judges split — NLM: ${j1.action}${j1.target ? ' → ' + j1.target : ''} | Gemini: ${j2.action}${j2.target ? ' → ' + j2.target : ''} — you chose ${cap(yourDec)}`;
          }
        }
        continue;
      }

      if (actionsAgree && targetsAgree) {
        data[i][cDecision] = cap(j1.action);
        if (finalTarget && cUpdated > -1) data[i][cUpdated] = finalTarget;
        if (cVerdict > -1) data[i][cVerdict] = '✅ Panel unanimous';
        autoFilled++;
      } else {
        disagreements.push({
          rowIndex: i,
          name,
          dictMatches: (cDictMatch > -1 ? data[i][cDictMatch] : '').toString(),
          phase1: { action: phase1Rec, target: phase1Target },
          notebooklm: j1,
          gemini: j2,
          targetRecords: lookupDictRecords(dictIndex, [j1.target, j2.target, phase1Target, name], (cDictMatch > -1 ? data[i][cDictMatch] : ''))
        });
      }
    }

    log(`📊 Unanimous auto-filled: ${autoFilled} | Already decided: ${skippedExisting} | Missing from judges: ${missingJudge} | Disagreements: ${disagreements.length}`);

    // ── Debate the disagreements over up to DEBATE_MAX_ROUNDS rounds ──
    const stalemates = [];
    if (disagreements.length > 0) {
      const outcomes = debateDisagreements(disagreements);
      let resolved = 0, needsYou = 0;

      outcomes.forEach(o => {
        const d = disagreements.find(x => x.name.toLowerCase() === (o.attribute || '').toLowerCase());
        if (!d) return;
        const i = d.rowIndex;
        const verdict = normaliseAction(o.verdict);

        // Always record the full three-way reasoning so it shows up in Review
        if (cNLM > -1)    data[i][cNLM]    = fmtJudge(d.notebooklm);
        if (cGem > -1)    data[i][cGem]    = fmtJudge(d.gemini);
        if (cDebate > -1) data[i][cDebate] = (o.transcript || '').substring(0, 4000);

        if (verdict && o.converged && (o.confidence || '').toLowerCase() === 'high') {
          data[i][cDecision] = verdict.charAt(0).toUpperCase() + verdict.slice(1);
          if ((verdict === 'merge' || verdict === 'rename') && o.target && cUpdated > -1) {
            data[i][cUpdated] = o.target;
          }
          if (cVerdict > -1) data[i][cVerdict] = `🤖 Debated → ${verdict.charAt(0).toUpperCase() + verdict.slice(1)} (${o.rounds} rounds, high conf): ${o.rationale || ''}`.substring(0, 900);
          resolved++;
        } else {
          // Still split after the debate — leave decision blank, surface everything
          if (cVerdict > -1) {
            data[i][cVerdict] = `🔶 NEEDS YOU — after ${o.rounds} round(s) NLM: ${d.notebooklm.action}→${d.notebooklm.target || '—'} | Gemini: ${d.gemini.action}→${d.gemini.target || '—'} | arb leans ${verdict || '?'}`.substring(0, 900);
          }
          needsYou++;
          stalemates.push({ name: d.name, notebooklm: d.notebooklm, gemini: d.gemini, dictMatches: d.dictMatches });
        }
      });
      log(`⚖️ Debate: ${resolved} resolved (high confidence) | ${needsYou} left for your review`);
    }

    // Expand the sheet grid if we added columns, then write everything back in one shot
    if (headers.length > reviewSheet.getMaxColumns()) {
      reviewSheet.insertColumnsAfter(reviewSheet.getMaxColumns(), headers.length - reviewSheet.getMaxColumns());
    }
    reviewSheet.getRange(1, 1, data.length, headers.length).setValues(data);

    // Hybrid: anything still unresolved after the final round gets a
    // fresh-opinion prompt you can paste into a NEW NotebookLM / Gemini chat.
    if (stalemates.length > 0) {
      writeDebateReprompts(stalemates);
      log(`📝 ${stalemates.length} attribute(s) still split after ${DEBATE_MAX_ROUNDS} rounds — fresh-opinion prompts written to ${SHEETS.DEBATE_REPROMPT} (see Judges tab).`);
    } else {
      clearDebateReprompts();
    }

    log(`✅ Arbitration complete — REVIEW tab updated. The Panel Verdict column shows the outcome (🔶 = needs you); your Notes column is left untouched.`);
    clearState();

  } catch(e) {
    saveState({ phase: 'Arbitration', status: 'error', error: e.message, time: new Date().toISOString() });
    log(`❌ Arbitration error: ${e.message}`);
    throw e;
  }
}

// ── STEP C (DEPRECATED): single-shot arbitration ────────────
// Superseded by debateDisagreements() below (multi-round debate).
// Kept for reference only — no longer called by runJudgeArbitration.
function arbitrateDisagreements_DEPRECATED(disagreements) {
  const caseBlocks = disagreements.map(d => {
    return [
      `ATTRIBUTE: ${d.name}`,
      `Existing dictionary matches: ${d.dictMatches || 'None'}`,
      `Phase 1AB recommendation: ${d.phase1.action || 'n/a'}${d.phase1.target ? ' → ' + d.phase1.target : ''}`,
      `NotebookLM verdict: ${d.notebooklm.action}${d.notebooklm.target ? ' → ' + d.notebooklm.target : ''}`,
      `NotebookLM rationale: ${d.notebooklm.rationale || 'n/a'}`,
      `Gemini verdict: ${d.gemini.action}${d.gemini.target ? ' → ' + d.gemini.target : ''}`,
      `Gemini rationale: ${d.gemini.rationale || 'n/a'}`
    ].join('\n');
  }).join('\n\n---\n\n');

  const prompt = `You are the final arbiter for a taxonomy review panel for our Akeneo PIM. Two AI judges and an initial recommendation have disagreed on the attributes below. Your job: issue one final verdict per attribute, applying these rules strictly.

CRITICAL RULES (The Akeneo Reality Check):
1. Dropdown Pollution: Do not merge Simple Select or Multi-Select attributes unless the option lists are practically identical. Keep domain-specific lists isolated.
2. Facet Collision: Do not merge attributes if they would cause bizarre cross-shopping filter results on the website.
3. Metric Free-Pass: Merge Metric (number + unit) and Boolean (Yes/No) attributes aggressively, as long as it doesn't break Rule 4.
4. Meaning Shift: Do not merge if the word's actual physical definition changes depending on the product family.
5. Component-Specific Fitment vs. Whole-Product Identity: Do not merge an attribute into a generic global field if it represents a critical component dimension or fitment variable of a larger assembly.
6. Logistics vs. Engineering Boundary: Never merge functional fitment dimensions into generic fields that act as ERP/WMS shipping bounding boxes.
7. Global Search & Data Type Test: If merging creates a chaotic global filter of unrelated L2 categories, prefix to isolate. Never merge faceted Dropdowns into descriptive Text fields.
8. Naming Convention: Domain-specific attributes must be prefixed/suffixed to make their scope clear.
9. The Akeneo 1:1 Label Constraint: Never merge across product families if it forces a generic or incorrect storefront label.
10. The Customer Vocabulary Test: Reject merges that erase the specific industry term a tradesperson or hobbyist would search for.
11. Acceptable Controlled Bloat: A slightly redundant attribute is acceptable to protect the shopping experience — prefix it with its L2 category.

ARBITRATION PRINCIPLES:
- When judges split between Merge and Keep, default to Keep unless the merge is clearly safe under all rules above. Bad merges are expensive to undo post-go-live; controlled bloat is cheap.
- Mark confidence "high" only when one position clearly violates a numbered rule or one judge's rationale is factually wrong. Mark "low" when both positions are defensible — a human will decide those.
- Verdict must be one of: Keep, Rename, Merge, Delete.

OUTPUT FORMAT (CRITICAL):
Respond with ONLY a JSON array, no markdown fences, no preamble. One object per attribute:
[{"attribute": "...", "verdict": "Keep|Rename|Merge|Delete", "target": "target name or empty string", "confidence": "high|low", "rationale": "one sentence citing the rule number that decided it"}]

DISAGREEMENTS TO ARBITRATE:

${caseBlocks}`;

  const output = callClaudeAPI(prompt, null);
  const clean = output.replace(/```json|```/g, '').trim();
  try {
    const parsed = JSON.parse(clean);
    if (!Array.isArray(parsed)) throw new Error('not an array');
    return parsed;
  } catch(e) {
    log(`⚠️ Could not parse arbitration JSON: ${e.message}`);
    saveRawOutput(output); // dump to RAW_OUTPUT_DEBUG for inspection
    return [];
  }
}

// ============================================================
// MULTI-ROUND DEBATE MODULE
// NotebookLM vs Gemini, adjudicated by Claude across up to
// DEBATE_MAX_ROUNDS rounds. NotebookLM's merge recommendations
// carry higher priority (better at spotting real existing matches).
// Anything still split after the final round gets a fresh-opinion
// re-paste prompt (hybrid loop).
// ============================================================

const DEBATE_MAX_ROUNDS = 3;

function cap(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : s; }

// One judge's position as a compact, human-readable block for the Review card
function fmtJudge(j) {
  if (!j) return '';
  const head = `${cap(j.action) || '?'}${j.target ? ' → ' + j.target : ''}`;
  return j.rationale ? `${head}\n${j.rationale}` : head;
}

// Index the live dictionary by lowercased attribute name so debated cases can
// carry the real target record (datatype, options, samples) into the prompt.
function buildDictIndex() {
  const idx = {};
  try {
    getDictionary().forEach(r => {
      const n = (r['Attribute Name'] || '').toString().trim();
      if (n) idx[n.toLowerCase()] = r;
    });
  } catch(e) {
    log(`⚠️ Could not load dictionary for debate enrichment: ${e.message}`);
  }
  return idx;
}

// Resolve the dictionary record(s) for the target name(s) a case references.
// `names` = judge/phase1 targets + the attribute itself; dictMatchStr = the
// REVIEW "Existing Dictionary Match" cell (may list several, comma/;/|/ separated).
function lookupDictRecords(idx, names, dictMatchStr) {
  const candidates = [];
  (names || []).forEach(n => { const t = (n || '').toString().trim(); if (t) candidates.push(t); });
  (dictMatchStr || '').toString().split(/[,;|\/]+/).forEach(n => { const t = n.trim(); if (t) candidates.push(t); });
  const seen = {}, out = [];
  candidates.forEach(n => {
    const k = n.toLowerCase();
    if (seen[k]) return;
    seen[k] = true;
    if (idx[k]) out.push(idx[k]);
  });
  return out;
}

// Render one dictionary record as a single compact line for the debate prompt.
function fmtDictRecord(r) {
  const parts = [`• ${r['Attribute Name'] || ''}`];
  if (r['Datatype']) parts.push(`datatype: ${r['Datatype']}`);
  if (r['Dropdown Name']) parts.push(`dropdown: ${r['Dropdown Name']}`);
  if (r['Unit Of Measure Metric']) parts.push(`unit: ${r['Unit Of Measure Metric']}`);
  if (r['Sample Values']) parts.push(`samples: ${r['Sample Values']}`);
  if (r['Description']) parts.push(`desc: ${r['Description']}`);
  return parts.join(' | ');
}

// Add columns to the in-memory REVIEW matrix if missing, padding every row so
// the matrix stays rectangular for setValues(). Returns the (mutated) headers.
function ensureReviewColumns(data, names) {
  const headers = data[0];
  names.forEach(n => { if (headers.indexOf(n) === -1) headers.push(n); });
  const width = headers.length;
  for (let i = 0; i < data.length; i++) {
    while (data[i].length < width) data[i].push('');
  }
  return headers;
}

// Runs the panel debate. Each round Claude sees the running transcript and
// either converges (high confidence) or escalates to the next round.
// Returns one outcome per attribute.
function debateDisagreements(disagreements) {
  const transcripts = {};   // nameLower -> [round strings]
  const results = {};       // nameLower -> latest round result
  disagreements.forEach(d => { transcripts[d.name.toLowerCase()] = []; });

  let pending = disagreements.slice();
  for (let round = 1; round <= DEBATE_MAX_ROUNDS && pending.length > 0; round++) {
    log(`⚖️ Debate round ${round} — ${pending.length} unresolved...`);
    const roundResults = runDebateRound(pending, round, transcripts);
    const stillPending = [];
    pending.forEach(d => {
      const k = d.name.toLowerCase();
      const r = roundResults[k];
      if (!r) { stillPending.push(d); return; } // parse gap → retry next round
      transcripts[k].push(
        `── Round ${round} ──\n` +
        `NotebookLM argues: ${r.nlm_argument || '(none)'}\n` +
        `Gemini argues: ${r.gemini_argument || '(none)'}\n` +
        `Claude adjudicates → ${cap(normaliseAction(r.verdict)) || '?'}${r.target ? ' (' + r.target + ')' : ''} [${r.confidence || 'low'}]: ${r.rationale || ''}`
      );
      results[k] = r;
      const done = r.converged && (r.confidence || '').toLowerCase() === 'high';
      if (!done) stillPending.push(d);
    });
    pending = stillPending;
  }

  const unresolved = {};
  pending.forEach(d => { unresolved[d.name.toLowerCase()] = true; });

  return disagreements.map(d => {
    const k = d.name.toLowerCase();
    const r = results[k] || {};
    const resolved = !unresolved[k] && r.converged && (r.confidence || '').toLowerCase() === 'high';
    return {
      attribute: d.name,
      verdict: r.verdict || '',
      target: r.target || '',
      confidence: r.confidence || 'low',
      rationale: r.rationale || '',
      transcript: (transcripts[k] || []).join('\n\n'),
      rounds: (transcripts[k] || []).length,
      converged: !!resolved
    };
  });
}

// One Claude call adjudicating every still-unresolved case for a given round.
function runDebateRound(cases, round, transcripts) {
  const caseBlocks = cases.map(d => {
    const prior = (transcripts[d.name.toLowerCase()] || []).join('\n');
    return [
      `ATTRIBUTE: ${d.name}`,
      `Existing dictionary matches: ${d.dictMatches || 'None'}`,
      (d.targetRecords && d.targetRecords.length)
        ? `Target attribute record(s) from the live dictionary (use these for the validity rules):\n${d.targetRecords.map(fmtDictRecord).join('\n')}`
        : `Target attribute record(s): (not found in the dictionary export)`,
      `Phase 1AB recommendation: ${d.phase1.action || 'n/a'}${d.phase1.target ? ' → ' + d.phase1.target : ''}`,
      `NotebookLM position: ${d.notebooklm.action}${d.notebooklm.target ? ' → ' + d.notebooklm.target : ''} — ${d.notebooklm.rationale || 'n/a'}`,
      `Gemini position: ${d.gemini.action}${d.gemini.target ? ' → ' + d.gemini.target : ''} — ${d.gemini.rationale || 'n/a'}`,
      prior ? `Debate so far:\n${prior}` : `Debate so far: (this is round 1)`
    ].join('\n');
  }).join('\n\n---\n\n');

  const prompt = `You are moderating and adjudicating a taxonomy review debate for our Akeneo PIM. Two AI judges — NotebookLM and Gemini — disagree on the attributes below. This is round ${round} of up to ${DEBATE_MAX_ROUNDS}.

Your job each round, per attribute:
1. Construct NotebookLM's STRONGEST argument for its position (steelman it).
2. Construct Gemini's STRONGEST counter-argument (steelman it).
3. Adjudicate under the rules below: pick the verdict, and say whether the panel has CONVERGED (settled) or should continue to another round.

JUDGE AUTHORITY (IMPORTANT — split existence from validity):
- EXISTENCE OF THE MATCH: NotebookLM is almost never wrong about whether an existing attribute exists and which one is the match. Treat its identification of the existing target attribute as AUTHORITATIVE and settled — do NOT re-litigate whether that attribute exists or whether it is the right match candidate. If NotebookLM and Gemini disagree on whether a match exists at all, NotebookLM wins outright.
- VALIDITY OF THE MERGE: Whether the attribute should actually be MERGED into that confirmed target (versus kept separate or prefixed) is a genuine question. Debate that on the merits with Gemini under the Akeneo rules below. A real existing match can still be a bad merge. Do NOT auto-converge to Merge just because NotebookLM proposed it — the merge must survive Rules 1-11.
- Use the "Target attribute record(s) from the live dictionary" provided per attribute (datatype, dropdown/option list, unit, sample values, description) to apply the validity rules concretely — e.g. do not merge Select attributes whose option lists differ (Rule 1), do not merge across incompatible datatypes or faceted dropdowns into text fields (Rule 7), and check for meaning shift against the target's description (Rule 4). If no record was found, say the target could not be verified in the dictionary export and lean conservative.

CRITICAL RULES (The Akeneo Reality Check):
1. Dropdown Pollution: Do not merge Simple Select or Multi-Select attributes unless the option lists are practically identical. Keep domain-specific lists isolated.
2. Facet Collision: Do not merge attributes if they would cause bizarre cross-shopping filter results on the website.
3. Metric Free-Pass: Merge Metric (number + unit) and Boolean (Yes/No) attributes aggressively, as long as it doesn't break Rule 4.
4. Meaning Shift: Do not merge if the word's actual physical definition changes depending on the product family.
5. Component-Specific Fitment vs. Whole-Product Identity: Do not merge an attribute into a generic global field if it represents a critical component dimension or fitment variable of a larger assembly.
6. Logistics vs. Engineering Boundary: Never merge functional fitment dimensions into generic fields that act as ERP/WMS shipping bounding boxes.
7. Global Search & Data Type Test: If merging creates a chaotic global filter of unrelated L2 categories, prefix to isolate. Never merge faceted Dropdowns into descriptive Text fields.
8. Naming Convention: Domain-specific attributes must be prefixed/suffixed to make their scope clear.
9. The Akeneo 1:1 Label Constraint: Never merge across product families if it forces a generic or incorrect storefront label.
10. The Customer Vocabulary Test: Reject merges that erase the specific industry term a tradesperson or hobbyist would search for.
11. Acceptable Controlled Bloat: A slightly redundant attribute is acceptable to protect the shopping experience — prefix it with its L2 category.

CONVERGENCE RULES:
- Set "converged": true and "confidence": "high" only when the debate is genuinely settled — one position clearly wins under the rules. When NotebookLM has identified a real existing match, take that match's EXISTENCE as given and let convergence turn entirely on whether the merge is VALID under Rules 1-11.
- Set "converged": false when both positions are still defensible after this round's exchange; the debate will run again next round with your arguments as input.
- On the final round (${DEBATE_MAX_ROUNDS}), still give your best verdict, but keep "converged": false / "confidence": "low" if the panel truly cannot agree — a human will decide it.
- Verdict must be one of: Keep, Rename, Merge, Delete.

OUTPUT FORMAT (CRITICAL):
Respond with ONLY a JSON array, no markdown fences, no preamble. One object per attribute:
[{"attribute":"...","nlm_argument":"NotebookLM's strongest point this round","gemini_argument":"Gemini's strongest counter this round","verdict":"Keep|Rename|Merge|Delete","target":"target name or empty string","converged":true|false,"confidence":"high|low","rationale":"one sentence citing the rule number or the merge-authority that decided it"}]

ATTRIBUTES:

${caseBlocks}`;

  const output = callClaudeAPI(prompt, null);
  const clean = output.replace(/```json|```/g, '').trim();
  const map = {};
  try {
    const parsed = JSON.parse(clean);
    if (!Array.isArray(parsed)) throw new Error('not an array');
    parsed.forEach(v => { if (v && v.attribute) map[v.attribute.toLowerCase()] = v; });
  } catch(e) {
    log(`⚠️ Round ${round}: could not parse debate JSON: ${e.message}`);
    saveRawOutput(output); // dump to RAW_OUTPUT_DEBUG for inspection
  }
  return map;
}

// Hybrid step: for attributes the panel could not settle, write a ready-to-paste
// prompt you can drop into a FRESH NotebookLM / Gemini chat for a tie-breaker.
function writeDebateReprompts(stalemates) {
  const sheet = getOrCreateSheet(SHEETS.DEBATE_REPROMPT);
  sheet.clearContents();
  const rows = [['Attribute', 'Fresh-opinion prompt']];
  stalemates.forEach(s => {
    const prompt =
      `We are reviewing the PIM attribute "${s.name}" and our review panel is split.\n` +
      `Existing dictionary matches: ${s.dictMatches || 'None'}\n` +
      `Position A (NotebookLM): ${cap(s.notebooklm.action)}${s.notebooklm.target ? ' → ' + s.notebooklm.target : ''} — ${s.notebooklm.rationale || 'n/a'}\n` +
      `Position B (Gemini): ${cap(s.gemini.action)}${s.gemini.target ? ' → ' + s.gemini.target : ''} — ${s.gemini.rationale || 'n/a'}\n\n` +
      `Give your independent verdict (Keep / Rename / Merge / Delete), the target attribute if merging or renaming, and a one-paragraph rationale. Do not just agree — tell us which position is stronger and why.`;
    rows.push([s.name, prompt]);
  });
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
}

function clearDebateReprompts() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.DEBATE_REPROMPT);
    if (sheet) sheet.clearContents();
  } catch(e) {}
}