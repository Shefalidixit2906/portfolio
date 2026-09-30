// Made by Sanidhay Kumar - https://techessentials.in

// ============================================================
//  MENU
// ============================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('BMP Formulas')
    .addItem('Setup (First Time)', 'setupWizard')
    .addSeparator()
    .addItem('Add First Step (creates Form)', 'addFirstStepWizard')
    .addItem('Add Step', 'addStepWizard')
    .addItem('Link Existing Form', 'linkFormWizard')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Apply Single Formula')
        .addItem('TAT (Working Hours)', 'applyTAT')
        .addItem('Days Before Future Date', 'applyTX')
        .addItem('Specific Clock Time', 'applySpecificTime')
        .addSeparator()
        .addItem('Show planned only when YES', 'applyShowYes')
        .addItem('Show planned only when NO', 'applyShowNo')
        .addSeparator()
        .addItem('Time Delay', 'applyTimeDelay')
        .addItem('Set Actual Time', 'setActualTime')
    )
    .addToUi();
}

// ============================================================
//  TIER 2 — SETUP WIZARD
// ============================================================

function setupWizard() {
  var html = HtmlService.createHtmlOutputFromFile('SetupDialog')
    .setWidth(440).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'BMP Formulas — Setup');
}

/**
 * Server callback invoked from SetupDialog.html.
 * payload: { open: 'HH:MM', close: 'HH:MM', workingDays: [Mon..Sun, true=working] }
 */
function setupWizardSave(payload) {
  var openFrac = parseTimeToFraction(payload.open);
  var closeFrac = parseTimeToFraction(payload.close);
  if (openFrac == null) throw new Error('Opening Time is not valid HH:MM (e.g. 10:00)');
  if (closeFrac == null) throw new Error('Closing Time is not valid HH:MM (e.g. 18:00)');
  if (closeFrac <= openFrac) throw new Error('Closing Time must be after Opening Time on the same day');
  if (!payload.workingDays || payload.workingDays.length !== 7) throw new Error('Working days array must have 7 entries (Mon..Sun)');

  // Build the WORKDAY.INTL bitstring: 0 = working day, 1 = off
  var bitstring = '';
  for (var i = 0; i < 7; i++) bitstring += payload.workingDays[i] ? '0' : '1';
  if (!/^[01]{7}$/.test(bitstring)) throw new Error('Working days bitstring is invalid');

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getActiveSheet();

  // FMS tab: only A1 = NOW() goes here. A1 is safe from column inserts
  // because column A is the leftmost column. Time Delay formulas
  // reference $A$1 for the live "now" value.
  sh.getRange('A1').setFormula('=NOW()');
  sh.getRange('B1').clearContent();
  sh.hideRows(1);

  // Config tab: holds opening time, closing time, working-days pattern.
  //
  // Why a separate tab instead of C1/D1/E1 of the FMS tab (the v1 / v2.1
  // through v2.1.7 approach): column inserts triggered by Add Step shift
  // any data in row 1 to the right. With form-data zones smaller than
  // 5 columns (e.g. 3 form fields → form ends at D), an Add Step that
  // inserts at column 5 displaces E1's working-days pattern to J1, K1,
  // etc. The TAT formula's hardcoded $E$1 reference then points at an
  // empty cell and WORKDAY.INTL returns #VALUE!. Putting config in a
  // separate tab makes its location immune to FMS column inserts.
  //
  // Layout in Config tab:
  //   A1  Opening time  (day fraction, formatted HH:mm)
  //   A2  Closing time  (day fraction, formatted HH:mm)
  //   A3  Working-days  (7-char 0/1 string, leading apostrophe forces text)
  //   B1-B3  Human-readable labels
  var configSheet = ss.getSheetByName('Config');
  if (!configSheet) {
    configSheet = ss.insertSheet('Config');
  }
  configSheet.getRange('A1').setValue(openFrac).setNumberFormat('HH:mm');
  configSheet.getRange('A2').setValue(closeFrac).setNumberFormat('HH:mm');
  configSheet.getRange('A3').setValue("'" + bitstring);
  configSheet.getRange('B1').setValue('Opening Time').setFontWeight('bold');
  configSheet.getRange('B2').setValue('Closing Time').setFontWeight('bold');
  configSheet.getRange('B3').setValue('Working Days (0=working, 1=off; Mon-Sun)').setFontWeight('bold');
  // Hide the Config tab so the user doesn't accidentally edit it
  configSheet.hideSheet();

  // Re-activate the FMS tab in case insertSheet/hideSheet shifted focus
  ss.setActiveSheet(sh);

  setIterativeCalc();

  // Tier 2 / Tier 4 #5: install the onChange trigger as part of setup so
  // the user does not need to remember to run "Set Actual Time" separately.
  ensureOnChangeTrigger();

  return {
    open: payload.open,
    close: payload.close,
    workingDays: bitstring,
    message: 'Setup complete. Add holidays manually in the Holidays tab as DD/MM/YYYY dates.'
  };
}

// ============================================================
//  TIER 2 — ADD FIRST STEP WIZARD (creates Form + first step)
// ============================================================

function addFirstStepWizard() {
  var html = HtmlService.createHtmlOutputFromFile('AddFirstStepDialog')
    .setWidth(620).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'BMP Formulas — Add First Step');
}

/**
 * Server callback invoked from AddFirstStepDialog.html.
 *
 * Creates a Google Form with the user's chosen fields, links it to this
 * spreadsheet (auto-creating a Form Responses tab), writes form-data
 * column headers + INDEX formulas in the FMS tab, and lays out the
 * *entry-event header band* (rows 2-5, merged across columns B..lastFormCol).
 *
 * The entry event is NOT a step block — it has no Planned/Actual/Status/
 * Time Delay columns. The form Timestamp in column A serves as its
 * Actual, and the When is always "Whenever Needed". This matches the
 * canonical Purchase FMS layout where Step 1 (Generate PO) is just a
 * header in B2:E2 above the form data, and Step 2 (Issue PO) is the
 * first step with a real timing formula.
 *
 * After this wizard completes, the user runs "Add Step" to add the
 * first step with a deadline (TAT / Days Before / Specific Time).
 *
 * payload: {
 *   processName,
 *   formFields: [{ label, type: 'text' | 'number' | 'date' }],
 *   stepName,    // e.g. "Generate PO", "Receive Order Inquiry"
 *   who, how
 * }
 */
function addFirstStepWizardSave(payload) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getActiveSheet();

  // Validate inputs
  if (!payload.processName || !payload.processName.trim()) throw new Error('Process name is required.');
  if (!payload.formFields || payload.formFields.length === 0) throw new Error('Add at least one form field (e.g. PO Number).');
  if (!payload.stepName || !payload.stepName.trim()) throw new Error('First step name is required (e.g. "Generate PO" or "Receive Order Inquiry").');

  // Validate sheet doesn't already have step blocks
  var lastCol = sh.getLastColumn();
  if (lastCol > 0) {
    var row6Existing = sh.getRange(6, 1, 1, lastCol).getValues()[0];
    var hasStep = row6Existing.some(function (h) {
      return /^(planned|actual|status|yes\/no|time delay)$/i.test(String(h || '').trim());
    });
    if (hasStep) {
      throw new Error('This sheet already has step blocks. Use "Add Step" to add more steps, not "Add First Step".');
    }
  }

  // Validate setup is done (Config tab must exist and be populated).
  // v2.1.8: was checking C1/D1/E1 of the FMS tab, which broke when
  // column inserts displaced E1's working-days pattern.
  var configSheet = ss.getSheetByName('Config');
  if (!configSheet ||
      !configSheet.getRange('A1').getValue() ||
      !configSheet.getRange('A2').getValue() ||
      !configSheet.getRange('A3').getValue()) {
    throw new Error('Run "Setup (First Time)" before adding the first step. Office hours and working days need to be configured.');
  }

  // 1. Snapshot existing tab names so we can find the new Form Responses tab afterward
  var beforeNames = ss.getSheets().map(function (s) { return s.getName(); });

  // 2. Create the Google Form
  var formInfo = createLinkedFormForFMS(payload.processName.trim(), payload.formFields, ss);

  // 3. Wait for the linked Form Responses tab to appear, then find it
  SpreadsheetApp.flush();
  Utilities.sleep(2500);
  var formTab = null;
  var afterSheets = ss.getSheets();
  for (var i = 0; i < afterSheets.length; i++) {
    if (beforeNames.indexOf(afterSheets[i].getName()) === -1) {
      formTab = afterSheets[i];
      break;
    }
  }
  if (!formTab) {
    throw new Error('Form was created but the linked Form Responses tab could not be found. Check the spreadsheet — the form may still appear shortly. You can then run "Link Existing Form" to wire it up.');
  }
  ss.setActiveSheet(sh);  // re-activate the FMS tab

  var formTabName = formTab.getName();
  var numFormCols = payload.formFields.length + 1;  // +1 for Timestamp

  // 4. Write column headers in row 6
  sh.getRange(6, 1).setValue('Timestamp').setFontWeight('bold');
  for (var fi = 0; fi < payload.formFields.length; fi++) {
    sh.getRange(6, fi + 2).setValue(payload.formFields[fi].label).setFontWeight('bold');
  }

  // 5. Write a single ARRAYFORMULA in A7 that pulls the entire form-data
  // zone from the linked Form Responses tab.
  //
  // The range is wrapped in INDIRECT() — this is critical. When a form
  // response arrives, Google Sheets inserts a row at the top of the
  // Form Responses tab, which auto-shifts absolute references like
  // `$A$2` to `$A$3` in any formula that references them. INDIRECT
  // takes a STRING reference instead of a direct one, so Sheets has
  // no way to "see" the row reference and doesn't adjust it.
  //
  // Final formula shape:
  //
  //   =ARRAYFORMULA(INDIRECT("'Form Responses 1'!A2:<lastFormColLetter>"))
  //
  // For a 4-question form: 4 questions + 1 timestamp = 5 columns,
  // lastFormColLetter = "E", formula:
  //
  //   =ARRAYFORMULA(INDIRECT("'Form Responses 1'!A2:E"))
  //
  // This is the same INDIRECT-wrapped pattern the FMS Builder Web App
  // monolith uses (line 1345 of FMS_Builder_Monolith.gs), simplified by
  // hardcoding the column letter from the form-field count instead of
  // dynamically computing it via COUNTA('...'!1:1).
  //
  // The single formula spills:
  //   - Rightward across all form-data columns (A through <lastCol>)
  //   - Downward as new responses arrive (no row 200 cap, no manual
  //     drag-down, no offset gymnastics, no auto-shifting on insert)
  //
  // Empty source rows correctly render as empty — ARRAYFORMULA's
  // default behaviour, no LET+IF wrap needed.
  var firstDataRow = 7;
  var safeTabName = formTabName.replace(/'/g, "''");
  var lastFormColLetter = colNumToLetter(numFormCols);
  var spillFormula = '=ARRAYFORMULA(INDIRECT("\'' + safeTabName +
                     '\'!A2:' + lastFormColLetter + '"))';
  sh.getRange(firstDataRow, 1).setFormula(spillFormula);

  // Format Timestamp (column A) as date-time across the data area.
  // ARRAYFORMULA places the formula in A7 only, but the spilled cells
  // in A8, A9, … still need the date format applied separately because
  // formatting doesn't auto-spill.
  var maxRowsForFormat = sh.getMaxRows();
  sh.getRange(firstDataRow, 1, maxRowsForFormat - firstDataRow + 1, 1)
    .setNumberFormat('dd/MM/yyyy HH:mm:ss');

  // 6. Write the row labels in column A (rows 2-5).
  //
  // These labels describe what each header row means. They are written
  // here (in Add First Step) because this is when the FMS structure
  // first takes shape. They persist for the life of the sheet — every
  // step block added later (Add Step) writes its values across rows 2-5
  // of its own columns, while column A stays as the labels column.
  sh.getRange(2, 1).setValue('What').setFontWeight('bold');
  sh.getRange(3, 1).setValue('Who').setFontWeight('bold');
  sh.getRange(4, 1).setValue('How').setFontWeight('bold');
  sh.getRange(5, 1).setValue('When').setFontWeight('bold');

  // 7. Write the entry-event header band.
  //
  // Per the canonical FMS layout (see fms-knowledge/03-sheet-anatomy/
  // 07-purchase-fms-column-map.md, Step 1: Generate PO), the first step
  // is the *entry event*. It has NO Planned/Actual/Status/Time Delay
  // columns of its own — the form Timestamp in column A serves as its
  // Actual time, and the When is always "Whenever Needed".
  //
  // The header band sits in rows 2-5, merged across columns B through
  // the last form column (canonical: B2:E2). Column A is left out of the
  // merge so it can carry the row labels (What/Who/How/When) and the
  // row-6 "Timestamp" column header.
  var entryStartCol = 2;             // skip A (labels column)
  var entryEndCol = numFormCols;     // last form column
  var entryWidth = entryEndCol - entryStartCol + 1;

  if (entryWidth >= 1) {
    var hRow2 = sh.getRange(2, entryStartCol, 1, entryWidth);
    hRow2.breakApart();
    if (entryWidth > 1) hRow2.merge();
    hRow2.setValue(payload.stepName.trim())
      .setHorizontalAlignment('center').setFontWeight('bold');

    var hRow3 = sh.getRange(3, entryStartCol, 1, entryWidth);
    hRow3.breakApart();
    if (entryWidth > 1) hRow3.merge();
    hRow3.setValue(payload.who || '').setHorizontalAlignment('center');

    var hRow4 = sh.getRange(4, entryStartCol, 1, entryWidth);
    hRow4.breakApart();
    if (entryWidth > 1) hRow4.merge();
    hRow4.setValue(payload.how || '').setHorizontalAlignment('center');

    var hRow5 = sh.getRange(5, entryStartCol, 1, entryWidth);
    hRow5.breakApart();
    if (entryWidth > 1) hRow5.merge();
    hRow5.setValue('Whenever Needed')
      .setHorizontalAlignment('center').setFontStyle('italic');
  }

  // 8. Thick right border at the end of the form-data zone, separating
  // form columns from any future step blocks the user adds via Add Step.
  var maxRowsForBorder = sh.getMaxRows();
  sh.getRange(2, numFormCols, maxRowsForBorder - 1, 1)
    .setBorder(null, null, null, true, null, null,
               '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 9. Apply the header band styling (light gray background, frozen rows,
  // bottom border under row 6) and rebuild step-CF (no steps yet, but
  // the rebuild also clears any stale CF rules left over from prior tests).
  applyHeaderStyling(sh);
  rebuildAllStepCF(sh);

  // 10. Make sure the onChange trigger is installed (Setup also installs
  // it, but no harm running ensureOnChangeTrigger again).
  ensureOnChangeTrigger();

  ss.toast(
    'Form created and entry event "' + payload.stepName + '" added. ' +
    'Run "Add Step" to add the first step that has a deadline.',
    'BMP Formulas', 8
  );

  return {
    formUrl: formInfo.publishedUrl,
    formEditUrl: formInfo.editUrl,
    formTabName: formTabName,
    entryHeaderRange: colNumToLetter(entryStartCol) + '2:' +
      colNumToLetter(entryEndCol) + '5',
    message: 'Form created and linked. Now run "Add Step" to add the first ' +
      'step with a deadline (TAT / Days Before / Specific Time).'
  };
}

/**
 * Creates a Google Form with the given fields and links it to the
 * given spreadsheet. Returns the form's URLs.
 *
 * Field types:
 *   'text'   → addTextItem (free text, no validation)
 *   'number' → addTextItem with numeric validation
 *   'date'   → addDateItem
 */
function createLinkedFormForFMS(processName, formFields, ss) {
  var title = processName + ' — FMS Entry Form';
  var form = FormApp.create(title);

  formFields.forEach(function (f) {
    var label = (f.label || '').trim();
    if (!label) return;
    var item;
    if (f.type === 'number') {
      item = form.addTextItem();
      item.setValidation(FormApp.createTextValidation().requireNumber().build());
    } else if (f.type === 'date') {
      item = form.addDateItem();
    } else {
      item = form.addTextItem();
    }
    item.setTitle(label);
  });

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  return {
    title: title,
    formId: form.getId(),
    publishedUrl: form.getPublishedUrl(),
    editUrl: form.getEditUrl()
  };
}

// ============================================================
//  TIER 2 — ADD STEP WIZARD
// ============================================================

function addStepWizard() {
  var template = HtmlService.createTemplateFromFile('AddStepDialog');
  template.context = JSON.stringify(getAddStepContext());
  var html = template.evaluate().setWidth(620).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'BMP Formulas — Add Step');
}

/**
 * Server-side helper that the wizard uses to populate its dropdowns:
 *   - decisionSteps: steps whose row-2 header ends with "?" (for the
 *     "Only run on Yes / No path" advanced option)
 *   - dataColumns: a unified list of all non-step-internal columns —
 *     form columns, additional info columns, and previous Actual columns.
 *     Each entry is shaped { column, columnLetter, header, type, label }
 *     where type is 'data' (form/info col) or 'actual' (a previous step's
 *     Actual cell). The Add Step wizard's timing pickers populate from this
 *     list and let the user choose which column to anchor against.
 *   - lastStepEndCol: rightmost occupied column. New columns insert here.
 *   - hasFormLink: true if A7's formula references a Form Responses tab
 *     First Step has been run). Used by addStepWizardSave to refuse
 *     Add Step on a sheet that has no entry form yet.
 */
function getAddStepContext() {
  var sh = SpreadsheetApp.getActive().getActiveSheet();
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) {
    return {
      decisionSteps: [],
      dataColumns: [],
      lastStepEndCol: 0,
      hasFormLink: false,
      sheetName: sh.getName()
    };
  }

  var row2 = sh.getRange(2, 1, 1, lastCol).getValues()[0];
  var row6 = sh.getRange(6, 1, 1, lastCol).getValues()[0];

  var decisionSteps = [];
  var dataColumns = [];
  var visited = {};
  var lastStepEndCol = 0;

  for (var c = 1; c <= lastCol; c++) {
    var headerRow6 = String(row6[c - 1] || '').trim();
    if (!headerRow6) continue;

    // Track the rightmost occupied column regardless of header type
    lastStepEndCol = c;

    // Skip step-internal headers other than Actual — they aren't useful as
    // anchors / lead times. Actual IS useful (e.g. Issue PO Actual is the
    // anchor for "Days Before delivery" formulas).
    if (/^(planned|status|yes\/no|time delay)$/i.test(headerRow6)) {
      // still counts toward lastStepEndCol; just don't add to dataColumns
    } else if (/^actual$/i.test(headerRow6)) {
      var merge = sh.getRange(2, c).getMergedRanges();
      var stepStartCol = c;
      if (merge && merge.length > 0) stepStartCol = merge[0].getColumn();
      var stepNameForActual = String(row2[stepStartCol - 1] || '').trim();
      var displayName = stepNameForActual || ('Column ' + colNumToLetter(c));
      dataColumns.push({
        column: c,
        columnLetter: colNumToLetter(c),
        header: 'Actual',
        type: 'actual',
        stepName: stepNameForActual,
        label: displayName + ' — Actual (' + colNumToLetter(c) + ')'
      });
    } else {
      // Form column (e.g. Timestamp, PO Number, Vendor) or additional
      // info column (e.g. Lead Time). Treat them all as one bucket.
      dataColumns.push({
        column: c,
        columnLetter: colNumToLetter(c),
        header: headerRow6,
        type: 'data',
        label: headerRow6 + ' (' + colNumToLetter(c) + ')'
      });
    }

    // Decision-step detection (row-2 step header ends with ?)
    var merge2 = sh.getRange(2, c).getMergedRanges();
    if (merge2 && merge2.length > 0) {
      var dStartCol = merge2[0].getColumn();
      if (!visited[dStartCol]) {
        visited[dStartCol] = true;
        var stepHeader = String(row2[dStartCol - 1] || '').trim();
        if (stepHeader && /\?\s*$/.test(stepHeader)) {
          var dEndCol = dStartCol + merge2[0].getNumColumns() - 1;
          var statusCol = -1;
          for (var sc = dStartCol; sc <= dEndCol; sc++) {
            var dH6 = String(row6[sc - 1] || '').trim();
            if (/yes\/no/i.test(dH6) || /^status$/i.test(dH6)) {
              statusCol = sc;
              break;
            }
          }
          decisionSteps.push({
            stepName: stepHeader,
            statusColumn: statusCol > 0 ? statusCol : dStartCol + 2,
            statusColumnLetter: colNumToLetter(statusCol > 0 ? statusCol : dStartCol + 2)
          });
        }
      }
    }
  }

  // Has the entry form been linked? Look for a Form-Responses reference
  // anywhere inside A7's formula. We can't just match `^=INDEX\(` —
  // v2.1.3 wrapped the INDEX in `=LET(v, INDEX(...), IF(v="","",v))`,
  // so the formula starts with `=LET(`. We also can't only match
  // "Form Responses" with a capital R: some Google account locales
  // produce the tab as "Form responses 1" (lowercase r).
  //
  // The robust signal is: A7's formula references a `Form responses N`
  // tab in any case. That holds for INDEX, LET-wrapped INDEX, INDIRECT,
  // ARRAYFORMULA, and the older `Link Existing Form` flow alike.
  var a7Formula = '';
  try { a7Formula = sh.getRange('A7').getFormula() || ''; } catch (e) { /* empty */ }
  var hasFormLink = /Form\s+responses?\s*\d*/i.test(a7Formula);

  return {
    decisionSteps: decisionSteps,
    dataColumns: dataColumns,
    lastStepEndCol: lastStepEndCol,
    hasFormLink: hasFormLink,
    sheetName: sh.getName()
  };
}

/**
 * Server callback invoked from AddStepDialog.html.
 *
 * Workflow:
 *   1. Refuse if no entry form linked yet (user must run Add First Step).
 *   2. Optionally insert M "additional info" columns at the right end —
 *      these are declared in the wizard's Advanced section (e.g. a
 *      Lead Time column needed for the Days-Before formula being added
 *      in this same call). The wizard refers to these via _new_<i>
 *      placeholder column refs in the payload's pickers.
 *   3. Insert the 4 step columns (Planned/Actual/Status/Time Delay).
 *   4. Build the Planned formula from explicit picker values, with
 *      placeholder refs resolved to actual column numbers.
 *   5. Wrap with decision condition if requested.
 *   6. Hand off to writeStepBlock for header band, formulas, dropdown,
 *      conditional formatting, and autofill-down.
 *
 * payload: {
 *   stepName, who, how,
 *   stepType: 'normal' | 'decision',
 *   timing: 'TAT' | 'TX' | 'SPECIFIC_TIME',
 *   tatHours, tatAnchorCol,
 *   txStartDateCol, txLeadTimeCol, txDaysBefore,
 *   stAnchorCol, stDaysAfter, stTimeOfDay,
 *   conditional, conditionalStatusCol, conditionalOn,
 *   additionalInfoColumns: [{ name }]   // Advanced section
 * }
 *
 * Column refs in pickers can be either a number (existing column) or
 * a string '_new_<i>' referring to the i-th newly-declared info column.
 */
function addStepWizardSave(payload) {
  var sh = SpreadsheetApp.getActive().getActiveSheet();
  var ctx = getAddStepContext();

  if (!payload.stepName || !payload.stepName.trim()) throw new Error('Step name is required.');
  if (!payload.timing) throw new Error('Pick a timing method (TAT / Days Before / Specific Time).');

  // Refuse if the entry form hasn't been linked yet — Add Step needs A7
  // to contain the form timestamp (via INDEX) so step formulas have
  // something to anchor against.
  if (!ctx.hasFormLink) {
    throw new Error('Run "Add First Step (creates Form)" first. The entry form needs to be set up before more steps can be added.');
  }

  var insertAfter = ctx.lastStepEndCol;

  // Step 1: Insert additional info columns (Advanced section) if any.
  var infoCols = (payload.additionalInfoColumns || []).filter(function (ic) {
    return ic && ic.name && ic.name.trim();
  });
  var infoColNumbers = [];   // maps array index → actual column number
  if (infoCols.length > 0) {
    sh.insertColumnsAfter(insertAfter, infoCols.length);
    for (var i = 0; i < infoCols.length; i++) {
      var col = insertAfter + 1 + i;
      sh.getRange(6, col).setValue(infoCols[i].name.trim()).setFontWeight('bold');
      infoColNumbers.push(col);
    }
    insertAfter += infoCols.length;
  }

  // Step 2: Insert the 4 step columns.
  sh.insertColumnsAfter(insertAfter, 4);
  var plannedCol = insertAfter + 1;
  var plannedColLetter = colNumToLetter(plannedCol);
  var firstDataRow = 7;
  // Pre-fill step formulas to at least row 200 (or further if data
  // already extends beyond that). This matters because the form-data
  // ARRAYFORMULA in A7 spills new rows automatically as responses
  // arrive, but step formulas do NOT auto-spill — each row needs its
  // own copy. Without this pre-fill, the 2nd, 3rd, ... form responses
  // would land in FMS rows 8, 9, ... but those rows would have no
  // Planned/Time-Delay formulas, so timing stays empty.
  var STEP_PREFILL_ROWS = 200;
  var lastRow = Math.max(sh.getLastRow(), firstDataRow, STEP_PREFILL_ROWS);
  var rowSuffix = String(firstDataRow);

  // The merged step header band in row 2 spans the info columns (if any)
  // PLUS the 4 step block columns. So the leftmost cell of the merge is
  // the first info column (or plannedCol if no info cols). The TAT-cell
  // reference at $5 of the leftmost merge col is where the timing value
  // lives — matches canonical Purchase FMS Step 2 (F$5 = 3 hours).
  var mergeStartCol = (infoColNumbers.length > 0) ? infoColNumbers[0] : plannedCol;
  var mergeWidth = infoCols.length + 4;
  var tatHoursRef = colNumToLetter(mergeStartCol) + '$5';

  // Resolve a column ref from the payload — either a numeric column
  // number (existing column) or '_new_<i>' (i-th info col we just inserted).
  function resolveCol(ref) {
    if (ref == null || ref === '') return 0;
    if (typeof ref === 'string' && /^_new_(\d+)$/.test(ref)) {
      var idx = parseInt(RegExp.$1, 10);
      return infoColNumbers[idx] || 0;
    }
    var n = Number(ref);
    return isNaN(n) ? 0 : n;
  }

  // Step 3: Build the Planned formula.
  var plannedFormula = '';
  var timingValue;

  if (payload.timing === 'TAT') {
    var tatHours = Number(payload.tatHours);
    if (isNaN(tatHours) || tatHours < 0) throw new Error('TAT hours must be 0 or greater.');
    timingValue = tatHours;

    var anchorCol = resolveCol(payload.tatAnchorCol);
    var timestampRef;
    if (anchorCol > 0) {
      timestampRef = colNumToLetter(anchorCol) + rowSuffix;
    } else {
      timestampRef = autoPickPreviousActual(sh, plannedCol, firstDataRow);
    }
    plannedFormula = buildTATFormula(timestampRef, tatHoursRef);

  } else if (payload.timing === 'TX') {
    var startCol = resolveCol(payload.txStartDateCol);
    if (!startCol) throw new Error('Pick a Start Date column for "Days Before Future Date".');
    var leadCol = resolveCol(payload.txLeadTimeCol);
    if (!leadCol) throw new Error('Pick a Lead Time column for "Days Before Future Date" (or add one in the Advanced section).');
    var anchorRef = colNumToLetter(startCol) + rowSuffix;
    var leadTimeRef = colNumToLetter(leadCol) + rowSuffix;
    var daysBefore = Number(payload.txDaysBefore || 0);
    if (isNaN(daysBefore) || daysBefore < 0) throw new Error('Days Before must be 0 or greater.');
    timingValue = 'T-' + daysBefore;
    plannedFormula = buildTXFormula(anchorRef, leadTimeRef, daysBefore);

  } else if (payload.timing === 'SPECIFIC_TIME') {
    var stCol = resolveCol(payload.stAnchorCol);
    if (!stCol) throw new Error('Pick a Reference Date column for "Specific Clock Time".');
    var stAnchorRef = colNumToLetter(stCol) + rowSuffix;
    var stDays = Number(payload.stDaysAfter || 0);
    if (isNaN(stDays) || stDays < 0) throw new Error('Days After must be 0 or greater.');
    var todFrac = parseTimeToFraction(payload.stTimeOfDay);
    if (todFrac == null) throw new Error('Time of Day must be HH:MM (e.g. 18:00).');
    timingValue = 'By ' + payload.stTimeOfDay + ' (+' + stDays + ' days)';
    plannedFormula = buildSpecificTimeFormula(stAnchorRef, stDays, todFrac);

  } else {
    throw new Error('Unknown timing method: ' + payload.timing);
  }

  // Step 4: Decision-conditional wrapper, if requested.
  if (payload.conditional && payload.conditionalStatusCol &&
      (payload.conditionalOn === 'Yes' || payload.conditionalOn === 'No')) {
    var statusWrapRef = colNumToLetter(Number(payload.conditionalStatusCol)) + rowSuffix;
    plannedFormula = '=IF(' + statusWrapRef + '="' + payload.conditionalOn + '",' +
      plannedFormula.substring(1) + ',"")';
  }

  // Step 5: Hand off to the shared step-block writer. mergeStartCol and
  // mergeWidth tell it to extend the row 2-5 step header band leftward
  // over the info columns (canonical layout — info cols sit UNDER the
  // step's merged header).
  writeStepBlock(sh, plannedCol, {
    stepName: payload.stepName.trim(),
    who: payload.who || '',
    how: payload.how || '',
    timingValue: timingValue,
    plannedFormula: plannedFormula,
    isDecision: payload.stepType === 'decision',
    firstDataRow: firstDataRow,
    lastRow: lastRow,
    mergeStartCol: mergeStartCol,
    mergeWidth: mergeWidth
  });

  // Apply header styling to cover the newly-added columns and rebuild
  // CF for ALL steps (cleans up any stale CF on Planned/Status columns
  // left over from earlier script versions).
  applyHeaderStyling(sh);
  rebuildAllStepCF(sh);

  ensureOnChangeTrigger();

  var msg = 'Step "' + payload.stepName + '" added at column ' + plannedColLetter + '.';
  if (infoCols.length > 0) {
    msg += ' Added ' + infoCols.length + ' info column(s): ' +
      infoCols.map(function (ic) { return ic.name; }).join(', ') + '.';
  }
  SpreadsheetApp.getActive().toast(msg, 'BMP Formulas', 6);

  return {
    plannedCol: plannedColLetter,
    infoColumns: infoColNumbers.map(function (n) { return colNumToLetter(n); }),
    message: msg
  };
}

// ============================================================
//  TIER 2 — LINK FORM RESPONSES
// ============================================================

function linkFormWizard() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  var formTabs = ss.getSheets().filter(function (s) {
    return /^Form Responses ?\d*$/.test(s.getName());
  });

  if (formTabs.length === 0) {
    ui.alert(
      'No Form Responses tab found',
      'Create a Google Form first via Tools > Create a new form. Once a Form Responses tab exists, run this again.',
      ui.ButtonSet.OK
    );
    return;
  }

  var picked;
  if (formTabs.length === 1) {
    picked = formTabs[0];
  } else {
    var listing = formTabs.map(function (t, i) { return (i + 1) + '. ' + t.getName(); }).join('\n');
    var resp = ui.prompt(
      'Pick Form Responses tab',
      'Multiple Form Responses tabs found. Enter the number of the one to link:\n\n' + listing,
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    var idx = parseInt(resp.getResponseText(), 10) - 1;
    if (isNaN(idx) || idx < 0 || idx >= formTabs.length) {
      ui.alert('Invalid selection', 'Try again.', ui.ButtonSet.OK);
      return;
    }
    picked = formTabs[idx];
  }

  var cols = picked.getLastColumn();
  if (cols < 1) {
    ui.alert('Empty form tab',
      'The Form Responses tab has no columns yet. Submit at least one form response, then run this again.',
      ui.ButtonSet.OK);
    return;
  }

  var fmsTab = ss.getActiveSheet();
  if (fmsTab.getName() === picked.getName()) {
    ui.alert('Wrong tab',
      'Run this from your main FMS tab, not the Form Responses tab.',
      ui.ButtonSet.OK);
    return;
  }

  var tabName = picked.getName().replace(/'/g, "''");
  for (var c = 1; c <= cols; c++) {
    var letter = colNumToLetter(c);
    var formula = "=INDEX('" + tabName + "'!" + letter + ":" + letter + ", ROW())";
    fmsTab.getRange(7, c).setFormula(formula);
  }
  fmsTab.getRange(7, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');

  ui.alert(
    'Form linked',
    'INDEX formulas placed in row 7, columns A–' + colNumToLetter(cols) +
      '. New form responses will populate automatically as they arrive.',
    ui.ButtonSet.OK
  );
}

// ============================================================
//  TIER 1 — ENHANCED SINGLE-FORMULA OPERATIONS
//
//  Each of these:
//    1. Confirms the active cell is the right type (or warns the user)
//    2. Auto-detects every cell reference it can
//    3. Shows ConfirmRefsDialog with editable prefilled fields
//    4. ConfirmRefsDialog calls executeOperation(...) on submit
// ============================================================

function applyTAT() {
  var ss = SpreadsheetApp.getActive();
  var cell = ss.getCurrentCell();
  if (!requireActiveCell(cell, 'Click into a Planned cell first, then run TAT.')) return;
  warnIfNotPlannedCell(cell, 'TAT');

  var defaults = {
    timestampCell: detectTimestampRef(cell),
    tatHoursCell: detectTatHoursRef(cell)
  };
  showConfirmRefsDialog(
    'Apply TAT (Working Hours)',
    'Auto-detected values are filled in below. Review or edit before applying. The script will not change anything until you click Apply.',
    [
      { name: 'timestampCell', label: 'Trigger Cell',
        value: defaults.timestampCell, isCellRef: true,
        hint: 'When this step starts counting from. First step: A7 (form timestamp). Later steps: previous step\'s Actual cell (e.g. H7).' },
      { name: 'tatHoursCell', label: 'TAT Hours Cell',
        value: defaults.tatHoursCell, isCellRef: true,
        hint: 'Where the TAT hours value lives. Typically the leftmost column of this step\'s header band, row 5. Example: F$5 or S$5.' }
    ],
    'TAT',
    cell
  );
}

function applyTX() {
  var ss = SpreadsheetApp.getActive();
  var cell = ss.getCurrentCell();
  if (!requireActiveCell(cell, 'Click into a Planned cell first, then run this.')) return;
  warnIfNotPlannedCell(cell, 'Days Before Future Date');

  var defaults = {
    startDateCell: '',
    leadTimeCell: detectLeadTimeRef(cell)
  };
  showConfirmRefsDialog(
    'Apply Days Before Future Date',
    'Plan a step that should happen X days before a future delivery date. The Start Date is required; the lead time is auto-detected from the Lead Time column.',
    [
      { name: 'startDateCell', label: 'Start Date Cell',
        value: defaults.startDateCell, isCellRef: true,
        hint: 'The date you add the lead time to. Usually the Actual cell of the step where the order/PO was issued (e.g. H7).' },
      { name: 'leadTimeCell', label: 'Lead Time Cell',
        value: defaults.leadTimeCell, isCellRef: true,
        hint: 'Auto-detected from the Lead Time column header in row 6, same row as the Planned cell.' },
      { name: 'daysBefore', label: 'Days Before (X)',
        value: '2', isCellRef: false,
        hint: 'Plain number. 2 means "follow up 2 days before delivery". 0 means "the delivery date itself".' }
    ],
    'TX',
    cell
  );
}

function applySpecificTime() {
  var ss = SpreadsheetApp.getActive();
  var cell = ss.getCurrentCell();
  if (!requireActiveCell(cell, 'Click into a Planned cell first, then run this.')) return;
  warnIfNotPlannedCell(cell, 'Specific Clock Time');

  // Note: dateCell auto-derivation is unreliable for Specific Time (just like T-X),
  // so the user picks. Days After is a small set of common choices; Time of Day is HH:MM.
  showConfirmRefsDialog(
    'Apply Specific Clock Time',
    'Plan a step that must be done by a fixed clock-time on a specific day. Pick the reference date cell, choose how many working days after, and enter the time of day.',
    [
      { name: 'dateCell', label: 'Date Cell',
        value: '', isCellRef: true,
        hint: 'The reference date this deadline counts from. Usually the previous step\'s Actual cell (e.g. P7).' },
      { name: 'daysAfter', label: 'Days After',
        value: '0', isCellRef: false,
        kind: 'radio',
        options: [
          { value: '0', label: 'Same day' },
          { value: '1', label: 'Next working day' },
          { value: '2', label: 'Day after next' }
        ],
        allowCustom: true,
        hint: 'How many working days after the reference date. Pick a common one or enter a custom number.' },
      { name: 'timeOfDay', label: 'Time of Day (HH:MM)',
        value: '18:00', isCellRef: false,
        hint: 'Use 24-hour HH:MM format. Examples: 09:30, 14:00, 18:00.' }
    ],
    'SPECIFIC_TIME',
    cell
  );
}

function applyShowYes() { applyShowCondition_('Yes'); }
function applyShowNo()  { applyShowCondition_('No');  }

function applyShowCondition_(condition) {
  var ss = SpreadsheetApp.getActive();
  var cell = ss.getCurrentCell();
  if (!requireActiveCell(cell, 'Click into the Planned cell that has a TAT / T-X / Specific Time formula already, then run this.')) return;

  var existing = cell.getFormula();
  if (!existing || !/^=IF\(/.test(existing)) {
    SpreadsheetApp.getUi().alert(
      'Apply a timing formula first',
      'This cell does not have a TAT, Days Before, or Specific Time formula yet. Apply one of those first, then come back here to wrap it with the YES/NO condition.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var defaults = { statusCell: detectStatusRef(cell) };
  showConfirmRefsDialog(
    'Show planned only when status is ' + condition.toUpperCase(),
    'The current Planned formula will be wrapped so that it only shows a value when the chosen Status cell equals "' + condition + '".',
    [
      { name: 'statusCell', label: 'Status Cell of the decision step',
        value: defaults.statusCell, isCellRef: true,
        hint: 'The Yes/No column of the decision step (e.g. U7).' }
    ],
    condition === 'Yes' ? 'SHOW_YES' : 'SHOW_NO',
    cell
  );
}

function applyTimeDelay() {
  var ss = SpreadsheetApp.getActive();
  var cell = ss.getCurrentCell();
  if (!requireActiveCell(cell, 'Click into the Time Delay cell first, then run this.')) return;

  var pair = detectTimeDelayPair(cell);
  showConfirmRefsDialog(
    'Apply Time Delay',
    'The Time Delay formula compares Planned and Actual to show overdue / late time. Cells are auto-detected; review or edit as needed.',
    [
      { name: 'plannedCell', label: 'Planned Cell',
        value: pair.planned, isCellRef: true,
        hint: 'The Planned cell of this step (typically 3 columns left of this Time Delay cell).' },
      { name: 'actualCell', label: 'Actual Cell',
        value: pair.actual, isCellRef: true,
        hint: 'The Actual cell of this step (typically 2 columns left of this Time Delay cell).' }
    ],
    'TIME_DELAY',
    cell
  );
}

function setActualTime() {
  ensureOnChangeTrigger();
  SpreadsheetApp.getUi().alert(
    'Set Actual Time',
    'Trigger installed. When someone types "Done", "Yes", or "No" in a Status column, the Actual cell to its left will be auto-stamped with the current time.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================
//  TIER 1 / TIER 2 — DIALOG SHOW + DISPATCH
// ============================================================

/**
 * Renders ConfirmRefsDialog.html with the given fields and remembers the
 * cell location so the server callback can find it after the modal closes.
 */
function showConfirmRefsDialog(title, intro, fields, opType, cell) {
  var template = HtmlService.createTemplateFromFile('ConfirmRefsDialog');
  template.dialogTitle = title;
  template.intro = intro;
  template.fields = fields;
  template.opType = opType;
  template.ctx = {
    sheetName: cell.getSheet().getName(),
    row: cell.getRow(),
    col: cell.getColumn()
  };
  var html = template.evaluate().setWidth(560).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

/**
 * Dispatcher — every Tier 1 modal calls this with its opType + payload.
 * Validates cell references, then routes to the executor.
 */
function executeOperation(opType, payload) {
  var ctx = payload.ctx;
  if (!ctx) throw new Error('Internal error: missing context');
  var sh = SpreadsheetApp.getActive().getSheetByName(ctx.sheetName);
  if (!sh) throw new Error('Sheet "' + ctx.sheetName + '" not found.');
  var cell = sh.getRange(ctx.row, ctx.col);

  // Validate every payload field that looks like a cell reference (Tier 4 #6)
  Object.keys(payload).forEach(function (key) {
    if (key === 'ctx') return;
    if (!/Cell$/.test(key)) return;
    var v = payload[key];
    if (!v) return;
    try { sh.getRange(v); } catch (e) {
      throw new Error('Invalid cell reference for ' + key + ': "' + v + '"');
    }
  });

  switch (opType) {
    case 'TAT': return executeTAT(sh, cell, payload);
    case 'TX': return executeTX(sh, cell, payload);
    case 'SPECIFIC_TIME': return executeSpecificTime(sh, cell, payload);
    case 'SHOW_YES': return executeShowCondition(sh, cell, payload, 'Yes');
    case 'SHOW_NO':  return executeShowCondition(sh, cell, payload, 'No');
    case 'TIME_DELAY': return executeTimeDelay(sh, cell, payload);
    default: throw new Error('Unknown operation: ' + opType);
  }
}

// ============================================================
//  TIER 1 — EXECUTORS
// ============================================================

function executeTAT(sh, cell, payload) {
  var formula = buildTATFormula(payload.timestampCell, payload.tatHoursCell);
  cell.setFormula(formula).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  autofillFormulaDown(cell);
  SpreadsheetApp.getActive().toast('TAT formula applied to ' + cell.getA1Notation() + '.', 'BMP Formulas', 4);
  return { ok: true };
}

function executeTX(sh, cell, payload) {
  var daysBefore = Number(payload.daysBefore);
  if (isNaN(daysBefore)) throw new Error('Days Before must be a number.');
  var formula = buildTXFormula(payload.startDateCell, payload.leadTimeCell, daysBefore);
  cell.setFormula(formula).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  autofillFormulaDown(cell);
  SpreadsheetApp.getActive().toast('Days-Before formula applied to ' + cell.getA1Notation() + '.', 'BMP Formulas', 4);
  return { ok: true };
}

function executeSpecificTime(sh, cell, payload) {
  var daysAfter = Number(payload.daysAfter);
  if (isNaN(daysAfter)) throw new Error('Days After must be a number.');
  var todFrac = parseTimeToFraction(payload.timeOfDay);
  if (todFrac == null) throw new Error('Time of Day must be HH:MM (e.g. 18:00).');
  var formula = buildSpecificTimeFormula(payload.dateCell, daysAfter, todFrac);
  cell.setFormula(formula).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  autofillFormulaDown(cell);
  SpreadsheetApp.getActive().toast('Specific Time formula applied to ' + cell.getA1Notation() + '.', 'BMP Formulas', 4);
  return { ok: true };
}

function executeShowCondition(sh, cell, payload, condition) {
  var existing = cell.getFormula();
  if (!existing || !/^=IF\(/.test(existing)) {
    throw new Error('This cell no longer has a timing formula. Apply TAT / T-X / Specific Time first, then re-run this.');
  }
  var newFormula = '=IF(' + payload.statusCell + '="' + condition + '",' + existing.substring(1) + ',"")';
  cell.setFormula(newFormula).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  autofillFormulaDown(cell);
  SpreadsheetApp.getActive().toast(
    'Wrapped: ' + cell.getA1Notation() + ' shows planned only when ' + payload.statusCell + ' = "' + condition + '".',
    'BMP Formulas', 4
  );
  return { ok: true };
}

function executeTimeDelay(sh, cell, payload) {
  var formula = buildTimeDelayFormula(payload.plannedCell, payload.actualCell);
  cell.setFormula(formula).setNumberFormat('[h]:mm:ss');
  autofillFormulaDown(cell);

  // Apply the 2-rule conditional formatting (Tier 4 #3)
  var lastRow = Math.max(sh.getLastRow(), cell.getRow());
  applyTimeDelayConditionalFormatting(
    sh, cell.getColumn(), cell.getRow(), lastRow,
    payload.plannedCell, payload.actualCell
  );

  SpreadsheetApp.getActive().toast('Time Delay applied to ' + cell.getA1Notation() + ' with overdue/late color rules.', 'BMP Formulas', 4);
  return { ok: true };
}

// ============================================================
//  AUTO-DETECTION HELPERS
// ============================================================

/** Walk left in row 6 from the active column for the first "Actual" header. */
function detectTimestampRef(activeCell) {
  var sh = activeCell.getSheet();
  var col = activeCell.getColumn();
  var row = activeCell.getRow();
  if (col <= 1) return 'A' + row;
  var row6 = sh.getRange(6, 1, 1, col - 1).getValues()[0];
  for (var c = col - 2; c >= 0; c--) {  // 0-indexed
    var h = String(row6[c] || '').trim();
    if (/^actual$/i.test(h)) {
      return colNumToLetter(c + 1) + row;
    }
  }
  return 'A' + row;
}

/** Find the leftmost column of the row-2 merged step header that contains the active column. */
function detectTatHoursRef(activeCell) {
  var sh = activeCell.getSheet();
  var col = activeCell.getColumn();
  var merges = sh.getRange(2, col).getMergedRanges();
  if (merges && merges.length > 0) {
    return colNumToLetter(merges[0].getColumn()) + '$5';
  }
  return colNumToLetter(col) + '$5';
}

/** First column whose row-6 header contains "Lead Time" (case-insensitive). */
function detectLeadTimeRef(activeCell) {
  var sh = activeCell.getSheet();
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return '';
  var row6 = sh.getRange(6, 1, 1, lastCol).getValues()[0];
  for (var c = 0; c < row6.length; c++) {
    var h = String(row6[c] || '').trim();
    if (/lead\s*time/i.test(h)) {
      return colNumToLetter(c + 1) + activeCell.getRow();
    }
  }
  return '';
}

/** Walk left from active column for a Yes/No header, or a Status cell whose parent step ends with "?". */
function detectStatusRef(activeCell) {
  var sh = activeCell.getSheet();
  var col = activeCell.getColumn();
  var row = activeCell.getRow();
  if (col <= 1) return '';
  var row2 = sh.getRange(2, 1, 1, col - 1).getValues()[0];
  var row6 = sh.getRange(6, 1, 1, col - 1).getValues()[0];
  for (var c = col - 2; c >= 0; c--) {
    var h6 = String(row6[c] || '').trim();
    if (/yes\/no/i.test(h6)) return colNumToLetter(c + 1) + row;
    if (/^status$/i.test(h6)) {
      // Check if this Status belongs to a decision step (row-2 header ends with ?)
      var merges = sh.getRange(2, c + 1).getMergedRanges();
      var startCol = c + 1;
      if (merges && merges.length > 0) startCol = merges[0].getColumn();
      var stepHeader = String(row2[startCol - 1] || '').trim();
      if (/\?\s*$/.test(stepHeader)) return colNumToLetter(c + 1) + row;
    }
  }
  return '';
}

/** Canonical 4-column step block: Planned | Actual | Status | Time Delay. */
function detectTimeDelayPair(activeCell) {
  var col = activeCell.getColumn();
  var row = activeCell.getRow();
  if (col < 4) return { planned: '', actual: '' };
  return {
    planned: colNumToLetter(col - 3) + row,
    actual: colNumToLetter(col - 2) + row
  };
}

/** Used by the Add Step wizard to pick a TAT timestamp anchor when none is supplied. */
function autoPickPreviousActual(sh, plannedCol, dataRow) {
  if (plannedCol <= 1) return 'A' + dataRow;
  var row6 = sh.getRange(6, 1, 1, plannedCol - 1).getValues()[0];
  for (var c = plannedCol - 2; c >= 0; c--) {
    if (/^actual$/i.test(String(row6[c] || '').trim())) {
      return colNumToLetter(c + 1) + dataRow;
    }
  }
  return 'A' + dataRow;
}

function autoPickLeadTime(sh, dataRow) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return '';
  var row6 = sh.getRange(6, 1, 1, lastCol).getValues()[0];
  for (var c = 0; c < row6.length; c++) {
    if (/lead\s*time/i.test(String(row6[c] || '').trim())) {
      return colNumToLetter(c + 1) + dataRow;
    }
  }
  return '';
}

// ============================================================
//  FORMULA BUILDERS
// ============================================================

/**
 * TAT formula (working hours, weekends, holidays).
 * Preserves the v1 LET-formula logic. The exact-multiple-day path in the
 * inner LET is the same as v1 — see the // TAT edge case comment for
 * the non-working-day trigger area flagged in Tier 4 #7.
 *
 * v2.1.8: opening time / closing time / working-days pattern moved from
 * $C$1, $D$1, $E$1 of the FMS tab to A1, A2, A3 of a dedicated Config
 * tab. This makes the references immune to FMS column-insert shifts
 * (which previously displaced the working-days pattern when Add Step
 * inserted columns at or before column E on a small form).
 */
function buildTATFormula(startRef, tatHoursRef) {
  var holRange = 'Holidays!$A$1:$A$300';  // Tier 4 #4 — proper $A$1:$A$300 syntax
  return '=IF(OR(ISBLANK(' + startRef + '),ISBLANK(' + tatHoursRef + ')),"",' +
    'LET(' +
      '_start,' + startRef + ',' +
      '_tat,' + tatHoursRef + '/24,' +
      '_open,Config!$A$1,' +
      '_close,Config!$A$2,' +
      '_wd,Config!$A$3,' +
      '_hol,' + holRange + ',' +
      '_sd,INT(_start),' +
      '_st,MOD(_start,1),' +
      // TAT edge case (Tier 4 #7): _iswd is true if the start day itself is a
      // working day; if not we jump to the next working day's opening time.
      '_iswd,WORKDAY.INTL(_sd-1,1,_wd,_hol)=_sd,' +
      '_first,' +
        'IF(_iswd,' +
           'IF(_st<_open,_sd+_open,' +
              'IF(_st>=_close,WORKDAY.INTL(_sd,1,_wd,_hol)+_open,_start)),' +
           'WORKDAY.INTL(_sd,1,_wd,_hol)+_open),' +
      '_ft,MOD(_first,1),' +
      '_avail,_close - MAX(_ft,_open),' +
      'IF(_tat<=_avail,' +
         '_first + _tat,' +
         'LET(' +
           '_rem1,_tat - _avail,' +
           '_daylen,_close - _open,' +
           '_k,INT(_rem1/_daylen),' +
           '_rem2,MOD(_rem1,_daylen),' +
           '_base,WORKDAY.INTL(INT(_first),1+_k,_wd,_hol),' +
           'IF(_rem2=0,' +
             'WORKDAY.INTL(INT(_first),_k,_wd,_hol) + _close,' +
             '_base + _open + _rem2' +
           ')' +
         ')' +
      ')' +
    '))';
}

/**
 * Days Before Future Date (T-X) formula.
 * Result = startDate + leadTime - daysBefore.
 * Guard requires both startDate and leadTime to be present.
 */
function buildTXFormula(startRef, leadTimeRef, daysBefore) {
  return '=IF(AND(' + startRef + '<>"",' + leadTimeRef + '<>""),' +
    startRef + '+' + leadTimeRef + '-' + Number(daysBefore) + ',"")';
}

/**
 * Specific Clock Time formula.
 *
 * Tier 4 #1: reads the working-days bitstring from a settings cell
 * instead of v1's hardcoded "0000001" so a Mon-Fri business doesn't
 * silently get Saturday treated as a working day.
 *
 * v2.1.8: that settings cell moved from FMS!$E$1 to Config!$A$3 to
 * avoid being displaced by Add Step's column inserts.
 */
function buildSpecificTimeFormula(dateRef, daysAfter, timeFraction) {
  var holRange = 'Holidays!$A$1:$A$300';  // Tier 4 #4
  return '=IF(' + dateRef + ',WORKDAY.INTL(INT(' + dateRef + '),' + Number(daysAfter) + ',Config!$A$3,' + holRange + ')+' + Number(timeFraction) + ',"")';
}

/**
 * Time Delay formula. Same logic as v1.
 */
function buildTimeDelayFormula(plannedRef, actualRef) {
  // 4 cases:
  //   - Planned empty                              → "" (blank, nothing to compute)
  //   - Planned set, Actual set, Actual <= Planned → "" (on time / early — no delay)
  //   - Planned set, Actual set, Actual > Planned  → Actual - Planned (completed-late delta)
  //   - Planned set, Actual empty                  → $A$1 - Planned
  //         Planned < NOW                          → positive (running overdue)
  //         Planned >= NOW                         → negative (countdown to deadline)
  //
  // v2.1.10: reverted the v2.1.9 `IF($A$1 > planned, …, "")` guard that
  // suppressed the negative form. Per user direction, the negative
  // `-hh:mm:ss` for not-yet-due tasks is the *intended* behavior — it
  // acts as a countdown timer to the deadline. The conditional
  // formatting (ISNUMBER-guarded since v2.1.9) does not light empty
  // rows yellow, so the CF noise that triggered v2.1.9 was a separate
  // bug and is independently fixed.
  return '=IF(' + plannedRef + ',IF(' + actualRef + '<>"",' +
    'IF(' + actualRef + '>' + plannedRef + ',' + actualRef + '-' + plannedRef + ',""),' +
    '$A$1-' + plannedRef + '),"")';
}

/**
 * Applies the 2-rule conditional formatting (red completed-late + yellow
 * overdue-pending) for a single Time Delay column.
 *
 * Three things differ from v2.1.3:
 *
 *  (a) Formula style ported from the FMS Builder Web App monolith:
 *        Red:    =IF(actualRef, IF(actualRef > plannedRef, 1, 0), 0)
 *        Yellow: =IF(actualRef, 0, IF(plannedRef < $A$1, 1, 0))
 *      Simpler than v2.1.3's AND(LEN(...)>0, ...) form, matches what
 *      the workshop teaches.
 *
 *  (b) Full-column range. Applies to startRow..maxRows so newly arriving
 *      form responses get colored automatically.
 *
 *  (c) Self-cleaning. Drops any existing CF rules whose range overlaps
 *      this Time Delay column before adding the fresh rules. Without
 *      this, repeated runs accumulate stale rules — and old rules left
 *      over from v1's getActiveRange()-based timeDelay() (which could
 *      end up scoped to a Planned column) keep coloring the wrong cells.
 *
 * For full-sheet rebuilds (the wizards' canonical path), prefer
 * rebuildAllStepCF below — it cleans rules across all step columns at
 * once, not just this one.
 */
function applyTimeDelayConditionalFormatting(sh, col, startRow, endRow, plannedRefForRow1, actualRefForRow1) {
  // (b) full column down to maxRows
  var maxRows = sh.getMaxRows();
  var numRows = Math.max(1, maxRows - startRow + 1);
  var range = sh.getRange(startRow, col, numRows, 1);

  // (c) drop existing rules that touch THIS column
  var existing = sh.getConditionalFormatRules();
  var preserved = [];
  for (var i = 0; i < existing.length; i++) {
    var rule = existing[i];
    var keep = true;
    var ranges = rule.getRanges();
    for (var r = 0; r < ranges.length && keep; r++) {
      if (col >= ranges[r].getColumn() && col <= ranges[r].getLastColumn()) {
        keep = false;
      }
    }
    if (keep) preserved.push(rule);
  }

  // (a) ISNUMBER-guarded formulas (v2.1.9) — red first (higher priority).
  //
  // Why ISNUMBER instead of the v2.1.4 monolith-style `IF(actualRef, …)`:
  // the older form fires yellow on EVERY empty data row because Sheets
  // coerces an empty cell to 0 when comparing against a date number,
  // and 0 < (any date in 2026) is TRUE — so `plannedRef<$A$1` evaluates
  // to TRUE for empty `plannedRef`. ISNUMBER returns FALSE for empty
  // cells, which makes the AND short-circuit cleanly.
  //
  //   Red:    Planned IS a date AND Actual IS a date AND Actual > Planned
  //   Yellow: Planned IS a date AND Actual is NOT a date AND Planned < NOW
  //
  // Empty rows: ISNUMBER on empty = FALSE → AND fails → no color.
  // Future planned, not done: Planned<NOW = FALSE → no color.
  // Done early: Actual<=Planned → red FALSE → no color.
  var redRule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([range])
    .whenFormulaSatisfied('=AND(ISNUMBER(' + plannedRefForRow1 + '),ISNUMBER(' + actualRefForRow1 + '),' + actualRefForRow1 + '>' + plannedRefForRow1 + ')')
    .setBackground('#F4C7C3')
    .build();

  var yellowRule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([range])
    .whenFormulaSatisfied('=AND(ISNUMBER(' + plannedRefForRow1 + '),NOT(ISNUMBER(' + actualRefForRow1 + ')),' + plannedRefForRow1 + '<$A$1)')
    .setBackground('#FCE8B2')
    .build();

  preserved.push(redRule);
  preserved.push(yellowRule);
  sh.setConditionalFormatRules(preserved);
}

/**
 * Walks row 6 to discover every step block (4-column run of
 * Planned / Actual / Status-or-Yes-No / Time Delay), drops any CF rules
 * touching ANY step column, then adds fresh red + yellow rules for
 * every step's Time Delay column. The canonical CF path for the
 * wizards — robust against orphan rules from earlier script versions
 * that may have scoped CF to Planned/Status columns by mistake.
 *
 * Ported in spirit from the Web App monolith's
 * rebuildTimeDelayConditionalFormatting (lines 2104-2154).
 */
function rebuildAllStepCF(sh) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 4) return;
  var startRow = 7;
  var maxRows = sh.getMaxRows();
  var row6 = sh.getRange(6, 1, 1, lastCol).getValues()[0];

  // Discover step blocks: 4 contiguous cols with the canonical headers
  var steps = [];
  var stepColumns = {};
  for (var c = 1; c <= lastCol - 3; c++) {
    var h0 = String(row6[c - 1] || '').trim();
    var h1 = String(row6[c]     || '').trim();
    var h2 = String(row6[c + 1] || '').trim();
    var h3 = String(row6[c + 2] || '').trim();
    if (/^planned$/i.test(h0) &&
        /^actual$/i.test(h1) &&
        /^(status|yes\/no)$/i.test(h2) &&
        /^time delay$/i.test(h3)) {
      steps.push({
        plannedCol: c,
        actualCol: c + 1,
        statusCol: c + 2,
        timeDelayCol: c + 3
      });
      stepColumns[c] = true;
      stepColumns[c + 1] = true;
      stepColumns[c + 2] = true;
      stepColumns[c + 3] = true;
    }
  }

  // Drop any existing CF rule that touches a step column
  var existing = sh.getConditionalFormatRules();
  var preserved = [];
  for (var i = 0; i < existing.length; i++) {
    var rule = existing[i];
    var overlap = false;
    var ranges = rule.getRanges();
    for (var r = 0; r < ranges.length && !overlap; r++) {
      var rStart = ranges[r].getColumn();
      var rEnd = ranges[r].getLastColumn();
      for (var rc = rStart; rc <= rEnd; rc++) {
        if (stepColumns[rc]) { overlap = true; break; }
      }
    }
    if (!overlap) preserved.push(rule);
  }

  // Add fresh red + yellow rules for every step's Time Delay column.
  // ISNUMBER-guarded formulas (v2.1.9) — see applyTimeDelayConditionalFormatting
  // for the full rationale. Short version: the older `IF(actualRef, …)`
  // form fired yellow on every empty data row because empty<NOW evaluates
  // to TRUE in Sheets' loose type coercion.
  steps.forEach(function (s) {
    var plannedRef = colNumToLetter(s.plannedCol) + startRow;
    var actualRef = colNumToLetter(s.actualCol) + startRow;
    var range = sh.getRange(startRow, s.timeDelayCol, maxRows - startRow + 1, 1);

    preserved.push(SpreadsheetApp.newConditionalFormatRule()
      .setRanges([range])
      .whenFormulaSatisfied('=AND(ISNUMBER(' + plannedRef + '),ISNUMBER(' + actualRef + '),' + actualRef + '>' + plannedRef + ')')
      .setBackground('#F4C7C3')
      .build());

    preserved.push(SpreadsheetApp.newConditionalFormatRule()
      .setRanges([range])
      .whenFormulaSatisfied('=AND(ISNUMBER(' + plannedRef + '),NOT(ISNUMBER(' + actualRef + ')),' + plannedRef + '<$A$1)')
      .setBackground('#FCE8B2')
      .build());
  });

  sh.setConditionalFormatRules(preserved);
}

/**
 * Visual styling for the FMS header band:
 *   - Light gray (#F3F3F3) background on rows 1-6 across the used width
 *   - Bold-medium bottom border on row 6 to separate header from data
 *   - Freeze top 6 rows so headers stay visible while scrolling
 *
 * Called at the end of addFirstStepWizardSave and addStepWizardSave so
 * newly-added columns inherit the styling.
 */
function applyHeaderStyling(sh) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return;

  // Light gray background on the header band (rows 1-6)
  sh.getRange(1, 1, 6, lastCol).setBackground('#F3F3F3');

  // Re-assert bold on row 6 (the column-header row) across the full
  // width — writeStepBlock already does this for new step columns,
  // but doing it here too covers any extra columns the user may have
  // added by hand.
  sh.getRange(6, 1, 1, lastCol).setFontWeight('bold');

  // Bottom border under row 6 separates header from data
  sh.getRange(6, 1, 1, lastCol)
    .setBorder(null, null, true, null, null, null,
               '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Freeze top 6 rows so headers stay visible while scrolling
  sh.setFrozenRows(6);
}

// ============================================================
//  COMMON HELPERS
// ============================================================

/**
 * Shared step-block writer used by both Add First Step and Add Step.
 * Caller has already inserted the 4 columns starting at plannedCol and
 * built the Planned formula. This helper handles everything else:
 * the merged step header band, row-6 column headers, formula placement
 * with autofill-down, Actual column formatting, Status data validation,
 * and Time Delay conditional formatting.
 *
 * opts: {
 *   stepName, who, how,
 *   timingValue,            // value or label put in row 5 of mergeStartCol
 *   plannedFormula,         // pre-built formula starting with =
 *   isDecision,             // boolean — Yes/No vs Done dropdown
 *   firstDataRow, lastRow,  // typically 7 and sh.getLastRow()
 *
 *   // Optional — controls the row 2-5 merged step header band:
 *   mergeStartCol,          // leftmost col of the merge (default = plannedCol)
 *   mergeWidth              // total width of the merge (default = 4)
 * }
 *
 * The merge can extend leftward to cover additional info columns (e.g.
 * a Lead Time column declared via the Add Step wizard's Advanced
 * section). When that happens, the timing value lives in the LEFTMOST
 * cell of the merge so the Planned formula's TAT-cell reference at
 * <mergeStartCol>$5 picks it up — matching canonical Purchase FMS
 * where Step 2's TAT value sits in F$5 even though the Planned column
 * is G.
 */
function writeStepBlock(sh, plannedCol, opts) {
  var actualCol = plannedCol + 1;
  var statusCol = plannedCol + 2;
  var timeDelayCol = plannedCol + 3;

  var plannedColLetter = colNumToLetter(plannedCol);
  var actualColLetter = colNumToLetter(actualCol);

  // Merge range covers info columns (if any) + the 4 step block columns
  var mergeStartCol = opts.mergeStartCol || plannedCol;
  var mergeWidth = opts.mergeWidth || 4;
  if (mergeStartCol > plannedCol) mergeStartCol = plannedCol;  // sanity
  if (mergeWidth < 4) mergeWidth = 4;

  // Row 2: step name (merged across mergeStartCol..mergeStartCol+mergeWidth-1).
  // breakApart-then-merge defends against neighbouring merges leaking
  // into newly-inserted columns (memory: addStep phantom-columns bug).
  var hRow2 = sh.getRange(2, mergeStartCol, 1, mergeWidth);
  hRow2.breakApart();
  if (mergeWidth > 1) hRow2.merge();
  hRow2.setValue(opts.stepName).setHorizontalAlignment('center').setFontWeight('bold');

  var hRow3 = sh.getRange(3, mergeStartCol, 1, mergeWidth);
  hRow3.breakApart();
  if (mergeWidth > 1) hRow3.merge();
  hRow3.setValue(opts.who || '').setHorizontalAlignment('center');

  var hRow4 = sh.getRange(4, mergeStartCol, 1, mergeWidth);
  hRow4.breakApart();
  if (mergeWidth > 1) hRow4.merge();
  hRow4.setValue(opts.how || '').setHorizontalAlignment('center');

  // Row 5: timing value at the LEFTMOST cell of the merge so the Planned
  // formula's TAT-cell reference at <mergeStartCol>$5 finds it. This
  // matches canonical (F$5 for Step 2 in Purchase FMS).
  if (opts.timingValue !== undefined && opts.timingValue !== null && opts.timingValue !== '') {
    sh.getRange(5, mergeStartCol).setValue(opts.timingValue);
  }

  // Row 6: column headers for the 4 step columns. (Info columns inserted
  // before plannedCol have their row-6 headers written by the caller —
  // addStepWizardSave does this when it inserts the info columns.)
  sh.getRange(6, plannedCol).setValue('Planned').setFontWeight('bold');
  sh.getRange(6, actualCol).setValue('Actual').setFontWeight('bold');
  sh.getRange(6, statusCol).setValue(opts.isDecision ? 'Yes/No' : 'Status').setFontWeight('bold');
  sh.getRange(6, timeDelayCol).setValue('Time Delay').setFontWeight('bold');

  var firstDataRow = opts.firstDataRow || 7;
  var lastRow = opts.lastRow || sh.getLastRow();
  if (lastRow < firstDataRow) lastRow = firstDataRow;

  // Planned formula in firstDataRow, then autofill down
  var plannedCell = sh.getRange(firstDataRow, plannedCol);
  plannedCell.setFormula(opts.plannedFormula).setNumberFormat('dd/MM/yyyy HH:mm:ss');

  // Time Delay formula in firstDataRow, then autofill down
  var timeDelayCell = sh.getRange(firstDataRow, timeDelayCol);
  timeDelayCell.setFormula(buildTimeDelayFormula(
    plannedColLetter + firstDataRow, actualColLetter + firstDataRow
  )).setNumberFormat('[h]:mm:ss');

  // Format the Actual column for the data range
  sh.getRange(firstDataRow, actualCol, lastRow - firstDataRow + 1, 1)
    .setNumberFormat('dd/MM/yyyy HH:mm:ss');

  // Autofill formulas down to lastRow. We MUST pass lastRow explicitly —
  // on a fresh FMS getLastRow() returns 7 (only A7's ARRAYFORMULA has
  // content), which would otherwise stop the autofill at row 7 and leave
  // rows 8..STEP_PREFILL_ROWS without step formulas. Form responses
  // landing in those empty rows would then show no Planned/Time-Delay.
  if (lastRow > firstDataRow) {
    autofillFormulaDown(plannedCell, lastRow);
    autofillFormulaDown(timeDelayCell, lastRow);
  }

  // Status data validation dropdown — apply down to maxRows so newly-
  // arriving form responses get the dropdown automatically
  var statusValues = opts.isDecision ? ['Yes', 'No'] : ['Done'];
  var statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusValues, true).build();
  var maxRows = sh.getMaxRows();
  sh.getRange(firstDataRow, statusCol, maxRows - firstDataRow + 1, 1)
    .setDataValidation(statusValidation);

  // Thick right border on Time Delay column visually separates this
  // step from the next. Spans rows 2..maxRows so it reaches all data.
  sh.getRange(2, timeDelayCol, maxRows - 1, 1)
    .setBorder(null, null, null, true, null, null,
               '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // NOTE: Time Delay conditional formatting is intentionally NOT applied
  // here. The wizards (addFirstStepWizardSave, addStepWizardSave) call
  // rebuildAllStepCF at the end to clean up stale rules and rebuild
  // fresh CF for every step in one pass. That avoids accumulation of
  // orphan CF rules across calls.

  return {
    plannedCol: plannedCol,
    actualCol: actualCol,
    statusCol: statusCol,
    timeDelayCol: timeDelayCol
  };
}

function parseTimeToFraction(txt) {
  if (txt == null) return null;
  var s = String(txt).trim();
  var m = /^([01]?\d|2[0-3]):([0-5]\d)$/.exec(s);
  if (!m) return null;
  return (parseInt(m[1], 10) * 60 + parseInt(m[2], 10)) / 1440;
}

function colNumToLetter(n) {
  var s = '';
  while (n > 0) {
    var rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function setIterativeCalc() {
  var ss = SpreadsheetApp.getActive();
  ss.setRecalculationInterval(SpreadsheetApp.RecalculationInterval.ON_CHANGE);
  ss.setIterativeCalculationEnabled(true);
  ss.setMaxIterativeCalculationCycles(1);
  ss.setIterativeCalculationConvergenceThreshold(0.05);
}

/**
 * Tier 4 #2: removes ALL onChange_new triggers (not just one) so duplicate
 * triggers — a known v1 bug — are eliminated. Replaces the v1 removeTrigger
 * which had an always-true comparison and a misplaced `break`.
 */
function removeOnChangeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onChange_new') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function ensureOnChangeTrigger() {
  removeOnChangeTriggers();
  ScriptApp.newTrigger('onChange_new')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange().create();
}

/**
 * Autofill the formula in `cell` down to the last row with data.
 * Cleaner replacement for v1's double-getNextDataRange trick.
 */
/**
 * Autofill the formula in `cell` down to `lastRow`. If `lastRow` is not
 * provided (or not a positive number above the cell's row), falls back
 * to the sheet's last-row-with-content via getLastRow().
 *
 * The optional `lastRow` parameter exists because writeStepBlock wants to
 * pre-fill ~200 rows of step formulas on a fresh FMS. At that point the
 * sheet only has content in row 7 (the form-data ARRAYFORMULA in A7),
 * so getLastRow() returns 7 and a no-arg autofill would no-op — leaving
 * the new step's formula in row 7 alone with rows 8..200 empty.
 *
 * Tier 1 (legacy) callers (executeTAT, executeTX, executeSpecificTime,
 * executeShowCondition, executeTimeDelay) continue to call this with one
 * argument and behave as in v1 — autofilling to whatever range the sheet
 * already has content for.
 */
function autofillFormulaDown(cell, lastRow) {
  var sheet = cell.getSheet();
  var startRow = cell.getRow();
  var col = cell.getColumn();
  var endRow = (typeof lastRow === 'number' && lastRow > startRow)
    ? lastRow
    : sheet.getLastRow();
  if (endRow > startRow) {
    var fillRange = sheet.getRange(startRow + 1, col, endRow - startRow, 1);
    cell.copyTo(fillRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  }
}

function requireActiveCell(cell, msg) {
  if (cell) return true;
  SpreadsheetApp.getUi().alert(msg);
  return false;
}

/** Soft warning — non-blocking — if the active cell's row-6 header isn't "Planned". */
function warnIfNotPlannedCell(cell, opName) {
  try {
    var sh = cell.getSheet();
    var h = String(sh.getRange(6, cell.getColumn()).getValue() || '').trim();
    if (h && !/^planned$/i.test(h)) {
      SpreadsheetApp.getActive().toast(
        'Heads up: this column\'s row-6 header is "' + h + '", not "Planned". ' + opName + ' is meant to apply to a Planned cell. Continue at your own risk.',
        'BMP Formulas', 7
      );
    }
  } catch (e) { /* row 6 may be empty in fresh sheets — ignore */ }
}

// ============================================================
//  onChange TRIGGER (unchanged behavior from v1)
// ============================================================

function onChange_new(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var active = ss.getActiveCell();
  if (!active) return;
  var row = active.getRow();
  var col = active.getColumn();
  var val = active.getValue();
  if (col <= 1) return;
  if (val === 'Done' || val === 'Yes' || val === 'No' || val === true) {
    var target = sheet.getRange(row, col - 1);
    if (target.getValue() === '' || target.getValue() === null) {
      target.setValue(new Date());
    }
  }
}