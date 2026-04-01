/**
 * 50/30/20 Personal Budget Tracker — Google Apps Script Builder
 *
 * HOW TO USE:
 * 1. Open a brand-new Google Sheet
 * 2. Go to Extensions → Apps Script
 * 3. Paste this entire script and click Run → buildBudgetTracker()
 * 4. Authorize when prompted — the script only touches this spreadsheet
 * 5. Delete the Apps Script project afterward if you like; the sheet is self-contained
 *
 * Every cell formula is native Google Sheets — no Apps Script runs at open time.
 */

// ─── Color Palette (hex) ────────────────────────────────────────────
var DARK_SAGE   = '#4A5E52';
var DUSTY_ROSE  = '#C49A8A';
var WARM_BEIGE  = '#F0EBE3';
var SOFT_SAGE   = '#D6DDD8';
var WHITE       = '#FFFFFF';
var DARK_TEXT   = '#2C2C2C';
var LIGHT_YELLOW = '#FFF9C4';
var LIGHT_GREEN = '#D6F0DC';
var LIGHT_RED   = '#F5C6C0';
var RED_TEXT    = '#C0392B';

// ─── MAIN ENTRY POINT ───────────────────────────────────────────────
function buildBudgetTracker() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename('50/30/20 Budget Tracker');

  // ── Step 1: Create 4 tabs with specific sheetIds ──────────────────
  setupTabs_(ss);

  var dashboard    = ss.getSheetByName('Dashboard');
  var transactions = ss.getSheetByName('Transactions');
  var budgetSetup  = ss.getSheetByName('Budget Setup');
  var reference    = ss.getSheetByName('Reference');

  // ── Step 2: Reference tab ─────────────────────────────────────────
  populateReference_(reference);

  // ── Step 3: Budget Setup tab ──────────────────────────────────────
  populateBudgetSetup_(budgetSetup);

  // ── Step 4 & 5: Transactions tab (structure + seed data) ──────────
  populateTransactions_(transactions);

  // ── Step 6: Dashboard tab ─────────────────────────────────────────
  populateDashboard_(dashboard);

  // ── Step 7: Charts ────────────────────────────────────────────────
  buildCharts_(dashboard);

  // ── Step 8: Formatting ────────────────────────────────────────────
  formatDashboard_(dashboard);
  formatTransactions_(transactions);
  formatBudgetSetup_(budgetSetup);

  // ── Step 9: Conditional formatting ────────────────────────────────
  addConditionalFormatting_(dashboard, transactions);

  // ── Step 10: Freeze, hide, protect ────────────────────────────────
  finalSettings_(dashboard, transactions, budgetSetup, reference);

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('✅  Budget Tracker built successfully!');
}

// =====================================================================
//  STEP 1 — TAB SETUP
// =====================================================================
function setupTabs_(ss) {
  // Rename the default sheet to Dashboard
  var sheets = ss.getSheets();
  sheets[0].setName('Dashboard');

  // Insert the other three tabs
  ss.insertSheet('Transactions', 1);
  ss.insertSheet('Budget Setup', 2);
  ss.insertSheet('Reference', 3);
}

// =====================================================================
//  STEP 2 — REFERENCE TAB
// =====================================================================
function populateReference_(sheet) {
  // Columns A:B — Sub-Category → Category
  var catLookup = [
    ['Groceries',        'Expenses'],
    ['Gas',              'Expenses'],
    ['Health',           'Expenses'],
    ['Dining Out',       'Expenses'],
    ['Clothes',          'Expenses'],
    ['Drugstore',        'Expenses'],
    ['Transport',        'Expenses'],
    ['Entertainment',    'Expenses'],
    ['Other',            'Expenses'],
    ['Rent/Mortgage',    'Bills'],
    ['Utilities',        'Bills'],
    ['Insurance',        'Bills'],
    ['Phone',            'Bills'],
    ['Internet',         'Bills'],
    ['Streaming',        'Subscriptions'],
    ['Gym',              'Subscriptions'],
    ['Loan Payment',     'Debts'],
    ['Credit Card',      'Debts'],
    ['401k',             'Savings'],
    ['Savings Transfer', 'Savings']
  ];
  sheet.getRange(1, 1, catLookup.length, 2).setValues(catLookup);

  // Columns D:E — Sub-Category → Bucket
  var bucketLookup = [
    ['Groceries',        'Need'],
    ['Gas',              'Need'],
    ['Health',           'Need'],
    ['Rent/Mortgage',    'Need'],
    ['Utilities',        'Need'],
    ['Insurance',        'Need'],
    ['Phone',            'Need'],
    ['Internet',         'Need'],
    ['Dining Out',       'Want'],
    ['Clothes',          'Want'],
    ['Drugstore',        'Want'],
    ['Entertainment',    'Want'],
    ['Transport',        'Want'],
    ['Streaming',        'Want'],
    ['Gym',              'Want'],
    ['Other',            'Want'],
    ['Loan Payment',     'Savings / Debts'],
    ['Credit Card',      'Savings / Debts'],
    ['401k',             'Savings / Debts'],
    ['Savings Transfer', 'Savings / Debts']
  ];
  sheet.getRange(1, 4, bucketLookup.length, 2).setValues(bucketLookup);

  // Column G — Month names
  var months = [
    ['January'],['February'],['March'],['April'],['May'],['June'],
    ['July'],['August'],['September'],['October'],['November'],['December']
  ];
  sheet.getRange(1, 7, 12, 1).setValues(months);
}

// =====================================================================
//  STEP 3 — BUDGET SETUP TAB
// =====================================================================
function populateBudgetSetup_(sheet) {
  // Section 1 — Income & Ratios
  sheet.getRange('A1:B1').merge().setValue('INCOME & RATIO SETTINGS')
       .setBackground(DARK_SAGE).setFontColor(WHITE).setFontWeight('bold');

  sheet.getRange('A2').setValue('Monthly Income');
  sheet.getRange('B2').setValue(5000).setNumberFormat('$#,##0.00').setBackground(LIGHT_YELLOW);
  sheet.getRange('A3').setValue('Need %');
  sheet.getRange('B3').setValue(50).setBackground(LIGHT_YELLOW);
  sheet.getRange('A4').setValue('Want %');
  sheet.getRange('B4').setValue(30).setBackground(LIGHT_YELLOW);
  sheet.getRange('A5').setValue('Savings/Debts %');
  sheet.getRange('B5').setValue(20).setBackground(LIGHT_YELLOW);
  sheet.getRange('A6').setValue('Ratio Check');
  sheet.getRange('B6').setFormula('=IF(B3+B4+B5=100,"✓ OK","⚠ Must = 100")');

  // Section 2 — Fixed Monthly Bills
  sheet.getRange('D8').setValue('FIXED ITEM');
  sheet.getRange('E8').setValue('AMOUNT');
  sheet.getRange('F8').setValue('CATEGORY');
  sheet.getRange('D8:F8').setBackground(DUSTY_ROSE).setFontColor(WHITE).setFontWeight('bold');

  var bills = [
    ['Rent/Mortgage', 1200, 'Bills'],
    ['Phone',         55,   'Bills'],
    ['Internet',      45,   'Bills'],
    ['Utilities',     80,   'Bills'],
    ['Streaming',     45,   'Subscriptions'],
    ['Gym',           60,   'Subscriptions'],
    ['Loan Payment',  500,  'Debts']
  ];
  sheet.getRange(9, 4, bills.length, 3).setValues(bills);
  sheet.getRange(9, 5, bills.length, 1).setNumberFormat('$#,##0.00');

  // Data validation dropdown for column F (F9:F30)
  var catRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Bills', 'Subscriptions', 'Debts'], true)
      .setAllowInvalid(false)
      .build();
  sheet.getRange('F9:F30').setDataValidation(catRule);
}

// =====================================================================
//  STEP 4 & 5 — TRANSACTIONS TAB (structure + seed data)
// =====================================================================
function populateTransactions_(sheet) {
  // Row 1: Headers
  var headers = ['DATE', 'AMOUNT', 'SUB-CATEGORY', 'CATEGORY', 'DESCRIPTION', 'BUCKET'];
  sheet.getRange(1, 1, 1, 6).setValues([headers])
       .setBackground(DARK_SAGE).setFontColor(WHITE).setFontWeight('bold');

  // Column C dropdown (C2:C200)
  var subCatRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(sheet.getParent().getSheetByName('Reference').getRange('A1:A20'), true)
      .setAllowInvalid(false)
      .build();
  sheet.getRange('C2:C200').setDataValidation(subCatRule);

  // Formulas for D and F columns (rows 2–200)
  var dFormulas = [];
  var fFormulas = [];
  for (var r = 2; r <= 200; r++) {
    dFormulas.push(['=IFERROR(VLOOKUP(C' + r + ',Reference!$A:$B,2,0),"")']);
    fFormulas.push(['=IFERROR(VLOOKUP(C' + r + ',Reference!$D:$E,2,0),"")']);
  }
  sheet.getRange(2, 4, 199, 1).setFormulas(dFormulas);
  sheet.getRange(2, 6, 199, 1).setFormulas(fFormulas);

  // ── Seed data (columns A, B, C, E only) ───────────────────────────
  var seed = [
    // JANUARY (20 rows)
    ['1/1/2024',  1200, 'Rent/Mortgage',    'Monthly rent'],
    ['1/1/2024',  80,   'Utilities',        'Electric bill'],
    ['1/2/2024',  100,  'Groceries',        'Weekly shop'],
    ['1/3/2024',  45,   'Internet',         ''],
    ['1/4/2024',  55,   'Phone',            ''],
    ['1/5/2024',  500,  'Loan Payment',     'Car loan'],
    ['1/5/2024',  60,   'Gym',              ''],
    ['1/6/2024',  45,   'Streaming',        'Netflix + Spotify'],
    ['1/7/2024',  150,  'Groceries',        ''],
    ['1/8/2024',  90,   'Dining Out',       'Birthday dinner'],
    ['1/10/2024', 200,  'Clothes',          'Winter jacket'],
    ['1/12/2024', 120,  'Gas',              ''],
    ['1/14/2024', 75,   'Drugstore',        ''],
    ['1/15/2024', 250,  'Savings Transfer', 'Monthly savings'],
    ['1/16/2024', 40,   'Entertainment',    'Cinema'],
    ['1/18/2024', 500,  'Other',            'Boat repair'],
    ['1/20/2024', 150,  'Health',           'Dentist'],
    ['1/22/2024', 50,   'Transport',        'Uber rides'],
    ['1/25/2024', 130,  'Dining Out',       ''],
    ['1/28/2024', 15,   'Clothes',          ''],
    // FEBRUARY (10 rows)
    ['2/1/2024',  1200, 'Rent/Mortgage',    ''],
    ['2/2/2024',  80,   'Utilities',        ''],
    ['2/3/2024',  120,  'Groceries',        ''],
    ['2/5/2024',  45,   'Streaming',        ''],
    ['2/7/2024',  200,  'Clothes',          "Valentine's gift"],
    ['2/10/2024', 75,   'Dining Out',       ''],
    ['2/14/2024', 500,  'Loan Payment',     ''],
    ['2/15/2024', 300,  'Savings Transfer', ''],
    ['2/18/2024', 55,   'Phone',            ''],
    ['2/22/2024', 90,   'Gas',              ''],
    // MARCH (10 rows)
    ['3/1/2024',  1200, 'Rent/Mortgage',    ''],
    ['3/3/2024',  95,   'Groceries',        ''],
    ['3/5/2024',  500,  'Loan Payment',     ''],
    ['3/8/2024',  60,   'Gym',              ''],
    ['3/10/2024', 130,  'Clothes',          ''],
    ['3/12/2024', 45,   'Streaming',        ''],
    ['3/15/2024', 250,  'Savings Transfer', ''],
    ['3/18/2024', 110,  'Dining Out',       ''],
    ['3/20/2024', 80,   'Utilities',        ''],
    ['3/25/2024', 55,   'Phone',            '']
  ];

  // Write columns A (date), B (amount), C (sub-category), E (description)
  for (var i = 0; i < seed.length; i++) {
    var row = i + 2;
    sheet.getRange(row, 1).setValue(seed[i][0]);  // date
    sheet.getRange(row, 2).setValue(seed[i][1]);  // amount
    sheet.getRange(row, 3).setValue(seed[i][2]);  // sub-category
    sheet.getRange(row, 5).setValue(seed[i][3]);  // description
  }

  // Format columns
  sheet.getRange('A2:A200').setNumberFormat('M/d/yyyy');
  sheet.getRange('B2:B200').setNumberFormat('$#,##0.00');
}

// =====================================================================
//  STEP 6 — DASHBOARD TAB
// =====================================================================
function populateDashboard_(sheet) {
  // Ensure enough columns
  if (sheet.getMaxColumns() < 9) sheet.insertColumnsAfter(sheet.getMaxColumns(), 9 - sheet.getMaxColumns());

  // ── ROW 1-2: Title banner ─────────────────────────────────────────
  sheet.getRange('A1:I2').merge()
       .setValue('50/30/20 DASHBOARD')
       .setBackground(WARM_BEIGE).setFontColor(DARK_SAGE)
       .setFontWeight('bold').setFontSize(20)
       .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── ROW 3-4: Month selector ───────────────────────────────────────
  sheet.getRange('A3:I3').merge()
       .setValue('WHICH MONTH DO YOU WANT TO IMPROVE?')
       .setBackground(DUSTY_ROSE).setFontColor(WHITE)
       .setFontWeight('bold').setFontSize(10)
       .setHorizontalAlignment('center');

  sheet.getRange('H4').setValue('SELECT MONTH →').setFontWeight('bold');

  // Month dropdown in I4
  var monthRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(sheet.getParent().getSheetByName('Reference').getRange('G1:G12'), true)
      .setAllowInvalid(false)
      .build();
  sheet.getRange('I4').setDataValidation(monthRule).setValue('January')
       .setFontWeight('bold').setFontSize(11);

  // ── ROW 6-13: Budget Summary Table ────────────────────────────────

  // Helper: month number in I6
  sheet.getRange('I6').setFormula('=MATCH(I4,Reference!$G$1:$G$12,0)')
       .setFontSize(8).setFontColor('#999999');

  // Row 6: Headers
  sheet.getRange('A6').setValue('SOURCE');
  sheet.getRange('C6').setValue('BUDGET');
  sheet.getRange('E6').setValue('REAL');
  sheet.getRange('A6:E6').setBackground(DUSTY_ROSE).setFontColor(WHITE).setFontWeight('bold');

  // Row 7: INCOME
  sheet.getRange('A7').setValue('INCOME');
  sheet.getRange('C7').setFormula("='Budget Setup'!B2");
  sheet.getRange('E7').setFormula("='Budget Setup'!B2");

  // Row 8: BILLS
  sheet.getRange('A8').setValue('BILLS');
  sheet.getRange('C8').setFormula('=IFERROR(SUMIFS(\'Budget Setup\'!E:E,\'Budget Setup\'!F:F,"Bills"),0)');
  sheet.getRange('E8').setFormula('=IFERROR(SUMIFS(Transactions!B:B,Transactions!D:D,"Bills",MONTH(Transactions!A:A),I$6),0)');

  // Row 9: SUBSCRIPTIONS
  sheet.getRange('A9').setValue('SUBSCRIPTIONS');
  sheet.getRange('C9').setFormula('=IFERROR(SUMIFS(\'Budget Setup\'!E:E,\'Budget Setup\'!F:F,"Subscriptions"),0)');
  sheet.getRange('E9').setFormula('=IFERROR(SUMIFS(Transactions!B:B,Transactions!D:D,"Subscriptions",MONTH(Transactions!A:A),I$6),0)');

  // Row 10: EXPENSES
  sheet.getRange('A10').setValue('EXPENSES');
  sheet.getRange('C10').setFormula("=IFERROR('Budget Setup'!B2*('Budget Setup'!B3/100)-C8-C9,0)");
  sheet.getRange('E10').setFormula('=IFERROR(SUMIFS(Transactions!B:B,Transactions!D:D,"Expenses",MONTH(Transactions!A:A),I$6),0)');

  // Row 11: SAVINGS
  sheet.getRange('A11').setValue('SAVINGS');
  sheet.getRange('C11').setFormula("=IFERROR('Budget Setup'!B2*('Budget Setup'!B5/100)*0.5,0)");
  sheet.getRange('E11').setFormula('=IFERROR(SUMIFS(Transactions!B:B,Transactions!D:D,"Savings",MONTH(Transactions!A:A),I$6),0)');

  // Row 12: DEBTS
  sheet.getRange('A12').setValue('DEBTS');
  sheet.getRange('C12').setFormula("=IFERROR('Budget Setup'!B2*('Budget Setup'!B5/100)*0.5,0)");
  sheet.getRange('E12').setFormula('=IFERROR(SUMIFS(Transactions!B:B,Transactions!D:D,"Debts",MONTH(Transactions!A:A),I$6),0)');

  // Row 13: AMOUNT LEFT
  sheet.getRange('A13').setValue('AMOUNT LEFT');
  sheet.getRange('E13').setFormula('=IFERROR(E7-SUM(E8:E12),0)').setFontWeight('bold');

  // Currency formatting for summary table
  sheet.getRange('C7:C13').setNumberFormat('$#,##0.00');
  sheet.getRange('E7:E13').setNumberFormat('$#,##0.00');

  // ── ROW 15-20: Goal Breakdown Table ───────────────────────────────
  sheet.getRange('A15:E15').merge()
       .setValue('MY GOAL BREAKDOWN')
       .setBackground(DARK_SAGE).setFontColor(WHITE).setFontWeight('bold');

  sheet.getRange('A16').setValue('CATEGORY');
  sheet.getRange('C16').setValue('%');
  sheet.getRange('E16').setValue('$');
  sheet.getRange('A16:E16').setBackground(DUSTY_ROSE).setFontColor(WHITE).setFontWeight('bold');

  sheet.getRange('A17').setValue('Need');
  sheet.getRange('C17').setFormula("=TEXT('Budget Setup'!B3/100,\"0%\")");
  sheet.getRange('E17').setFormula("='Budget Setup'!B2*('Budget Setup'!B3/100)");

  sheet.getRange('A18').setValue('Want');
  sheet.getRange('C18').setFormula("=TEXT('Budget Setup'!B4/100,\"0%\")");
  sheet.getRange('E18').setFormula("='Budget Setup'!B2*('Budget Setup'!B4/100)");

  sheet.getRange('A19').setValue('Savings / Debts');
  sheet.getRange('C19').setFormula("=TEXT('Budget Setup'!B5/100,\"0%\")");
  sheet.getRange('E19').setFormula("='Budget Setup'!B2*('Budget Setup'!B5/100)");

  sheet.getRange('A20').setValue('TOTAL');
  sheet.getRange('C20').setValue('100%');
  sheet.getRange('E20').setFormula('=SUM(E17:E19)');

  sheet.getRange('E17:E20').setNumberFormat('$#,##0.00');

  // ── ROW 22-27: Real Breakdown Table ───────────────────────────────
  sheet.getRange('A22:E22').merge()
       .setValue('MY REAL BUDGET BREAKDOWN')
       .setBackground(DARK_SAGE).setFontColor(WHITE).setFontWeight('bold');

  sheet.getRange('A23').setValue('CATEGORY');
  sheet.getRange('C23').setValue('%');
  sheet.getRange('E23').setValue('$');
  sheet.getRange('A23:E23').setBackground(DUSTY_ROSE).setFontColor(WHITE).setFontWeight('bold');

  sheet.getRange('A24').setValue('Need');
  sheet.getRange('E24').setFormula('=IFERROR(SUMIFS(Transactions!B:B,Transactions!F:F,"Need",MONTH(Transactions!A:A),I$6),0)');
  sheet.getRange('C24').setFormula('=IFERROR(E24/E7,0)').setNumberFormat('0%');

  sheet.getRange('A25').setValue('Want');
  sheet.getRange('E25').setFormula('=IFERROR(SUMIFS(Transactions!B:B,Transactions!F:F,"Want",MONTH(Transactions!A:A),I$6),0)');
  sheet.getRange('C25').setFormula('=IFERROR(E25/E7,0)').setNumberFormat('0%');

  sheet.getRange('A26').setValue('Savings / Debts');
  sheet.getRange('E26').setFormula('=IFERROR(SUMIFS(Transactions!B:B,Transactions!F:F,"Savings / Debts",MONTH(Transactions!A:A),I$6),0)');
  sheet.getRange('C26').setFormula('=IFERROR(E26/E7,0)').setNumberFormat('0%');

  sheet.getRange('A27').setValue('TOTAL');
  sheet.getRange('C27').setFormula('=IFERROR(SUM(C24:C26),0)').setNumberFormat('0%');
  sheet.getRange('E27').setFormula('=SUM(E24:E26)');

  sheet.getRange('E24:E27').setNumberFormat('$#,##0.00');

  // ── ROW 29+: Transaction Viewer ───────────────────────────────────
  sheet.getRange('A29:F29').merge()
       .setValue('TRANSACTION TRACKER')
       .setBackground(DARK_SAGE).setFontColor(WHITE).setFontWeight('bold');

  var txHeaders = ['DATE', 'AMOUNT', 'SUB-CATEGORY', 'CATEGORY', 'DESCRIPTION', 'BUCKET'];
  sheet.getRange(30, 1, 1, 6).setValues([txHeaders])
       .setBackground(DUSTY_ROSE).setFontColor(WHITE).setFontWeight('bold');

  // QUERY formula to pull filtered transactions
  sheet.getRange('A31').setFormula(
    '=IFERROR(QUERY(Transactions!A:F,' +
    '"SELECT A,B,C,D,E,F WHERE A IS NOT NULL AND MONTH(A)+1 = "&I$6&"' +
    ' ORDER BY A DESC LABEL A \'\',B \'\',C \'\',D \'\',E \'\',F \'\'"),"")'
  );
}

// =====================================================================
//  STEP 7 — CHARTS
// =====================================================================
function buildCharts_(sheet) {
  // Chart 1: Grouped Column — Goal vs Real
  var chart1 = sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sheet.getRange('A17:A19'))   // labels
      .addRange(sheet.getRange('E17:E19'))   // Goal series
      .addRange(sheet.getRange('E24:E26'))   // Real series
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(0)
      .setOption('title', 'MY GOAL vs MY REAL BUDGET')
      .setOption('legend', {position: 'top'})
      .setOption('series', {
        0: {color: DARK_SAGE,  labelInLegend: 'MY GOAL'},
        1: {color: DUSTY_ROSE, labelInLegend: 'MY BUDGET'}
      })
      .setOption('hAxis', {format: '$#,##0'})
      .setOption('isStacked', false)
      .setPosition(6, 7, 0, 0)
      .setOption('width', 400)
      .setOption('height', 280)
      .build();
  sheet.insertChart(chart1);

  // Chart 2: Donut — Goal Breakdown
  var chart2 = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange('A17:A19'))
      .addRange(sheet.getRange('E17:E19'))
      .setNumHeaders(0)
      .setOption('title', 'MY GOAL BREAKDOWN')
      .setOption('pieHole', 0.5)
      .setOption('legend', {position: 'right'})
      .setOption('slices', {
        0: {color: DARK_SAGE},
        1: {color: DUSTY_ROSE},
        2: {color: '#8FA898'}
      })
      .setPosition(15, 7, 0, 0)
      .setOption('width', 320)
      .setOption('height', 260)
      .build();
  sheet.insertChart(chart2);

  // Chart 3: Donut — Real Breakdown
  var chart3 = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange('A24:A26'))
      .addRange(sheet.getRange('E24:E26'))
      .setNumHeaders(0)
      .setOption('title', 'MY REAL BUDGET BREAKDOWN')
      .setOption('pieHole', 0.5)
      .setOption('legend', {position: 'right'})
      .setOption('slices', {
        0: {color: DARK_SAGE},
        1: {color: DUSTY_ROSE},
        2: {color: '#8FA898'}
      })
      .setPosition(22, 7, 0, 0)
      .setOption('width', 320)
      .setOption('height', 260)
      .build();
  sheet.insertChart(chart3);
}

// =====================================================================
//  STEP 8 — FORMATTING
// =====================================================================
function formatDashboard_(sheet) {
  // Column widths
  sheet.setColumnWidth(1, 160);  // A
  sheet.setColumnWidth(2, 20);   // B
  sheet.setColumnWidth(3, 100);  // C
  sheet.setColumnWidth(4, 20);   // D
  sheet.setColumnWidth(5, 100);  // E
  sheet.setColumnWidth(6, 20);   // F (spacer)
  sheet.setColumnWidth(7, 140);  // G
  sheet.setColumnWidth(8, 140);  // H
  sheet.setColumnWidth(9, 140);  // I

  // Alternating row colors for budget summary (rows 7-13)
  for (var r = 7; r <= 13; r++) {
    var bg = (r % 2 === 1) ? WHITE : WARM_BEIGE;
    sheet.getRange(r, 1, 1, 5).setBackground(bg);
  }

  // Alternating row colors for goal breakdown (rows 17-20)
  for (var r = 17; r <= 20; r++) {
    var bg = (r % 2 === 1) ? WHITE : WARM_BEIGE;
    sheet.getRange(r, 1, 1, 5).setBackground(bg);
  }

  // Alternating row colors for real breakdown (rows 24-27)
  for (var r = 24; r <= 27; r++) {
    var bg = (r % 2 === 0) ? WHITE : WARM_BEIGE;
    sheet.getRange(r, 1, 1, 5).setBackground(bg);
  }

  // Borders for tables
  sheet.getRange('A6:E13').setBorder(true, true, true, true, true, true,
      '#C8C8C8', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('A15:E20').setBorder(true, true, true, true, true, true,
      '#C8C8C8', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('A22:E27').setBorder(true, true, true, true, true, true,
      '#C8C8C8', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('A29:F30').setBorder(true, true, true, true, true, true,
      '#C8C8C8', SpreadsheetApp.BorderStyle.SOLID);

  // Percentage cells
  sheet.getRange('C17:C20').setNumberFormat('0%');
  sheet.getRange('C24:C27').setNumberFormat('0%');
}

function formatTransactions_(sheet) {
  // Column widths
  sheet.setColumnWidth(1, 100);  // A
  sheet.setColumnWidth(2, 90);   // B
  sheet.setColumnWidth(3, 140);  // C
  sheet.setColumnWidth(4, 110);  // D
  sheet.setColumnWidth(5, 180);  // E
  sheet.setColumnWidth(6, 130);  // F

  // Alternating rows (2–41 for seed data)
  for (var r = 2; r <= 41; r++) {
    var bg = (r % 2 === 0) ? WHITE : WARM_BEIGE;
    sheet.getRange(r, 1, 1, 6).setBackground(bg);
  }
}

function formatBudgetSetup_(sheet) {
  // Section header already styled in populate
  // Input cells yellow bg already set
  // Fixed items header row 8 already styled
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 120);
}

// =====================================================================
//  STEP 9 — CONDITIONAL FORMATTING
// =====================================================================
function addConditionalFormatting_(dashboard, transactions) {
  // Rule 1: AMOUNT LEFT positive (E13)
  var rule1 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground(LIGHT_GREEN)
      .setFontColor(DARK_SAGE)
      .setBold(true)
      .setRanges([dashboard.getRange('E13')])
      .build();

  // Rule 2: AMOUNT LEFT negative (E13)
  var rule2 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground(LIGHT_RED)
      .setFontColor(RED_TEXT)
      .setBold(true)
      .setRanges([dashboard.getRange('E13')])
      .build();

  // Rule 3: Real breakdown over budget (each row)
  var rule3a = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=E24>E17')
      .setBackground(LIGHT_RED)
      .setRanges([dashboard.getRange('E24')])
      .build();
  var rule3b = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=E25>E18')
      .setBackground(LIGHT_RED)
      .setRanges([dashboard.getRange('E25')])
      .build();
  var rule3c = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=E26>E19')
      .setBackground(LIGHT_RED)
      .setRanges([dashboard.getRange('E26')])
      .build();

  // Rule 4: Budget summary real > budget (E8:E12)
  var rule4 = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=E8>C8')
      .setFontColor(RED_TEXT)
      .setRanges([dashboard.getRange('E8:E12')])
      .build();

  dashboard.setConditionalFormatRules([rule1, rule2, rule3a, rule3b, rule3c, rule4]);

  // Rule 5: Large transactions on Transactions tab (B2:B200)
  var rule5 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(300)
      .setBackground(WARM_BEIGE)
      .setBold(true)
      .setRanges([transactions.getRange('B2:B200')])
      .build();

  transactions.setConditionalFormatRules([rule5]);
}

// =====================================================================
//  STEP 10 — FREEZE, HIDE, PROTECT
// =====================================================================
function finalSettings_(dashboard, transactions, budgetSetup, reference) {
  // Freeze rows
  dashboard.setFrozenRows(4);
  transactions.setFrozenRows(1);
  budgetSetup.setFrozenRows(1);

  // Hide Reference tab
  reference.hideSheet();

  // Protect formula columns D and F on Transactions
  var protectD = transactions.getRange('D1:D200').protect()
      .setDescription('Formula column — Category (auto-calculated)')
      .setWarningOnly(true);
  var protectF = transactions.getRange('F1:F200').protect()
      .setDescription('Formula column — Bucket (auto-calculated)')
      .setWarningOnly(true);
}
