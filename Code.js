// Code.gs in Template - ONE TIME SETUP, never touch again

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Finance Toolbox")
    .addItem("New Rental Assessment", "setupRentalSheet")
    .addItem("Analyze Rental", "analyzeRental")
    .addSeparator()
    .addItem("Help & Documentation", "showDocs")
    .addToUi();
}

/**
 * Universal dispatcher - routes any function call to the Library
 * @param {string} fnName Function name to call
 * @param {...any} args Arguments to pass
 * @return {any} Result
 * @customfunction
 */
function FT(fnName, ...args) {
  return FinanceToolbox[fnName](...args);
}

// ─────────────────────────────────────────────────────────────────────────────
// RENTAL ASSESSMENT — SHEET SETUP & ANALYZE
// ─────────────────────────────────────────────────────────────────────────────

// Fixed row numbers (1-based) for every input and output cell in column B.
var ROWS = {
  // ── inputs ──
  purchasePrice: 2,
  purchaseDate: 3,
  loanAmount: 5,
  loanRate: 6,
  loanYears: 7,
  appreciation: 9,
  saleDate: 10,
  salePriceOverride: 11,
  monthlyRent: 13,
  vacancy: 14,
  annualMaintenance: 15,
  hoaStart: 17,
  hoaAnnualIncrease: 18,
  // ── outputs ──
  hMonths: 21,
  hYears: 22,
  estSalePrice: 23,
  monthlyPayment: 24,
  loanBalance: 25,
  principalPaid: 26,
  interestPaid: 27,
  hoaCost: 28,
  effectiveRent: 29,
  totalRent: 30,
  totalMaintenance: 31,
  netOpCashflow: 32,
  netSaleProceeds: 33,
  initialCash: 34,
  estGain: 35,
  totalROI: 36,
  annualizedROI: 37,
};

/**
 * Creates a NEW sheet and writes the Rental Assessment input/output layout.
 * Run via Finance Toolbox > New Rental Assessment.
 */
function setupRentalSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz =
    ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || "Etc/UTC";
  var stamp = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm");
  var sheet = ss.insertSheet("Rental Assessment " + stamp);
  ss.setActiveSheet(sheet);
  var A = 1,
    B = 2;

  function label(row, text, bold, bg) {
    var cell = sheet.getRange(row, A);
    cell.setValue(text);
    if (bold) cell.setFontWeight("bold");
    if (bg) cell.setBackground(bg);
  }
  function inputRow(row, lbl, defaultVal) {
    label(row, lbl, false, null);
    var c = sheet.getRange(row, B);
    if (defaultVal !== undefined) c.setValue(defaultVal);
    c.setBackground("#fff9c4"); // yellow = user input
    c.setNote("Enter value here");
  }
  function outputRow(row, lbl) {
    label(row, lbl, false, null);
    sheet.getRange(row, B).setBackground("#e8f5e9"); // green = calculated
  }

  // section headers
  label(1, "Rental Assessment Inputs", true, "#90caf9");
  label(20, "Assessment (Results)", true, "#a5d6a7");

  // input rows
  inputRow(ROWS.purchasePrice, "Purchase price");
  inputRow(ROWS.purchaseDate, "Purchase date");
  inputRow(ROWS.loanAmount, "Loan amount");
  inputRow(ROWS.loanRate, "Loan rate (annual)", 0.0675);
  inputRow(ROWS.loanYears, "Loan duration (years)", 30);
  inputRow(ROWS.appreciation, "Assumed annual appreciation", 0.03);
  inputRow(ROWS.saleDate, "Sale date");
  inputRow(ROWS.salePriceOverride, "Sale price (optional override)");
  inputRow(ROWS.monthlyRent, "Monthly rent");
  inputRow(ROWS.vacancy, "Annual vacancy", 0.05);
  inputRow(ROWS.annualMaintenance, "Annual maintenance (amount)", 0);
  inputRow(ROWS.hoaStart, "HOA at start (monthly)", 0);
  inputRow(ROWS.hoaAnnualIncrease, "Annual HOA increase", 0.03);

  // output rows
  outputRow(ROWS.hMonths, "Holding months");
  outputRow(ROWS.hYears, "Holding years");
  outputRow(ROWS.estSalePrice, "Estimated sale price");
  outputRow(ROWS.monthlyPayment, "Monthly mortgage payment");
  outputRow(ROWS.loanBalance, "Loan balance by sale date");
  outputRow(ROWS.principalPaid, "Principal paid by sale date");
  outputRow(ROWS.interestPaid, "Loan interest paid by sale date");
  outputRow(ROWS.hoaCost, "HOA cost by sale date");
  outputRow(ROWS.effectiveRent, "Effective monthly rent after vacancy");
  outputRow(ROWS.totalRent, "Total rent by sale date");
  outputRow(ROWS.totalMaintenance, "Total maintenance by sale date");
  outputRow(ROWS.netOpCashflow, "Net operating cashflow by sale date");
  outputRow(ROWS.netSaleProceeds, "Net sale proceeds");
  outputRow(ROWS.initialCash, "Initial cash invested");
  outputRow(ROWS.estGain, "Estimated gain by sale date");
  outputRow(ROWS.totalROI, "Total ROI on initial cash");
  outputRow(ROWS.annualizedROI, "Annualized ROI");

  // column widths
  sheet.setColumnWidth(A, 280);
  sheet.setColumnWidth(B, 200);

  SpreadsheetApp.getUi().alert(
    "Setup complete!\n\n" +
      "To add an Analyze button:\n" +
      "1. Insert > Drawing\n" +
      '2. Draw a shape and type "Analyze"\n' +
      "3. Save & close\n" +
      "4. Click the ⋮ menu on the drawing > Assign script > type: analyzeRental\n\n" +
      "Or use Finance Toolbox > Analyze Rental from the menu anytime.",
  );
}

/**
 * Reads inputs from fixed cell positions, computes all rental figures
 * via the FinanceToolbox library, and writes results back to the sheet.
 * Assign to a button drawing: right-click > Assign script > analyzeRental
 */
function analyzeRental() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  function get(row) {
    var v = sheet.getRange(row, 2).getValue();
    return v === "" || v === null || v === undefined ? null : v;
  }
  function set(row, val) {
    sheet.getRange(row, 2).setValue(val === null ? "" : val);
  }
  function ok() {
    for (var i = 0; i < arguments.length; i++) {
      if (arguments[i] === null) return false;
    }
    return true;
  }

  var purchasePrice = get(ROWS.purchasePrice);
  var purchaseDate = get(ROWS.purchaseDate);
  var loanAmount = get(ROWS.loanAmount);
  var loanRate = get(ROWS.loanRate);
  var loanYears = get(ROWS.loanYears);
  var appreciation = get(ROWS.appreciation);
  var saleDate = get(ROWS.saleDate);
  var salePriceOverride = get(ROWS.salePriceOverride);
  var monthlyRent = get(ROWS.monthlyRent);
  var vacancy = get(ROWS.vacancy);
  var annualMaintenance = get(ROWS.annualMaintenance) || 0;
  var hoaStart = get(ROWS.hoaStart);
  var hoaAnnualIncrease = get(ROWS.hoaAnnualIncrease);

  // holding period
  var hMonths = null;
  if (ok(purchaseDate, saleDate)) {
    var a = new Date(purchaseDate),
      b = new Date(saleDate);
    hMonths =
      (b.getFullYear() - a.getFullYear()) * 12 + (b.getMonth() - a.getMonth());
  }
  var hYears = hMonths !== null ? hMonths / 12 : null;

  // sale price
  var estSalePrice = null;
  if (ok(salePriceOverride) && salePriceOverride !== 0) {
    estSalePrice = salePriceOverride;
  } else if (ok(purchasePrice, appreciation, purchaseDate, saleDate)) {
    estSalePrice = FinanceToolbox.GROW_VALUE_BETWEEN_DATES(
      purchasePrice,
      appreciation,
      new Date(purchaseDate),
      new Date(saleDate),
    );
  }

  // loan
  var monthlyPayment = null;
  if (ok(loanAmount, loanRate, loanYears))
    monthlyPayment = FinanceToolbox.LOAN_MONTHLY_PAYMENT(
      loanAmount,
      loanRate,
      loanYears,
    );

  var loanBalance = null;
  if (ok(loanAmount, loanRate, loanYears, hMonths))
    loanBalance = FinanceToolbox.LOAN_BALANCE_AFTER_PAYMENTS(
      loanAmount,
      loanRate,
      loanYears,
      hMonths,
    );

  var principalPaid = ok(loanAmount, loanBalance)
    ? loanAmount - loanBalance
    : null;
  var interestPaid = ok(monthlyPayment, hMonths, principalPaid)
    ? monthlyPayment * hMonths - principalPaid
    : null;

  // HOA
  var hoaCost = null;
  if (ok(hoaStart, hoaAnnualIncrease, purchaseDate, saleDate))
    hoaCost = FinanceToolbox.TOTAL_HOA_COST(
      hoaStart,
      hoaAnnualIncrease,
      new Date(purchaseDate),
      new Date(saleDate),
    );

  // rent & cashflow
  var effectiveRent = ok(monthlyRent, vacancy)
    ? monthlyRent * (1 - vacancy)
    : null;
  var totalRent = ok(effectiveRent, hMonths) ? effectiveRent * hMonths : null;
  var totalMaintenance = hYears !== null ? annualMaintenance * hYears : null;
  var netOpCashflow = ok(totalRent, hoaCost, totalMaintenance)
    ? totalRent - hoaCost - totalMaintenance
    : null;

  // returns
  var netSaleProceeds = ok(estSalePrice, loanBalance)
    ? estSalePrice - loanBalance
    : null;
  var initialCash = ok(purchasePrice, loanAmount)
    ? purchasePrice - loanAmount
    : null;
  var estGain = ok(netSaleProceeds, netOpCashflow, initialCash)
    ? netSaleProceeds + netOpCashflow - initialCash
    : null;
  var totalROI =
    ok(estGain, initialCash) && initialCash !== 0
      ? estGain / initialCash
      : null;
  var annualizedROI =
    ok(totalROI, hYears) && hYears > 0
      ? Math.pow(1 + totalROI, 1 / hYears) - 1
      : null;

  // write results
  set(ROWS.hMonths, hMonths);
  set(ROWS.hYears, hYears);
  set(ROWS.estSalePrice, estSalePrice);
  set(ROWS.monthlyPayment, monthlyPayment);
  set(ROWS.loanBalance, loanBalance);
  set(ROWS.principalPaid, principalPaid);
  set(ROWS.interestPaid, interestPaid);
  set(ROWS.hoaCost, hoaCost);
  set(ROWS.effectiveRent, effectiveRent);
  set(ROWS.totalRent, totalRent);
  set(ROWS.totalMaintenance, totalMaintenance);
  set(ROWS.netOpCashflow, netOpCashflow);
  set(ROWS.netSaleProceeds, netSaleProceeds);
  set(ROWS.initialCash, initialCash);
  set(ROWS.estGain, estGain);
  set(ROWS.totalROI, totalROI);
  set(ROWS.annualizedROI, annualizedROI);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Analysis complete!",
    "Rental Assessment",
    3,
  );
}

function showDocs() {
  const html = HtmlService.createHtmlOutput(FinanceToolbox.getDocsHtml())
    .setTitle("Finance Toolbox Docs")
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Generates a rental assessment table with all calculations computed server-side.
 *
 * Usage:
 *   =RENTAL_ASSESSMENT_TABLE(purchasePrice, purchaseDate, loanAmount, loanRate, loanYears,
 *                             appreciation, saleDate, salePriceOverride,
 *                             monthlyRent, vacancy, annualMaintenance,
 *                             hoaStart, hoaAnnualIncrease)
 *
 * Leave any input as "" or omit trailing args to get blank results for dependent rows.
 *
 * @param {number} purchasePrice
 * @param {Date}   purchaseDate
 * @param {number} loanAmount
 * @param {number} loanRate        Annual rate, e.g. 0.0675
 * @param {number} loanYears
 * @param {number} appreciation    Annual appreciation rate, e.g. 0.03
 * @param {Date}   saleDate
 * @param {number} salePriceOverride  Leave blank to use appreciation estimate
 * @param {number} monthlyRent
 * @param {number} vacancy         Annual vacancy rate, e.g. 0.05
 * @param {number} annualMaintenance
 * @param {number} hoaStart        Monthly HOA at start
 * @param {number} hoaAnnualIncrease
 * @return {any[][]}
 * @customfunction
 */
function RENTAL_ASSESSMENT_TABLE(
  purchasePrice,
  purchaseDate,
  loanAmount,
  loanRate,
  loanYears,
  appreciation,
  saleDate,
  salePriceOverride,
  monthlyRent,
  vacancy,
  annualMaintenance,
  hoaStart,
  hoaAnnualIncrease,
) {
  var NA = "";

  // ── helpers ──────────────────────────────────────────────────────────────
  function isSet(v) {
    return v !== "" && v !== null && v !== undefined;
  }

  function holdingMonths(pd, sd) {
    if (!isSet(pd) || !isSet(sd)) return NA;
    var a = new Date(pd),
      b = new Date(sd);
    return (
      (b.getFullYear() - a.getFullYear()) * 12 + (b.getMonth() - a.getMonth())
    );
  }

  // ── derived values ────────────────────────────────────────────────────────
  var hMonths = holdingMonths(purchaseDate, saleDate);
  var hYears = isSet(hMonths) ? hMonths / 12 : NA;

  var estSalePrice = NA;
  if (isSet(salePriceOverride) && salePriceOverride !== 0) {
    estSalePrice = salePriceOverride;
  } else if (
    isSet(purchasePrice) &&
    isSet(appreciation) &&
    isSet(purchaseDate) &&
    isSet(saleDate)
  ) {
    estSalePrice = FinanceToolbox.GROW_VALUE_BETWEEN_DATES(
      purchasePrice,
      appreciation,
      new Date(purchaseDate),
      new Date(saleDate),
    );
  }

  var monthlyPayment = NA;
  if (isSet(loanAmount) && isSet(loanRate) && isSet(loanYears)) {
    monthlyPayment = FinanceToolbox.LOAN_MONTHLY_PAYMENT(
      loanAmount,
      loanRate,
      loanYears,
    );
  }

  var loanBalance = NA;
  if (
    isSet(loanAmount) &&
    isSet(loanRate) &&
    isSet(loanYears) &&
    isSet(hMonths)
  ) {
    loanBalance = FinanceToolbox.LOAN_BALANCE_AFTER_PAYMENTS(
      loanAmount,
      loanRate,
      loanYears,
      hMonths,
    );
  }

  var principalPaid =
    isSet(loanAmount) && isSet(loanBalance) ? loanAmount - loanBalance : NA;

  var interestPaid =
    isSet(monthlyPayment) && isSet(hMonths) && isSet(principalPaid)
      ? monthlyPayment * hMonths - principalPaid
      : NA;

  var hoaCost = NA;
  if (
    isSet(hoaStart) &&
    isSet(hoaAnnualIncrease) &&
    isSet(purchaseDate) &&
    isSet(saleDate)
  ) {
    hoaCost = FinanceToolbox.TOTAL_HOA_COST(
      hoaStart,
      hoaAnnualIncrease,
      new Date(purchaseDate),
      new Date(saleDate),
    );
  }

  var effectiveRent =
    isSet(monthlyRent) && isSet(vacancy) ? monthlyRent * (1 - vacancy) : NA;

  var totalRent =
    isSet(effectiveRent) && isSet(hMonths) ? effectiveRent * hMonths : NA;

  var maintenance =
    annualMaintenance === "" || !isSet(annualMaintenance)
      ? 0
      : annualMaintenance;
  var totalMaintenance = isSet(hYears) ? maintenance * hYears : NA;

  var netOpCashflow =
    isSet(totalRent) && isSet(hoaCost) && isSet(totalMaintenance)
      ? totalRent - hoaCost - totalMaintenance
      : NA;

  var netSaleProceeds =
    isSet(estSalePrice) && isSet(loanBalance) ? estSalePrice - loanBalance : NA;

  var initialCash =
    isSet(purchasePrice) && isSet(loanAmount) ? purchasePrice - loanAmount : NA;

  var estGain =
    isSet(netSaleProceeds) && isSet(netOpCashflow) && isSet(initialCash)
      ? netSaleProceeds + netOpCashflow - initialCash
      : NA;

  var totalROI =
    isSet(estGain) && isSet(initialCash) && initialCash !== 0
      ? estGain / initialCash
      : NA;

  var annualizedROI =
    isSet(totalROI) && isSet(hYears) && hYears > 0
      ? Math.pow(1 + totalROI, 1 / hYears) - 1
      : NA;

  // ── table ─────────────────────────────────────────────────────────────────
  return [
    ["Rental Assessment Inputs", ""],
    ["Purchase price", isSet(purchasePrice) ? purchasePrice : NA],
    ["Purchase date", isSet(purchaseDate) ? new Date(purchaseDate) : NA],
    ["", ""],
    ["Loan amount", isSet(loanAmount) ? loanAmount : NA],
    ["Loan rate (annual)", isSet(loanRate) ? loanRate : NA],
    ["Loan duration (years)", isSet(loanYears) ? loanYears : NA],
    ["", ""],
    ["Assumed annual appreciation", isSet(appreciation) ? appreciation : NA],
    ["Sale date", isSet(saleDate) ? new Date(saleDate) : NA],
    [
      "Sale price (optional override)",
      isSet(salePriceOverride) && salePriceOverride !== 0
        ? salePriceOverride
        : NA,
    ],
    ["", ""],
    ["Monthly rent", isSet(monthlyRent) ? monthlyRent : NA],
    ["Annual vacancy", isSet(vacancy) ? vacancy : NA],
    ["Annual maintenance (amount)", maintenance],
    ["", ""],
    ["HOA at start (monthly)", isSet(hoaStart) ? hoaStart : NA],
    ["Annual HOA increase", isSet(hoaAnnualIncrease) ? hoaAnnualIncrease : NA],
    ["", ""],
    ["Assessment", ""],
    ["Holding months", hMonths],
    ["Holding years", hYears],
    ["Estimated sale price", estSalePrice],
    ["Monthly mortgage payment", monthlyPayment],
    ["Loan balance by sale date", loanBalance],
    ["Principal paid by sale date", principalPaid],
    ["Loan interest paid by sale date", interestPaid],
    ["HOA cost by sale date", hoaCost],
    ["Effective monthly rent after vacancy", effectiveRent],
    ["Total rent by sale date", totalRent],
    ["Total maintenance by sale date", totalMaintenance],
    ["Net operating cashflow by sale date", netOpCashflow],
    ["Net sale proceeds", netSaleProceeds],
    ["Initial cash invested", initialCash],
    ["Estimated gain by sale date", estGain],
    ["Total ROI on initial cash", totalROI],
    ["Annualized ROI", annualizedROI],
  ];
}
