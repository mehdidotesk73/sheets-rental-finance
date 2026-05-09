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
  createRentalAssessmentSheetLayout(sheet);

  SpreadsheetApp.getUi().alert(
    "Setup complete!\n\n" +
      "A play-button image was added in column C and linked to analyzeRental.\n\n" +
      "You can also use Finance Toolbox > Analyze Rental from the menu anytime.",
  );
}

/**
 * Reads inputs from fixed cell positions, computes all rental figures
 * via the FinanceToolbox library, and writes results back to the sheet.
 * Assign to a button drawing: right-click > Assign script > analyzeRental
 */
function analyzeRental() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  clearRentalOutputValues(sheet);
  var inputs = getRentalInputValues(sheet);
  var outputs = calculateRentalAssessment(inputs);
  writeRentalOutputValues(sheet, outputs);

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
