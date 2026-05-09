var ANALYZE_BUTTON_URL =
  "https://raw.githubusercontent.com/mehdidotesk73/sheets-rental-finance/main/src/assets/square-play.png";

var RENTAL_INPUT_FIELDS = [
  {
    key: "purchasePrice",
    label: "Purchase price",
    defaultValue: 250000,
    format: "$#,##0.00",
  },
  {
    key: "purchaseDate",
    label: "Purchase date",
    defaultValue: "",
    format: "yyyy-mm-dd",
  },
  {
    key: "loanToValueRatio",
    label: "Loan-to-value ratio",
    defaultValue: 0.95,
    format: "0.00%",
  },
  {
    key: "downPaymentAssistance",
    label: "Down payment assistance",
    defaultValue: 7500,
    format: "$#,##0.00",
  },
  {
    key: "dpaForgivenessMonths",
    label: "DPA forgiveness months",
    defaultValue: 36,
    format: "0",
  },
  {
    key: "annualInterestRate",
    label: "Annual interest rate",
    defaultValue: 0.035,
    format: "0.00%",
  },
  {
    key: "loanTermYears",
    label: "Loan term years",
    defaultValue: 30,
    format: "0",
  },
  {
    key: "monthlyEscrow",
    label: "Monthly escrow",
    defaultValue: 295,
    format: "$#,##0.00",
  },
  {
    key: "closingCosts",
    label: "Closing costs",
    defaultValue: 0,
    format: "$#,##0.00",
    note: "All out-of-pocket closing costs (lender fees, title, prepaid items, etc.). Added to initial cash invested.",
  },
  {
    key: "hoaInitialMonthly",
    label: "HOA initial monthly",
    defaultValue: 390,
    format: "$#,##0.00",
  },
  {
    key: "hoaAnnualGrowthRate",
    label: "HOA annual growth rate",
    defaultValue: 0.055,
    format: "0.00%",
  },
  {
    key: "initialMonthlyRent",
    label: "Initial monthly rent",
    defaultValue: 1900,
    format: "$#,##0.00",
  },
  {
    key: "rentAnnualGrowthRate",
    label: "Rent annual growth rate",
    defaultValue: 0.02,
    format: "0.00%",
  },
  {
    key: "tenantTurnoverIntervalMonths",
    label: "Tenant turnover interval months",
    defaultValue: 24,
    format: "0",
  },
  {
    key: "vacancyMonthsPerTurnover",
    label: "Vacancy months per turnover",
    defaultValue: 1,
    format: "0.00",
  },
  {
    key: "costPerTurnover",
    label: "Cost per turnover",
    defaultValue: 1500,
    format: "$#,##0.00",
  },
  {
    key: "annualMaintenanceCost",
    label: "Annual maintenance cost",
    defaultValue: 1500,
    format: "$#,##0.00",
  },
  {
    key: "annualAppreciationRate",
    label: "Annual appreciation rate",
    defaultValue: 0.025,
    format: "0.00%",
  },
  {
    key: "saleDate",
    label: "Sale date",
    defaultValue: "",
    format: "yyyy-mm-dd",
  },
  {
    key: "longTermCapitalGainsRate",
    label: "Long-term capital gains rate",
    defaultValue: 0.15,
    format: "0.00%",
  },
  {
    key: "depreciationRecaptureRate",
    label: "Depreciation recapture rate",
    defaultValue: 0.25,
    format: "0.00%",
  },
  {
    key: "totalDepreciationTaken",
    label: "Total depreciation taken",
    defaultValue: 20000,
    format: "$#,##0.00",
  },
  {
    key: "agentFeeRate",
    label: "Agent fee rate",
    defaultValue: 0.06,
    format: "0.00%",
  },
  {
    key: "inflationDiscountRate",
    label: "Inflation discount rate",
    defaultValue: 0.07,
    format: "0.00%",
  },
];

var RENTAL_OUTPUT_FIELDS = [
  {
    key: "loanPrincipal",
    label: "Loan principal",
    format: "$#,##0.00",
    note: "Purchase price × loan-to-value ratio.",
  },
  {
    key: "cashAtClose",
    label: "Cash at close",
    format: "$#,##0.00",
    note: "Down payment (purchase price − loan principal − DPA) + closing costs. Total initial cash invested.",
  },
  {
    key: "monthlyInterestRate",
    label: "Monthly interest rate",
    format: "0.0000%",
    note: "Annual interest rate ÷ 12.",
  },
  {
    key: "loanTermMonths",
    label: "Loan term months",
    format: "0",
    note: "Loan term years × 12.",
  },
  {
    key: "monthlyPrincipalAndInterest",
    label: "Monthly principal & interest",
    format: "$#,##0.00",
    note: "Monthly amortized payment from loan inputs.",
  },
  {
    key: "totalMonthlyPayment",
    label: "Total monthly payment",
    format: "$#,##0.00",
    note: "Monthly principal & interest + monthly escrow.",
  },
  {
    key: "ownershipMonths",
    label: "Ownership months",
    format: "0",
    note: "Months between purchase date and sale date.",
  },
  {
    key: "ownershipYears",
    label: "Ownership years",
    format: "0.00",
    note: "Ownership months ÷ 12.",
  },
  {
    key: "remainingBalanceAtSale",
    label: "Remaining balance at sale",
    format: "$#,##0.00",
    note: "Loan balance after ownership months of payments.",
  },
  {
    key: "totalInterestPaid",
    label: "Total interest paid",
    format: "$#,##0.00",
    note: "Total principal & interest paid − principal repaid.",
  },
  {
    key: "totalPrincipalPaid",
    label: "Total principal paid",
    format: "$#,##0.00",
    note: "Loan principal − remaining balance at sale.",
  },
  {
    key: "totalEscrowPaid",
    label: "Total escrow paid",
    format: "$#,##0.00",
    note: "Monthly escrow × ownership months.",
  },
  {
    key: "dpaMonthsForgiven",
    label: "DPA months forgiven",
    format: "0",
    note: "Minimum of ownership months and DPA forgiveness months.",
  },
  {
    key: "dpaRemainingAtSale",
    label: "DPA remaining at sale",
    format: "$#,##0.00",
    note: "Unforgiven down payment assistance balance at sale.",
  },
  {
    key: "averageMonthlyInterest",
    label: "Average monthly interest",
    format: "$#,##0.00",
    note: "Total interest paid ÷ ownership months.",
  },
  {
    key: "totalVacancyCost",
    label: "Total vacancy cost",
    format: "$#,##0.00",
    note: "Vacancy cost per turnover × number of turnovers. Covers PM fees, travel, touch-ups — excludes lost rent (accounted separately).",
  },
  {
    key: "totalHoaPaid",
    label: "Total HOA paid",
    format: "$#,##0.00",
    note: "Annual HOA sum with annual growth over ownership horizon.",
  },
  {
    key: "totalGrossRentCollected",
    label: "Total gross rent collected",
    format: "$#,##0.00",
    note: "Annual gross rent sum with annual rent growth and vacancy months.",
  },
  {
    key: "numberOfTurnovers",
    label: "Number of turnovers",
    format: "0",
    note: "Floor of ownership months ÷ turnover interval.",
  },
  {
    key: "totalTurnoverCost",
    label: "Total turnover cost",
    format: "$#,##0.00",
    note: "Sum of turnover-month rents using rent growth by year.",
  },
  {
    key: "totalMaintenanceCost",
    label: "Total maintenance cost",
    format: "$#,##0.00",
    note: "Annual maintenance cost × ownership years.",
  },
  {
    key: "netCashReceived",
    label: "Net cash received (operations)",
    format: "$#,##0.00",
    note: "Total rent collected − all cash outflows (interest, principal, escrow, HOA, maintenance, turnover, vacancy). Matches actual cash landing in bank account.",
  },
  {
    key: "totalCashOutlaid",
    label: "Total cash outlaid (net)",
    format: "$#,##0.00",
    note: "Cash at close − net operating cash. Negative means rental income more than covered all costs including initial investment.",
  },
  {
    key: "projectedSaleValue",
    label: "Projected sale value",
    format: "$#,##0.00",
    note: "Purchase price grown by annual appreciation over ownership years.",
  },
  {
    key: "grossSaleProceeds",
    label: "Gross sale proceeds",
    format: "$#,##0.00",
    note: "Projected sale value − loan balance − DPA remaining − agent fee.",
  },
  {
    key: "adjustedCostBasis",
    label: "Adjusted cost basis",
    format: "$#,##0.00",
    note: "Purchase price − total depreciation taken.",
  },
  {
    key: "capitalGain",
    label: "Capital gain",
    format: "$#,##0.00",
    note: "Sale price − agent fee − adjusted cost basis. Loan repayment is excluded (it doesn't reduce taxable gain).",
  },
  {
    key: "depreciationRecaptureTax",
    label: "Depreciation recapture tax",
    format: "$#,##0.00",
    note: "Total depreciation taken × recapture rate.",
  },
  {
    key: "capitalGainsTax",
    label: "Capital gains tax",
    format: "$#,##0.00",
    note: "Max(0, capital gain − depreciation) × LTCG rate.",
  },
  {
    key: "totalTaxOnSale",
    label: "Total tax on sale",
    format: "$#,##0.00",
    note: "Depreciation recapture tax + capital gains tax.",
  },
  {
    key: "netSaleProceeds",
    label: "Net sale proceeds",
    format: "$#,##0.00",
    note: "Gross sale proceeds − total tax on sale.",
  },
  {
    key: "inflationAdjustedProceeds",
    label: "Inflation-adjusted proceeds",
    format: "$#,##0.00",
    note: "Net sale proceeds adjusted to purchase-date dollars.",
  },
  {
    key: "returnOnInvestment",
    label: "Return on investment",
    format: "0.00%",
    note: "(Net cash received + net sale proceeds − cash at close) ÷ cash at close.",
  },
  {
    key: "inflationAdjustedRoi",
    label: "Inflation-adjusted ROI",
    format: "0.00%",
    note: "(Net cash received + inflation-adjusted proceeds − cash at close) ÷ cash at close.",
  },
  {
    key: "annualizedReturnCagr",
    label: "Annualized return (CAGR)",
    format: "0.00%",
    note: "Annualized growth from cash at close to (net cash received + net sale proceeds).",
  },
  {
    key: "inflationAdjustedCagr",
    label: "Inflation-adjusted CAGR",
    format: "0.00%",
    note: "Annualized growth from cash at close to (net cash received + inflation-adjusted proceeds).",
  },
];

var INPUT_CATEGORY_BY_KEY = {
  purchasePrice: "Purchase & Financing",
  purchaseDate: "Purchase & Financing",
  loanToValueRatio: "Purchase & Financing",
  downPaymentAssistance: "Purchase & Financing",
  dpaForgivenessMonths: "Purchase & Financing",
  annualInterestRate: "Purchase & Financing",
  loanTermYears: "Purchase & Financing",
  monthlyEscrow: "Purchase & Financing",
  closingCosts: "Purchase & Financing",
  hoaInitialMonthly: "HOA",
  hoaAnnualGrowthRate: "HOA",
  initialMonthlyRent: "Rental Income",
  rentAnnualGrowthRate: "Rental Income",
  tenantTurnoverIntervalMonths: "Vacancy & Maintenance",
  vacancyMonthsPerTurnover: "Vacancy & Maintenance",
  costPerTurnover: "Vacancy & Maintenance",
  annualMaintenanceCost: "Vacancy & Maintenance",
  annualAppreciationRate: "Sale & Tax Assumptions",
  saleDate: "Sale & Tax Assumptions",
  longTermCapitalGainsRate: "Sale & Tax Assumptions",
  depreciationRecaptureRate: "Sale & Tax Assumptions",
  totalDepreciationTaken: "Sale & Tax Assumptions",
  agentFeeRate: "Sale & Tax Assumptions",
  inflationDiscountRate: "Sale & Tax Assumptions",
};

var OUTPUT_STEP_BY_KEY = {
  loanPrincipal: "Step 1 — Financing Setup",
  cashAtClose: "Step 1 — Financing Setup",
  monthlyInterestRate: "Step 1 — Financing Setup",
  loanTermMonths: "Step 1 — Financing Setup",
  monthlyPrincipalAndInterest: "Step 1 — Financing Setup",
  totalMonthlyPayment: "Step 1 — Financing Setup",
  ownershipMonths: "Step 2 — Ownership Period",
  ownershipYears: "Step 2 — Ownership Period",
  remainingBalanceAtSale: "Step 3 — Amortization",
  totalInterestPaid: "Step 3 — Amortization",
  totalPrincipalPaid: "Step 3 — Amortization",
  totalEscrowPaid: "Step 3 — Amortization",
  dpaMonthsForgiven: "Step 4 — DPA Forgiveness",
  dpaRemainingAtSale: "Step 4 — DPA Forgiveness",
  averageMonthlyInterest: "Step 5 — Vacancy Cost",
  totalVacancyCost: "Step 5 — Vacancy Cost",
  totalHoaPaid: "Step 6 — HOA",
  totalGrossRentCollected: "Step 7 — Rental Income",
  numberOfTurnovers: "Step 8 — Turnover",
  totalTurnoverCost: "Step 8 — Turnover",
  totalMaintenanceCost: "Step 9 — Maintenance",
  netCashReceived: "Step 10 — Total Cash Outlaid",
  totalCashOutlaid: "Step 10 — Total Cash Outlaid",
  projectedSaleValue: "Step 11 — Property Value At Sale",
  grossSaleProceeds: "Step 12 — Gross Sale Proceeds",
  adjustedCostBasis: "Step 13 — Tax",
  capitalGain: "Step 13 — Tax",
  depreciationRecaptureTax: "Step 13 — Tax",
  capitalGainsTax: "Step 13 — Tax",
  totalTaxOnSale: "Step 13 — Tax",
  netSaleProceeds: "Step 14 — Net Proceeds",
  inflationAdjustedProceeds: "Step 15 — Inflation Adjusted",
  returnOnInvestment: "Step 16 — Return Metrics",
  inflationAdjustedRoi: "Step 16 — Return Metrics",
  annualizedReturnCagr: "Step 16 — Return Metrics",
  inflationAdjustedCagr: "Step 16 — Return Metrics",
};

// ─── Category & Step visual metadata ───────────────────────────────────────

var INPUT_CATEGORY_META = {
  "Purchase & Financing": {
    label: "🏠  Purchase & Financing",
    headerBg: "#1565C0",
    rowBg: "#E3F2FD",
  },
  HOA: { label: "🏢  HOA", headerBg: "#6A1B9A", rowBg: "#EDE7F6" },
  "Rental Income": {
    label: "💵  Rental Income",
    headerBg: "#1B5E20",
    rowBg: "#E8F5E9",
  },
  "Vacancy & Maintenance": {
    label: "🔧  Vacancy & Maintenance",
    headerBg: "#BF360C",
    rowBg: "#FBE9E7",
  },
  "Sale & Tax Assumptions": {
    label: "📈  Sale & Tax Assumptions",
    headerBg: "#E65100",
    rowBg: "#FFF3E0",
  },
};

// Alternating row-tint flavors for output steps (A=lighter, B=slightly grey)
var OUTPUT_STEP_META = (function buildStepMeta() {
  var uniqueSteps = [];
  var seen = {};
  for (var i = 0; i < RENTAL_OUTPUT_FIELDS.length; i++) {
    var s = OUTPUT_STEP_BY_KEY[RENTAL_OUTPUT_FIELDS[i].key];
    if (!seen[s]) {
      seen[s] = true;
      uniqueSteps.push(s);
    }
  }
  var meta = {};
  for (var j = 0; j < uniqueSteps.length; j++) {
    meta[uniqueSteps[j]] = { rowBg: j % 2 === 0 ? "#FFFFFF" : "#F5F7FA" };
  }
  return meta;
})();

// ─── Row layout (includes category & step header rows) ─────────────────────

var RENTAL_LAYOUT = (function buildRentalLayout() {
  var rows = {};
  var categoryHeaderRows = {}; // category name → row
  var stepHeaderRows = {}; // step name → row
  var row = 1;

  var inputMainHeaderRow = row++;

  var currentCat = null;
  var inputStartRow = -1;
  var inputEndRow = -1;
  for (var i = 0; i < RENTAL_INPUT_FIELDS.length; i++) {
    var f = RENTAL_INPUT_FIELDS[i];
    var cat = INPUT_CATEGORY_BY_KEY[f.key];
    if (cat !== currentCat) {
      currentCat = cat;
      categoryHeaderRows[cat] = row++;
    }
    var dataRow = row++;
    rows[f.key] = dataRow;
    if (inputStartRow === -1) inputStartRow = dataRow;
    inputEndRow = dataRow;
  }

  var spacerRow = row++;
  var outputMainHeaderRow = row++;

  var currentStep = null;
  var outputStartRow = -1;
  var outputEndRow = -1;
  for (var j = 0; j < RENTAL_OUTPUT_FIELDS.length; j++) {
    var of_ = RENTAL_OUTPUT_FIELDS[j];
    var step = OUTPUT_STEP_BY_KEY[of_.key];
    if (step !== currentStep) {
      currentStep = step;
      stepHeaderRows[step] = row++;
    }
    var oRow = row++;
    rows[of_.key] = oRow;
    if (outputStartRow === -1) outputStartRow = oRow;
    outputEndRow = oRow;
  }

  return {
    rows: rows,
    categoryHeaderRows: categoryHeaderRows,
    stepHeaderRows: stepHeaderRows,
    inputMainHeaderRow: inputMainHeaderRow,
    inputStartRow: inputStartRow,
    inputEndRow: inputEndRow,
    spacerRow: spacerRow,
    outputMainHeaderRow: outputMainHeaderRow,
    outputStartRow: outputStartRow,
    outputEndRow: outputEndRow,
    lastRow: outputEndRow,
  };
})();

function insertAnalyzeButton(sheet) {
  var maxSize = 25;
  var button = sheet.insertImage(
    ANALYZE_BUTTON_URL,
    2,
    RENTAL_LAYOUT.outputMainHeaderRow,
    0,
    0,
  );
  var w = button.getWidth();
  var h = button.getHeight();
  if (w > 0 && h > 0) {
    var scale = Math.min(maxSize / w, maxSize / h);
    button.setWidth(Math.max(1, Math.round(w * scale)));
    button.setHeight(Math.max(1, Math.round(h * scale)));
  }
  button.assignScript("analyzeRental");
}

function createRentalAssessmentSheetLayout(sheet) {
  var A = 1,
    B = 2,
    C = 3;

  // ── helpers ──────────────────────────────────────────────────────────────

  function setMainHeader(row, text, bg) {
    sheet
      .getRange(row, A, 1, 3)
      .setBackground(bg)
      .setFontColor("#FFFFFF")
      .setFontWeight("bold")
      .setFontSize(11)
      .setVerticalAlignment("middle");
    sheet.getRange(row, A).setValue("  " + text);
    sheet.getRange(row, B, 1, 2).setValue("");
    sheet.setRowHeight(row, 30);
  }

  function setCategoryHeader(row, text, bg) {
    sheet
      .getRange(row, A, 1, 3)
      .setBackground(bg)
      .setFontColor("#FFFFFF")
      .setFontWeight("bold")
      .setFontSize(10)
      .setVerticalAlignment("middle");
    sheet.getRange(row, A).setValue("  " + text);
    sheet.getRange(row, B, 1, 2).setValue("");
    sheet.setRowHeight(row, 26);
  }

  function setStepHeader(row, text) {
    sheet
      .getRange(row, A, 1, 3)
      .setBackground("#37474F")
      .setFontColor("#FFFFFF")
      .setFontWeight("bold")
      .setFontSize(10)
      .setVerticalAlignment("middle");
    sheet.getRange(row, A).setValue("  " + text);
    sheet.getRange(row, B, 1, 2).setValue("");
    sheet.setRowHeight(row, 26);
  }

  function setInputDataRow(row, field, rowBg) {
    sheet
      .getRange(row, A)
      .setValue(field.label)
      .setBackground(rowBg)
      .setFontColor("#212121")
      .setVerticalAlignment("middle");

    var valCell = sheet.getRange(row, B);
    valCell
      .setBackground("#FFF9C4")
      .setNumberFormat(field.format)
      .setNote("Enter value here")
      .setVerticalAlignment("middle");
    if (
      field.defaultValue !== "" &&
      field.defaultValue !== null &&
      field.defaultValue !== undefined
    ) {
      valCell.setValue(field.defaultValue);
    }

    sheet
      .getRange(row, C)
      .setValue(field.note || "")
      .setBackground(rowBg)
      .setFontColor("#9E9E9E")
      .setFontStyle("italic")
      .setFontSize(9)
      .setVerticalAlignment("middle");

    sheet.setRowHeight(row, 21);
  }

  function setOutputDataRow(row, field, rowBg) {
    sheet
      .getRange(row, A)
      .setValue(field.label)
      .setBackground(rowBg)
      .setFontColor("#212121")
      .setVerticalAlignment("middle");

    sheet
      .getRange(row, B)
      .setBackground("#C8E6C9")
      .setNumberFormat(field.format)
      .setVerticalAlignment("middle");

    sheet
      .getRange(row, C)
      .setValue(field.note || "")
      .setBackground(rowBg)
      .setFontColor("#9E9E9E")
      .setFontStyle("italic")
      .setFontSize(9)
      .setVerticalAlignment("middle");

    sheet.setRowHeight(row, 21);
  }

  function hairlineBottom(row) {
    sheet
      .getRange(row, A, 1, 3)
      .setBorder(
        false,
        false,
        true,
        false,
        false,
        false,
        "#E0E0E0",
        SpreadsheetApp.BorderStyle.SOLID,
      );
  }

  // ── SPACER row ────────────────────────────────────────────────────────────
  sheet.getRange(RENTAL_LAYOUT.spacerRow, A, 1, 3).setBackground("#F5F5F5");
  sheet.setRowHeight(RENTAL_LAYOUT.spacerRow, 12);

  // ── INPUT SECTION ─────────────────────────────────────────────────────────
  setMainHeader(
    RENTAL_LAYOUT.inputMainHeaderRow,
    "RENTAL ASSESSMENT — INPUTS",
    "#0D47A1",
  );

  var currentInputCat = null;
  for (var i = 0; i < RENTAL_INPUT_FIELDS.length; i++) {
    var f = RENTAL_INPUT_FIELDS[i];
    var cat = INPUT_CATEGORY_BY_KEY[f.key];
    var catMeta = INPUT_CATEGORY_META[cat];
    if (cat !== currentInputCat) {
      currentInputCat = cat;
      setCategoryHeader(
        RENTAL_LAYOUT.categoryHeaderRows[cat],
        catMeta.label,
        catMeta.headerBg,
      );
    }
    setInputDataRow(RENTAL_LAYOUT.rows[f.key], f, catMeta.rowBg);
    hairlineBottom(RENTAL_LAYOUT.rows[f.key]);
  }

  // ── OUTPUT SECTION ────────────────────────────────────────────────────────
  setMainHeader(
    RENTAL_LAYOUT.outputMainHeaderRow,
    "ASSESSMENT RESULTS",
    "#1B5E20",
  );

  var currentStep = null;
  for (var j = 0; j < RENTAL_OUTPUT_FIELDS.length; j++) {
    var of_ = RENTAL_OUTPUT_FIELDS[j];
    var step = OUTPUT_STEP_BY_KEY[of_.key];
    var stepMeta = OUTPUT_STEP_META[step] || { rowBg: "#FFFFFF" };
    if (step !== currentStep) {
      currentStep = step;
      setStepHeader(RENTAL_LAYOUT.stepHeaderRows[step], step);
    }
    setOutputDataRow(RENTAL_LAYOUT.rows[of_.key], of_, stepMeta.rowBg);
    hairlineBottom(RENTAL_LAYOUT.rows[of_.key]);
  }

  // ── Column widths ─────────────────────────────────────────────────────────
  sheet.setColumnWidth(A, 300);
  sheet.setColumnWidth(B, 165);
  sheet.setColumnWidth(C, 540);

  // ── Freeze header ─────────────────────────────────────────────────────────
  sheet.setFrozenRows(1);

  // ── Thin vertical divider between B and C ─────────────────────────────────
  sheet
    .getRange(1, B, RENTAL_LAYOUT.lastRow, 1)
    .setBorder(
      false,
      false,
      false,
      true,
      false,
      false,
      "#BDBDBD",
      SpreadsheetApp.BorderStyle.SOLID,
    );

  insertAnalyzeButton(sheet);
}

function getRentalInputValues(sheet) {
  var inputs = {};
  for (var i = 0; i < RENTAL_INPUT_FIELDS.length; i++) {
    var field = RENTAL_INPUT_FIELDS[i];
    var row = RENTAL_LAYOUT.rows[field.key];
    var value = sheet.getRange(row, 2).getValue();
    inputs[field.key] =
      value === "" || value === null || value === undefined ? null : value;
  }
  return inputs;
}

function clearRentalOutputValues(sheet) {
  sheet
    .getRange(
      RENTAL_LAYOUT.outputStartRow,
      2,
      RENTAL_LAYOUT.outputEndRow - RENTAL_LAYOUT.outputStartRow + 1,
      1,
    )
    .clearContent();
}

function writeRentalOutputValues(sheet, outputs) {
  for (var i = 0; i < RENTAL_OUTPUT_FIELDS.length; i++) {
    var field = RENTAL_OUTPUT_FIELDS[i];
    var row = RENTAL_LAYOUT.rows[field.key];
    var value = outputs[field.key];
    sheet
      .getRange(row, 2)
      .setValue(value === null || value === undefined ? "" : value);
  }
}
