function calculateRentalAssessment(inputs) {
  function toNumber(value) {
    if (value === null || value === undefined || value === "") return null;
    var number = Number(value);
    return isNaN(number) ? null : number;
  }

  function toDate(value) {
    if (!value) return null;
    var date = value instanceof Date ? value : new Date(value);
    return isNaN(date.getTime()) ? null : date;
  }

  function has() {
    for (var i = 0; i < arguments.length; i++) {
      if (
        arguments[i] === null ||
        arguments[i] === undefined ||
        arguments[i] === ""
      )
        return false;
    }
    return true;
  }

  function monthsBetween(startDate, endDate) {
    if (!has(startDate, endDate)) return null;
    return (
      (endDate.getFullYear() - startDate.getFullYear()) * 12 +
      (endDate.getMonth() - startDate.getMonth())
    );
  }

  var purchasePrice = toNumber(inputs.purchasePrice);
  var purchaseDate = toDate(inputs.purchaseDate);
  var loanToValueRatio = toNumber(inputs.loanToValueRatio);
  var downPaymentAssistance = toNumber(inputs.downPaymentAssistance) || 0;
  var closingCosts = toNumber(inputs.closingCosts) || 0;
  var dpaForgivenessMonths = toNumber(inputs.dpaForgivenessMonths);
  var annualInterestRate = toNumber(inputs.annualInterestRate);
  var loanTermYears = toNumber(inputs.loanTermYears);
  var monthlyEscrow = toNumber(inputs.monthlyEscrow) || 0;
  var hoaInitialMonthly = toNumber(inputs.hoaInitialMonthly);
  var hoaAnnualGrowthRate = toNumber(inputs.hoaAnnualGrowthRate);
  var initialMonthlyRent = toNumber(inputs.initialMonthlyRent);
  var rentAnnualGrowthRate = toNumber(inputs.rentAnnualGrowthRate);
  var vacancyMonthsPerTurnover = toNumber(inputs.vacancyMonthsPerTurnover) || 0;
  var costPerTurnover = toNumber(inputs.costPerTurnover) || 0;
  var annualMaintenanceCost = toNumber(inputs.annualMaintenanceCost) || 0;
  var tenantTurnoverIntervalMonths = toNumber(
    inputs.tenantTurnoverIntervalMonths,
  );
  var annualAppreciationRate = toNumber(inputs.annualAppreciationRate);
  var saleDate = toDate(inputs.saleDate);
  var longTermCapitalGainsRate = toNumber(inputs.longTermCapitalGainsRate) || 0;
  var depreciationRecaptureRate =
    toNumber(inputs.depreciationRecaptureRate) || 0;
  var totalDepreciationTaken = toNumber(inputs.totalDepreciationTaken) || 0;
  var agentFeeRate = toNumber(inputs.agentFeeRate) || 0;
  var inflationDiscountRate = toNumber(inputs.inflationDiscountRate) || 0;

  var ownershipMonths = monthsBetween(purchaseDate, saleDate);
  var ownershipYears = ownershipMonths !== null ? ownershipMonths / 12 : null;
  var ceilOwnershipYears =
    ownershipYears !== null ? Math.ceil(ownershipYears) : null;

  var loanPrincipal = has(purchasePrice, loanToValueRatio)
    ? purchasePrice * loanToValueRatio
    : null;
  var cashAtClose = has(purchasePrice, loanPrincipal)
    ? purchasePrice -
      loanPrincipal -
      (downPaymentAssistance || 0) +
      (closingCosts || 0)
    : null;
  var monthlyInterestRate = has(annualInterestRate)
    ? annualInterestRate / 12
    : null;
  var loanTermMonths = has(loanTermYears) ? loanTermYears * 12 : null;

  var monthlyPrincipalAndInterest = has(
    loanPrincipal,
    annualInterestRate,
    loanTermYears,
  )
    ? FinanceToolbox.LOAN_MONTHLY_PAYMENT(
        loanPrincipal,
        annualInterestRate,
        loanTermYears,
      )
    : null;

  var totalMonthlyPayment = has(monthlyPrincipalAndInterest)
    ? monthlyPrincipalAndInterest + monthlyEscrow
    : null;

  var remainingBalanceAtSale = has(
    loanPrincipal,
    annualInterestRate,
    loanTermYears,
    ownershipMonths,
  )
    ? FinanceToolbox.LOAN_BALANCE_AFTER_PAYMENTS(
        loanPrincipal,
        annualInterestRate,
        loanTermYears,
        Math.max(0, ownershipMonths),
      )
    : null;

  var totalPrincipalPaid = has(loanPrincipal, remainingBalanceAtSale)
    ? loanPrincipal - remainingBalanceAtSale
    : null;

  var totalInterestPaid = has(
    monthlyPrincipalAndInterest,
    ownershipMonths,
    totalPrincipalPaid,
  )
    ? monthlyPrincipalAndInterest * ownershipMonths - totalPrincipalPaid
    : null;

  var totalEscrowPaid = has(ownershipMonths)
    ? monthlyEscrow * ownershipMonths
    : null;

  var dpaMonthsForgiven = has(ownershipMonths, dpaForgivenessMonths)
    ? Math.min(ownershipMonths, dpaForgivenessMonths)
    : null;

  var dpaRemainingAtSale =
    has(downPaymentAssistance, dpaMonthsForgiven, dpaForgivenessMonths) &&
    dpaForgivenessMonths > 0
      ? Math.max(
          0,
          downPaymentAssistance *
            (1 - dpaMonthsForgiven / dpaForgivenessMonths),
        )
      : null;

  var averageMonthlyInterest =
    has(totalInterestPaid, ownershipMonths) && ownershipMonths > 0
      ? totalInterestPaid / ownershipMonths
      : null;

  var numberOfTurnovers =
    has(ownershipMonths, tenantTurnoverIntervalMonths) &&
    tenantTurnoverIntervalMonths > 0
      ? Math.floor(ownershipMonths / tenantTurnoverIntervalMonths)
      : null;

  var totalVacancyCost = has(numberOfTurnovers)
    ? costPerTurnover * numberOfTurnovers
    : null;

  var totalHoaPaid = has(
    hoaInitialMonthly,
    hoaAnnualGrowthRate,
    ceilOwnershipYears,
  )
    ? FinanceToolbox.TOTAL_GROWING_PAYMENTS(
        hoaInitialMonthly * 12,
        hoaAnnualGrowthRate,
        Math.max(0, ceilOwnershipYears),
      )
    : null;

  var vacancyMonthsPerYear =
    has(vacancyMonthsPerTurnover, tenantTurnoverIntervalMonths) &&
    tenantTurnoverIntervalMonths > 0
      ? vacancyMonthsPerTurnover * (12 / tenantTurnoverIntervalMonths)
      : 0;

  var totalGrossRentCollected = has(
    initialMonthlyRent,
    rentAnnualGrowthRate,
    ceilOwnershipYears,
  )
    ? FinanceToolbox.TOTAL_GROWING_PAYMENTS(
        initialMonthlyRent * (12 - vacancyMonthsPerYear),
        rentAnnualGrowthRate,
        Math.max(0, ceilOwnershipYears),
      )
    : null;

  var totalTurnoverCost = null;
  if (has(numberOfTurnovers, initialMonthlyRent, rentAnnualGrowthRate)) {
    totalTurnoverCost = 0;
    for (var k = 1; k <= numberOfTurnovers; k++) {
      var turnoverMonth = k * tenantTurnoverIntervalMonths;
      var turnoverYear = Math.floor(turnoverMonth / 12);
      var turnoverCost = FinanceToolbox.GROW_VALUE_OVER_TIME(
        initialMonthlyRent,
        rentAnnualGrowthRate,
        turnoverYear,
        1,
      );
      totalTurnoverCost += turnoverCost;
    }
  }

  var totalMaintenanceCost = has(annualMaintenanceCost, ownershipYears)
    ? annualMaintenanceCost * ownershipYears
    : null;

  // Operating cash: rent minus all cash outflows during ownership (principal is real cash to lender)
  var netCashReceived = has(
    totalGrossRentCollected,
    totalInterestPaid,
    totalPrincipalPaid,
    totalEscrowPaid,
    totalHoaPaid,
    totalMaintenanceCost,
    totalTurnoverCost,
    totalVacancyCost,
  )
    ? totalGrossRentCollected -
      totalInterestPaid -
      totalPrincipalPaid -
      totalEscrowPaid -
      totalHoaPaid -
      totalMaintenanceCost -
      totalTurnoverCost -
      totalVacancyCost
    : null;

  var totalCashOutlaid = has(cashAtClose, netCashReceived)
    ? cashAtClose - netCashReceived
    : null;

  var projectedSaleValue = has(
    purchasePrice,
    annualAppreciationRate,
    ownershipYears,
  )
    ? FinanceToolbox.GROW_VALUE_OVER_TIME(
        purchasePrice,
        annualAppreciationRate,
        ownershipYears,
        1,
      )
    : null;

  var grossSaleProceeds = has(
    projectedSaleValue,
    remainingBalanceAtSale,
    dpaRemainingAtSale,
    agentFeeRate,
  )
    ? projectedSaleValue -
      remainingBalanceAtSale -
      dpaRemainingAtSale -
      projectedSaleValue * agentFeeRate
    : null;

  var adjustedCostBasis = has(purchasePrice)
    ? purchasePrice - totalDepreciationTaken
    : null;

  // Capital gain uses the sale price (minus selling costs) as amount realized,
  // NOT grossSaleProceeds — loan repayment does not reduce taxable gain.
  var amountRealized = has(projectedSaleValue, agentFeeRate)
    ? projectedSaleValue - projectedSaleValue * agentFeeRate
    : null;

  var capitalGain = has(amountRealized, adjustedCostBasis)
    ? amountRealized - adjustedCostBasis
    : null;

  var depreciationRecaptureTax =
    totalDepreciationTaken * depreciationRecaptureRate;

  var capitalGainsTax = has(capitalGain)
    ? Math.max(0, capitalGain - totalDepreciationTaken) *
      longTermCapitalGainsRate
    : null;

  var totalTaxOnSale = has(depreciationRecaptureTax, capitalGainsTax)
    ? depreciationRecaptureTax + capitalGainsTax
    : null;

  var netSaleProceeds = has(grossSaleProceeds, totalTaxOnSale)
    ? grossSaleProceeds - totalTaxOnSale
    : null;

  var inflationAdjustedProceeds = has(
    netSaleProceeds,
    inflationDiscountRate,
    purchaseDate,
    saleDate,
  )
    ? FinanceToolbox.INFLATION_ADJUSTED_VALUE(
        netSaleProceeds,
        inflationDiscountRate,
        purchaseDate,
        saleDate,
      )
    : null;

  // ROI denominator is cashAtClose (actual out-of-pocket), not totalCashOutlaid
  // (which can be negative when rent surplus exceeds operating costs, inverting the sign).
  // totalReturn = net sale proceeds PLUS any operating surplus (or minus deficit).
  var totalReturn = has(netSaleProceeds, netCashReceived, cashAtClose)
    ? netSaleProceeds + netCashReceived - cashAtClose
    : null;

  var returnOnInvestment =
    has(totalReturn, cashAtClose) && cashAtClose !== 0
      ? totalReturn / cashAtClose
      : null;

  var inflationAdjustedRoi =
    has(inflationAdjustedProceeds, totalCashOutlaid, cashAtClose) &&
    cashAtClose !== 0
      ? (inflationAdjustedProceeds - totalCashOutlaid) / cashAtClose
      : null;

  var annualizedReturnCagr =
    has(totalReturn, cashAtClose, purchaseDate, saleDate) &&
    cashAtClose > 0 &&
    totalReturn > 0
      ? FinanceToolbox.YEARLY_RETURN_RATE(
          cashAtClose,
          totalReturn,
          purchaseDate,
          saleDate,
        )
      : null;

  var inflationAdjustedTotalReturn = has(
    inflationAdjustedProceeds,
    totalCashOutlaid,
  )
    ? inflationAdjustedProceeds - totalCashOutlaid
    : null;

  var inflationAdjustedCagr =
    has(inflationAdjustedTotalReturn, cashAtClose, purchaseDate, saleDate) &&
    cashAtClose > 0 &&
    inflationAdjustedTotalReturn > 0
      ? FinanceToolbox.YEARLY_RETURN_RATE(
          cashAtClose,
          inflationAdjustedTotalReturn,
          purchaseDate,
          saleDate,
        )
      : null;

  return {
    loanPrincipal: loanPrincipal,
    cashAtClose: cashAtClose,
    monthlyInterestRate: monthlyInterestRate,
    loanTermMonths: loanTermMonths,
    monthlyPrincipalAndInterest: monthlyPrincipalAndInterest,
    totalMonthlyPayment: totalMonthlyPayment,
    ownershipMonths: ownershipMonths,
    ownershipYears: ownershipYears,
    remainingBalanceAtSale: remainingBalanceAtSale,
    totalInterestPaid: totalInterestPaid,
    totalPrincipalPaid: totalPrincipalPaid,
    totalEscrowPaid: totalEscrowPaid,
    dpaMonthsForgiven: dpaMonthsForgiven,
    dpaRemainingAtSale: dpaRemainingAtSale,
    averageMonthlyInterest: averageMonthlyInterest,
    totalVacancyCost: totalVacancyCost,
    totalHoaPaid: totalHoaPaid,
    totalGrossRentCollected: totalGrossRentCollected,
    numberOfTurnovers: numberOfTurnovers,
    totalTurnoverCost: totalTurnoverCost,
    totalMaintenanceCost: totalMaintenanceCost,
    netCashReceived: netCashReceived,
    totalCashOutlaid: totalCashOutlaid,
    projectedSaleValue: projectedSaleValue,
    grossSaleProceeds: grossSaleProceeds,
    adjustedCostBasis: adjustedCostBasis,
    capitalGain: capitalGain,
    depreciationRecaptureTax: depreciationRecaptureTax,
    capitalGainsTax: capitalGainsTax,
    totalTaxOnSale: totalTaxOnSale,
    netSaleProceeds: netSaleProceeds,
    inflationAdjustedProceeds: inflationAdjustedProceeds,
    returnOnInvestment: returnOnInvestment,
    inflationAdjustedRoi: inflationAdjustedRoi,
    annualizedReturnCagr: annualizedReturnCagr,
    inflationAdjustedCagr: inflationAdjustedCagr,
  };
}
