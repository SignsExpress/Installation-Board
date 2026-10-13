function calculateEstimateFromMetrics(metrics, settings) {
    const standardPrintArea = Number(metrics.standardPrintArea) || 0;
    const contraArea = Number(metrics.contraArea) || 0;
    const reflectiveArea = Number(metrics.reflectiveArea) || 0;
    const wrapArea = Number(metrics.wrapArea) || 0;
    const standardArea = standardPrintArea + contraArea + reflectiveArea;
    const totalArea = standardArea + wrapArea;
    const vehicleArea = Number(metrics.vehicleArea) || 0;
    const wrapSectionCount = Number(metrics.wrapSectionCount) || 0;
    const wrapSectionFactor = wrapArea > 0 ? getWrapSectionFactor(wrapSectionCount, settings) : 1;
    const weightedDifficulty = Number(metrics.weightedDifficulty) || 1;

    if (totalArea <= 0) return 0;

    function calculateScaledEstimate(scale = 1) {
      const scaledStandardArea = standardArea * scale;
      const scaledStandardMaterialMultiplierArea =
        (standardPrintArea * (settings.materialMultipliers.standard || 1) +
          contraArea * (settings.materialMultipliers.contra || 1) +
          reflectiveArea * (settings.materialMultipliers.reflective || 1)) *
        scale;
      const scaledWrapArea = wrapArea * scale;
      const scaledTotalArea = totalArea * scale;
      const scaledCoverage = vehicleArea > 0 ? scaledTotalArea / vehicleArea : 0;
      const scaledWrapCoverage = vehicleArea > 0 ? scaledWrapArea / vehicleArea : 0;
      const standardSell = scaledStandardMaterialMultiplierArea * settings.standardVinylRate;
      const wrapSell = scaledWrapArea * getFormulaWrapRate(scaledWrapCoverage, settings);
      const standardLabourHours =
        scaledStandardArea <= 0
          ? 0
          : scaledCoverage < 0.15
            ? clampNumber(
                scaledStandardArea * settings.standardSmallHoursPerM2,
                settings.standardSmallMinHours,
                settings.standardSmallMaxHours
              )
            : scaledStandardArea * settings.standardLargeHoursPerM2;
      const wrapLabourHours =
        scaledWrapArea * getFormulaWrapLabourHours(scaledWrapCoverage, settings) * wrapSectionFactor * weightedDifficulty;
      const basePrice = standardSell + wrapSell + (standardLabourHours + wrapLabourHours) * settings.labourSellRate;
      const blend = getBlendWeights(scaledCoverage, scaledWrapArea, settings);
      let estimate = blend.calculated * basePrice + blend.anchor * getMarketAnchor(scaledCoverage, settings);

      estimate = Math.max(estimate, settings.minPrice);
      if (scaledWrapArea > 0) estimate = Math.max(estimate, settings.minAnyWrapPrice);
      if (scaledWrapArea > 0 && scaledCoverage >= 0.15) estimate = Math.max(estimate, settings.minPartialWrapPrice);
      if (scaledWrapArea > 0 && scaledCoverage >= 0.85) estimate = Math.max(estimate, settings.minFullWrapPrice);
      return estimate;
    }

    const currentPrice = calculateScaledEstimate(1);
    const monotonicFloor = Array.from({ length: 80 }, (_, index) => (index + 1) / 80).reduce(
      (highest, scale) => Math.max(highest, calculateScaledEstimate(scale)),
      currentPrice
    );
    return Math.round(Math.max(currentPrice, monotonicFloor) / 50) * 50;
  }
function getWrapSectionFactor(sectionCount, settings) {
    if (sectionCount <= 1) return settings.sectionFactors.one;
    if (sectionCount <= 3) return settings.sectionFactors.twoToThree;
    if (sectionCount <= 5) return settings.sectionFactors.fourToFive;
    return settings.sectionFactors.moreThanFive;
  }
function getFormulaWrapRate(wrapCoverage, settings) {
    return Math.max(settings.wrapRateFloor, settings.wrapRateStart - settings.wrapRateTaper * Math.sqrt(Math.max(0, wrapCoverage)));
  }
function clampNumber(value, min, max) {
    return Math.min(max, Math.max(min, value));
  }
function getFormulaWrapLabourHours(wrapCoverage, settings) {
    return Math.max(
      settings.wrapLabourFloorHoursPerM2,
      settings.wrapLabourStartHoursPerM2 - settings.wrapLabourTaper * Math.sqrt(Math.max(0, wrapCoverage))
    );
  }
function getBlendWeights(coverage, wrapArea, settings) {
    if (wrapArea <= 0) return settings.blendWeights.noWrap;
    if (coverage < 0.35) return settings.blendWeights.wrapUnder35;
    if (coverage < 0.7) return settings.blendWeights.wrapUnder70;
    return settings.blendWeights.wrapFull;
  }
function getMarketAnchor(coverage, settings) {
    return interpolatePricingPoint(
      [
        { coverage: 0, value: settings.marketAnchors.c0 },
        { coverage: 0.05, value: settings.marketAnchors.c05 },
        { coverage: 0.1, value: settings.marketAnchors.c10 },
        { coverage: 0.15, value: settings.marketAnchors.c15 },
        { coverage: 0.22, value: settings.marketAnchors.c22 },
        { coverage: 0.35, value: settings.marketAnchors.c35 },
        { coverage: 0.55, value: settings.marketAnchors.c55 },
        { coverage: 0.85, value: settings.marketAnchors.c85 },
        { coverage: 1, value: settings.marketAnchors.c100 }
      ],
      coverage
    );
  }
function interpolatePricingPoint(points, value) {
    if (value <= points[0].coverage) return points[0].value;
    for (let index = 1; index < points.length; index += 1) {
      const previous = points[index - 1];
      const next = points[index];
      if (value > next.coverage) continue;
      const progress = (value - previous.coverage) / (next.coverage - previous.coverage || 1);
      return previous.value + (next.value - previous.value) * progress;
    }
    return points[points.length - 1].value;
  }
export {calculateEstimateFromMetrics};
