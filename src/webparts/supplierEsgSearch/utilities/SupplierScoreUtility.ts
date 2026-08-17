export interface ISupplierScoreResult {
  obtainedPoints: number;
  maxPoints: number;
  overallPercentage: number;
  weightagePercentage: number;
  weightedScore: number;
}

const roundToTwoDecimals = (value: number): number => {
  return Math.round(value * 100) / 100;
};

const normalizeNumber = (
  value: string | number | undefined
): number => {
  if (typeof value === 'number') {
    return isFinite(value) ? value : 0;
  }

  if (typeof value !== 'string') {
    return 0;
  }

  const normalizedValue: string = value
    .replace('%', '')
    .replace(',', '.')
    .trim();

  const parsedValue: number = parseFloat(normalizedValue);

  return isFinite(parsedValue) ? parsedValue : 0;
};

export const normalizeStoredPercentage = (
  value: string | number | undefined
): number => {
  const parsedValue: number = normalizeNumber(value);

  const percentage: number =
    parsedValue > 0 && parsedValue <= 1
      ? parsedValue * 100
      : parsedValue;

  if (percentage < 0) {
    return 0;
  }

  if (percentage > 100) {
    return 100;
  }

  return roundToTwoDecimals(percentage);
};

export const calculateSupplierScore = (
  obtainedPoints: string | number | undefined,
  maxPoints: number,
  weightagePercentage: number
): ISupplierScoreResult => {
  const normalizedObtainedPoints: number =
    normalizeNumber(obtainedPoints);

  if (!isFinite(maxPoints) || maxPoints <= 0) {
    throw new Error(
      'Maximum points must be greater than zero.'
    );
  }

  if (
    !isFinite(weightagePercentage) ||
    weightagePercentage < 0
  ) {
    throw new Error(
      'Weightage percentage cannot be negative.'
    );
  }

  const safeObtainedPoints: number = Math.min(
    Math.max(normalizedObtainedPoints, 0),
    maxPoints
  );

  const overallPercentage: number =
    roundToTwoDecimals(
      (safeObtainedPoints / maxPoints) * 100
    );

  const weightedScore: number =
    roundToTwoDecimals(
      (overallPercentage / 100) *
      weightagePercentage
    );

  return {
    obtainedPoints: safeObtainedPoints,
    maxPoints,
    overallPercentage,
    weightagePercentage,
    weightedScore
  };
};

export const calculateWeightedScoreFromPercentage = (
  overallPercentageValue: string | number | undefined,
  weightagePercentage: number
): number => {
  const overallPercentage: number =
    normalizeStoredPercentage(
      overallPercentageValue
    );

  if (
    !isFinite(weightagePercentage) ||
    weightagePercentage < 0
  ) {
    return 0;
  }

  return roundToTwoDecimals(
    (overallPercentage / 100) *
    weightagePercentage
  );
};