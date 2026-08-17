import { IQualificationResult } from '../models/IQualificationResult';

export interface IQualificationThresholds {
  qualifiedMinimum: number;
  conditionalMinimum: number;
}

export class SupplierQualificationUtility {
  public static parsePercentage(
    value: string | number | undefined
  ): number {
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

    if (!normalizedValue) {
      return 0;
    }

    const parsedValue: number = parseFloat(normalizedValue);

    return isFinite(parsedValue) ? parsedValue : 0;
  }

  public static normalizePercentage(
    value: string | number | undefined
  ): number {
    let parsedValue: number =
      SupplierQualificationUtility.parsePercentage(value);

    /*
     * If SharePoint returns 0.87, convert it to 87.
     * Zero remains zero.
     */
    if (parsedValue > 0 && parsedValue <= 1) {
      parsedValue = parsedValue * 100;
    }

    if (parsedValue < 0) {
      return 0;
    }

    if (parsedValue > 100) {
      return 100;
    }

    return Math.round(parsedValue * 100) / 100;
  }

  public static evaluate(
    overallPercentage: number,
    thresholds: IQualificationThresholds
  ): IQualificationResult {
    if (overallPercentage >= thresholds.qualifiedMinimum) {
      return {
        qualification: 'Qualified',
        riskRating: 'Low Risk',
        recommendation: 'Approve'
      };
    }

    if (overallPercentage >= thresholds.conditionalMinimum) {
      return {
        qualification: 'Conditionally Qualified',
        riskRating: 'Medium Risk',
        recommendation: 'Corrective Action'
      };
    }

    return {
      qualification: 'Not Qualified',
      riskRating: 'High Risk',
      recommendation: 'Requires Review'
    };
  }

  public static formatPercentage(value: number): string {
    return `${value.toFixed(2).replace(/\.00$/, '')}%`;
  }
}
/*export type QualificationLevel =
  | 'Qualified'
  | 'Conditionally Qualified'
  | 'Not Qualified';

export interface IQualificationThresholds {
  qualifiedMinimum: number;
  conditionalMinimum: number;
}

export interface IQualificationResult {
  qualification: QualificationLevel;
  overallPercentage: number;
}

const clampPercentage = (value: number): number => {
  if (!isFinite(value)) {
    return 0;
  }

  if (value < 0) {
    return 0;
  }

  if (value > 100) {
    return 100;
  }

  return value;
};

export const evaluateSupplierQualification = (
  overallPercentage: number,
  thresholds: IQualificationThresholds
): IQualificationResult => {
  const normalizedPercentage: number =
    clampPercentage(overallPercentage);

  if (
    thresholds.qualifiedMinimum <
    thresholds.conditionalMinimum
  ) {
    throw new Error(
      'The qualified minimum cannot be lower than ' +
      'the conditional minimum.'
    );
  }

  if (
    thresholds.qualifiedMinimum > 100 ||
    thresholds.conditionalMinimum < 0
  ) {
    throw new Error(
      'Qualification thresholds must be between 0 and 100.'
    );
  }

  if (
    normalizedPercentage >=
    thresholds.qualifiedMinimum
  ) {
    return {
      qualification: 'Qualified',
      overallPercentage: normalizedPercentage
    };
  }

  if (
    normalizedPercentage >=
    thresholds.conditionalMinimum
  ) {
    return {
      qualification: 'Conditionally Qualified',
      overallPercentage: normalizedPercentage
    };
  }

  return {
    qualification: 'Not Qualified',
    overallPercentage: normalizedPercentage
  };
};*/