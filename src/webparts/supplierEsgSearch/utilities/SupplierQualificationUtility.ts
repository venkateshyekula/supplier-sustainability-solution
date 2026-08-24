export class SupplierQualificationUtility {
  /**
   * Parses the exact numeric value returned by SharePoint.
   *
   * This method does not normalize, multiply, divide, clamp,
   * or otherwise change the numeric scale.
   *
   * Examples:
   * 10       -> 10
   * "10"     -> 10
   * "10%"    -> 10
   * "9.99%"  -> 9.99
   * "4,99%"  -> 4.99
   * undefined -> 0
   */
  public static parseSharePointValue(
    value: string | number | undefined
  ): number {
    if (typeof value === 'number') {
      return (
        isFinite(value) &&
        value >= 0
      )
        ? value
        : 0;
    }

    if (typeof value !== 'string') {
      return 0;
    }

    const normalizedValue: string =
      value
        .replace(/%/g, '')
        .replace(',', '.')
        .trim();

    if (!normalizedValue) {
      return 0;
    }

    const parsedValue: number =
      parseFloat(normalizedValue);

    if (
      isNaN(parsedValue) ||
      !isFinite(parsedValue) ||
      parsedValue < 0
    ) {
      return 0;
    }

    return parsedValue;
  }

  /**
   * Displays the exact SharePoint value with a percentage symbol.
   *
   * Unnecessary trailing zeros are removed:
   * 10    -> 10%
   * 9.99  -> 9.99%
   * 5     -> 5%
   * 4.9   -> 4.9%
   * 1.99  -> 1.99%
   */
  public static formatPercentage(
    value: number
  ): string {
    if (
      !isFinite(value) ||
      value < 0
    ) {
      return '0%';
    }

    const roundedValue: number =
      Math.round(value * 100) / 100;

    return `${roundedValue.toString()}%`;
  }
}