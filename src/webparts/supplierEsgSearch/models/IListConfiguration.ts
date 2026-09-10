export type SupplierTier =
  | 'Tier 1'
  | 'Tier 2'
  | 'Tier 3';

export interface IListConfiguration {
  listTitle: string;
  tier: SupplierTier;

  supplierNameDisplayName: string;
  supplierNameInternalName: string;
  emailInternalName: string;
  contactNameInternalName: string;

  /**
   * Internal name of the SharePoint column containing
   * the Sustainability percentage displayed in the UI.
   *
   * Pre-Production:
   * OverallQuestionsPercentage
   *
   * Production:
   * Tier 1: Tier_x0020_1_x0020_Sustainabilit
   * Tier 2: Tier_x0020_2_x0020_Sustainabilit
   * Tier 3: Tier_x0020_3_x0020_Sustainabilit
   */
  overallPercentageInternalName: string;

  /**
   * Internal name of the SharePoint column containing
   * the weighted score used for qualification.
   *
   * Pre-Production:
   * OverallQuestionsPercentage
   *
   * Production:
   * Tier 1: Tier_x0020_1_x0020_Weighting
   * Tier 2: Tier_x0020_2_x0020_Weighting
   * Tier 3: Tier_x0020_3_x0020_Weighting
   */
  weightingInternalName: string;

  qualifiedWeightedMinimum: number;
  conditionalWeightedMinimum: number;
}