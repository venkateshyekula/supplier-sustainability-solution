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
  overallPercentageInternalName: string;

  maximumWeightedScore: number;
}