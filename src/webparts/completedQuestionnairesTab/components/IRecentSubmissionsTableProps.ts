import { ISupplierSubmission } from '../../supplierEsgSearch/models/ISupplierSubmission';

export interface ITierListUrls {
  tier1?: string;
  tier2?: string;
  tier3?: string;
}

export interface IRecentSubmissionsTableProps {
  submissions: readonly ISupplierSubmission[];
  isLoading: boolean;
  recentItemsLimit: number;
  tierListUrls: ITierListUrls;
  onOpenSubmission(item: ISupplierSubmission): void;
  onOpenTierList(url?: string): void;
}