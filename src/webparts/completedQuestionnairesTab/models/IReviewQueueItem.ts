import { SupplierTier } from '../../supplierEsgSearch/models/IListConfiguration';

export interface IReviewQueueItem {
  key: string;
  supplierName: string;
  tier: SupplierTier;
  submittedDate: string;
  dueDate?: string;
  reviewer: string;
  qualification: string;
  sourceItemUrl: string;
}