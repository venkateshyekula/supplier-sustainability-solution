import { SupplierTier } from '../../supplierEsgSearch/models/IListConfiguration';

export interface ISupplierDocumentFilters {
  searchText: string;
  tier?: SupplierTier;
  uploadedBy?: string;
  startDate?: Date;
  endDate?: Date;
}