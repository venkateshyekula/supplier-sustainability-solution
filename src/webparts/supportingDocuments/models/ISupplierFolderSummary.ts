import { SupplierTier } from '../../supplierEsgSearch/models/IListConfiguration';

export interface ISupplierFolderSummary {
  key: string;
  supplierName: string;
  tier: SupplierTier;
  libraryTitle: string;
  documentCount: number;
  totalFileSize: number;
  latestUploadDate?: string;
  latestModifiedBy?: string;
  folderServerRelativeUrl: string;
}