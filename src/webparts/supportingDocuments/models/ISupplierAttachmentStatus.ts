import {
  SupplierTier
} from '../../supplierEsgSearch/models/IListConfiguration';

export type SupplierAttachmentStatus =
  | 'Submitted'
  | 'Pending'
  | 'Missing';

export interface ISupplierAttachmentStatus {
  key: string;

  supplierName: string;
  tier: SupplierTier;

  status: SupplierAttachmentStatus;

  submissionId: number;
  submissionDate: string;
  submissionItemUrl: string;

  documentCount: number;

  folderExists: boolean;
  folderServerRelativeUrl?: string;

  latestDocumentDate?: string;
  latestDocumentName?: string;
}