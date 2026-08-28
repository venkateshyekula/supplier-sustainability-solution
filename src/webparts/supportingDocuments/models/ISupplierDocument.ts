import { SupplierTier } from '../../supplierEsgSearch/models/IListConfiguration';
import {  SupplierAttachmentStatus } from './ISupplierAttachmentStatus';

export interface ISupplierDocument {
  key: string;
  itemId: number;
  supplierName: string;
  tier: SupplierTier;
  libraryTitle: string;
  fileName: string;
  fileExtension: string;
  fileSize: number;
  created: string;
  createdBy: string;
  modified: string;
  modifiedBy: string;
  folderServerRelativeUrl: string;
  fileServerRelativeUrl: string;
  status: SupplierAttachmentStatus;
}