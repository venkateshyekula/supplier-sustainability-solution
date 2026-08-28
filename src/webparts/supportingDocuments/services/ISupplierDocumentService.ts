import { IDocumentLibraryConfiguration } from '../models/IDocumentLibraryConfiguration';
import { ISupplierDocument } from '../models/ISupplierDocument';
export interface ISupplierDocumentService { 
  getAllDocuments(
    configurations: readonly IDocumentLibraryConfiguration[]
  ): Promise<ISupplierDocument[]>; 
  getLibraryUrls(
    configurations: readonly IDocumentLibraryConfiguration[]
  ): 
  Promise<Record<string, string>>; 
}