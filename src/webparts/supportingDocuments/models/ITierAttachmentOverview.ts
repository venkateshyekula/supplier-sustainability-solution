import { SupplierTier } from '../../supplierEsgSearch/models/IListConfiguration';
export interface ITierAttachmentOverview {
  tier: SupplierTier; 
  submissions: number; 
  suppliersWithDocuments: number; 
  missingAttachments: number;
  latestSupplier?: string; 
  latestSubmissionDate?: string; 
  latestQualification?: string;
  latestRiskRating?: string; 
  latestRecommendation?: string; 
  latestSubmissionUrl?: string;
}