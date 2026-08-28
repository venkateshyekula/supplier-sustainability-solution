import { SupplierTier } from '../../supplierEsgSearch/models/IListConfiguration';
export interface IMissingAttachment {
    key: string; 
    supplierName: string; 
    tier: SupplierTier; 
    submissionId: number;
    submissionDate: string; 
    qualification: string; 
    riskRating: string; 
    recommendation: string; 
    submissionItemUrl: string;
}