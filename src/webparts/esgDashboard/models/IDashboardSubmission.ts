import { QualificationStatus, Recommendation, RiskRating } from '../../supplierEsgSearch/models/IQualificationResult';
import { SupplierTier } from '../../supplierEsgSearch/models/IListConfiguration';
export interface IDashboardSubmission {
    key: string; id: number; supplierName: string; tier: SupplierTier; score: number;
    qualification: QualificationStatus; riskRating: RiskRating; recommendation: Recommendation;
    submittedOn: string; submittedBy: string; sourceItemUrl: string;
}