export type QualificationStatus =
  | 'Qualified'
  | 'Conditionally Qualified'
  | 'Not Qualified';

export type RiskRating =
  | 'Low Risk'
  | 'Medium Risk'
  | 'High Risk';

export type Recommendation =
  | 'Approve'
  | 'Corrective Action'
  | 'Requires Review';

export interface IQualificationResult {
  qualification: QualificationStatus;
  riskRating: RiskRating;
  recommendation: Recommendation;
}