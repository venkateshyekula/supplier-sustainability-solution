import {
  QualificationStatus,
  Recommendation,
  RiskRating
} from './IQualificationResult';

import { SupplierTier } from './IListConfiguration';

export interface ISupplierSubmission {
  key: string;
  id: number;
  sourceListTitle: string;
  sourceItemUrl: string;
  tier: SupplierTier;
  supplierName: string;
  email: string;
  submittedByName: string;

  /**
   * Sustainability percentage displayed in the UI.
   */
  overallPercentage: number;

  /**
   * Weighted score used for qualification.
   */
  weightedScore: number;

  qualification: QualificationStatus;
  riskRating: RiskRating;
  recommendation: Recommendation;
  created: string;
  modified: string;
}