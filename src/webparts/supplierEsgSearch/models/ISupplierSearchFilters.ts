import { QualificationStatus } from './IQualificationResult';
import { SupplierTier } from './IListConfiguration';

export interface ISupplierSearchFilters {
  searchText: string;
  tier?: SupplierTier;
  qualification?: QualificationStatus;
  minimumPercentage?: number;
  maximumPercentage?: number;
  startDate?: Date;
  endDate?: Date;
}