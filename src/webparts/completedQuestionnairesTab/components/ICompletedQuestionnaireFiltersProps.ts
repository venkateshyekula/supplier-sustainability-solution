import { ISupplierSearchFilters } from '../../supplierEsgSearch/models/ISupplierSearchFilters';

export interface ICompletedQuestionnaireFiltersProps {
  filters: ISupplierSearchFilters;
  isLoading: boolean;
  onFiltersChange(filters: ISupplierSearchFilters): void;
  onSearch(): void;
  onClear(): void;
}