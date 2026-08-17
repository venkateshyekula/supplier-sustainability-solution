import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISupplierEsgSearchProps {
  context: WebPartContext;

  tier1ListTitle: string;
  tier2ListTitle: string;
  tier3ListTitle: string;

  tier1SupplierNameField: string;
  tier2SupplierNameField: string;
  tier3SupplierNameField: string;

  qualifiedMinimum: number;
  conditionalMinimum: number;
}