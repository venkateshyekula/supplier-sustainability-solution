import {
  WebPartContext
} from '@microsoft/sp-webpart-base';

export interface ICompletedQuestionnairesTabProps {
  context: WebPartContext;

  tier1ListTitle: string;
  tier2ListTitle: string;
  tier3ListTitle: string;

  tier1SupplierNameField: string;
  tier2SupplierNameField: string;
  tier3SupplierNameField: string;

  tier1DocumentLibraryTitle: string;
  tier2DocumentLibraryTitle: string;
  tier3DocumentLibraryTitle: string;
}