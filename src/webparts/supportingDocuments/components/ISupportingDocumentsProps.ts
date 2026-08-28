import {
  WebPartContext
} from '@microsoft/sp-webpart-base';

export interface ISupportingDocumentsProps {
  context: WebPartContext;

  tier1QuestionnaireListTitle: string;
  tier2QuestionnaireListTitle: string;
  tier3QuestionnaireListTitle: string;

  tier1SupplierNameField: string;
  tier2SupplierNameField: string;
  tier3SupplierNameField: string;

  tier1DocumentLibraryTitle: string;
  tier2DocumentLibraryTitle: string;
  tier3DocumentLibraryTitle: string;

  documentGuidanceUrl?: string;
}