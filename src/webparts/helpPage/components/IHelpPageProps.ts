import {
  WebPartContext
} from '@microsoft/sp-webpart-base';

export interface IHelpPageProps {
  context: WebPartContext;

  supportEmail: string;
  procurementEmail: string;
  supportPhone: string;
  supportHours: string;

  serviceNowUrl: string;

  esgGuidelinesUrl: string;
  userGuideUrl: string;
  supplierPortalUrl: string;
  policyComplianceUrl: string;

  supplierGuideUrl: string;
  procurementGuideUrl: string;

  faqPageUrl: string;

  documentLinksLoading?: boolean;
  documentLinksError?: string;
}