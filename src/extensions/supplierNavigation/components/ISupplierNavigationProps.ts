export interface ISupplierNavigationItem {
  key: string;
  text: string;
  iconName: string;
  url: string;
  ariaLabel?: string;
  label?: string;
}

export type SupplierNavigationRenderMode =
  | 'navigation'
  | 'footer';

export interface ISupplierNavigationProps {
  renderMode: SupplierNavigationRenderMode;
  
  organizationName: string;

  homeUrl: string;
  completedQuestionnairesUrl: string;
  supportingDocumentsUrl: string;
  dashboardUrl: string;
  helpUrl: string;

  currentPath: string;

  footerText?: string;
  privacyUrl?: string;
  accessibilityUrl?: string;

  navigationItems?:
  readonly ISupplierNavigationItem[];
}

/*export interface ISupplierNavigationProps {
  siteTitle: string;
  organizationName: string;

  homeUrl: string;
  completedQuestionnairesUrl: string;
  supportingDocumentsUrl: string;
  dashboardUrl: string;
  helpUrl: string;

  currentPath: string;
  navigationItems:
  readonly ISupplierNavigationItem[];
}*/