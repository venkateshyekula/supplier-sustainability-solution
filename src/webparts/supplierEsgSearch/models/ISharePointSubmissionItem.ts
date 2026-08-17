export interface ISharePointSubmissionItem {
  Id: number;
  Created: string;
  Modified: string;

  /**
   * Additional SharePoint fields are accessed dynamically
   * using their internal names.
   */
  [key: string]: string | number | undefined;
}

export interface ISharePointItemsResponse {
  value?: ISharePointSubmissionItem[];

  /**
   * OData v4 next-page link.
   */
  '@odata.nextLink'?: string;

  /**
   * Older SharePoint REST next-page link.
   */
  'odata.nextLink'?: string;
}