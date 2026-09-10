export interface ISharePointSubmissionItem {
  Id: number;
  Created: string;
  Modified: string;

  [fieldName: string]:
    | string
    | number
    | boolean
    | undefined;
}

export interface ISharePointItemsResponse {
  value: ISharePointSubmissionItem[];
  '@odata.nextLink'?: string;
  'odata.nextLink'?: string;
}