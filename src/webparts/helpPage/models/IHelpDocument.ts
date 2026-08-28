export interface IHelpDocument {
  id: number;
  resourceKey: string;
  title: string;
  fileName: string;
  serverRelativeUrl: string;
  modified?: string;
}