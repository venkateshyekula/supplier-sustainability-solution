import {
  IListConfiguration
} from '../models/IListConfiguration';

import {
  ISupplierSubmission
} from '../models/ISupplierSubmission';

export interface ISupplierSubmissionService {
  getAllSubmissions(
    listConfigurations:
      readonly IListConfiguration[]
  ): Promise<ISupplierSubmission[]>;
}