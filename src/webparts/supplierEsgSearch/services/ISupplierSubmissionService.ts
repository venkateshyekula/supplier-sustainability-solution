import {
  IListConfiguration
} from '../models/IListConfiguration';

import {
  ISupplierSubmission
} from '../models/ISupplierSubmission';

import {
  IQualificationThresholds
} from '../utilities/SupplierQualificationUtility';

export interface ISupplierSubmissionService {
  getAllSubmissions(
    listConfigurations: readonly IListConfiguration[],
    thresholds: IQualificationThresholds
  ): Promise<ISupplierSubmission[]>;
}