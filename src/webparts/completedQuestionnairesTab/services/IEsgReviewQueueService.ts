import { ISupplierSubmission } from '../../supplierEsgSearch/models/ISupplierSubmission';
import { IReviewQueueItem } from '../models/IReviewQueueItem';

export interface IEsgReviewQueueService {
  buildQueue(
    submissions: readonly ISupplierSubmission[],
    maximumItems?: number
  ): IReviewQueueItem[];
}