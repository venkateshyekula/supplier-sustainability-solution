import { ISupplierSubmission } from '../../supplierEsgSearch/models/ISupplierSubmission';
import { IReviewQueueItem } from '../models/IReviewQueueItem';
import { IEsgReviewQueueService } from './IEsgReviewQueueService';

export class EsgReviewQueueService implements IEsgReviewQueueService {
  public buildQueue(
    submissions: readonly ISupplierSubmission[],
    maximumItems: number = 5
  ): IReviewQueueItem[] {
    return submissions
      .filter((item: ISupplierSubmission): boolean => {
        return item.qualification !== 'Qualified';
      })
      .slice()
      .sort((a: ISupplierSubmission, b: ISupplierSubmission): number => {
        return new Date(b.created).getTime() - new Date(a.created).getTime();
      })
      .slice(0, maximumItems)
      .map((item: ISupplierSubmission): IReviewQueueItem => {
        return {
          key: item.key,
          supplierName: item.supplierName,
          tier: item.tier,
          submittedDate: item.created,
          dueDate: undefined,
          reviewer: 'Unassigned',
          qualification: item.qualification,
          sourceItemUrl: item.sourceItemUrl
        };
      });
  }
}