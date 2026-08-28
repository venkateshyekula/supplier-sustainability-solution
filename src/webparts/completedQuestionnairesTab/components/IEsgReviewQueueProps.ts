import { IReviewQueueItem } from '../models/IReviewQueueItem';

export interface IEsgReviewQueueProps {
  items: readonly IReviewQueueItem[];
  isLoading: boolean;
  errorMessage?: string;
  onOpenItem(item: IReviewQueueItem): void;
}