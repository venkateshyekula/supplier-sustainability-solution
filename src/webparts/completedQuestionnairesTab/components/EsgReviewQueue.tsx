import * as React from 'react';

import {
  DetailsListLayoutMode,
  IColumn,
  Icon,
  Link,
  SelectionMode,
  ShimmeredDetailsList
} from '@fluentui/react';

import {
  IReviewQueueItem
} from '../models/IReviewQueueItem';

import {
  IEsgReviewQueueProps
} from './IEsgReviewQueueProps';

import styles from './EsgReviewQueue.module.scss';

export function EsgReviewQueue(
  props: IEsgReviewQueueProps
): React.ReactElement<IEsgReviewQueueProps> {
  const formatDate = (
    value?: string
  ): string => {
    if (!value) {
      return 'Not set';
    }

    const date: Date =
      new Date(value);

    if (isNaN(date.getTime())) {
      return 'Not set';
    }

    return date.toLocaleDateString(
      'en-GB',
      {
        day: '2-digit',
        month: 'short',
        year: 'numeric'
      }
    );
  };

  const getTierClassName = (
    tier: string
  ): string => {
    switch (tier) {
      case 'Tier 1':
        return styles.tier1;

      case 'Tier 2':
        return styles.tier2;

      case 'Tier 3':
        return styles.tier3;

      default:
        return '';
    }
  };

  const getDueDateClassName = (
    value?: string
  ): string => {
    if (!value) {
      return styles.due;
    }

    const dueDate: Date =
      new Date(value);

    if (isNaN(dueDate.getTime())) {
      return styles.due;
    }

    const today: Date =
      new Date();

    today.setHours(
      0,
      0,
      0,
      0
    );

    dueDate.setHours(
      0,
      0,
      0,
      0
    );

    const differenceInMilliseconds: number =
      dueDate.getTime() -
      today.getTime();

    const differenceInDays: number =
      Math.ceil(
        differenceInMilliseconds /
        (
          1000 *
          60 *
          60 *
          24
        )
      );

    if (differenceInDays < 0) {
      return (
        `${styles.due} ` +
        `${styles.overdue}`
      );
    }

    if (differenceInDays <= 7) {
      return (
        `${styles.due} ` +
        `${styles.dueSoon}`
      );
    }

    return styles.due;
  };

  const columns: IColumn[] = [
    {
      key: 'supplierName',
      name: 'Supplier',
      fieldName: 'supplierName',
      minWidth: 150,
      maxWidth: 220,
      isResizable: true,

      onRender: (
        item: IReviewQueueItem
      ): React.ReactNode => {
        return (
          <div className={styles.supplierCell}>
            <Link
              className={styles.supplierName}
              title={item.supplierName}
              onClick={(): void => {
                props.onOpenItem(item);
              }}
            >
              {item.supplierName}
            </Link>
          </div>
        );
      }
    },
    {
      key: 'tier',
      name: 'Tier',
      fieldName: 'tier',
      minWidth: 56,
      maxWidth: 72,
      isResizable: true,

      onRender: (
        item: IReviewQueueItem
      ): React.ReactNode => {
        return (
          <span
            className={
              `${styles.tier} ` +
              `${getTierClassName(item.tier)}`
            }
          >
            {item.tier}
          </span>
        );
      }
    },
    {
      key: 'submittedDate',
      name: 'Submitted',
      fieldName: 'submittedDate',
      minWidth: 90,
      maxWidth: 112,
      isResizable: true,

      onRender: (
        item: IReviewQueueItem
      ): React.ReactNode => {
        return (
          <span className={styles.submittedDate}>
            {formatDate(item.submittedDate)}
          </span>
        );
      }
    },
    {
      key: 'dueDate',
      name: 'Due Date',
      fieldName: 'dueDate',
      minWidth: 90,
      maxWidth: 112,
      isResizable: true,

      onRender: (
        item: IReviewQueueItem
      ): React.ReactNode => {
        return (
          <span
            className={
              getDueDateClassName(
                item.dueDate
              )
            }
          >
            {formatDate(item.dueDate)}
          </span>
        );
      }
    },
    {
      key: 'reviewer',
      name: 'Reviewer',
      fieldName: 'reviewer',
      minWidth: 95,
      maxWidth: 130,
      isResizable: true,

      onRender: (
        item: IReviewQueueItem
      ): React.ReactNode => {
        const reviewer: string =
          item.reviewer ||
          'Unassigned';

        const reviewerClassName: string =
          reviewer === 'Unassigned'
            ? (
              `${styles.reviewerBadge} ` +
              `${styles.unassignedReviewer}`
            )
            : (
              `${styles.reviewerBadge} ` +
              `${styles.assignedReviewer}`
            );

        return (
          <span
            className={reviewerClassName}
            title={reviewer}
          >
            {reviewer}
          </span>
        );
      }
    }
  ];

  return (
    <div className={styles.esgReviewQueue}>
      <section
        className={styles.panel}
        aria-labelledby="esg-review-queue-title"
      >
        <header className={styles.header}>
          <div className={styles.headerContent}>
            <span
              className={styles.headerIcon}
              aria-hidden="true"
            >
              <Icon iconName="ClipboardList" />
            </span>

            <h2
              id="esg-review-queue-title"
              className={styles.headerTitle}
            >
              ESG Review Queue
            </h2>
          </div>

          <button
            type="button"
            className={styles.viewAllButton}
            disabled={
              props.isLoading ||
              props.items.length === 0
            }
          >
            View all

            <Icon
              iconName="OpenInNewWindow"
              aria-hidden="true"
            />
          </button>
        </header>

        {props.errorMessage ? (
          <div
            className={
              `${styles.message} ` +
              `${styles.error}`
            }
            role="alert"
          >
            <div className={styles.messageContent}>
              <Icon
                className={styles.errorIcon}
                iconName="ErrorBadge"
                aria-hidden="true"
              />

              <h3 className={styles.messageTitle}>
                Unable to load ESG review queue
              </h3>

              <p className={styles.messageDescription}>
                {props.errorMessage}
              </p>
            </div>
          </div>
        ) : (
          <>
            <div className={styles.table}>
              <ShimmeredDetailsList
                items={
                  props.items as IReviewQueueItem[]
                }
                columns={columns}
                setKey="esg-review-queue"
                enableShimmer={props.isLoading}
                selectionMode={SelectionMode.none}
                layoutMode={
                  DetailsListLayoutMode.justified
                }
                compact={true}
                onItemInvoked={(
                  item?: IReviewQueueItem
                ): void => {
                  if (item) {
                    props.onOpenItem(item);
                  }
                }}
                ariaLabelForGrid={
                  'Supplier ESG review queue'
                }
                ariaLabelForShimmer={
                  'Loading supplier ESG review queue'
                }
              />
            </div>

            {!props.isLoading &&
              props.items.length === 0 && (
                <div className={styles.message}>
                  <div className={styles.messageContent}>
                    <Icon
                      className={styles.messageIcon}
                      iconName="ClipboardList"
                      aria-hidden="true"
                    />

                    <h3 className={styles.messageTitle}>
                      No submissions require ESG review
                    </h3>

                    <p className={styles.messageDescription}>
                      New submissions requiring ESG review
                      will appear in this queue.
                    </p>
                  </div>
                </div>
              )}
          </>
        )}

        <footer className={styles.footer}>
          <button
            type="button"
            className={styles.footerLink}
            disabled={
              props.isLoading ||
              props.items.length === 0
            }
          >
            Go to ESG Review

            <span
              className={styles.footerArrow}
              aria-hidden="true"
            >
              &rarr;
            </span>
          </button>

          <span className={styles.resultCount}>
            {props.items.length}
            {' '}
            item
            {props.items.length === 1
              ? ''
              : 's'}
          </span>
        </footer>
      </section>
    </div>
  );
}

export default EsgReviewQueue;