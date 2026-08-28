import * as React from 'react';

import {
  Icon,
  Link,
  Shimmer
} from '@fluentui/react';

import {
  ITierAttachmentOverview
} from '../models/ITierAttachmentOverview';

import styles
  from './TierAttachmentOverview.module.scss';

export interface ITierAttachmentOverviewProps {
  items:
    readonly ITierAttachmentOverview[];

  isLoading:
    boolean;

  onOpenSubmission(
    url?: string
  ): void;
}

const LOADING_TIERS:
  readonly string[] = [
    'Tier 1',
    'Tier 2',
    'Tier 3'
  ];

export function TierAttachmentOverview(
  props: ITierAttachmentOverviewProps
): React.ReactElement<
  ITierAttachmentOverviewProps
> {
  const formatDate = (
    value?: string
  ): string => {
    if (!value) {
      return 'No submissions';
    }

    const date: Date =
      new Date(value);

    if (isNaN(date.getTime())) {
      return 'No submissions';
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
        return styles.defaultTier;
    }
  };

  const renderLoadingCards =
    (): React.ReactElement[] => {
      return LOADING_TIERS.map(
        (
          tier: string
        ): React.ReactElement => {
          return (
            <div
              key={tier}
              className={
                styles.loadingCard
              }
              aria-hidden="true"
            >
              <Shimmer
                isDataLoaded={false}
                width="100%"
              />
            </div>
          );
        }
      );
    };

  const renderTierCards =
    (): React.ReactElement[] => {
      return props.items.map(
        (
          item:
            ITierAttachmentOverview
        ): React.ReactElement => {
          const hasMissingAttachments:
            boolean =
              item.missingAttachments > 0;

          return (
            <article
              key={item.tier}
              className={styles.tierCard}
              aria-label={
                `${item.tier} attachment overview`
              }
            >
              <div className={styles.cardTop}>
                <span
                  className={
                    `${styles.tier} ` +
                    `${getTierClassName(
                      item.tier
                    )}`
                  }
                >
                  {item.tier}
                </span>

                {item.latestSubmissionUrl && (
                  <Link
                    onClick={(): void => {
                      props.onOpenSubmission(
                        item.latestSubmissionUrl
                      );
                    }}
                    ariaLabel={
                      `Open latest ${item.tier} submission`
                    }
                  >
                    Open latest
                  </Link>
                )}
              </div>

              <div
                className={styles.stats}
                aria-label={
                  `${item.tier} submission statistics`
                }
              >
                <div
                  className={
                    styles.statItem
                  }
                >
                  <strong>
                    {item.submissions}
                  </strong>

                  <small>
                    Submissions
                  </small>
                </div>

                <div
                  className={
                    `${styles.statItem} ` +
                    `${styles.withDocs}`
                  }
                >
                  <strong>
                    {
                      item
                        .suppliersWithDocuments
                    }
                  </strong>

                  <small>
                    With documents
                  </small>
                </div>

                <div
                  className={
                    `${styles.statItem} ` +
                    `${
                      hasMissingAttachments
                        ? styles.missing
                        : styles.withDocs
                    }`
                  }
                >
                  <strong>
                    {
                      item
                        .missingAttachments
                    }
                  </strong>

                  <small>
                    Missing
                  </small>
                </div>
              </div>

              <div className={styles.latest}>
                <span
                  className={
                    styles.latestIcon
                  }
                  aria-hidden="true"
                >
                  <Icon iconName="Recent" />
                </span>

                <span
                  className={
                    styles.latestContent
                  }
                >
                  <b>
                    {
                      item.latestSupplier ||
                      'No submissions'
                    }
                  </b>

                  <small>
                    {
                      formatDate(
                        item
                          .latestSubmissionDate
                      )
                    }

                    <span
                      className={
                        styles.detailSeparator
                      }
                      aria-hidden="true"
                    >
                      {' '}·{' '}
                    </span>

                    {
                      item
                        .latestQualification ||
                      '-'
                    }

                    <span
                      className={
                        styles.detailSeparator
                      }
                      aria-hidden="true"
                    >
                      {' '}·{' '}
                    </span>

                    {
                      item.latestRiskRating ||
                      '-'
                    }
                  </small>
                </span>
              </div>
            </article>
          );
        }
      );
    };

  return (
    <section
      className={styles.panel}
      aria-labelledby={
        'tier-attachment-overview-title'
      }
    >
      <header className={styles.header}>
        <h2
          id={
            'tier-attachment-overview-title'
          }
        >
          Submission Attachment Overview
        </h2>

        <p>
          Latest submission and attachment
          status by Tier
        </p>
      </header>

      <div
        className={styles.cardList}
        aria-live="polite"
        aria-busy={props.isLoading}
      >
        {props.isLoading
          ? renderLoadingCards()
          : renderTierCards()}
      </div>

      {!props.isLoading &&
        props.items.length === 0 && (
          <div className={styles.emptyState}>
            <Icon
              iconName="OpenFolderHorizontal"
              className={styles.emptyIcon}
              aria-hidden="true"
            />

            <p>
              No Tier attachment information
              is currently available.
            </p>
          </div>
        )}
    </section>
  );
}

export default TierAttachmentOverview;