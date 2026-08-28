import * as React from 'react';

import {
  Icon,
  Link
} from '@fluentui/react';

import {
  IDashboardSubmission
} from '../models/IDashboardSubmission';

import styles
  from './RecentSubmissionsSnapshot.module.scss';

export interface IRecentSubmissionsSnapshotProps {
  items:
    readonly IDashboardSubmission[];

  isLoading?: boolean;

  onOpen(
    item: IDashboardSubmission
  ): void;

  onViewAll?(): void;

  onOpenCompletedQuestionnaires?(): void;
}

const MAX_VISIBLE_ITEMS: number = 5;

export function RecentSubmissionsSnapshot(
  props:
    IRecentSubmissionsSnapshotProps
): React.ReactElement<
  IRecentSubmissionsSnapshotProps
> {
  const visibleItems:
    readonly IDashboardSubmission[] =
      props.items.slice(
        0,
        MAX_VISIBLE_ITEMS
      );

  const formatDate = (
    value?: string
  ): string => {
    if (!value) {
      return '-';
    }

    const date: Date =
      new Date(value);

    if (isNaN(date.getTime())) {
      return '-';
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

  const formatScore = (
    value: number
  ): string => {
    if (!isFinite(value)) {
      return '0.0%';
    }

    return `${value.toFixed(1)}%`;
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

  const getQualificationClassName = (
    qualification: string
  ): string => {
    switch (qualification) {
      case 'Qualified':
        return styles.qualified;

      case 'Conditionally Qualified':
        return styles.conditionallyQualified;

      case 'Not Qualified':
        return styles.notQualified;

      default:
        return styles.defaultStatus;
    }
  };

  const getRiskClassName = (
    riskRating: string
  ): string => {
    switch (riskRating) {
      case 'Low Risk':
        return styles.lowRisk;

      case 'Medium Risk':
        return styles.mediumRisk;

      case 'High Risk':
        return styles.highRisk;

      default:
        return styles.defaultRisk;
    }
  };

  const getRecommendationClassName = (
    recommendation: string
  ): string => {
    switch (recommendation) {
      case 'Approve':
        return styles.approve;

      case 'Corrective Action':
        return styles.correctiveAction;

      case 'Do Not Approve':
        return styles.doNotApprove;

      default:
        return styles.defaultRecommendation;
    }
  };

  return (
    <section
      className={styles.panel}
      aria-labelledby=
        "recent-submissions-snapshot-title"
    >
      <header className={styles.header}>
        <div className={styles.headerContent}>
          <h2
            id=
              "recent-submissions-snapshot-title"
          >
            Recent Submissions Snapshot
          </h2>

          <span className={styles.resultSummary}>
            Showing {visibleItems.length} of{' '}
            {props.items.length}
          </span>
        </div>

        <Link
          className={styles.viewAllLink}
          disabled={
            props.isLoading ||
            props.items.length === 0 ||
            !props.onViewAll
          }
          onClick={(): void => {
            if (props.onViewAll) {
              props.onViewAll();
            }
          }}
        >
          View all

          <Icon
            iconName="OpenInNewWindow"
            aria-hidden="true"
          />
        </Link>
      </header>

      {props.isLoading ? (
        <div
          className={styles.loading}
          role="status"
          aria-live="polite"
        >
          Loading recent submissions...
        </div>
      ) : visibleItems.length === 0 ? (
        <div className={styles.emptyState}>
          <span
            className={styles.emptyIcon}
            aria-hidden="true"
          >
            <Icon iconName="ClipboardList" />
          </span>

          <strong>
            No recent submissions
          </strong>

          <p>
            Completed questionnaire submissions
            will appear here.
          </p>
        </div>
      ) : (
        <div className={styles.scroll}>
          <table
            className={styles.table}
            aria-label=
              "Recent supplier questionnaire submissions"
          >
            <thead>
              <tr>
                <th scope="col">
                  Supplier Name
                </th>

                <th scope="col">
                  Tier
                </th>

                <th scope="col">
                  Score
                </th>

                <th scope="col">
                  Qualification
                </th>

                <th scope="col">
                  Risk Rating
                </th>

                <th scope="col">
                  Submitted On
                </th>

                <th scope="col">
                  Recommendation
                </th>
              </tr>
            </thead>

            <tbody>
              {visibleItems.map(
                (
                  item:
                    IDashboardSubmission
                ): React.ReactElement => {
                  return (
                    <tr key={item.key}>
                      <td>
                        <Link
                          className={
                            styles.supplierLink
                          }
                          title={
                            item.supplierName
                          }
                          onClick={(): void => {
                            props.onOpen(item);
                          }}
                        >
                          {item.supplierName}
                        </Link>
                      </td>

                      <td>
                        <span
                          className={
                            `${styles.tierBadge} ` +
                            `${getTierClassName(
                              item.tier
                            )}`
                          }
                        >
                          {item.tier}
                        </span>
                      </td>

                      <td>
                        <span
                          className={
                            styles.score
                          }
                        >
                          {formatScore(
                            item.score
                          )}
                        </span>
                      </td>

                      <td>
                        <span
                          className={
                            `${styles.statusBadge} ` +
                            `${getQualificationClassName(
                              item.qualification
                            )}`
                          }
                        >
                          {item.qualification}
                        </span>
                      </td>

                      <td>
                        <span
                          className={
                            `${styles.risk} ` +
                            `${getRiskClassName(
                              item.riskRating
                            )}`
                          }
                        >
                          <i aria-hidden="true" />

                          {item.riskRating}
                        </span>
                      </td>

                      <td>
                        <span
                          className={
                            styles.submittedDate
                          }
                        >
                          {formatDate(
                            item.submittedOn
                          )}
                        </span>
                      </td>

                      <td>
                        <span
                          className={
                            `${styles.recommendationBadge} ` +
                            `${getRecommendationClassName(
                              item.recommendation
                            )}`
                          }
                        >
                          {item.recommendation}
                        </span>
                      </td>
                    </tr>
                  );
                }
              )}
            </tbody>
          </table>
        </div>
      )}

      <footer className={styles.footer}>
        <button
          type="button"
          className={styles.footerLink}
          disabled={
            props.isLoading ||
            !props
              .onOpenCompletedQuestionnaires
          }
          onClick={(): void => {
            if (
              props
                .onOpenCompletedQuestionnaires
            ) {
              props
                .onOpenCompletedQuestionnaires();
            }
          }}
        >
          Go to Completed Questionnaires

          <Icon
            className={styles.footerArrow}
            iconName="ChevronRight"
            aria-hidden="true"
          />
        </button>
      </footer>
    </section>
  );
}

export default RecentSubmissionsSnapshot;