import * as React from 'react';

import {
  Icon,
  Shimmer
} from '@fluentui/react';

import {
  IDashboardMetrics
} from '../models/IDashboardMetrics';

import styles
  from './DashboardKpiCards.module.scss';

export interface IDashboardKpiCardsProps {
  metrics: IDashboardMetrics;
  loading: boolean;
}

interface IDashboardKpiCard {
  key: string;
  title: string;
  value: number;
  iconName: string;
  change: number;
  colorClassName: string;
}

export function DashboardKpiCards(
  props: IDashboardKpiCardsProps
): React.ReactElement<IDashboardKpiCardsProps> {
  const cards: readonly IDashboardKpiCard[] = [
    {
      key: 'total',
      title: 'Total Submissions',
      value: props.metrics.total,
      iconName: 'TextDocument',
      change: props.metrics.totalChange,
      colorClassName: styles.blue
    },
    {
      key: 'approved',
      title: 'Approved by ESG',
      value: props.metrics.approved,
      iconName: 'Completed',
      change: props.metrics.approvedChange,
      colorClassName: styles.green
    },
    {
      key: 'pending',
      title: 'Pending ESG Review',
      value: props.metrics.pending,
      iconName: 'Clock',
      change: props.metrics.pendingChange,
      colorClassName: styles.orange
    },
    {
      key: 'requiresAction',
      title: 'Requires Action',
      value: props.metrics.requiresAction,
      iconName: 'Warning',
      change: props.metrics.actionChange,
      colorClassName: styles.red
    },
    {
      key: 'highRisk',
      title: 'High Risk Suppliers',
      value: props.metrics.highRisk,
      iconName: 'Shield',
      change: props.metrics.highRiskChange,
      colorClassName: styles.purple
    }
  ];

  return (
    <section
      className={styles.grid}
      aria-label="Supplier ESG dashboard metrics"
    >
      {cards.map(
        (
          card: IDashboardKpiCard
        ): React.ReactElement => {
          const percentage: number =
            calculatePercentage(
              card.value,
              props.metrics.total
            );

          const isPositiveChange: boolean =
            card.change >= 0;

          return (
            <Shimmer
              key={card.key}
              isDataLoaded={!props.loading}
              ariaLabel={`Loading ${card.title}`}
            >
              <article
                className={
                  `${styles.card} ` +
                  `${card.colorClassName}`
                }
                aria-label={
                  `${card.title}: ${card.value}, ` +
                  `${percentage}% of total`
                }
              >
                <span
                  className={styles.iconContainer}
                  aria-hidden="true"
                >
                  <Icon
                    iconName={card.iconName}
                  />
                </span>

                <div className={styles.cardContent}>
                  <span className={styles.cardTitle}>
                    {card.title}
                  </span>

                  <strong className={styles.cardValue}>
                    {formatNumber(card.value)}
                  </strong>

                  <span className={styles.percentage}>
                    {percentage}% of total
                  </span>

                  <span
                    className={
                      isPositiveChange
                        ? styles.up
                        : styles.down
                    }
                    aria-label={
                      getChangeAriaLabel(
                        card.change
                      )
                    }
                  >
                    <span aria-hidden="true">
                      {isPositiveChange
                        ? '\u2191'
                        : '\u2193'}
                    </span>

                    {' '}

                    {formatChange(card.change)}% vs last 30 days
                  </span>
                </div>
              </article>
            </Shimmer>
          );
        }
      )}
    </section>
  );
}

function calculatePercentage(
  value: number,
  total: number
): number {
  if (
    !isFinite(value) ||
    !isFinite(total) ||
    value < 0 ||
    total <= 0
  ) {
    return 0;
  }

  return Math.round(
    (value / total) * 100
  );
}

function formatNumber(
  value: number
): string {
  if (
    !isFinite(value) ||
    value < 0
  ) {
    return '0';
  }

  return Math.round(value).toLocaleString(
    'en-GB'
  );
}

function formatChange(
  value: number
): string {
  if (!isFinite(value)) {
    return '0';
  }

  const absoluteValue: number =
    Math.abs(value);

  const roundedValue: number =
    Math.round(
      absoluteValue * 10
    ) / 10;

  return roundedValue.toString();
}

function getChangeAriaLabel(
  change: number
): string {
  if (
    !isFinite(change) ||
    change === 0
  ) {
    return 'No change compared with the last 30 days';
  }

  const direction: string =
    change > 0
      ? 'increase'
      : 'decrease';

  return (
    `${formatChange(change)} percent ${direction} ` +
    'compared with the last 30 days'
  );
}

export default DashboardKpiCards;