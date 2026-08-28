import * as React from 'react';

import {
  IAgingDatum
} from '../models/IDashboardChartData';

import {
  Panel
} from './BasePanel';

import styles
  from './ChartShared.module.scss';

export interface IPendingReviewAgingProps {
  items: readonly IAgingDatum[];
}

export function PendingReviewAging(
  props: IPendingReviewAgingProps
): React.ReactElement<IPendingReviewAgingProps> {
  const validItems: readonly IAgingDatum[] =
    props.items.map(
      (
        item: IAgingDatum
      ): IAgingDatum => {
        return {
          ...item,

          value:
            isFinite(item.value) &&
            item.value > 0
              ? item.value
              : 0
        };
      }
    );

  const total: number =
    validItems.reduce(
      (
        sum: number,
        item: IAgingDatum
      ): number => {
        return sum + item.value;
      },
      0
    );

  const maximumValue: number =
    Math.max(
      1,
      ...validItems.map(
        (
          item: IAgingDatum
        ): number => {
          return item.value;
        }
      )
    );

  return (
    <Panel
      title="Pending Review Aging"
      subtitle="Pending submissions by aging"
    >
      <div
        className={styles.bars}
        role="list"
        aria-label="Pending review aging breakdown"
      >
        {validItems.map(
          (
            item: IAgingDatum
          ): React.ReactElement => {
            const widthPercentage: number =
              calculateBarWidth(
                item.value,
                maximumValue
              );

            const totalPercentage: number =
              calculatePercentage(
                item.value,
                total
              );

            return (
              <div
                className={styles.barRow}
                key={item.label}
                role="listitem"
                aria-label={
                  `${item.label}: ` +
                  `${item.value} pending submissions, ` +
                  `${totalPercentage}% of pending reviews`
                }
              >
                <span
                  className={styles.barLabel}
                  title={item.label}
                >
                  {item.label}
                </span>

                <div
                  className={styles.barTrack}
                  aria-hidden="true"
                >
                  <div
                    className={styles.barFill}
                    style={{
                      width:
                        `${widthPercentage}%`,

                      backgroundColor:
                        item.color
                    }}
                    title={
                      `${item.label}: ` +
                      `${item.value}`
                    }
                  />
                </div>

                <strong
                  className={styles.barValue}
                >
                  {item.value}
                </strong>

                <b
                  className={
                    styles.barPercentage
                  }
                >
                  {totalPercentage}%
                </b>
              </div>
            );
          }
        )}

        {validItems.length === 0 && (
          <div className={styles.chartEmptyState}>
            No pending review aging data is available.
          </div>
        )}

        <div className={styles.barTotal}>
          <span>Total pending reviews</span>

          <strong>{total}</strong>
        </div>
      </div>
    </Panel>
  );
}

function calculateBarWidth(
  value: number,
  maximumValue: number
): number {
  if (
    !isFinite(value) ||
    value <= 0 ||
    !isFinite(maximumValue) ||
    maximumValue <= 0
  ) {
    return 0;
  }

  return Math.min(
    100,
    Math.max(
      0,
      (
        value /
        maximumValue
      ) * 100
    )
  );
}

function calculatePercentage(
  value: number,
  total: number
): number {
  if (
    !isFinite(value) ||
    value <= 0 ||
    !isFinite(total) ||
    total <= 0
  ) {
    return 0;
  }

  return Math.round(
    (
      value /
      total
    ) * 100
  );
}

export default PendingReviewAging;