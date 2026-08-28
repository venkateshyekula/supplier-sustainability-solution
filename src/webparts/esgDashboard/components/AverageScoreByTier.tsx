import * as React from 'react';

import {
  IChartDatum
} from '../models/IDashboardChartData';

import { Panel } from './BasePanel';

import styles from './ChartShared.module.scss';

export interface IAverageScoreByTierProps {
  items: readonly IChartDatum[];
}

export function AverageScoreByTier(
  props: IAverageScoreByTierProps
): React.ReactElement<IAverageScoreByTierProps> {
  const normalizeScore = (
    value: number
  ): number => {
    if (
      !isFinite(value) ||
      isNaN(value)
    ) {
      return 0;
    }

    return Math.min(
      100,
      Math.max(
        0,
        value
      )
    );
  };

  return (
    <Panel
      title="Average Score by Tier"
      subtitle="Average normalized weighted score"
    >
      <div
        className={styles.vertical}
        role="img"
        aria-label="Average ESG score by supplier tier"
      >
        {props.items.map(
          (
            item: IChartDatum
          ): React.ReactElement => {
            const score: number =
              normalizeScore(
                item.value
              );

            return (
              <div
                key={item.label}
                className={styles.vItem}
                aria-label={
                  `${item.label}: ` +
                  `${score.toString()}%`
                }
              >
                <strong
                  className={styles.vValue}
                >
                  {score}%
                </strong>

                <div
                  className={styles.vBarTrack}
                  aria-hidden="true"
                >
                  <div
                    className={styles.vBar}
                    style={{
                      height:
                        `${score.toString()}%`,

                      backgroundColor:
                        item.color
                    }}
                  />
                </div>

                <span
                  className={styles.vLabel}
                >
                  {item.label}
                </span>
              </div>
            );
          }
        )}
      </div>
    </Panel>
  );
}

export default AverageScoreByTier;