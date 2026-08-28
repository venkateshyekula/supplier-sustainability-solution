import * as React from 'react';

import {
  Icon
} from '@fluentui/react';

import styles from './OverallEsgScore.module.scss';

export interface IOverallEsgScoreProps {
  score: number;
  change: number;
  target: number;
}

export function OverallEsgScore(
  props: IOverallEsgScoreProps
): React.ReactElement<IOverallEsgScoreProps> {
  const normalizePercentage = (
    value: number
  ): number => {
    if (!isFinite(value)) {
      return 0;
    }

    return Math.max(
      0,
      Math.min(
        100,
        Math.round(value * 100) / 100
      )
    );
  };

  const score: number =
    normalizePercentage(
      props.score
    );

  const target: number =
    normalizePercentage(
      props.target
    );

  const change: number =
    isFinite(props.change)
      ? Math.round(props.change * 100) / 100
      : 0;

  const isPositiveChange: boolean =
    change >= 0;

  const changeClassName: string =
    isPositiveChange
      ? styles.up
      : styles.down;

  const changeIconName: string =
    isPositiveChange
      ? 'Up'
      : 'Down';

  const changeLabel: string =
    isPositiveChange
      ? 'increase'
      : 'decrease';

  const formattedScore: string =
    `${score.toString()}%`;

  const formattedChange: string =
    `${Math.abs(change).toString()}%`;

  const formattedTarget: string =
    `${target.toString()}%`;

  const targetDifference: number =
    Math.round(
      (
        score -
        target
      ) *
      100
    ) / 100;

  const targetStatus: string =
    targetDifference >= 0
      ? 'Target achieved'
      : `${Math.abs(targetDifference)}% below target`;

  return (
    <section
      className={styles.panel}
      aria-labelledby="overall-esg-score-title"
    >
      <header className={styles.header}>
        <h2 id="overall-esg-score-title">
          Overall ESG Score
        </h2>

        <span className={styles.averageLabel}>
          Average
        </span>
      </header>

      <div className={styles.row}>
        <div
          className={styles.ring}
          style={{
            background:
              `conic-gradient(` +
              `#36a852 0% ${score}%, ` +
              `#e6ebef ${score}% 100%)`
          }}
          role="img"
          aria-label={
            `Overall ESG score: ${formattedScore}`
          }
        >
          <div className={styles.ringInner}>
            <strong className={styles.score}>
              {formattedScore}
            </strong>

            <span className={styles.scoreLabel}>
              ESG Score
            </span>
          </div>
        </div>

        <div className={styles.details}>
          <div className={styles.changeSection}>
            <strong
              className={
                `${styles.change} ` +
                `${changeClassName}`
              }
              aria-label={
                `${formattedChange} ${changeLabel} ` +
                'compared with the last 30 days'
              }
            >
              <Icon
                iconName={changeIconName}
                className={styles.changeIcon}
                aria-hidden="true"
              />

              <span>
                {formattedChange}
              </span>
            </strong>

            <span className={styles.changeDescription}>
              vs last 30 days
            </span>
          </div>

          <div className={styles.divider} />

          <div className={styles.targetSection}>
            <div className={styles.targetHeading}>
              <span>
                Target
              </span>

              <Icon
                iconName="Info"
                className={styles.infoIcon}
                title="Configured ESG target"
                aria-label="Configured ESG target"
              />
            </div>

            <strong className={styles.targetValue}>
              {formattedTarget}
            </strong>

            <span
              className={
                targetDifference >= 0
                  ? styles.targetAchieved
                  : styles.targetBelow
              }
            >
              {targetStatus}
            </span>
          </div>
        </div>
      </div>
    </section>
  );
}

export default OverallEsgScore;