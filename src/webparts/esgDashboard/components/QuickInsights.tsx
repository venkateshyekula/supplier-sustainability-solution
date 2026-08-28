import * as React from 'react';

import {
  Icon
} from '@fluentui/react';

import {
  DashboardInsightTone,
  IDashboardInsight
} from '../models/IDashboardInsight';

import styles from './QuickInsights.module.scss';

export interface IQuickInsightsProps {
  items: readonly IDashboardInsight[];
}

export function QuickInsights(
  props: IQuickInsightsProps
): React.ReactElement<IQuickInsightsProps> {
  const getToneClassName = (
    tone: DashboardInsightTone
  ): string => {
    switch (tone) {
      case 'positive':
        return styles.positive;

      case 'warning':
        return styles.warning;

      case 'danger':
        return styles.danger;

      case 'info':
        return styles.info;

      default:
        return styles.info;
    }
  };

  return (
    <section
      className={styles.panel}
      aria-labelledby="quick-insights-title"
    >
      <header className={styles.header}>
        <h2 id="quick-insights-title">
          Quick Insights
        </h2>
      </header>

      {props.items.length > 0 ? (
        <div className={styles.insights}>
          {props.items.map(
            (
              insight: IDashboardInsight
            ): React.ReactElement => {
              return (
                <article
                  key={insight.key}
                  className={
                    `${styles.insight} ` +
                    `${getToneClassName(
                      insight.tone
                    )}`
                  }
                >
                  <span
                    className={styles.iconContainer}
                    aria-hidden="true"
                  >
                    <Icon
                      className={styles.icon}
                      iconName={insight.icon}
                    />
                  </span>

                  <div className={styles.insightContent}>
                    <strong className={styles.insightTitle}>
                      {insight.title}
                    </strong>

                    <p className={styles.insightDescription}>
                      {insight.description}
                    </p>
                  </div>
                </article>
              );
            }
          )}
        </div>
      ) : (
        <div
          className={styles.emptyState}
          role="status"
        >
          <span
            className={styles.emptyIcon}
            aria-hidden="true"
          >
            <Icon iconName="Info" />
          </span>

          <strong>
            No insights available
          </strong>

          <p>
            Insights will appear after ESG submission data
            has been loaded and analyzed.
          </p>
        </div>
      )}
    </section>
  );
}

export default QuickInsights;