import * as React from 'react';

import {
  Icon
} from '@fluentui/react';

import {
  IActionDatum
} from '../models/IDashboardChartData';

import {
  Panel
} from './BasePanel';

import styles
  from './TopPendingActions.module.scss';

export interface ITopPendingActionsProps {
  items: readonly IActionDatum[];
  isLoading?: boolean;
}

export function TopPendingActions(
  props: ITopPendingActionsProps
): React.ReactElement<ITopPendingActionsProps> {
  const validItems: readonly IActionDatum[] =
    props.items.filter(
      (
        item: IActionDatum
      ): boolean => {
        return Boolean(
          item &&
          item.label &&
          item.label.trim()
        );
      }
    );

  return (
    <Panel
      title="Top Pending Actions"
      subtitle="Most common actions required"
    >
      {props.isLoading ? (
        <div
          className={styles.loading}
          role="status"
          aria-live="polite"
        >
          <span className={styles.loadingLine} />
          <span className={styles.loadingLine} />
          <span className={styles.loadingLine} />

          <span className={styles.screenReaderText}>
            Loading pending actions
          </span>
        </div>
      ) : validItems.length > 0 ? (
        <div
          className={styles.list}
          role="list"
          aria-label="Top pending actions"
        >
          {validItems.map(
            (
              item: IActionDatum,
              index: number
            ): React.ReactElement => {
              const value: number =
                getSafeValue(
                  item.value
                );

              const itemKey: string =
                `${item.label}-${index.toString()}`;

              return (
                <div
                  key={itemKey}
                  className={styles.listItem}
                  role="listitem"
                  aria-label={
                    `${item.label}: ` +
                    `${value.toString()}`
                  }
                >
                  <span
                    className={styles.iconContainer}
                    style={{
                      color:
                        getSafeColor(
                          item.color
                        )
                    }}
                    aria-hidden="true"
                  >
                    <Icon
                      iconName={
                        item.icon ||
                        'TaskLogo'
                      }
                    />
                  </span>

                  <span
                    className={styles.label}
                    title={item.label}
                  >
                    {item.label}
                  </span>

                  <strong
                    className={styles.value}
                    aria-label={
                      `${value.toString()} ` +
                      `${
                        value === 1
                          ? 'action'
                          : 'actions'
                      }`
                    }
                  >
                    {formatValue(value)}
                  </strong>
                </div>
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
            <Icon iconName="Completed" />
          </span>

          <strong>
            No pending actions
          </strong>

          <p>
            There are currently no supplier actions requiring
            attention.
          </p>
        </div>
      )}
    </Panel>
  );
}

function getSafeValue(
  value: number
): number {
  if (
    typeof value !== 'number' ||
    !isFinite(value) ||
    value < 0
  ) {
    return 0;
  }

  return Math.round(value);
}

function formatValue(
  value: number
): string {
  return value.toLocaleString(
    'en-GB'
  );
}

function getSafeColor(
  value?: string
): string {
  if (
    !value ||
    !value.trim()
  ) {
    return '#005dc7';
  }

  return value.trim();
}

export default TopPendingActions;