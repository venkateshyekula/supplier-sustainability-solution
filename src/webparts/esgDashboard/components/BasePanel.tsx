import * as React from 'react';

import {
  Link
} from '@fluentui/react';

import styles from './ChartShared.module.scss';

export interface IPanelProps {
  title: string;
  subtitle?: string;
  children: React.ReactNode;

  viewDetailsText?: string;

  onViewDetails?(): void;
}

export function Panel(
  props: IPanelProps
): React.ReactElement<IPanelProps> {
  const viewDetailsText: string =
    props.viewDetailsText ||
    'View details';

  const handleViewDetails = (
    event: React.MouseEvent<HTMLElement>
  ): void => {
    event.preventDefault();

    if (!props.onViewDetails) {
      return;
    }

    props.onViewDetails();
  };

  return (
    <section className={styles.panel}>
      <header className={styles.header}>
        <div className={styles.headerContent}>
          <h2 className={styles.panelTitle}>
            {props.title}
          </h2>

          {props.subtitle && (
            <p className={styles.panelSubtitle}>
              {props.subtitle}
            </p>
          )}
        </div>

        {props.onViewDetails && (
          <Link
            className={styles.view}
            title={viewDetailsText}
            ariaLabel={viewDetailsText}
            onClick={handleViewDetails}
          >
            <span>
              {viewDetailsText}
            </span>

            <span
              className={styles.viewIcon}
              aria-hidden="true"
            >
              ↗
            </span>
          </Link>
        )}
      </header>

      <div className={styles.body}>
        {props.children}
      </div>
    </section>
  );
}

export default Panel;