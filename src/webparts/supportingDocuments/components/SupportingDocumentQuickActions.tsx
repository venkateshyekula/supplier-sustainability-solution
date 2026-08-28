import * as React from 'react';

import {
  Icon
} from '@fluentui/react';

import styles
  from './SupportingDocumentQuickActions.module.scss';

export interface ISupportingDocumentQuickActionsProps {
  tier1Url?: string;
  tier2Url?: string;
  tier3Url?: string;
  guidanceUrl?: string;

  onOpen(
    url?: string
  ): void;
}

interface IQuickAction {
  key: string;
  title: string;
  description: string;
  iconName: string;
  url?: string;
}

export function SupportingDocumentQuickActions(
  props:
    ISupportingDocumentQuickActionsProps
): React.ReactElement<
  ISupportingDocumentQuickActionsProps
> {
  const actions:
    readonly IQuickAction[] = [
      {
        key:
          'tier1',

        title:
          'Open Tier 1 Documents',

        description:
          'Browse Tier 1 supplier folders',

        iconName:
          'FabricFolder',

        url:
          props.tier1Url
      },
      {
        key:
          'tier2',

        title:
          'Open Tier 2 Documents',

        description:
          'Browse Tier 2 supplier folders',

        iconName:
          'FabricFolder',

        url:
          props.tier2Url
      },
      {
        key:
          'tier3',

        title:
          'Open Tier 3 Documents',

        description:
          'Browse Tier 3 supplier folders',

        iconName:
          'FabricFolder',

        url:
          props.tier3Url
      },
      {
        key:
          'guidance',

        title:
          'Document Guidance',

        description:
          'Help with supporting documents',

        iconName:
          'Help',

        url:
          props.guidanceUrl
      }
    ];

  return (
    <section
      className={styles.panel}
      aria-labelledby=
        "supporting-document-quick-actions-title"
    >
      <header className={styles.header}>
        <h2
          id=
            "supporting-document-quick-actions-title"
        >
          Quick Actions
        </h2>
      </header>

      <div className={styles.actions}>
        {actions.map(
          (
            action:
              IQuickAction
          ): React.ReactElement => {
            const isDisabled: boolean =
              !action.url;

            return (
              <button
                key={action.key}
                type="button"
                className={styles.action}
                disabled={isDisabled}
                title={
                  isDisabled
                    ? `${action.title} is not configured`
                    : action.title
                }
                onClick={(): void => {
                  if (
                    action.url
                  ) {
                    props.onOpen(
                      action.url
                    );
                  }
                }}
              >
                <span
                  className={
                    styles.actionIcon
                  }
                  aria-hidden="true"
                >
                  <Icon
                    iconName={
                      action.iconName
                    }
                  />
                </span>

                <span
                  className={
                    styles.actionText
                  }
                >
                  <strong>
                    {action.title}
                  </strong>

                  <small>
                    {action.description}
                  </small>
                </span>

                <Icon
                  className={styles.arrow}
                  iconName="ChevronRight"
                  aria-hidden="true"
                />
              </button>
            );
          }
        )}
      </div>
    </section>
  );
}

export default SupportingDocumentQuickActions;