import * as React from 'react';
import { useEffect, useState } from 'react';

import { Icon } from '@fluentui/react';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import { IQuestionnaireQuickLinksProps } from './IQuestionnaireQuickLinksProps';
import styles from './QuestionnaireQuickLinks.module.scss';

type LinkKind = 'questionnaire' | 'documents';
type TierNumber = 1 | 2 | 3;

interface ILinkConfiguration {
  key: string;
  tier: TierNumber;
  kind: LinkKind;
  title: string;
  description: string;
  sourceTitle: string;
  iconName: string;
  cardClassName: string;
  iconClassName: string;
}

interface IResolvedLink extends ILinkConfiguration {
  url?: string;
  errorMessage?: string;
}

interface ISharePointMetadataResponse {
  Title?: string;
  BaseTemplate?: number;
  RootFolder?: {
    ServerRelativeUrl?: string;
  };
}

const PREPROD_DEFAULTS = {
  tier1QuestionnaireListTitle: 'CASSTECH_SSQ',
  tier2QuestionnaireListTitle:
    'Tier 2 ESG Procurement Questionnaire',
  tier3QuestionnaireListTitle:
    'Supplier Sustainability Questionnaires Tier 3',
  tier1DocumentLibraryTitle: 'TestSupply',
  tier2DocumentLibraryTitle: 'SSQ-Tier2 Attachments',
  tier3DocumentLibraryTitle: 'SSQ-Tier3 Attachments'
};

const getConfiguredTitle = (
  value: string,
  fallback: string
): string => {
  const trimmedValue: string = value ? value.trim() : '';
  return trimmedValue || fallback;
};

const joinClasses = (...classNames: string[]): string => {
  return classNames
    .filter((className: string): boolean => Boolean(className))
    .join(' ');
};

const escapeODataString = (value: string): string => {
  return value.replace(/'/g, "''");
};

const getConfigurations = (
  props: IQuestionnaireQuickLinksProps
): readonly ILinkConfiguration[] => {
  return [
    {
      key: 'tier1-questionnaire',
      tier: 1,
      kind: 'questionnaire',
      title: 'Tier 1 Responses',
      description: 'View all completed Tier 1 questionnaires',
      sourceTitle: getConfiguredTitle(
        props.tier1QuestionnaireListTitle,
        PREPROD_DEFAULTS.tier1QuestionnaireListTitle
      ),
      iconName: 'OpenFolderHorizontal',
      cardClassName: styles.blueCard,
      iconClassName: styles.blueIcon
    },
    {
      key: 'tier2-questionnaire',
      tier: 2,
      kind: 'questionnaire',
      title: 'Tier 2 Responses',
      description: 'View all completed Tier 2 questionnaires',
      sourceTitle: getConfiguredTitle(
        props.tier2QuestionnaireListTitle,
        PREPROD_DEFAULTS.tier2QuestionnaireListTitle
      ),
      iconName: 'OpenFolderHorizontal',
      cardClassName: styles.goldCard,
      iconClassName: styles.goldIcon
    },
    {
      key: 'tier3-questionnaire',
      tier: 3,
      kind: 'questionnaire',
      title: 'Tier 3 Responses',
      description: 'View all completed Tier 3 questionnaires',
      sourceTitle: getConfiguredTitle(
        props.tier3QuestionnaireListTitle,
        PREPROD_DEFAULTS.tier3QuestionnaireListTitle
      ),
      iconName: 'OpenFolderHorizontal',
      cardClassName: styles.redCard,
      iconClassName: styles.redIcon
    },
    {
      key: 'tier1-documents',
      tier: 1,
      kind: 'documents',
      title: 'Tier 1 Documents',
      description: 'View all supporting documents',
      sourceTitle: getConfiguredTitle(
        props.tier1DocumentLibraryTitle,
        PREPROD_DEFAULTS.tier1DocumentLibraryTitle
      ),
      iconName: 'Attach',
      cardClassName: styles.blueCard,
      iconClassName: styles.blueIcon
    },
    {
      key: 'tier2-documents',
      tier: 2,
      kind: 'documents',
      title: 'Tier 2 Documents',
      description: 'View all supporting documents',
      sourceTitle: getConfiguredTitle(
        props.tier2DocumentLibraryTitle,
        PREPROD_DEFAULTS.tier2DocumentLibraryTitle
      ),
      iconName: 'Attach',
      cardClassName: styles.goldCard,
      iconClassName: styles.goldIcon
    },
    {
      key: 'tier3-documents',
      tier: 3,
      kind: 'documents',
      title: 'Tier 3 Documents',
      description: 'View all supporting documents',
      sourceTitle: getConfiguredTitle(
        props.tier3DocumentLibraryTitle,
        PREPROD_DEFAULTS.tier3DocumentLibraryTitle
      ),
      iconName: 'Attach',
      cardClassName: styles.redCard,
      iconClassName: styles.redIcon
    }
  ];
};

const buildAllItemsUrl = (
  serverRelativeUrl: string,
  kind: LinkKind
): string => {
  const normalizedUrl: string = serverRelativeUrl.replace(/\/$/, '');

  return kind === 'documents'
    ? `${normalizedUrl}/Forms/AllItems.aspx`
    : `${normalizedUrl}/AllItems.aspx`;
};

export function QuestionnaireQuickLinks(
  props: IQuestionnaireQuickLinksProps
): React.ReactElement {
  const [links, setLinks] = useState<IResolvedLink[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [errorMessage, setErrorMessage] = useState<string>('');

  useEffect((): (() => void) => {
    let isMounted: boolean = true;

    const resolveLink = async (
      configuration: ILinkConfiguration
    ): Promise<IResolvedLink> => {
      try {
        const webUrl: string =
          props.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
        const escapedTitle: string = escapeODataString(
          configuration.sourceTitle
        );
        const endpoint: string =
          `${webUrl}/_api/web/lists/getbytitle('${escapedTitle}')` +
          '?$select=Title,BaseTemplate,RootFolder/ServerRelativeUrl' +
          '&$expand=RootFolder';

        const response: SPHttpClientResponse =
          await props.context.spHttpClient.get(
            endpoint,
            SPHttpClient.configurations.v1,
            {
              headers: {
                Accept: 'application/json;odata.metadata=none'
              }
            }
          );

        if (!response.ok) {
          const responseText: string = await response.text();

          throw new Error(
            `HTTP ${response.status} ${response.statusText}. ` +
            responseText
          );
        }

        const metadata: ISharePointMetadataResponse =
          await response.json() as ISharePointMetadataResponse;
        const serverRelativeUrl: string =
          metadata.RootFolder?.ServerRelativeUrl || '';

        if (!serverRelativeUrl) {
          throw new Error(
            'Root folder URL was not returned by SharePoint.'
          );
        }

        return {
          ...configuration,
          url: buildAllItemsUrl(
            serverRelativeUrl,
            configuration.kind
          )
        };
      } catch (error: unknown) {
        const detail: string =
          error instanceof Error ? error.message : String(error);

        console.error('[Questionnaire Quick Links]', {
          webUrl: props.context.pageContext.web.absoluteUrl,
          sourceTitle: configuration.sourceTitle,
          kind: configuration.kind,
          tier: configuration.tier,
          error: detail
        });

        return {
          ...configuration,
          errorMessage:
            `Unable to open "${configuration.sourceTitle}". ` + detail
        };
      }
    };

    const loadLinks = async (): Promise<void> => {
      setLoading(true);
      setErrorMessage('');

      const configurations: readonly ILinkConfiguration[] =
        getConfigurations(props);

      console.info('[Questionnaire Quick Links Configuration]', {
        webUrl: props.context.pageContext.web.absoluteUrl,
        sources: configurations.map(
          (configuration: ILinkConfiguration): string =>
            configuration.sourceTitle
        )
      });

      const resolvedLinks: IResolvedLink[] = await Promise.all(
        configurations.map(
          (
            configuration: ILinkConfiguration
          ): Promise<IResolvedLink> => resolveLink(configuration)
        )
      );

      if (!isMounted) {
        return;
      }

      setLinks(resolvedLinks);

      const failedLinks: IResolvedLink[] = resolvedLinks.filter(
        (link: IResolvedLink): boolean => Boolean(link.errorMessage)
      );

      if (failedLinks.length > 0) {
        setErrorMessage(
          `${failedLinks.length} SharePoint source` +
          `${failedLinks.length === 1 ? '' : 's'} could not be resolved: ` +
          failedLinks
            .map((link: IResolvedLink): string => link.sourceTitle)
            .join(', ') +
          '. Check the dashboard property-pane titles and site location.'
        );
      }

      setLoading(false);
    };

    loadLinks().catch((error: unknown): void => {
      console.error(
        'Unable to load questionnaire quick links.',
        error
      );

      if (isMounted) {
        setLinks([]);
        setLoading(false);
        setErrorMessage(
          'The questionnaire links could not be loaded from SharePoint.'
        );
      }
    });

    return (): void => {
      isMounted = false;
    };
  }, [
    props.context,
    props.tier1QuestionnaireListTitle,
    props.tier2QuestionnaireListTitle,
    props.tier3QuestionnaireListTitle,
    props.tier1DocumentLibraryTitle,
    props.tier2DocumentLibraryTitle,
    props.tier3DocumentLibraryTitle
  ]);

  const openLink = (link: IResolvedLink): void => {
    if (!link.url) {
      return;
    }

    window.open(link.url, '_blank', 'noopener,noreferrer');
  };

  const renderCards = (kind: LinkKind): React.ReactElement => {
    const sectionLinks: IResolvedLink[] = links.filter(
      (link: IResolvedLink): boolean => link.kind === kind
    );

    return (
      <div className={styles.cardGrid}>
        {sectionLinks.map(
          (link: IResolvedLink): React.ReactElement => (
            <button
              key={link.key}
              type="button"
              className={joinClasses(
                styles.card,
                link.cardClassName
              )}
              disabled={!link.url}
              onClick={(): void => openLink(link)}
              aria-label={`${link.title}. ${link.description}`}
              title={link.errorMessage || link.sourceTitle}
            >
              <Icon
                iconName={link.iconName}
                className={joinClasses(
                  styles.icon,
                  link.iconClassName
                )}
                aria-hidden="true"
              />

              <h3 className={styles.cardTitle}>{link.title}</h3>
              <p className={styles.cardDescription}>
                {link.description}
              </p>

              {link.errorMessage && (
                <span className={styles.status}>Source unavailable</span>
              )}

              <span className={styles.arrow} aria-hidden="true">
                ›
              </span>
            </button>
          )
        )}
      </div>
    );
  };

  if (loading) {
    return (
      <div
        className={styles.loading}
        role="status"
        aria-live="polite"
      >
        Loading questionnaire links...
      </div>
    );
  }

  return (
    <section
      className={styles.questionnaireQuickLinks}
      aria-label="Supplier sustainability questionnaire links"
    >
      {errorMessage && (
        <div className={styles.errorMessage} role="alert">
          {errorMessage}
        </div>
      )}

      <div className={styles.contentGrid}>
        <section
          className={styles.section}
          aria-labelledby="completed-questionnaires-title"
        >
          <h2
            id="completed-questionnaires-title"
            className={styles.sectionTitle}
          >
            Completed Questionnaires
          </h2>
          {renderCards('questionnaire')}
        </section>

        <div className={styles.divider} aria-hidden="true" />

        <section
          className={styles.section}
          aria-labelledby="supporting-documents-title"
        >
          <h2
            id="supporting-documents-title"
            className={styles.sectionTitle}
          >
            Supporting Documents
          </h2>
          {renderCards('documents')}
        </section>
      </div>
    </section>
  );
}

export default QuestionnaireQuickLinks;