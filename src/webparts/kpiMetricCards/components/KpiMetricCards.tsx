import * as React from 'react';
import { useEffect, useState } from 'react';

import { Icon } from '@fluentui/react';

import { getSP } from '../../../pnpConfig';
import { IKpiMetricCardsProps } from './IKpiMetricCardsProps';

interface IKpiData {
  total: number;
  pending: number;
  approved: number;
  actionRequired: number;
  highRisk: number;
}

interface IResponseItem {
  Id?: number;
  normalizedScore?: number;
}

interface IListConfiguration {
  title: string;
  percentageInternalName: string;
  maximumWeightedScore: number;
}

interface IListFetchResult {
  listTitle: string;
  items: IResponseItem[];
  succeeded: boolean;
  errorMessage?: string;
}

interface IKpiCardConfiguration {
  key: keyof IKpiData;
  title: string;
  description: string;
  color: string;
  iconName: string;
  iconBackgroundColor: string;
  cardBackgroundColor: string;
  borderColor: string;
}

const EMPTY_KPI_DATA: IKpiData = {
  total: 0,
  pending: 0,
  approved: 0,
  actionRequired: 0,
  highRisk: 0
};

const DEFAULT_LIST_CONFIGURATIONS:
  readonly IListConfiguration[] = [
    {
      title: 'CASSTECH_SSQ',
      percentageInternalName:
        'OverallQuestionsPercentage',
      maximumWeightedScore: 10
    },
    {
      title:
        'Tier 2 ESG Procurement Questionnaire',
      percentageInternalName:
        'OverallQuestionsPercentage',
      maximumWeightedScore: 5
    },
    {
      title:
        'Supplier Sustainability Questionnaires Tier 3',
      percentageInternalName:
        'OverallQuestionsPercentage',
      maximumWeightedScore: 2
    }
  ];

/*
 * KPI bands applied to each tier's normalized completion ratio:
 * Approved: 80% and above
 * Pending: 50% to below 80%
 * Procurement action: 20% to below 50%
 * High risk: below 20%
 */
const APPROVED_MINIMUM = 0.8;
const PENDING_MINIMUM = 0.5;
const ACTION_REQUIRED_MINIMUM = 0.2;

const KPI_CARDS:
  readonly IKpiCardConfiguration[] = [
    {
      key: 'total',
      title: 'Total Submissions',
      description: 'All time',
      color: '#1665a8',
      iconName: 'TextDocument',
      iconBackgroundColor: '#e7f1fb',
      cardBackgroundColor: '#f6faff',
      borderColor: '#d9e9f8'
    },
    {
      key: 'pending',
      title: 'Pending ESG Review',
      description: 'Awaiting review',
      color: '#9a6700',
      iconName: 'Clock',
      iconBackgroundColor: '#fff1cc',
      cardBackgroundColor: '#fffaf1',
      borderColor: '#f7e6bd'
    },
    {
      key: 'approved',
      title: 'Approved by ESG',
      description: 'Ready to onboard',
      color: '#16833a',
      iconName: 'Completed',
      iconBackgroundColor: '#e0f4e5',
      cardBackgroundColor: '#f5fbf6',
      borderColor: '#d5ecd9'
    },
    {
      key: 'actionRequired',
      title: 'Requires Procurement Action',
      description: 'Attention needed',
      color: '#c4314b',
      iconName: 'ErrorBadge',
      iconBackgroundColor: '#fbe1e5',
      cardBackgroundColor: '#fff6f7',
      borderColor: '#f4d7dc'
    },
    {
      key: 'highRisk',
      title: 'High Risk Suppliers',
      description: 'Review required',
      color: '#5c2d91',
      iconName: 'ShieldAlert',
      iconBackgroundColor: '#eee5f8',
      cardBackgroundColor: '#faf7fd',
      borderColor: '#e4d9f1'
    }
  ];

/* V8 invalidates values cached by earlier scoring logic. */
const CACHE_KEY = 'SSQ_KPI_METRICS_DATA_V8';
const CACHE_TIME_KEY = 'SSQ_KPI_METRICS_TIMESTAMP_V8';
const TTL_MS = 10 * 60 * 1000;

const getErrorMessage = (error: unknown): string => {
  if (error instanceof Error) {
    return error.message;
  }

  return String(error);
};

const createCacheKeySegment = (
  value: string
): string => {
  return value
    .replace(/[^a-zA-Z0-9]/g, '_')
    .substring(0, 180);
};

/*
 * SharePoint already calculates the weighted value:
 * Tier 1 maximum = 10, Tier 2 maximum = 5, Tier 3 maximum = 2.
 * This function only converts that stored value into a 0-1 ratio for
 * consistent KPI classification. It does not replace the SharePoint value.
 */
const normalizeWeightedScore = (
  weightedScore: unknown,
  maximumWeightedScore: number
): number | undefined => {
  if (
    weightedScore === null ||
    weightedScore === undefined ||
    weightedScore === ''
  ) {
    return undefined;
  }

  if (
    !isFinite(maximumWeightedScore) ||
    maximumWeightedScore <= 0
  ) {
    return undefined;
  }

  let normalizedValue: unknown = weightedScore;

  if (typeof normalizedValue === 'string') {
    normalizedValue = normalizedValue
      .replace('%', '')
      .replace(',', '.')
      .trim();
  }

  const numericScore: number =
    typeof normalizedValue === 'number'
      ? normalizedValue
      : Number(normalizedValue);

  if (
    isNaN(numericScore) ||
    !isFinite(numericScore) ||
    numericScore < 0
  ) {
    return undefined;
  }

  const completionRatio: number =
    numericScore / maximumWeightedScore;

  if (
    !isFinite(completionRatio) ||
    completionRatio < 0
  ) {
    return undefined;
  }

  return Math.min(completionRatio, 1);
};

const isKpiData = (
  value: unknown
): value is IKpiData => {
  if (
    typeof value !== 'object' ||
    value === null
  ) {
    return false;
  }

  const candidate: Partial<IKpiData> =
    value as Partial<IKpiData>;

  return (
    typeof candidate.total === 'number' &&
    typeof candidate.pending === 'number' &&
    typeof candidate.approved === 'number' &&
    typeof candidate.actionRequired === 'number' &&
    typeof candidate.highRisk === 'number'
  );
};

const clearCachedKpi = (
  cacheKey: string,
  cacheTimeKey: string
): void => {
  window.sessionStorage.removeItem(cacheKey);
  window.sessionStorage.removeItem(cacheTimeKey);
};

const readCachedKpi = (
  cacheKey: string,
  cacheTimeKey: string
): IKpiData | undefined => {
  try {
    const cachedData: string | null =
      window.sessionStorage.getItem(cacheKey);

    const cachedTimestamp: string | null =
      window.sessionStorage.getItem(cacheTimeKey);

    if (!cachedData || !cachedTimestamp) {
      return undefined;
    }

    const timestamp: number =
      Number(cachedTimestamp);

    if (!isFinite(timestamp)) {
      clearCachedKpi(cacheKey, cacheTimeKey);
      return undefined;
    }

    const cacheAge: number =
      Date.now() - timestamp;

    if (
      cacheAge < 0 ||
      cacheAge >= TTL_MS
    ) {
      clearCachedKpi(cacheKey, cacheTimeKey);
      return undefined;
    }

    const parsedData: unknown =
      JSON.parse(cachedData);

    if (!isKpiData(parsedData)) {
      clearCachedKpi(cacheKey, cacheTimeKey);
      return undefined;
    }

    return parsedData;
  } catch (error: unknown) {
    console.warn(
      'Unable to read KPI metrics cache:',
      error
    );

    clearCachedKpi(cacheKey, cacheTimeKey);
    return undefined;
  }
};

const writeCachedKpi = (
  data: IKpiData,
  cacheKey: string,
  cacheTimeKey: string
): void => {
  try {
    window.sessionStorage.setItem(
      cacheKey,
      JSON.stringify(data)
    );

    window.sessionStorage.setItem(
      cacheTimeKey,
      Date.now().toString()
    );
  } catch (error: unknown) {
    console.warn(
      'Unable to save KPI metrics to sessionStorage:',
      error
    );
  }
};

const wrapperStyle: React.CSSProperties = {
  width: '100%',
  fontFamily: '"Segoe UI", sans-serif'
};

const containerStyle: React.CSSProperties = {
  display: 'grid',
  gridTemplateColumns:
    'repeat(auto-fit, minmax(210px, 1fr))',
  gap: '12px',
  width: '100%',
  boxSizing: 'border-box'
};

const cardStyle: React.CSSProperties = {
  minHeight: '130px',
  padding: '18px 16px',
  display: 'flex',
  alignItems: 'center',
  gap: '14px',
  borderRadius: '12px',
  boxShadow: '0 2px 8px rgba(0, 0, 0, 0.06)',
  boxSizing: 'border-box'
};

const cardTextStyle: React.CSSProperties = {
  minWidth: 0,
  display: 'flex',
  flexDirection: 'column',
  gap: '2px'
};

const titleStyle: React.CSSProperties = {
  fontSize: '14px',
  lineHeight: 1.3,
  fontWeight: 600,
  color: '#201f1e'
};

const descriptionStyle: React.CSSProperties = {
  fontSize: '12px',
  lineHeight: 1.3,
  color: '#605e5c'
};

const warningStyle: React.CSSProperties = {
  marginBottom: '14px',
  padding: '10px 14px',
  fontSize: '13px',
  lineHeight: 1.45,
  color: '#8a6d1d',
  backgroundColor: '#fff4ce',
  border: '1px solid #fce100',
  borderRadius: '6px'
};

export function KpiMetricCards(
  props: IKpiMetricCardsProps
): React.ReactElement {
  const [kpi, setKpi] =
    useState<IKpiData>(EMPTY_KPI_DATA);
  const [loading, setLoading] =
    useState<boolean>(true);
  const [errorMessage, setErrorMessage] =
    useState<string>('');

  useEffect((): (() => void) => {
    let isMounted: boolean = true;

    const listConfigurations:
      readonly IListConfiguration[] = [
        {
          ...DEFAULT_LIST_CONFIGURATIONS[0],
          title:
            props.tier1ListTitle ||
            DEFAULT_LIST_CONFIGURATIONS[0].title
        },
        {
          ...DEFAULT_LIST_CONFIGURATIONS[1],
          title:
            props.tier2ListTitle ||
            DEFAULT_LIST_CONFIGURATIONS[1].title
        },
        {
          ...DEFAULT_LIST_CONFIGURATIONS[2],
          title:
            props.tier3ListTitle ||
            DEFAULT_LIST_CONFIGURATIONS[2].title
        }
      ];

    const listConfigurationSignature: string =
      listConfigurations
        .map(
          (
            configuration: IListConfiguration
          ): string => {
            return configuration.title;
          }
        )
        .join('|');

    const cacheSignature: string =
      createCacheKeySegment(
        listConfigurationSignature
      );

    const cacheKey: string =
      `${CACHE_KEY}_${cacheSignature}`;

    const cacheTimeKey: string =
      `${CACHE_TIME_KEY}_${cacheSignature}`;

    const fetchListItems = async (
      configuration: IListConfiguration
    ): Promise<IListFetchResult> => {
      try {
        const sp = getSP(props.context);

        const items = await sp.web.lists
          .getByTitle(configuration.title)
          .items.select(
            'Id',
            configuration.percentageInternalName
          )
          .top(5000)();

        console.info(
          `[KPI RAW DATA] ${configuration.title} ` +
          '(first 3 items):',
          items.slice(0, 3)
        );

        const mappedItems: IResponseItem[] =
          items.map(
            (
              item: Record<string, unknown>
            ): IResponseItem => {
              const idValue: unknown = item.Id;
              const rawScore: unknown =
                item[
                  configuration.percentageInternalName
                ];

              return {
                Id:
                  typeof idValue === 'number'
                    ? idValue
                    : undefined,
                normalizedScore:
                  normalizeWeightedScore(
                    rawScore,
                    configuration.maximumWeightedScore
                  )
              };
            }
          );

        console.info(
          `[KPI] ${configuration.title}: ` +
          `${mappedItems.length} items received.`
        );

        return {
          listTitle: configuration.title,
          items: mappedItems,
          succeeded: true
        };
      } catch (error: unknown) {
        const message: string =
          getErrorMessage(error);

        console.error(
          `[KPI] Failed to retrieve ` +
          `"${configuration.title}".`,
          error
        );

        return {
          listTitle: configuration.title,
          items: [],
          succeeded: false,
          errorMessage: message
        };
      }
    };

    const fetchMetrics = async (): Promise<void> => {
      if (!isMounted) {
        return;
      }

      setLoading(true);
      setErrorMessage('');

      const cachedKpi:
        IKpiData | undefined =
          readCachedKpi(
            cacheKey,
            cacheTimeKey
          );

      if (cachedKpi) {
        setKpi(cachedKpi);
        setLoading(false);
        console.info(
          '[KPI] Metrics loaded from sessionStorage.'
        );
        return;
      }

      try {
        const results: IListFetchResult[] =
          await Promise.all(
            listConfigurations.map(
              (
                configuration: IListConfiguration
              ): Promise<IListFetchResult> => {
                return fetchListItems(configuration);
              }
            )
          );

        if (!isMounted) {
          return;
        }

        const successfulResults:
          IListFetchResult[] =
            results.filter(
              (
                result: IListFetchResult
              ): boolean => {
                return result.succeeded;
              }
            );

        const failedResults:
          IListFetchResult[] =
            results.filter(
              (
                result: IListFetchResult
              ): boolean => {
                return !result.succeeded;
              }
            );

        if (successfulResults.length === 0) {
          setKpi(EMPTY_KPI_DATA);
          setErrorMessage(
            'KPI records could not be loaded from any ' +
            'SharePoint list. Open the browser console ' +
            'for details.'
          );
          return;
        }

        const allItems: IResponseItem[] =
          successfulResults.reduce<IResponseItem[]>(
            (
              accumulatedItems: IResponseItem[],
              currentResult: IListFetchResult
            ): IResponseItem[] => {
              return accumulatedItems.concat(
                currentResult.items
              );
            },
            []
          );

        const scoredRatios: number[] = allItems
          .map(
            (
              item: IResponseItem
            ): number | undefined => {
              return item.normalizedScore;
            }
          )
          .filter(
            (
              score: number | undefined
            ): score is number => {
              return score !== undefined;
            }
          );

        const freshKpi: IKpiData = {
          total: allItems.length,

          approved: scoredRatios.filter(
            (score: number): boolean => {
              return score >= APPROVED_MINIMUM;
            }
          ).length,

          pending: scoredRatios.filter(
            (score: number): boolean => {
              return (
                score >= PENDING_MINIMUM &&
                score < APPROVED_MINIMUM
              );
            }
          ).length,

          actionRequired: scoredRatios.filter(
            (score: number): boolean => {
              return (
                score >= ACTION_REQUIRED_MINIMUM &&
                score < PENDING_MINIMUM
              );
            }
          ).length,

          highRisk: scoredRatios.filter(
            (score: number): boolean => {
              return score < ACTION_REQUIRED_MINIMUM;
            }
          ).length
        };

        setKpi(freshKpi);

        console.info(
          '[KPI] Calculated metrics summary:',
          freshKpi
        );

        if (failedResults.length === 0) {
          writeCachedKpi(
            freshKpi,
            cacheKey,
            cacheTimeKey
          );
        } else {
          const failedListNames: string =
            failedResults
              .map(
                (
                  result: IListFetchResult
                ): string => {
                  return result.listTitle;
                }
              )
              .join(', ');

          setErrorMessage(
            'Some KPI data could not be loaded. ' +
            `Failed lists: ${failedListNames}.`
          );
        }
      } catch (error: unknown) {
        console.error(
          'Error calculating KPI metrics:',
          error
        );

        if (isMounted) {
          setKpi(EMPTY_KPI_DATA);
          setErrorMessage(
            'An unexpected error occurred while ' +
            'calculating KPI metrics.'
          );
        }
      } finally {
        if (isMounted) {
          setLoading(false);
        }
      }
    };

    fetchMetrics().catch(
      (error: unknown): void => {
        console.error(
          'Unhandled KPI metrics error:',
          error
        );

        if (isMounted) {
          setKpi(EMPTY_KPI_DATA);
          setLoading(false);
          setErrorMessage(
            'An unexpected error occurred while ' +
            'loading KPI metrics.'
          );
        }
      }
    );

    return (): void => {
      isMounted = false;
    };
  }, [
    props.context,
    props.tier1ListTitle,
    props.tier2ListTitle,
    props.tier3ListTitle
  ]);

  if (loading) {
    return (
      <div
        role="status"
        aria-live="polite"
        style={{
          padding: '12px',
          fontFamily: '"Segoe UI", sans-serif',
          color: '#605e5c'
        }}
      >
        Loading KPI metrics...
      </div>
    );
  }

  return (
    <section
      aria-label="ESG questionnaire KPI metrics"
      style={wrapperStyle}
    >
      {errorMessage && (
        <div
          role="alert"
          aria-live="assertive"
          style={warningStyle}
        >
          {errorMessage}
        </div>
      )}

      <div style={containerStyle}>
        {KPI_CARDS.map(
          (
            card: IKpiCardConfiguration
          ): React.ReactElement => (
            <article
              key={card.key}
              style={{
                ...cardStyle,
                backgroundColor:
                  card.cardBackgroundColor,
                border:
                  `1px solid ${card.borderColor}`
              }}
            >
              <div
                aria-hidden="true"
                style={{
                  width: '54px',
                  height: '54px',
                  flex: '0 0 54px',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  color: card.color,
                  backgroundColor:
                    card.iconBackgroundColor,
                  borderRadius: '50%'
                }}
              >
                <Icon
                  iconName={card.iconName}
                  styles={{
                    root: {
                      fontSize: '27px',
                      lineHeight: '27px'
                    }
                  }}
                />
              </div>

              <div
                style={{
                  minWidth: '36px',
                  fontSize: '30px',
                  lineHeight: 1,
                  fontWeight: 600,
                  color: '#151515',
                  textAlign: 'center'
                }}
              >
                {kpi[card.key]}
              </div>

              <div style={cardTextStyle}>
                <div style={titleStyle}>
                  {card.title}
                </div>
                <div style={descriptionStyle}>
                  {card.description}
                </div>
              </div>
            </article>
          )
        )}
      </div>
    </section>
  );
}

export default KpiMetricCards;