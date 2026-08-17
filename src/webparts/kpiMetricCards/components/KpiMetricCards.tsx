import * as React from "react";
import { useEffect, useState } from "react";
import { Icon } from "@fluentui/react";

import { getSP } from "../../../pnpConfig";
import { IKpiMetricCardsProps } from "./IKpiMetricCardsProps";

interface IKpiData {
  total: number;
  pending: number;
  approved: number;
  actionRequired: number;
  highRisk: number;
}

interface IResponseItem {
  Id?: number;
  OverallQuestionsPercentage?: number;
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
}

const EMPTY_KPI_DATA: IKpiData = {
  total: 0,
  pending: 0,
  approved: 0,
  actionRequired: 0,
  highRisk: 0,
};

const LIST_CONFIGURATIONS: readonly IListConfiguration[] = [
  {
    title: "CASSTECH_SSQ",
    percentageInternalName: "OverallQuestionsPercentage",
    maximumWeightedScore: 10,
  },
  {
    title: "Tier 2 ESG Procurement Questionnaire",
    percentageInternalName: "OverallQuestionsPercentage",
    maximumWeightedScore: 5,
  },
  {
    title: "Supplier Sustainability Questionnaires Tier 3",
    percentageInternalName: "OverallQuestionsPercentage",
    maximumWeightedScore: 2,
  },
];

// Excel Alignment: Compliant (>80%), Need to Inform (50-80%), Need to Engage (20-50%), High Risk (<20%)
const APPROVED_MINIMUM = 0.8;
const PENDING_MINIMUM = 0.5;
const ACTION_REQUIRED_MINIMUM = 0.2;

const KPI_CARDS: readonly IKpiCardConfiguration[] = [
  {
    key: "total",
    title: "Total Submissions",
    description: "All tiers",
    color: "#201f1e",
    iconName: "BulletedList",
    iconBackgroundColor: "#edebe9",
  },
  {
    key: "pending",
    title: "Pending ESG Review",
    description: "50% to below 80%",
    color: "#8764b8",
    iconName: "Clock",
    iconBackgroundColor: "#f2ecf9",
  },
  {
    key: "approved",
    title: "Approved by ESG",
    description: "80% or higher",
    color: "#107c41",
    iconName: "CompletedSolid",
    iconBackgroundColor: "#dff6dd",
  },
  {
    key: "actionRequired",
    title: "Requires Procurement Action",
    description: "20% to below 50%",
    color: "#d13438",
    iconName: "Warning",
    iconBackgroundColor: "#fde7e9",
  },
  {
    key: "highRisk",
    title: "High Risk Suppliers",
    description: "Below 20%",
    color: "#a80000",
    iconName: "ShieldAlert",
    iconBackgroundColor: "#fdd8db",
  },
];

// Incremented cache key to V6 to automatically invalidate old stored session data
const CACHE_KEY =
  'SSQ_KPI_METRICS_DATA_V7';

const CACHE_TIME_KEY =
  'SSQ_KPI_METRICS_TIMESTAMP_V7';
const TTL_MS = 10 * 60 * 1000;

const getErrorMessage = (error: unknown): string => {
  if (error instanceof Error) {
    return error.message;
  }

  return String(error);
};

const normalizeWeightedScore = (
  weightedScore: unknown,
  maximumWeightedScore: number,
): number | undefined => {
  if (
    weightedScore === null ||
    weightedScore === undefined ||
    weightedScore === ""
  ) {
    return undefined;
  }

  if (!isFinite(maximumWeightedScore) || maximumWeightedScore <= 0) {
    return undefined;
  }

  let normalizedValue: unknown = weightedScore;

  if (typeof normalizedValue === "string") {
    normalizedValue = normalizedValue.replace("%", "").replace(",", ".").trim();
  }

  const numericScore: number =
    typeof normalizedValue === "number"
      ? normalizedValue
      : Number(normalizedValue);

  if (isNaN(numericScore) || !isFinite(numericScore) || numericScore < 0) {
    return undefined;
  }

  const completionRatio: number = numericScore / maximumWeightedScore;

  if (!isFinite(completionRatio) || completionRatio < 0) {
    return undefined;
  }

  /*
   * Clamp the result to 1 in case a SharePoint calculation
   * temporarily returns more than the configured maximum.
   */
  return Math.min(completionRatio, 1);
};

const isKpiData = (value: unknown): value is IKpiData => {
  if (typeof value !== "object" || value === undefined) {
    return false;
  }

  const candidate = value as Partial<IKpiData>;

  return (
    typeof candidate.total === "number" &&
    typeof candidate.pending === "number" &&
    typeof candidate.approved === "number" &&
    typeof candidate.actionRequired === "number" &&
    typeof candidate.highRisk === "number"
  );
};

const readCachedKpi = (): IKpiData | undefined => {
  try {
    const cachedData = window.sessionStorage.getItem(CACHE_KEY);
    const cachedTimestamp = window.sessionStorage.getItem(CACHE_TIME_KEY);

    if (!cachedData || !cachedTimestamp) {
      return undefined;
    }

    const timestamp = Number(cachedTimestamp);

    if (!isFinite(timestamp)) {
      window.sessionStorage.removeItem(CACHE_KEY);
      window.sessionStorage.removeItem(CACHE_TIME_KEY);
      return undefined;
    }

    const cacheAge = Date.now() - timestamp;

    if (cacheAge < 0 || cacheAge >= TTL_MS) {
      window.sessionStorage.removeItem(CACHE_KEY);
      window.sessionStorage.removeItem(CACHE_TIME_KEY);
      return undefined;
    }

    const parsedData: unknown = JSON.parse(cachedData);

    if (!isKpiData(parsedData)) {
      window.sessionStorage.removeItem(CACHE_KEY);
      window.sessionStorage.removeItem(CACHE_TIME_KEY);
      return undefined;
    }

    return parsedData;
  } catch (error: unknown) {
    console.warn("Unable to read KPI metrics cache:", error);
    window.sessionStorage.removeItem(CACHE_KEY);
    window.sessionStorage.removeItem(CACHE_TIME_KEY);
    return undefined;
  }
};

const writeCachedKpi = (data: IKpiData): void => {
  try {
    window.sessionStorage.setItem(CACHE_KEY, JSON.stringify(data));
    window.sessionStorage.setItem(CACHE_TIME_KEY, Date.now().toString());
  } catch (error: unknown) {
    console.warn("Unable to save KPI metrics to sessionStorage:", error);
  }
};

export function KpiMetricCards(
  props: IKpiMetricCardsProps,
): React.ReactElement {
  const [kpi, setKpi] = useState<IKpiData>(EMPTY_KPI_DATA);
  const [loading, setLoading] = useState<boolean>(true);
  const [errorMessage, setErrorMessage] = useState<string>("");

  useEffect((): (() => void) => {
    let isMounted = true;

    const fetchListItems = async (
      configuration: IListConfiguration,
    ): Promise<IListFetchResult> => {
      try {
        const sp = getSP(props.context);

        const items = await sp.web.lists
          .getByTitle(configuration.title)
          .items.select("Id", configuration.percentageInternalName)
          .top(5000)();

        console.info(
          `[KPI RAW DATA] ${configuration.title} (first 3 items):`,
          items.slice(0, 3),
        );

        const mappedItems: IResponseItem[] = items.map(
          (item: Record<string, unknown>): IResponseItem => {
            const idValue = item.Id;
            const rawPercentage = item[configuration.percentageInternalName];

            return {
              Id: typeof idValue === "number" ? idValue : undefined,
              OverallQuestionsPercentage: normalizeWeightedScore(
                rawPercentage,
                configuration.maximumWeightedScore,
              ),
            };
          },
        );

        console.info(
          `[KPI] ${configuration.title}: ${mappedItems.length} items received.`,
        );

        console.info(
          `[KPI] ${configuration.title}: parsed values`,
          mappedItems.map(
            (item: IResponseItem): number | string =>
              item.OverallQuestionsPercentage ?? "(empty/invalid)",
          ),
        );

        return {
          listTitle: configuration.title,
          items: mappedItems,
          succeeded: true,
        };
      } catch (error: unknown) {
        const message = getErrorMessage(error);
        console.error(
          `[KPI] Failed to retrieve "${configuration.title}".`,
          error,
        );

        return {
          listTitle: configuration.title,
          items: [],
          succeeded: false,
          errorMessage: message,
        };
      }
    };

    const fetchMetrics = async (): Promise<void> => {
      if (!isMounted) {
        return;
      }

      setLoading(true);
      setErrorMessage("");

      const cachedKpi = readCachedKpi();

      if (cachedKpi) {
        setKpi(cachedKpi);
        setLoading(false);
        console.info("[KPI] Metrics loaded from sessionStorage.");
        return;
      }

      try {
        const results: IListFetchResult[] = await Promise.all(
          LIST_CONFIGURATIONS.map(
            async (
              configuration: IListConfiguration,
            ): Promise<IListFetchResult> => fetchListItems(configuration),
          ),
        );

        if (!isMounted) {
          return;
        }

        const successfulResults = results.filter(
          (result: IListFetchResult): boolean => result.succeeded,
        );

        const failedResults = results.filter(
          (result: IListFetchResult): boolean => !result.succeeded,
        );

        if (successfulResults.length === 0) {
          setKpi(EMPTY_KPI_DATA);
          setErrorMessage(
            "KPI records could not be loaded from any SharePoint list. Open browser console for details.",
          );
          return;
        }

        const allItems: IResponseItem[] = successfulResults.reduce<
          IResponseItem[]
        >(
          (
            accumulatedItems: IResponseItem[],
            currentResult: IListFetchResult,
          ): IResponseItem[] => accumulatedItems.concat(currentResult.items),
          [],
        );

        const scoredPercentages: number[] = allItems
          .map(
            (item: IResponseItem): number | undefined =>
              item.OverallQuestionsPercentage,
          )
          .filter(
            (percentage: number | undefined): percentage is number =>
              percentage !== undefined,
          );

        const freshKpi: IKpiData = {
          total: allItems.length,

          approved: scoredPercentages.filter(
            (percentage: number): boolean => percentage >= APPROVED_MINIMUM,
          ).length,

          pending: scoredPercentages.filter(
            (percentage: number): boolean =>
              percentage >= PENDING_MINIMUM && percentage < APPROVED_MINIMUM,
          ).length,

          actionRequired: scoredPercentages.filter(
            (percentage: number): boolean =>
              percentage >= ACTION_REQUIRED_MINIMUM &&
              percentage < PENDING_MINIMUM,
          ).length,

          highRisk: scoredPercentages.filter(
            (percentage: number): boolean =>
              percentage < ACTION_REQUIRED_MINIMUM,
          ).length,
        };

        setKpi(freshKpi);

        console.info("[KPI] Calculated metrics summary:", freshKpi);

        if (failedResults.length === 0) {
          writeCachedKpi(freshKpi);
        } else {
          const failedListNames = failedResults
            .map((result: IListFetchResult): string => result.listTitle)
            .join(", ");

          setErrorMessage(
            `Some KPI data could not be loaded. Failed lists: ${failedListNames}.`,
          );
        }
      } catch (error: unknown) {
        console.error("Error calculating KPI metrics:", error);
        if (isMounted) {
          setKpi(EMPTY_KPI_DATA);
          setErrorMessage(
            "An unexpected error occurred while calculating KPI metrics.",
          );
        }
      } finally {
        if (isMounted) {
          setLoading(false);
        }
      }
    };

    fetchMetrics().catch((error: unknown): void => {
      console.error("Unhandled KPI metrics error:", error);
      if (isMounted) {
        setKpi(EMPTY_KPI_DATA);
        setLoading(false);
        setErrorMessage(
          "An unexpected error occurred while loading KPI metrics.",
        );
      }
    });

    return (): void => {
      isMounted = false;
    };
  }, [props.context]);

  const wrapperStyle: React.CSSProperties = {
    width: "100%",
    fontFamily: '"Segoe UI", sans-serif',
  };

  const containerStyle: React.CSSProperties = {
    display: "grid",
    gridTemplateColumns: "repeat(5, 1fr)",
    gap: "12px",
    width: "100%",
    boxSizing: "border-box",
  };

  const cardStyle: React.CSSProperties = {
    minHeight: "104px",
    padding: "14px 12px",
    display: "flex",
    alignItems: "center",
    gap: "10px",
    backgroundColor: "#ffffff",
    border: "1px solid #e1dfdd",
    borderRadius: "10px",
    boxShadow: "0 2px 7px rgba(0, 0, 0, 0.12)",
    boxSizing: "border-box",
  };

  const cardTextStyle: React.CSSProperties = {
    minWidth: 0,
    display: "flex",
    flexDirection: "column",
    gap: "2px",
  };

  const titleStyle: React.CSSProperties = {
    fontSize: "13px",
    lineHeight: 1.3,
    fontWeight: 600,
    color: "#201f1e",
  };

  const descriptionStyle: React.CSSProperties = {
    fontSize: "11px",
    lineHeight: 1.3,
    color: "#605e5c",
  };

  const warningStyle: React.CSSProperties = {
    marginBottom: "14px",
    padding: "10px 14px",
    fontSize: "13px",
    lineHeight: 1.45,
    color: "#8a6d1d",
    backgroundColor: "#fff4ce",
    border: "1px solid #fce100",
    borderRadius: "6px",
  };

  if (loading) {
    return (
      <div
        role="status"
        aria-live="polite"
        style={{
          padding: "12px",
          fontFamily: '"Segoe UI", sans-serif',
          color: "#605e5c",
        }}
      >
        Loading KPI metrics...
      </div>
    );
  }

  return (
    <section aria-label="ESG questionnaire KPI metrics" style={wrapperStyle}>
      {errorMessage && (
        <div role="alert" aria-live="assertive" style={warningStyle}>
          {errorMessage}
        </div>
      )}

      <div style={containerStyle}>
        {KPI_CARDS.map(
          (card: IKpiCardConfiguration): React.ReactElement => (
            <article key={card.key} style={cardStyle}>
              <div
                aria-hidden="true"
                style={{
                  width: "42px",
                  height: "42px",
                  flex: "0 0 42px",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  color: card.color,
                  backgroundColor: card.iconBackgroundColor,
                  borderRadius: "50%",
                }}
              >
                <Icon
                  iconName={card.iconName}
                  styles={{
                    root: {
                      fontSize: "20px",
                      lineHeight: "20px",
                    },
                  }}
                />
              </div>

              <div
                style={{
                  minWidth: "42px",
                  fontSize: "28px",
                  lineHeight: 1,
                  fontWeight: 600,
                  color: card.color,
                  textAlign: "center",
                }}
              >
                {kpi[card.key]}
              </div>

              <div style={cardTextStyle}>
                <div style={titleStyle}>{card.title}</div>
                <div style={descriptionStyle}>{card.description}</div>
              </div>
            </article>
          ),
        )}
      </div>
    </section>
  );
}

export default KpiMetricCards;
