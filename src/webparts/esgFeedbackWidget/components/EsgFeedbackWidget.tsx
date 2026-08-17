import * as React from "react";
import { useEffect, useState } from "react";
import {
  Cell as RechartsCell,
  Pie as RechartsPie,
  PieChart as RechartsPieChart,
  ResponsiveContainer as RechartsResponsiveContainer,
} from "recharts";

import { getSP } from "../../../pnpConfig";
import { IEsgFeedbackWidgetProps } from "./IEsgFeedbackWidgetProps";

interface ICategoryScore {
  name: string;
  value: number;
  color: string;
}

interface ITierItem {
  Id?: number;
  Author?: {
    EMail?: string;
  };
  Supplier_x0020_Name?: string;
  OverallQuestionsPercentage?: number | string;
  EnvironmentalQuestionsCount?: number | string;
  EnvQuestionsCount?: number | string;
  SocialQuestionsCount?: number | string;
  GovernanceQuestionsCount?: number | string;
}

interface IEsgCachePayload {
  supplierName: string;
  overallScore: number;
  scores: ICategoryScore[];
}

interface IResponsiveContainerProps {
  width?: string | number;
  height?: string | number;
  children?: React.ReactNode;
}

interface IPieChartProps {
  children?: React.ReactNode;
}

interface IPieProps {
  data: ICategoryScore[];
  cx?: string | number;
  cy?: string | number;
  innerRadius?: number | string;
  outerRadius?: number | string;
  paddingAngle?: number;
  dataKey: keyof ICategoryScore;
  children?: React.ReactNode;
}

interface ICellProps {
  fill?: string;
}

const ResponsiveContainer =
  RechartsResponsiveContainer as unknown as React.ComponentType<IResponsiveContainerProps>;

const PieChart =
  RechartsPieChart as unknown as React.ComponentType<IPieChartProps>;

const Pie =
  RechartsPie as unknown as React.ComponentType<IPieProps>;

const Cell =
  RechartsCell as unknown as React.ComponentType<ICellProps>;

const EMPTY_CATEGORY_COUNTS: ICategoryScore[] = [
  {
    name: "Environmental",
    value: 0,
    color: "#0088fe",
  },
  {
    name: "Social",
    value: 0,
    color: "#00c49f",
  },
  {
    name: "Governance",
    value: 0,
    color: "#ffbb28",
  },
];

const CACHE_KEY = "SSQ_ESG_FEEDBACK_DATA_V5";
const CACHE_TIME_KEY = "SSQ_ESG_FEEDBACK_TIMESTAMP_V5";
const TTL_MS = 10 * 60 * 1000;

const getValidCount = (value: unknown): number => {
  if (value === null || value === undefined || value === "") {
    return 0;
  }

  const num = typeof value === "number" ? value : Number(value);

  if (isNaN(num) || !isFinite(num) || num < 0) {
    return 0;
  }

  return Math.round(num);
};

const normalizePercentage = (value: unknown): number => {
  if (value === null || value === undefined || value === "") {
    return 0;
  }

  let val = value;
  if (typeof val === "string") {
    val = val.replace("%", "").trim();
  }

  const num = typeof val === "number" ? val : Number(val);

  if (isNaN(num) || !isFinite(num) || num < 0) {
    return 0;
  }

  const percentage = num <= 1 ? Math.round(num * 100) : Math.round(num);

  return Math.min(100, percentage);
};

const isCategoryScore = (value: unknown): value is ICategoryScore => {
  if (typeof value !== "object" || value === undefined) {
    return false;
  }

  const candidate = value as Partial<ICategoryScore>;

  return (
    typeof candidate.name === "string" &&
    typeof candidate.value === "number" &&
    typeof candidate.color === "string"
  );
};

const isEsgCachePayload = (value: unknown): value is IEsgCachePayload => {
  if (typeof value !== "object" || value === undefined) {
    return false;
  }

  const candidate = value as Partial<IEsgCachePayload>;

  return (
    typeof candidate.supplierName === "string" &&
    typeof candidate.overallScore === "number" &&
    Array.isArray(candidate.scores) &&
    candidate.scores.every((score: unknown): boolean => isCategoryScore(score))
  );
};

const readCachedPayload = (): IEsgCachePayload | undefined => {
  try {
    const cachedData = window.sessionStorage.getItem(CACHE_KEY);
    const cachedTimestamp = window.sessionStorage.getItem(CACHE_TIME_KEY);

    if (!cachedData || !cachedTimestamp) {
      return undefined;
    }

    const timestamp = Number(cachedTimestamp);

    if (isNaN(timestamp) || !isFinite(timestamp)) {
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

    if (!isEsgCachePayload(parsedData)) {
      window.sessionStorage.removeItem(CACHE_KEY);
      window.sessionStorage.removeItem(CACHE_TIME_KEY);
      return undefined;
    }

    return parsedData;
  } catch (error: unknown) {
    console.warn("Unable to read ESG feedback cache:", error);
    window.sessionStorage.removeItem(CACHE_KEY);
    window.sessionStorage.removeItem(CACHE_TIME_KEY);
    return undefined;
  }
};

const writeCachedPayload = (payload: IEsgCachePayload): void => {
  try {
    window.sessionStorage.setItem(CACHE_KEY, JSON.stringify(payload));
    window.sessionStorage.setItem(CACHE_TIME_KEY, Date.now().toString());
  } catch (error: unknown) {
    console.warn("Unable to save ESG feedback cache:", error);
  }
};

export function EsgFeedbackWidget(
  props: IEsgFeedbackWidgetProps,
): React.ReactElement {
  const [scores, setScores] = useState<ICategoryScore[]>(EMPTY_CATEGORY_COUNTS);
  const [overallScore, setOverallScore] = useState<number>(0);
  const [supplierName, setSupplierName] = useState<string>("");
  const [loading, setLoading] = useState<boolean>(true);
  const [errorMessage, setErrorMessage] = useState<string>("");

  useEffect((): (() => void) => {
    let isMounted = true;

    const fetchLatestScores = async (): Promise<void> => {
      if (!isMounted) {
        return;
      }

      setLoading(true);
      setErrorMessage("");

      const cachedPayload = readCachedPayload();

      if (cachedPayload) {
        setSupplierName(cachedPayload.supplierName);
        setOverallScore(cachedPayload.overallScore);
        setScores(cachedPayload.scores);
        setLoading(false);
        console.info("[ESG Feedback] Data loaded from sessionStorage.");
        return;
      }

      try {
        const sp = getSP(props.context);

        /*
         * CASSTECH_SSQ Internal Columns:
         * EnvironmentalQuestionsCount, SocialQuestionsCount, GovernanceQuestionsCount,
         * ESGinformationCount, OverallQuestionsCount, OverallQuestionsPercentage
         */
        const items: ITierItem[] = await sp.web.lists
          .getByTitle("CASSTECH_SSQ")
          .items.select(
            "Id",
            "Author/EMail",
            "Supplier_x0020_Name",
            "OverallQuestionsPercentage",
            "EnvironmentalQuestionsCount",
            "SocialQuestionsCount",
            "GovernanceQuestionsCount",
          )
          .expand("Author")
          .orderBy("Created", false)
          .top(1)();

        if (!isMounted) {
          return;
        }

        const latestItem: ITierItem | undefined = items[0];

        if (!latestItem) {
          const emptyPayload: IEsgCachePayload = {
            supplierName: "No submissions found",
            overallScore: 0,
            scores: EMPTY_CATEGORY_COUNTS,
          };

          setSupplierName(emptyPayload.supplierName);
          setOverallScore(emptyPayload.overallScore);
          setScores(emptyPayload.scores);
          writeCachedPayload(emptyPayload);
          return;
        }

        const supplier =
          latestItem.Supplier_x0020_Name?.trim() ||
          latestItem.Author?.EMail?.trim() ||
          "Latest supplier";

        const percentage = normalizePercentage(
          latestItem.OverallQuestionsPercentage,
        );

        const envCountValue =
          latestItem.EnvironmentalQuestionsCount ??
          latestItem.EnvQuestionsCount;

        const categoryCounts: ICategoryScore[] = [
          {
            name: "Environmental",
            value: getValidCount(envCountValue),
            color: "#0088fe",
          },
          {
            name: "Social",
            value: getValidCount(latestItem.SocialQuestionsCount),
            color: "#00c49f",
          },
          {
            name: "Governance",
            value: getValidCount(latestItem.GovernanceQuestionsCount),
            color: "#ffbb28",
          },
        ];

        const payload: IEsgCachePayload = {
          supplierName: supplier,
          overallScore: percentage,
          scores: categoryCounts,
        };

        setSupplierName(payload.supplierName);
        setOverallScore(payload.overallScore);
        setScores(payload.scores);
        writeCachedPayload(payload);

        console.info("[ESG Feedback] Latest item retrieved:", latestItem);
        console.info("[ESG Feedback] Calculated payload:", payload);
      } catch (error: unknown) {
        console.error("Error fetching ESG feedback:", error);

        if (isMounted) {
          setSupplierName("Unable to load supplier");
          setOverallScore(0);
          setScores(EMPTY_CATEGORY_COUNTS);
          setErrorMessage(
            "ESG feedback could not be loaded from SharePoint. Open browser console to review details.",
          );
        }
      } finally {
        if (isMounted) {
          setLoading(false);
        }
      }
    };

    fetchLatestScores().catch((error: unknown): void => {
      console.error("Unhandled ESG feedback error:", error);

      if (isMounted) {
        setLoading(false);
        setErrorMessage(
          "An unexpected error occurred while loading ESG feedback.",
        );
      }
    });

    return (): void => {
      isMounted = false;
    };
  }, [props.context]);

  const widgetStyle: React.CSSProperties = {
    width: "100%",
    padding: "16px",
    fontFamily: '"Segoe UI", sans-serif',
    color: "#323130",
    backgroundColor: "#ffffff",
    border: "1px solid #e1dfdd",
    borderRadius: "8px",
    boxSizing: "border-box",
    boxShadow: "0 1.6px 3.6px rgba(0, 0, 0, 0.11)",
  };

  const warningStyle: React.CSSProperties = {
    marginBottom: "12px",
    padding: "8px 10px",
    fontSize: "12px",
    lineHeight: 1.4,
    color: "#8a6d1d",
    backgroundColor: "#fff4ce",
    border: "1px solid #fce100",
    borderRadius: "4px",
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
        Loading ESG feedback...
      </div>
    );
  }

  return (
    <section aria-label="Latest ESG feedback" style={widgetStyle}>
      <h3
        style={{
          margin: "0 0 4px",
          fontSize: "16px",
          color: "#323130",
        }}
      >
        Latest ESG Feedback
      </h3>

      <div
        title={supplierName}
        style={{
          marginBottom: "12px",
          overflowWrap: "anywhere",
          fontSize: "13px",
          fontWeight: 600,
          color: "#605e5c",
        }}
      >
        {supplierName}
      </div>

      {errorMessage && (
        <div role="alert" aria-live="assertive" style={warningStyle}>
          {errorMessage}
        </div>
      )}

      <div
        style={{
          position: "relative",
          height: "180px",
        }}
      >
        <ResponsiveContainer width="100%" height="100%">
          <PieChart>
            <Pie
              data={scores}
              cx="50%"
              cy="50%"
              innerRadius={55}
              outerRadius={75}
              paddingAngle={4}
              dataKey="value"
            >
              {scores.map(
                (score: ICategoryScore, index: number): React.ReactElement => (
                  <Cell key={`${score.name}-${index}`} fill={score.color} />
                ),
              )}
            </Pie>
          </PieChart>
        </ResponsiveContainer>

        <div
          style={{
            position: "absolute",
            top: "50%",
            left: "50%",
            textAlign: "center",
            pointerEvents: "none",
            transform: "translate(-50%, -50%)",
          }}
        >
          <span
            style={{
              display: "block",
              fontSize: "26px",
              lineHeight: 1.1,
              fontWeight: 700,
              color: "#201f1e",
            }}
          >
            {overallScore}%
          </span>

          <div
            style={{
              marginTop: "2px",
              fontSize: "10px",
              color: "#605e5c",
            }}
          >
            Overall Score
          </div>
        </div>
      </div>

      <div
        style={{
          marginTop: "12px",
          paddingTop: "10px",
          borderTop: "1px solid #f3f2f1",
        }}
      >
        {scores.map(
          (score: ICategoryScore): React.ReactElement => (
            <div
              key={score.name}
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                padding: "4px 0",
                fontSize: "13px",
              }}
            >
              <span
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: "8px",
                }}
              >
                <span
                  aria-hidden="true"
                  style={{
                    display: "inline-block",
                    width: "10px",
                    height: "10px",
                    flexShrink: 0,
                    backgroundColor: score.color,
                    borderRadius: "50%",
                  }}
                />
                {score.name}
              </span>

              <span style={{ fontWeight: 600 }}>{score.value} questions</span>
            </div>
          ),
        )}
      </div>
    </section>
  );
}

export default EsgFeedbackWidget;