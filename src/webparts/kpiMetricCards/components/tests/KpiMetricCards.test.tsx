jest.mock("../../../../pnpConfig", () => ({
  getSP: jest.fn(),
}));

import * as React from "react";
import { render, screen, waitFor, within } from "@testing-library/react";
import "@testing-library/jest-dom";
import { initializeIcons } from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { getSP } from "../../../../pnpConfig";
import { KpiMetricCards } from "../KpiMetricCards";

interface IMockResponseItem {
  Id: number;
  OverallQuestionsPercentage?: number | undefined;
}

interface IMockQuery {
  list: IMockList;
  selectMock: jest.Mock;
  topMock: jest.Mock;
  executeQueryMock: jest.Mock;
}

interface IMockList {
  items: {
    select: jest.Mock;
  };
}

interface IMockSp {
  web: {
    lists: {
      getByTitle: jest.Mock;
    };
  };
}

interface IKpiCacheData {
  total: number;
  pending: number;
  approved: number;
  actionRequired: number;
  highRisk: number;
}

const mockGetByTitle = jest.fn();

const mockedGetSP = getSP as jest.MockedFunction<typeof getSP>;

const mockContext = {} as WebPartContext;

const CACHE_KEY = "SSQ_KPI_METRICS_DATA_V3";
const CACHE_TIME_KEY = "SSQ_KPI_METRICS_TIMESTAMP_V3";

const CACHE_DURATION_MS = 10 * 60 * 1000;

const LIST_TITLES: readonly string[] = [
  "CASSTECH_SSQ",
  "Tier 2 ESG Procurement Questionnaire",
  "Supplier Sustainability Questionnaires Tier 3",
];

/**
 * Creates a successful mock for the PnPjs query chain:
 *
 * sp.web.lists
 *   .getByTitle(listTitle)
 *   .items
 *   .select('Id', 'OverallQuestionsPercentage')
 *   .top(5000)();
 */
const createListMock = (data: IMockResponseItem[]): IMockQuery => {
  const executeQueryMock = jest.fn().mockResolvedValue(data);

  const topMock = jest.fn().mockReturnValue(executeQueryMock);

  const selectMock = jest.fn().mockReturnValue({
    top: topMock,
  });

  return {
    list: {
      items: {
        select: selectMock,
      },
    },
    selectMock,
    topMock,
    executeQueryMock,
  };
};

/**
 * Creates an unresolved SharePoint request.
 * This keeps the component in its loading state.
 */
const createPendingListMock = (): IMockQuery => {
  const pendingPromise = new Promise<IMockResponseItem[]>((): void => {
    // Intentionally unresolved.
  });

  const executeQueryMock = jest.fn().mockReturnValue(pendingPromise);

  const topMock = jest.fn().mockReturnValue(executeQueryMock);

  const selectMock = jest.fn().mockReturnValue({
    top: topMock,
  });

  return {
    list: {
      items: {
        select: selectMock,
      },
    },
    selectMock,
    topMock,
    executeQueryMock,
  };
};

/**
 * Creates a failed SharePoint request.
 */
const createFailedListMock = (
  message: string = "SharePoint request failed",
): IMockQuery => {
  const executeQueryMock = jest.fn().mockRejectedValue(new Error(message));

  const topMock = jest.fn().mockReturnValue(executeQueryMock);

  const selectMock = jest.fn().mockReturnValue({
    top: topMock,
  });

  return {
    list: {
      items: {
        select: selectMock,
      },
    },
    selectMock,
    topMock,
    executeQueryMock,
  };
};

/**
 * Finds a KPI card by its title.
 */
const getMetricCard = (title: string): HTMLElement => {
  const titleElement = screen.getByText(title);
  const cardElement = titleElement.closest("article");

  if (!cardElement) {
    throw new Error(`Unable to find the KPI article for "${title}".`);
  }

  return cardElement;
};

/**
 * Verifies the numeric value displayed inside a KPI card.
 */
const expectMetricValue = (title: string, expectedValue: number): void => {
  const cardElement = getMetricCard(title);

  expect(
    within(cardElement).getByText(expectedValue.toString()),
  ).toBeInTheDocument();
};

/**
 * Adds valid KPI data to sessionStorage.
 */
const setValidCache = (
  data: IKpiCacheData,
  timestamp: number = Date.now(),
): void => {
  sessionStorage.setItem(CACHE_KEY, JSON.stringify(data));

  sessionStorage.setItem(CACHE_TIME_KEY, timestamp.toString());
};

describe("KpiMetricCards Component", () => {
  beforeEach((): void => {
    initializeIcons(undefined, { disableWarnings: true });
    jest.clearAllMocks();
    sessionStorage.clear();

    const mockSp: IMockSp = {
      web: {
        lists: {
          getByTitle: mockGetByTitle,
        },
      },
    };

    mockedGetSP.mockReturnValue(mockSp as unknown as ReturnType<typeof getSP>);
  });

  afterEach((): void => {
    jest.restoreAllMocks();
  });

  test("renders the loading indicator initially", (): void => {
    const pendingListMock = createPendingListMock();

    mockGetByTitle.mockReturnValue(pendingListMock.list);

    render(<KpiMetricCards context={mockContext} />);

    expect(screen.getByText(/Loading KPI Metrics/i)).toBeInTheDocument();
  });

  test("fetches all three SharePoint lists and calculates the KPI metrics correctly", async (): Promise<void> => {
    const tier1Data: IMockResponseItem[] = [
      {
        Id: 1,
        OverallQuestionsPercentage: 0.85,
      },
      {
        Id: 2,
        OverallQuestionsPercentage: 0.5,
      },
    ];

    const tier2Data: IMockResponseItem[] = [
      {
        Id: 3,
        OverallQuestionsPercentage: 0.3,
      },
      {
        Id: 4,
        OverallQuestionsPercentage: 0.15,
      },
    ];

    const tier3Data: IMockResponseItem[] = [
      {
        Id: 5,
        OverallQuestionsPercentage: 0.75,
      },
    ];

    const tier1Mock = createListMock(tier1Data);
    const tier2Mock = createListMock(tier2Data);
    const tier3Mock = createListMock(tier3Data);

    mockGetByTitle
      .mockReturnValueOnce(tier1Mock.list)
      .mockReturnValueOnce(tier2Mock.list)
      .mockReturnValueOnce(tier3Mock.list);

    render(<KpiMetricCards context={mockContext} />);

    await waitFor((): void => {
      expect(screen.getByText("Total Submissions")).toBeInTheDocument();
    });

    expectMetricValue("Total Submissions", 5);
    expectMetricValue("Pending ESG Review", 1);
    expectMetricValue("Approved by ESG", 2);
    expectMetricValue("Requires Procurement Action", 1);
    expectMetricValue("High Risk Suppliers", 1);

    expect(mockedGetSP).toHaveBeenCalledWith(mockContext);

    expect(mockGetByTitle).toHaveBeenCalledTimes(3);

    expect(mockGetByTitle).toHaveBeenNthCalledWith(1, LIST_TITLES[0]);

    expect(mockGetByTitle).toHaveBeenNthCalledWith(2, LIST_TITLES[1]);

    expect(mockGetByTitle).toHaveBeenNthCalledWith(3, LIST_TITLES[2]);
  });

  test("uses the expected select fields and top limit for every SharePoint list", async (): Promise<void> => {
    const tier1Mock = createListMock([]);
    const tier2Mock = createListMock([]);
    const tier3Mock = createListMock([]);

    mockGetByTitle
      .mockReturnValueOnce(tier1Mock.list)
      .mockReturnValueOnce(tier2Mock.list)
      .mockReturnValueOnce(tier3Mock.list);

    render(<KpiMetricCards context={mockContext} />);

    await waitFor((): void => {
      expect(screen.getByText("Total Submissions")).toBeInTheDocument();
    });

    const queryMocks: IMockQuery[] = [tier1Mock, tier2Mock, tier3Mock];

    queryMocks.forEach((queryMock: IMockQuery): void => {
      expect(queryMock.selectMock).toHaveBeenCalledTimes(1);

      expect(queryMock.selectMock).toHaveBeenCalledWith(
        "Id",
        "OverallQuestionsPercentage",
      );

      expect(queryMock.topMock).toHaveBeenCalledTimes(1);

      expect(queryMock.topMock).toHaveBeenCalledWith(5000);

      expect(queryMock.executeQueryMock).toHaveBeenCalledTimes(1);
    });
  });

  test("loads valid KPI data from sessionStorage without calling SharePoint", async (): Promise<void> => {
    const cachedKpi: IKpiCacheData = {
      total: 10,
      pending: 3,
      approved: 5,
      actionRequired: 2,
      highRisk: 1,
    };

    setValidCache(cachedKpi);

    render(<KpiMetricCards context={mockContext} />);

    await waitFor((): void => {
      expectMetricValue("Total Submissions", 10);
    });

    expectMetricValue("Pending ESG Review", 3);

    expectMetricValue("Approved by ESG", 5);

    expectMetricValue("Requires Procurement Action", 2);

    expectMetricValue("High Risk Suppliers", 1);

    expect(mockGetByTitle).not.toHaveBeenCalled();
  });

  test("fetches fresh data when the sessionStorage cache has expired", async (): Promise<void> => {
    const expiredTimestamp = Date.now() - CACHE_DURATION_MS - 1000;

    const expiredCachedKpi: IKpiCacheData = {
      total: 99,
      pending: 99,
      approved: 99,
      actionRequired: 99,
      highRisk: 99,
    };

    setValidCache(expiredCachedKpi, expiredTimestamp);

    const tier1Mock = createListMock([
      {
        Id: 1,
        OverallQuestionsPercentage: 0.9,
      },
    ]);

    const tier2Mock = createListMock([]);
    const tier3Mock = createListMock([]);

    mockGetByTitle
      .mockReturnValueOnce(tier1Mock.list)
      .mockReturnValueOnce(tier2Mock.list)
      .mockReturnValueOnce(tier3Mock.list);

    render(<KpiMetricCards context={mockContext} />);

    await waitFor((): void => {
      expectMetricValue("Total Submissions", 1);
    });

    expectMetricValue("Approved by ESG", 1);

    expectMetricValue("Pending ESG Review", 0);

    expectMetricValue("Requires Procurement Action", 0);

    expectMetricValue("High Risk Suppliers", 0);

    expect(mockGetByTitle).toHaveBeenCalledTimes(3);

    expect(screen.queryByText("99")).not.toBeInTheDocument();
  });

  test("fetches fresh data when the cache timestamp is missing", async (): Promise<void> => {
    const cachedKpi: IKpiCacheData = {
      total: 50,
      pending: 10,
      approved: 20,
      actionRequired: 15,
      highRisk: 5,
    };

    sessionStorage.setItem(CACHE_KEY, JSON.stringify(cachedKpi));

    const tier1Mock = createListMock([
      {
        Id: 1,
        OverallQuestionsPercentage: 0.8,
      },
    ]);

    const tier2Mock = createListMock([]);
    const tier3Mock = createListMock([]);

    mockGetByTitle
      .mockReturnValueOnce(tier1Mock.list)
      .mockReturnValueOnce(tier2Mock.list)
      .mockReturnValueOnce(tier3Mock.list);

    render(<KpiMetricCards context={mockContext} />);

    await waitFor((): void => {
      expectMetricValue("Total Submissions", 1);
    });

    expect(mockGetByTitle).toHaveBeenCalledTimes(3);
  });

  test("fetches fresh data when the cached KPI JSON is invalid", async (): Promise<void> => {
    const consoleErrorSpy = jest
      .spyOn(console, "error")
      .mockImplementation((): void => {
        // Prevent expected cache parsing errors in Jest output.
      });

    sessionStorage.setItem(CACHE_KEY, "{invalid-json");

    sessionStorage.setItem(CACHE_TIME_KEY, Date.now().toString());

    const tier1Mock = createListMock([
      {
        Id: 1,
        OverallQuestionsPercentage: 0.75,
      },
    ]);

    const tier2Mock = createListMock([]);
    const tier3Mock = createListMock([]);

    mockGetByTitle
      .mockReturnValueOnce(tier1Mock.list)
      .mockReturnValueOnce(tier2Mock.list)
      .mockReturnValueOnce(tier3Mock.list);

    render(<KpiMetricCards context={mockContext} />);

    await waitFor((): void => {
      expectMetricValue("Total Submissions", 1);
    });

    expectMetricValue("Approved by ESG", 1);

    expect(mockGetByTitle).toHaveBeenCalledTimes(3);

    consoleErrorSpy.mockRestore();
  });

  test("handles missing and null OverallQuestionsPercentage values safely", async (): Promise<void> => {
    const tier1Mock = createListMock([
      {
        Id: 1,
        OverallQuestionsPercentage: undefined,
      },
      {
        Id: 2,
        OverallQuestionsPercentage: undefined,
      },
    ]);

    const tier2Mock = createListMock([
      {
        Id: 3,
        OverallQuestionsPercentage: 0.8,
      },
    ]);

    const tier3Mock = createListMock([]);

    mockGetByTitle
      .mockReturnValueOnce(tier1Mock.list)
      .mockReturnValueOnce(tier2Mock.list)
      .mockReturnValueOnce(tier3Mock.list);

    render(<KpiMetricCards context={mockContext} />);

    await waitFor((): void => {
      expectMetricValue("Total Submissions", 3);
    });

    expectMetricValue("Approved by ESG", 1);
  });

  test("calculates KPI values correctly at threshold boundaries", async (): Promise<void> => {
    const tier1Mock = createListMock([
      {
        Id: 1,
        OverallQuestionsPercentage: 0.7,
      },
      {
        Id: 2,
        OverallQuestionsPercentage: 0.69,
      },
      {
        Id: 3,
        OverallQuestionsPercentage: 0.4,
      },
      {
        Id: 4,
        OverallQuestionsPercentage: 0.39,
      },
      {
        Id: 5,
        OverallQuestionsPercentage: 0.2,
      },
      {
        Id: 6,
        OverallQuestionsPercentage: 0.19,
      },
    ]);

    const tier2Mock = createListMock([]);
    const tier3Mock = createListMock([]);

    mockGetByTitle
      .mockReturnValueOnce(tier1Mock.list)
      .mockReturnValueOnce(tier2Mock.list)
      .mockReturnValueOnce(tier3Mock.list);

    render(<KpiMetricCards context={mockContext} />);

    await waitFor((): void => {
      expectMetricValue("Total Submissions", 6);
    });

    /*
     * Approved:
     * 0.70
     */
    expectMetricValue("Approved by ESG", 1);

    /*
     * Pending:
     * 0.69, 0.40
     */
    expectMetricValue("Pending ESG Review", 2);

    /*
     * Procurement action:
     * 0.39, 0.20
     */
    expectMetricValue("Requires Procurement Action", 2);

    /*
     * High risk:
     * 0.19
     */
    expectMetricValue("High Risk Suppliers", 1);
  });

  test("continues loading available data when one SharePoint list request fails", async (): Promise<void> => {
    const consoleErrorSpy = jest
      .spyOn(console, "error")
      .mockImplementation((): void => {
        // Prevent expected SharePoint error output.
      });

    const tier1Mock = createListMock([
      {
        Id: 1,
        OverallQuestionsPercentage: 0.85,
      },
      {
        Id: 2,
        OverallQuestionsPercentage: 0.5,
      },
    ]);

    const tier2Mock = createFailedListMock("Tier 2 list request failed");

    const tier3Mock = createListMock([
      {
        Id: 3,
        OverallQuestionsPercentage: 0.15,
      },
    ]);

    mockGetByTitle
      .mockReturnValueOnce(tier1Mock.list)
      .mockReturnValueOnce(tier2Mock.list)
      .mockReturnValueOnce(tier3Mock.list);

    render(<KpiMetricCards context={mockContext} />);

    await waitFor((): void => {
      expectMetricValue("Total Submissions", 3);
    });

    expectMetricValue("Approved by ESG", 1);

    expectMetricValue("Pending ESG Review", 1);

    expectMetricValue("Requires Procurement Action", 0);

    expectMetricValue("High Risk Suppliers", 1);

    expect(mockGetByTitle).toHaveBeenCalledTimes(3);

    expect(consoleErrorSpy).toHaveBeenCalled();

    consoleErrorSpy.mockRestore();
  });

  test("renders zero metrics when all SharePoint lists return empty data", async (): Promise<void> => {
    const tier1Mock = createListMock([]);
    const tier2Mock = createListMock([]);
    const tier3Mock = createListMock([]);

    mockGetByTitle
      .mockReturnValueOnce(tier1Mock.list)
      .mockReturnValueOnce(tier2Mock.list)
      .mockReturnValueOnce(tier3Mock.list);

    render(<KpiMetricCards context={mockContext} />);

    await waitFor((): void => {
      expect(
        screen.queryByText(/Loading KPI Metrics/i),
      ).not.toBeInTheDocument();
    });

    expectMetricValue("Total Submissions", 0);

    expectMetricValue("Pending ESG Review", 0);

    expectMetricValue("Approved by ESG", 0);

    expectMetricValue("Requires Procurement Action", 0);

    expectMetricValue("High Risk Suppliers", 0);
  });

  test("saves newly calculated metrics into sessionStorage", async (): Promise<void> => {
    const tier1Mock = createListMock([
      {
        Id: 1,
        OverallQuestionsPercentage: 0.8,
      },
      {
        Id: 2,
        OverallQuestionsPercentage: 0.5,
      },
      {
        Id: 3,
        OverallQuestionsPercentage: 0.15,
      },
    ]);

    const tier2Mock = createListMock([]);
    const tier3Mock = createListMock([]);

    mockGetByTitle
      .mockReturnValueOnce(tier1Mock.list)
      .mockReturnValueOnce(tier2Mock.list)
      .mockReturnValueOnce(tier3Mock.list);

    render(<KpiMetricCards context={mockContext} />);

    await waitFor((): void => {
      expectMetricValue("Total Submissions", 3);
    });

    const cachedDataValue = sessionStorage.getItem(CACHE_KEY);

    const cachedTimestampValue = sessionStorage.getItem(CACHE_TIME_KEY);

    expect(cachedDataValue).not.toBeNull();
    expect(cachedTimestampValue).not.toBeNull();

    const cachedData = JSON.parse(cachedDataValue as string) as IKpiCacheData;

    expect(cachedData).toEqual({
      total: 3,
      pending: 1,
      approved: 1,
      actionRequired: 0,
      highRisk: 1,
    });

    expect(Number(cachedTimestampValue)).not.toBeNaN();

    expect(Number(cachedTimestampValue)).toBeGreaterThan(0);
  });
});