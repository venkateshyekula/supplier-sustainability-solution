jest.mock('../../../../pnpConfig', () => ({
  getSP: jest.fn()
}));

import * as React from 'react';

import {
  render,
  screen,
  waitFor,
  within
} from '@testing-library/react';

import '@testing-library/jest-dom';

import { WebPartContext } from '@microsoft/sp-webpart-base';

import { getSP } from '../../../../pnpConfig';
import { KpiMetricCards } from '../KpiMetricCards';

interface IMockResponseItem {
  Id: number;
  OverallQuestionsPercentage?: number | string;
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

const mockContext = ({} as unknown) as WebPartContext;

const CACHE_KEY_BASE = 'SSQ_KPI_METRICS_DATA_V8';
const CACHE_TIME_KEY_BASE = 'SSQ_KPI_METRICS_TIMESTAMP_V8';
const CACHE_DURATION_MS = 10 * 60 * 1000;

const LIST_TITLES: readonly string[] = [
  'CASSTECH_SSQ',
  'Tier 2 ESG Procurement Questionnaire',
  'Supplier Sustainability Questionnaires Tier 3'
];

const createCacheKeySegment = (value: string): string => {
  return value
    .replace(/[^a-zA-Z0-9]/g, '_')
    .substring(0, 180);
};

const getCacheKeys = (
  listTitles: readonly string[] = LIST_TITLES
): { cacheKey: string; cacheTimeKey: string } => {
  const signature: string = createCacheKeySegment(
    listTitles.join('|')
  );

  return {
    cacheKey: `${CACHE_KEY_BASE}_${signature}`,
    cacheTimeKey: `${CACHE_TIME_KEY_BASE}_${signature}`
  };
};

const createListMock = (
  data: IMockResponseItem[]
): IMockQuery => {
  const executeQueryMock = jest.fn().mockResolvedValue(data);
  const topMock = jest.fn().mockReturnValue(executeQueryMock);
  const selectMock = jest.fn().mockReturnValue({
    top: topMock
  });

  return {
    list: {
      items: {
        select: selectMock
      }
    },
    selectMock,
    topMock,
    executeQueryMock
  };
};

const createPendingListMock = (): IMockQuery => {
  const pendingPromise = new Promise<IMockResponseItem[]>(
    (): void => {
      // Intentionally unresolved for the loading-state test.
    }
  );

  const executeQueryMock = jest.fn().mockReturnValue(pendingPromise);
  const topMock = jest.fn().mockReturnValue(executeQueryMock);
  const selectMock = jest.fn().mockReturnValue({
    top: topMock
  });

  return {
    list: {
      items: {
        select: selectMock
      }
    },
    selectMock,
    topMock,
    executeQueryMock
  };
};

const createFailedListMock = (
  message: string = 'SharePoint request failed'
): IMockQuery => {
  const executeQueryMock = jest
    .fn()
    .mockRejectedValue(new Error(message));
  const topMock = jest.fn().mockReturnValue(executeQueryMock);
  const selectMock = jest.fn().mockReturnValue({
    top: topMock
  });

  return {
    list: {
      items: {
        select: selectMock
      }
    },
    selectMock,
    topMock,
    executeQueryMock
  };
};

const getMetricCard = (title: string): HTMLElement => {
  const titleElement: HTMLElement = screen.getByText(title);
  const cardElement: HTMLElement | null =
    titleElement.closest('article');

  if (!cardElement) {
    throw new Error(
      `Unable to find the KPI article for "${title}".`
    );
  }

  return cardElement;
};

const expectMetricValue = (
  title: string,
  expectedValue: number
): void => {
  const cardElement: HTMLElement = getMetricCard(title);

  expect(
    within(cardElement).getByText(expectedValue.toString())
  ).toBeInTheDocument();
};

const setValidCache = (
  data: IKpiCacheData,
  timestamp: number = Date.now(),
  listTitles: readonly string[] = LIST_TITLES
): void => {
  const cacheKeys = getCacheKeys(listTitles);

  sessionStorage.setItem(
    cacheKeys.cacheKey,
    JSON.stringify(data)
  );
  sessionStorage.setItem(
    cacheKeys.cacheTimeKey,
    timestamp.toString()
  );
};

const renderComponent = (
  listTitles: readonly string[] = LIST_TITLES
): ReturnType<typeof render> => {
  return render(
    <KpiMetricCards
      context={mockContext}
      tier1ListTitle={listTitles[0]}
      tier2ListTitle={listTitles[1]}
      tier3ListTitle={listTitles[2]}
    />
  );
};

describe('KpiMetricCards Component', (): void => {
  beforeEach((): void => {
    jest.clearAllMocks();
    sessionStorage.clear();

    const mockSp: IMockSp = {
      web: {
        lists: {
          getByTitle: mockGetByTitle
        }
      }
    };

    mockedGetSP.mockReturnValue(
      mockSp as unknown as ReturnType<typeof getSP>
    );
  });

  afterEach((): void => {
    jest.restoreAllMocks();
  });

  test('renders the loading indicator initially', (): void => {
    const pendingListMock: IMockQuery =
      createPendingListMock();

    mockGetByTitle.mockReturnValue(pendingListMock.list);

    renderComponent();

    expect(
      screen.getByText(/Loading KPI Metrics/i)
    ).toBeInTheDocument();
  });

  test(
    'fetches all three lists and categorizes direct SharePoint weighted values',
    async (): Promise<void> => {
      const tier1Mock: IMockQuery = createListMock([
        { Id: 1, OverallQuestionsPercentage: 8.5 },
        { Id: 2, OverallQuestionsPercentage: 5 }
      ]);
      const tier2Mock: IMockQuery = createListMock([
        { Id: 3, OverallQuestionsPercentage: 1.5 },
        { Id: 4, OverallQuestionsPercentage: 0.75 }
      ]);
      const tier3Mock: IMockQuery = createListMock([
        { Id: 5, OverallQuestionsPercentage: 1.5 }
      ]);

      mockGetByTitle
        .mockReturnValueOnce(tier1Mock.list)
        .mockReturnValueOnce(tier2Mock.list)
        .mockReturnValueOnce(tier3Mock.list);

      renderComponent();

      await screen.findByText('Total Submissions');

      expectMetricValue('Total Submissions', 5);
      expectMetricValue('Approved by ESG', 1);
      expectMetricValue('Pending ESG Review', 2);
      expectMetricValue('Requires Procurement Action', 1);
      expectMetricValue('High Risk Suppliers', 1);

      expect(mockedGetSP).toHaveBeenCalledWith(mockContext);
      expect(mockGetByTitle).toHaveBeenCalledTimes(3);
      expect(mockGetByTitle).toHaveBeenNthCalledWith(
        1,
        LIST_TITLES[0]
      );
      expect(mockGetByTitle).toHaveBeenNthCalledWith(
        2,
        LIST_TITLES[1]
      );
      expect(mockGetByTitle).toHaveBeenNthCalledWith(
        3,
        LIST_TITLES[2]
      );
    }
  );

  test(
    'uses the expected select fields and top limit',
    async (): Promise<void> => {
      const tier1Mock: IMockQuery = createListMock([]);
      const tier2Mock: IMockQuery = createListMock([]);
      const tier3Mock: IMockQuery = createListMock([]);

      mockGetByTitle
        .mockReturnValueOnce(tier1Mock.list)
        .mockReturnValueOnce(tier2Mock.list)
        .mockReturnValueOnce(tier3Mock.list);

      renderComponent();

      await screen.findByText('Total Submissions');

      const queryMocks: IMockQuery[] = [
        tier1Mock,
        tier2Mock,
        tier3Mock
      ];

      queryMocks.forEach((queryMock: IMockQuery): void => {
        expect(queryMock.selectMock).toHaveBeenCalledWith(
          'Id',
          'OverallQuestionsPercentage'
        );
        expect(queryMock.topMock).toHaveBeenCalledWith(5000);
        expect(queryMock.executeQueryMock).toHaveBeenCalledTimes(1);
      });
    }
  );

  test(
    'loads valid V8 KPI data from configuration-specific sessionStorage keys without SharePoint calls',
    async (): Promise<void> => {
      const cachedKpi: IKpiCacheData = {
        total: 10,
        pending: 3,
        approved: 5,
        actionRequired: 2,
        highRisk: 1
      };

      setValidCache(cachedKpi);
      renderComponent();

      await waitFor((): void => {
        expectMetricValue('Total Submissions', 10);
      });

      expectMetricValue('Pending ESG Review', 3);
      expectMetricValue('Approved by ESG', 5);
      expectMetricValue('Requires Procurement Action', 2);
      expectMetricValue('High Risk Suppliers', 1);
      expect(mockGetByTitle).not.toHaveBeenCalled();
    }
  );

  test(
    'fetches fresh data when the V8 cache has expired',
    async (): Promise<void> => {
      setValidCache(
        {
          total: 99,
          pending: 99,
          approved: 99,
          actionRequired: 99,
          highRisk: 99
        },
        Date.now() - CACHE_DURATION_MS - 1000
      );

      const tier1Mock: IMockQuery = createListMock([
        { Id: 1, OverallQuestionsPercentage: 9 }
      ]);
      const tier2Mock: IMockQuery = createListMock([]);
      const tier3Mock: IMockQuery = createListMock([]);

      mockGetByTitle
        .mockReturnValueOnce(tier1Mock.list)
        .mockReturnValueOnce(tier2Mock.list)
        .mockReturnValueOnce(tier3Mock.list);

      renderComponent();

      await waitFor((): void => {
        expectMetricValue('Total Submissions', 1);
      });

      expectMetricValue('Approved by ESG', 1);
      expect(screen.queryByText('99')).not.toBeInTheDocument();
      expect(mockGetByTitle).toHaveBeenCalledTimes(3);
    }
  );

  test(
    'fetches fresh data when cached JSON is invalid',
    async (): Promise<void> => {
      const consoleWarnSpy = jest
        .spyOn(console, 'warn')
        .mockImplementation((): void => {
          // Expected invalid cache warning.
        });

      const cacheKeys = getCacheKeys();

      sessionStorage.setItem(
        cacheKeys.cacheKey,
        '{invalid-json'
      );
      sessionStorage.setItem(
        cacheKeys.cacheTimeKey,
        Date.now().toString()
      );

      const tier1Mock: IMockQuery = createListMock([
        { Id: 1, OverallQuestionsPercentage: 7.5 }
      ]);
      const tier2Mock: IMockQuery = createListMock([]);
      const tier3Mock: IMockQuery = createListMock([]);

      mockGetByTitle
        .mockReturnValueOnce(tier1Mock.list)
        .mockReturnValueOnce(tier2Mock.list)
        .mockReturnValueOnce(tier3Mock.list);

      renderComponent();

      await waitFor((): void => {
        expectMetricValue('Total Submissions', 1);
      });

      expectMetricValue('Pending ESG Review', 1);
      expect(mockGetByTitle).toHaveBeenCalledTimes(3);
      expect(consoleWarnSpy).toHaveBeenCalled();
    }
  );

  test(
    'keeps invalid scores in total but excludes them from categories',
    async (): Promise<void> => {
      const tier1Mock: IMockQuery = createListMock([
        { Id: 1, OverallQuestionsPercentage: undefined },
        { Id: 2, OverallQuestionsPercentage: 'invalid' }
      ]);
      const tier2Mock: IMockQuery = createListMock([
        { Id: 3, OverallQuestionsPercentage: 4 }
      ]);
      const tier3Mock: IMockQuery = createListMock([]);

      mockGetByTitle
        .mockReturnValueOnce(tier1Mock.list)
        .mockReturnValueOnce(tier2Mock.list)
        .mockReturnValueOnce(tier3Mock.list);

      renderComponent();

      await waitFor((): void => {
        expectMetricValue('Total Submissions', 3);
      });

      expectMetricValue('Approved by ESG', 1);
      expectMetricValue('Pending ESG Review', 0);
      expectMetricValue('Requires Procurement Action', 0);
      expectMetricValue('High Risk Suppliers', 0);
    }
  );

  test(
    'uses tier-relative boundaries for all three maximum scores',
    async (): Promise<void> => {
      const tier1Mock: IMockQuery = createListMock([
        { Id: 1, OverallQuestionsPercentage: 8 },
        { Id: 2, OverallQuestionsPercentage: 5 },
        { Id: 3, OverallQuestionsPercentage: 2 },
        { Id: 4, OverallQuestionsPercentage: 1.99 }
      ]);
      const tier2Mock: IMockQuery = createListMock([
        { Id: 5, OverallQuestionsPercentage: 4 },
        { Id: 6, OverallQuestionsPercentage: 2.5 },
        { Id: 7, OverallQuestionsPercentage: 1 },
        { Id: 8, OverallQuestionsPercentage: 0.99 }
      ]);
      const tier3Mock: IMockQuery = createListMock([
        { Id: 9, OverallQuestionsPercentage: 1.6 },
        { Id: 10, OverallQuestionsPercentage: 1 },
        { Id: 11, OverallQuestionsPercentage: 0.4 },
        { Id: 12, OverallQuestionsPercentage: 0.39 }
      ]);

      mockGetByTitle
        .mockReturnValueOnce(tier1Mock.list)
        .mockReturnValueOnce(tier2Mock.list)
        .mockReturnValueOnce(tier3Mock.list);

      renderComponent();

      await waitFor((): void => {
        expectMetricValue('Total Submissions', 12);
      });

      expectMetricValue('Approved by ESG', 3);
      expectMetricValue('Pending ESG Review', 3);
      expectMetricValue('Requires Procurement Action', 3);
      expectMetricValue('High Risk Suppliers', 3);
    }
  );

  test(
    'continues with available data when one list fails',
    async (): Promise<void> => {
      const consoleErrorSpy = jest
        .spyOn(console, 'error')
        .mockImplementation((): void => {
          // Expected failed-list logging.
        });

      const tier1Mock: IMockQuery = createListMock([
        { Id: 1, OverallQuestionsPercentage: 8.5 },
        { Id: 2, OverallQuestionsPercentage: 5 }
      ]);
      const tier2Mock: IMockQuery = createFailedListMock(
        'Tier 2 list request failed'
      );
      const tier3Mock: IMockQuery = createListMock([
        { Id: 3, OverallQuestionsPercentage: 0.3 }
      ]);

      mockGetByTitle
        .mockReturnValueOnce(tier1Mock.list)
        .mockReturnValueOnce(tier2Mock.list)
        .mockReturnValueOnce(tier3Mock.list);

      renderComponent();

      await waitFor((): void => {
        expectMetricValue('Total Submissions', 3);
      });

      expectMetricValue('Approved by ESG', 1);
      expectMetricValue('Pending ESG Review', 1);
      expectMetricValue('High Risk Suppliers', 1);
      expect(
        screen.getByText(/Failed lists:/i)
      ).toBeInTheDocument();
      expect(consoleErrorSpy).toHaveBeenCalled();
    }
  );

  test(
    'renders zero metrics when all lists are empty',
    async (): Promise<void> => {
      const tier1Mock: IMockQuery = createListMock([]);
      const tier2Mock: IMockQuery = createListMock([]);
      const tier3Mock: IMockQuery = createListMock([]);

      mockGetByTitle
        .mockReturnValueOnce(tier1Mock.list)
        .mockReturnValueOnce(tier2Mock.list)
        .mockReturnValueOnce(tier3Mock.list);

      renderComponent();

      await waitFor((): void => {
        expect(
          screen.queryByText(/Loading KPI Metrics/i)
        ).not.toBeInTheDocument();
      });

      expectMetricValue('Total Submissions', 0);
      expectMetricValue('Pending ESG Review', 0);
      expectMetricValue('Approved by ESG', 0);
      expectMetricValue('Requires Procurement Action', 0);
      expectMetricValue('High Risk Suppliers', 0);
    }
  );

  test(
    'uses configurable list titles supplied through component props',
    async (): Promise<void> => {
      const customTitles: readonly string[] = [
        'Custom Tier 1 List',
        'Custom Tier 2 List',
        'Custom Tier 3 List'
      ];

      const tier1Mock: IMockQuery = createListMock([]);
      const tier2Mock: IMockQuery = createListMock([]);
      const tier3Mock: IMockQuery = createListMock([]);

      mockGetByTitle
        .mockReturnValueOnce(tier1Mock.list)
        .mockReturnValueOnce(tier2Mock.list)
        .mockReturnValueOnce(tier3Mock.list);

      renderComponent(customTitles);

      await screen.findByText('Total Submissions');

      expect(mockGetByTitle).toHaveBeenNthCalledWith(
        1,
        customTitles[0]
      );
      expect(mockGetByTitle).toHaveBeenNthCalledWith(
        2,
        customTitles[1]
      );
      expect(mockGetByTitle).toHaveBeenNthCalledWith(
        3,
        customTitles[2]
      );
    }
  );

  test(
    'saves newly calculated metrics using configuration-specific V8 cache keys',
    async (): Promise<void> => {
      const tier1Mock: IMockQuery = createListMock([
        { Id: 1, OverallQuestionsPercentage: 8 },
        { Id: 2, OverallQuestionsPercentage: 5 },
        { Id: 3, OverallQuestionsPercentage: 1 }
      ]);
      const tier2Mock: IMockQuery = createListMock([]);
      const tier3Mock: IMockQuery = createListMock([]);

      mockGetByTitle
        .mockReturnValueOnce(tier1Mock.list)
        .mockReturnValueOnce(tier2Mock.list)
        .mockReturnValueOnce(tier3Mock.list);

      renderComponent();

      await waitFor((): void => {
        expectMetricValue('Total Submissions', 3);
      });

      const cacheKeys = getCacheKeys();

      const cachedDataValue: string | null =
        sessionStorage.getItem(cacheKeys.cacheKey);
      const cachedTimestampValue: string | null =
        sessionStorage.getItem(cacheKeys.cacheTimeKey);

      expect(cachedDataValue).not.toBeNull();
      expect(cachedTimestampValue).not.toBeNull();

      const cachedData: IKpiCacheData = JSON.parse(
        cachedDataValue as string
      ) as IKpiCacheData;

      expect(cachedData).toEqual({
        total: 3,
        pending: 1,
        approved: 1,
        actionRequired: 0,
        highRisk: 1
      });
      expect(Number(cachedTimestampValue)).toBeGreaterThan(0);
    }
  );
});
