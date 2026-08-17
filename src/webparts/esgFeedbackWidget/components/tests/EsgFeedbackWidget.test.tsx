jest.mock("recharts", () => {
  const ReactActual =
    jest.requireActual<typeof import("react")>("react");

  const ResponsiveContainer = ({
    children,
  }: {
    children?: import("react").ReactNode;
  }): import("react").ReactElement =>
    ReactActual.createElement(
      "div",
      {
        "data-testid": "responsive-container",
      },
      children,
    );

  const PieChart = ({
    children,
  }: {
    children?: import("react").ReactNode;
  }): import("react").ReactElement =>
    ReactActual.createElement(
      "div",
      {
        "data-testid": "pie-chart",
      },
      children,
    );

  const Pie = ({
    children,
  }: {
    children?: import("react").ReactNode;
  }): import("react").ReactElement =>
    ReactActual.createElement(
      "div",
      {
        "data-testid": "pie",
      },
      children,
    );

  const Cell = (): import("react").ReactElement =>
    ReactActual.createElement("div", {
      "data-testid": "pie-cell",
    });

  return {
    ResponsiveContainer,
    PieChart,
    Pie,
    Cell,
  };
});

jest.mock("../../../../pnpConfig", () => ({
  getSP: jest.fn(),
}));

import * as React from "react";
import {
  render,
  screen,
  waitFor,
} from "@testing-library/react";
import "@testing-library/jest-dom";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import { getSP } from "../../../../pnpConfig";
import { EsgFeedbackWidget } from "../EsgFeedbackWidget";

interface IMockTierItem {
  Id?: number;
  Email?: string;
  Supplier_x0020_Name?: string;
  OverallQuestionsPercentage?: number;
  EnvQuestionsCount?: number;
  SocialQuestionsCount?: number;
  GovernanceQuestionsCount?: number;
  ComplianceCount?: number;
}

interface IMockCategoryScore {
  name: string;
  value: number;
  color: string;
}

interface IMockCachePayload {
  supplierName: string;
  overallScore: number;
  scores: IMockCategoryScore[];
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

const mockGetByTitle = jest.fn();

const mockedGetSP = getSP as jest.MockedFunction<
  typeof getSP
>;

const mockContext = {} as WebPartContext;

const CACHE_KEY = "SSQ_ESG_FEEDBACK_DATA_V2";

const CACHE_TIME_KEY =
  "SSQ_ESG_FEEDBACK_TIMESTAMP_V2";

/**
 * Creates a successful mock for the PnPjs query chain:
 *
 * items
 *   .select(...)
 *   .orderBy("Created", false)
 *   .top(1)()
 */
const createListMock = (
  data: IMockTierItem[],
): IMockList => {
  const executeQueryMock = jest
    .fn()
    .mockResolvedValue(data);

  const topMock = jest
    .fn()
    .mockReturnValue(executeQueryMock);

  const orderByMock = jest
    .fn()
    .mockReturnValue({
      top: topMock,
    });

  const selectMock = jest
    .fn()
    .mockReturnValue({
      orderBy: orderByMock,
    });

  return {
    items: {
      select: selectMock,
    },
  };
};

/**
 * Creates an unresolved SharePoint query so that the component
 * remains in the loading state.
 */
const createPendingListMock = (): IMockList => {
  const pendingPromise = new Promise<
    IMockTierItem[]
  >((): void => {
    // Intentionally unresolved for the loading-state test.
  });

  const executeQueryMock = jest
    .fn()
    .mockReturnValue(pendingPromise);

  const topMock = jest
    .fn()
    .mockReturnValue(executeQueryMock);

  const orderByMock = jest
    .fn()
    .mockReturnValue({
      top: topMock,
    });

  const selectMock = jest
    .fn()
    .mockReturnValue({
      orderBy: orderByMock,
    });

  return {
    items: {
      select: selectMock,
    },
  };
};

/**
 * Creates a rejected SharePoint query.
 */
const createFailedListMock = (
  message: string = "SharePoint API error",
): IMockList => {
  const executeQueryMock = jest
    .fn()
    .mockRejectedValue(new Error(message));

  const topMock = jest
    .fn()
    .mockReturnValue(executeQueryMock);

  const orderByMock = jest
    .fn()
    .mockReturnValue({
      top: topMock,
    });

  const selectMock = jest
    .fn()
    .mockReturnValue({
      orderBy: orderByMock,
    });

  return {
    items: {
      select: selectMock,
    },
  };
};

describe("EsgFeedbackWidget component", (): void => {
  beforeEach((): void => {
    jest.clearAllMocks();
    window.sessionStorage.clear();

    const mockSp: IMockSp = {
      web: {
        lists: {
          getByTitle: mockGetByTitle,
        },
      },
    };

    mockedGetSP.mockReturnValue(
      mockSp as unknown as ReturnType<typeof getSP>,
    );
  });

  test(
    "renders the loading indicator initially",
    (): void => {
      mockGetByTitle.mockReturnValue(
        createPendingListMock(),
      );

      render(
        <EsgFeedbackWidget context={mockContext} />,
      );

      expect(
        screen.getByText("Loading ESG feedback..."),
      ).toBeInTheDocument();
    },
  );

  test(
    "fetches the latest Tier 1 item and displays genuine SharePoint values",
    async (): Promise<void> => {
      const mockApiResponse: IMockTierItem[] = [
        {
          Id: 1,
          Supplier_x0020_Name: "Acme Corp Ltd",
          Email: "acme@test.com",
          OverallQuestionsPercentage: 0.82,
          EnvQuestionsCount: 24,
          SocialQuestionsCount: 20,
          GovernanceQuestionsCount: 18,
          ComplianceCount: 16,
        },
      ];

      mockGetByTitle.mockReturnValue(
        createListMock(mockApiResponse),
      );

      render(
        <EsgFeedbackWidget context={mockContext} />,
      );

      await waitFor((): void => {
        expect(
          screen.getByText("Latest ESG Feedback"),
        ).toBeInTheDocument();
      });

      expect(
        screen.getByText("Acme Corp Ltd"),
      ).toBeInTheDocument();

      expect(
        screen.getByText("82%"),
      ).toBeInTheDocument();

      expect(
        screen.getByText("Overall Score"),
      ).toBeInTheDocument();

      expect(
        screen.getByText("Environmental"),
      ).toBeInTheDocument();

      expect(
        screen.getByText("Social"),
      ).toBeInTheDocument();

      expect(
        screen.getByText("Governance"),
      ).toBeInTheDocument();

      expect(
        screen.getByText("24 questions"),
      ).toBeInTheDocument();

      expect(
        screen.getByText("20 questions"),
      ).toBeInTheDocument();

      expect(
        screen.getByText("18 questions"),
      ).toBeInTheDocument();

      expect(
        screen.getByTestId("responsive-container"),
      ).toBeInTheDocument();

      expect(
        screen.getByTestId("pie-chart"),
      ).toBeInTheDocument();

      expect(
        screen.getByTestId("pie"),
      ).toBeInTheDocument();

      expect(
        screen.getAllByTestId("pie-cell"),
      ).toHaveLength(3);

      expect(
        mockGetByTitle,
      ).toHaveBeenCalledTimes(1);

      expect(
        mockGetByTitle,
      ).toHaveBeenCalledWith("CASSTECH_SSQ");
    },
  );

  test(
    "uses the email when the supplier name is unavailable",
    async (): Promise<void> => {
      const mockApiResponse: IMockTierItem[] = [
        {
          Id: 2,
          Email: "fallback@test.com",
          OverallQuestionsPercentage: 0.75,
          EnvQuestionsCount: 20,
          SocialQuestionsCount: 18,
          GovernanceQuestionsCount: 15,
        },
      ];

      mockGetByTitle.mockReturnValue(
        createListMock(mockApiResponse),
      );

      render(
        <EsgFeedbackWidget context={mockContext} />,
      );

      await waitFor((): void => {
        expect(
          screen.getByText("fallback@test.com"),
        ).toBeInTheDocument();
      });

      expect(
        screen.getByText("75%"),
      ).toBeInTheDocument();

      expect(
        screen.getByText("20 questions"),
      ).toBeInTheDocument();

      expect(
        screen.getByText("18 questions"),
      ).toBeInTheDocument();

      expect(
        screen.getByText("15 questions"),
      ).toBeInTheDocument();

      expect(
        mockGetByTitle,
      ).toHaveBeenCalledTimes(1);
    },
  );

  test(
    "uses the fallback supplier label when name and email are unavailable",
    async (): Promise<void> => {
      const mockApiResponse: IMockTierItem[] = [
        {
          Id: 3,
          OverallQuestionsPercentage: 0.7,
          EnvQuestionsCount: 18,
          SocialQuestionsCount: 16,
          GovernanceQuestionsCount: 14,
        },
      ];

      mockGetByTitle.mockReturnValue(
        createListMock(mockApiResponse),
      );

      render(
        <EsgFeedbackWidget context={mockContext} />,
      );

      await waitFor((): void => {
        expect(
          screen.getByText("Latest supplier"),
        ).toBeInTheDocument();
      });

      expect(
        screen.getByText("70%"),
      ).toBeInTheDocument();
    },
  );

  test(
    "normalizes a whole-number overall percentage",
    async (): Promise<void> => {
      const mockApiResponse: IMockTierItem[] = [
        {
          Id: 4,
          Supplier_x0020_Name:
            "Whole Percentage Supplier",
          OverallQuestionsPercentage: 88,
          EnvQuestionsCount: 22,
          SocialQuestionsCount: 19,
          GovernanceQuestionsCount: 17,
        },
      ];

      mockGetByTitle.mockReturnValue(
        createListMock(mockApiResponse),
      );

      render(
        <EsgFeedbackWidget context={mockContext} />,
      );

      await waitFor((): void => {
        expect(
          screen.getByText(
            "Whole Percentage Supplier",
          ),
        ).toBeInTheDocument();
      });

      expect(
        screen.getByText("88%"),
      ).toBeInTheDocument();
    },
  );

  test(
    "uses ComplianceCount when GovernanceQuestionsCount is unavailable",
    async (): Promise<void> => {
      const mockApiResponse: IMockTierItem[] = [
        {
          Id: 5,
          Supplier_x0020_Name:
            "Compliance Supplier",
          OverallQuestionsPercentage: 0.65,
          EnvQuestionsCount: 19,
          SocialQuestionsCount: 17,
          ComplianceCount: 14,
        },
      ];

      mockGetByTitle.mockReturnValue(
        createListMock(mockApiResponse),
      );

      render(
        <EsgFeedbackWidget context={mockContext} />,
      );

      await waitFor((): void => {
        expect(
          screen.getByText(
            "Compliance Supplier",
          ),
        ).toBeInTheDocument();
      });

      expect(
        screen.getByText("65%"),
      ).toBeInTheDocument();

      expect(
        screen.getByText("14 questions"),
      ).toBeInTheDocument();
    },
  );

  test(
    "loads valid data from sessionStorage without querying SharePoint",
    async (): Promise<void> => {
      const cachedPayload: IMockCachePayload = {
        supplierName: "Cached Enterprise Inc",
        overallScore: 94,
        scores: [
          {
            name: "Environmental",
            value: 25,
            color: "#0088fe",
          },
          {
            name: "Social",
            value: 22,
            color: "#00c49f",
          },
          {
            name: "Governance",
            value: 20,
            color: "#ffbb28",
          },
        ],
      };

      window.sessionStorage.setItem(
        CACHE_KEY,
        JSON.stringify(cachedPayload),
      );

      window.sessionStorage.setItem(
        CACHE_TIME_KEY,
        Date.now().toString(),
      );

      render(
        <EsgFeedbackWidget context={mockContext} />,
      );

      await waitFor((): void => {
        expect(
          screen.getByText(
            "Cached Enterprise Inc",
          ),
        ).toBeInTheDocument();
      });

      expect(
        screen.getByText("94%"),
      ).toBeInTheDocument();

      expect(
        screen.getByText("25 questions"),
      ).toBeInTheDocument();

      expect(
        screen.getByText("22 questions"),
      ).toBeInTheDocument();

      expect(
        screen.getByText("20 questions"),
      ).toBeInTheDocument();

      expect(
        mockGetByTitle,
      ).not.toHaveBeenCalled();
    },
  );

  test(
    "retrieves fresh data when the sessionStorage cache is expired",
    async (): Promise<void> => {
      const expiredTimestamp =
        Date.now() - 15 * 60 * 1000;

      const expiredPayload: IMockCachePayload = {
        supplierName: "Stale Supplier",
        overallScore: 50,
        scores: [
          {
            name: "Environmental",
            value: 10,
            color: "#0088fe",
          },
          {
            name: "Social",
            value: 10,
            color: "#00c49f",
          },
          {
            name: "Governance",
            value: 10,
            color: "#ffbb28",
          },
        ],
      };

      window.sessionStorage.setItem(
        CACHE_KEY,
        JSON.stringify(expiredPayload),
      );

      window.sessionStorage.setItem(
        CACHE_TIME_KEY,
        expiredTimestamp.toString(),
      );

      const freshApiResponse: IMockTierItem[] = [
        {
          Id: 6,
          Supplier_x0020_Name:
            "Fresh Data Supplier",
          OverallQuestionsPercentage: 0.88,
          EnvQuestionsCount: 24,
          SocialQuestionsCount: 21,
          GovernanceQuestionsCount: 19,
        },
      ];

      mockGetByTitle.mockReturnValue(
        createListMock(freshApiResponse),
      );

      render(
        <EsgFeedbackWidget context={mockContext} />,
      );

      await waitFor((): void => {
        expect(
          screen.getByText(
            "Fresh Data Supplier",
          ),
        ).toBeInTheDocument();
      });

      expect(
        screen.getByText("88%"),
      ).toBeInTheDocument();

      expect(
        screen.queryByText("Stale Supplier"),
      ).not.toBeInTheDocument();

      expect(
        mockGetByTitle,
      ).toHaveBeenCalledTimes(1);
    },
  );

  test(
    "ignores an invalid cache payload and retrieves SharePoint data",
    async (): Promise<void> => {
      window.sessionStorage.setItem(
        CACHE_KEY,
        JSON.stringify({
          supplierName: "Invalid Cache Supplier",
        }),
      );

      window.sessionStorage.setItem(
        CACHE_TIME_KEY,
        Date.now().toString(),
      );

      const freshApiResponse: IMockTierItem[] = [
        {
          Id: 7,
          Supplier_x0020_Name:
            "Valid SharePoint Supplier",
          OverallQuestionsPercentage: 0.72,
          EnvQuestionsCount: 20,
          SocialQuestionsCount: 18,
          GovernanceQuestionsCount: 15,
        },
      ];

      mockGetByTitle.mockReturnValue(
        createListMock(freshApiResponse),
      );

      render(
        <EsgFeedbackWidget context={mockContext} />,
      );

      await waitFor((): void => {
        expect(
          screen.getByText(
            "Valid SharePoint Supplier",
          ),
        ).toBeInTheDocument();
      });

      expect(
        screen.getByText("72%"),
      ).toBeInTheDocument();

      expect(
        screen.queryByText(
          "Invalid Cache Supplier",
        ),
      ).not.toBeInTheDocument();

      expect(
        mockGetByTitle,
      ).toHaveBeenCalledTimes(1);
    },
  );

  test(
    "ignores a cache entry with an invalid timestamp",
    async (): Promise<void> => {
      const cachedPayload: IMockCachePayload = {
        supplierName:
          "Invalid Timestamp Supplier",
        overallScore: 99,
        scores: [],
      };

      window.sessionStorage.setItem(
        CACHE_KEY,
        JSON.stringify(cachedPayload),
      );

      window.sessionStorage.setItem(
        CACHE_TIME_KEY,
        "invalid-timestamp",
      );

      const freshApiResponse: IMockTierItem[] = [
        {
          Id: 8,
          Supplier_x0020_Name:
            "Timestamp Fresh Supplier",
          OverallQuestionsPercentage: 0.64,
          EnvQuestionsCount: 18,
          SocialQuestionsCount: 16,
          GovernanceQuestionsCount: 14,
        },
      ];

      mockGetByTitle.mockReturnValue(
        createListMock(freshApiResponse),
      );

      render(
        <EsgFeedbackWidget context={mockContext} />,
      );

      await waitFor((): void => {
        expect(
          screen.getByText(
            "Timestamp Fresh Supplier",
          ),
        ).toBeInTheDocument();
      });

      expect(
        screen.getByText("64%"),
      ).toBeInTheDocument();

      expect(
        screen.queryByText(
          "Invalid Timestamp Supplier",
        ),
      ).not.toBeInTheDocument();

      expect(
        mockGetByTitle,
      ).toHaveBeenCalledTimes(1);
    },
  );

  test(
    "displays the empty state when no submission exists",
    async (): Promise<void> => {
      mockGetByTitle.mockReturnValue(
        createListMock([]),
      );

      render(
        <EsgFeedbackWidget context={mockContext} />,
      );

      await waitFor((): void => {
        expect(
          screen.getByText(
            "No submissions found",
          ),
        ).toBeInTheDocument();
      });

      expect(
        screen.getByText("0%"),
      ).toBeInTheDocument();

      expect(
        screen.getAllByText("0 questions"),
      ).toHaveLength(3);

      expect(
        mockGetByTitle,
      ).toHaveBeenCalledTimes(1);
    },
  );

  test(
    "preserves legitimate zero values from SharePoint",
    async (): Promise<void> => {
      const mockApiResponse: IMockTierItem[] = [
        {
          Id: 9,
          Supplier_x0020_Name:
            "Zero Score Supplier",
          OverallQuestionsPercentage: 0,
          EnvQuestionsCount: 0,
          SocialQuestionsCount: 0,
          GovernanceQuestionsCount: 0,
        },
      ];

      mockGetByTitle.mockReturnValue(
        createListMock(mockApiResponse),
      );

      render(
        <EsgFeedbackWidget context={mockContext} />,
      );

      await waitFor((): void => {
        expect(
          screen.getByText(
            "Zero Score Supplier",
          ),
        ).toBeInTheDocument();
      });

      expect(
        screen.getByText("0%"),
      ).toBeInTheDocument();

      expect(
        screen.getAllByText("0 questions"),
      ).toHaveLength(3);
    },
  );

  test(
    "shows a warning when the SharePoint request fails",
    async (): Promise<void> => {
      mockGetByTitle.mockReturnValue(
        createFailedListMock(),
      );

      render(
        <EsgFeedbackWidget context={mockContext} />,
      );

      await waitFor((): void => {
        expect(
          screen.getByRole("alert"),
        ).toHaveTextContent(
          "ESG feedback could not be loaded from SharePoint.",
        );
      });

      expect(
        screen.getByText(
          "Unable to load supplier",
        ),
      ).toBeInTheDocument();

      expect(
        screen.getByText("0%"),
      ).toBeInTheDocument();

      expect(
        screen.getByText("Environmental"),
      ).toBeInTheDocument();

      expect(
        screen.getByText("Social"),
      ).toBeInTheDocument();

      expect(
        screen.getByText("Governance"),
      ).toBeInTheDocument();

      expect(
        screen.getAllByText("0 questions"),
      ).toHaveLength(3);

      expect(
        mockGetByTitle,
      ).toHaveBeenCalledTimes(1);
    },
  );
});