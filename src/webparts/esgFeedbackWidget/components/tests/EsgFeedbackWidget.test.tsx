jest.mock("../../../../pnpConfig", () => ({
  getSP: jest.fn(),
}));

import * as React from "react";

import { fireEvent, render, screen, waitFor } from "@testing-library/react";

import "@testing-library/jest-dom";

import { WebPartContext } from "@microsoft/sp-webpart-base";

import { getSP } from "../../../../pnpConfig";
import EsgFeedbackWidget from "../EsgFeedbackWidget";

interface IMockSharePointUser {
  Title?: string;
  EMail?: string;
}

interface IMockSubmission {
  Id?: number;
  Created?: string;
  Modified?: string;
  Author?: IMockSharePointUser;
  Editor?: IMockSharePointUser;
  Supplier_x0020_Name?: string;
  field_3?: string;
  field_4?: string;
  Email?: string;
  Name?: string;
  EnvironmentalQuestionsCount?: number;
  EnvQuestionsCount?: number;
  SocialQuestionsCount?: number;
  GovernanceQuestionsCount?: number;
  OverallQuestionsPercentage?: number | string;
}

interface IMockListItems {
  select: jest.Mock;
}

interface IMockList {
  items: IMockListItems;
}

interface IMockSp {
  web: {
    lists: {
      getByTitle: jest.Mock;
    };
  };
}

interface IListResponse {
  listTitle: string;
  items?: IMockSubmission[];
  error?: Error;
}

const TIER_1_LIST = "CASSTECH_SSQ";
const TIER_2_LIST = "Tier 2 ESG Procurement Questionnaire";
const TIER_3_LIST = "Supplier Sustainability Questionnaires Tier 3";

const mockedGetSP = getSP as jest.MockedFunction<typeof getSP>;
const mockGetByTitle = jest.fn();

const mockContext = {
  pageContext: {
    web: {
      absoluteUrl: "https://contoso.sharepoint.com/sites/esg",
    },
    cultureInfo: {
      currentUICultureName: "en-GB",
    },
  },
} as unknown as WebPartContext;

const renderWidget = (): ReturnType<typeof render> => {
  return render(
    <EsgFeedbackWidget
      context={mockContext}
      tier1ListTitle={TIER_1_LIST}
      tier2ListTitle={TIER_2_LIST}
      tier3ListTitle={TIER_3_LIST}
    />,
  );
};

const createQueryMock = (
  items: IMockSubmission[],
  error?: Error,
): IMockList => {
  const executeQuery = error
    ? jest.fn().mockRejectedValue(error)
    : jest.fn().mockResolvedValue(items);

  const top = jest.fn().mockReturnValue(executeQuery);
  const orderBy = jest.fn().mockReturnValue({ top });
  const expand = jest.fn().mockReturnValue({ orderBy });
  const select = jest.fn().mockReturnValue({ expand });

  return {
    items: {
      select,
    },
  };
};

const configureResponses = (responses: IListResponse[]): void => {
  mockGetByTitle.mockImplementation((listTitle: string): IMockList => {
    let matchingResponse: IListResponse | undefined;

    for (let index: number = 0; index < responses.length; index += 1) {
      if (responses[index].listTitle === listTitle) {
        matchingResponse = responses[index];
        break;
      }
    }

    if (!matchingResponse) {
      return createQueryMock([]);
    }

    return createQueryMock(
      matchingResponse.items || [],
      matchingResponse.error,
    );
  });
};

const configureEmptyResponses = (): void => {
  configureResponses([
    { listTitle: TIER_1_LIST, items: [] },
    { listTitle: TIER_2_LIST, items: [] },
    { listTitle: TIER_3_LIST, items: [] },
  ]);
};

const tier1Item: IMockSubmission = {
  Id: 11,
  Created: "2026-08-15T08:00:00Z",
  Modified: "2026-08-15T09:00:00Z",
  Author: {
    Title: "Tier One Author",
    EMail: "tier1.author@contoso.com",
  },
  Editor: {
    Title: "Tier One Editor",
    EMail: "tier1.editor@contoso.com",
  },
  Supplier_x0020_Name: "Tier One Supplier",
  field_3: "tier1.supplier@contoso.com",
  field_4: "Tier One Contact",
  EnvironmentalQuestionsCount: 13,
  SocialQuestionsCount: 12,
  GovernanceQuestionsCount: 7,
  OverallQuestionsPercentage: 10,
};

const tier2Item: IMockSubmission = {
  Id: 22,
  Created: "2026-08-16T08:00:00Z",
  Modified: "2026-08-16T09:00:00Z",
  Author: {
    Title: "Tier Two Author",
    EMail: "tier2.author@contoso.com",
  },
  Editor: {
    Title: "Tier Two Editor",
    EMail: "tier2.editor@contoso.com",
  },
  Supplier_x0020_Name: "Tier Two Supplier",
  Email: "tier2.supplier@contoso.com",
  Name: "Tier Two Contact",
  EnvQuestionsCount: 10,
  SocialQuestionsCount: 9,
  GovernanceQuestionsCount: 8,
  OverallQuestionsPercentage: 4,
};

const tier3Item: IMockSubmission = {
  Id: 33,
  Created: "2026-08-17T08:00:00Z",
  Modified: "2026-08-17T09:00:00Z",
  Author: {
    Title: "Tier Three Author",
    EMail: "tier3.author@contoso.com",
  },
  Editor: {
    Title: "Tier Three Editor",
    EMail: "tier3.editor@contoso.com",
  },
  Supplier_x0020_Name: "Tier Three Supplier",
  Email: "tier3.supplier@contoso.com",
  Name: "Tier Three Contact",
  EnvQuestionsCount: 6,
  SocialQuestionsCount: 5,
  GovernanceQuestionsCount: 4,
  OverallQuestionsPercentage: 1.3,
};

describe("EsgFeedbackWidget", (): void => {
  beforeEach((): void => {
    jest.clearAllMocks();

    const mockSp: IMockSp = {
      web: {
        lists: {
          getByTitle: mockGetByTitle,
        },
      },
    };

    mockedGetSP.mockReturnValue(mockSp as unknown as ReturnType<typeof getSP>);

    configureEmptyResponses();
  });

  afterEach((): void => {
    jest.restoreAllMocks();
  });

  test("retrieves the latest submission across all three tiers", async (): Promise<void> => {
    configureResponses([
      { listTitle: TIER_1_LIST, items: [tier1Item] },
      { listTitle: TIER_2_LIST, items: [tier2Item] },
      { listTitle: TIER_3_LIST, items: [tier3Item] },
    ]);

    renderWidget();

    expect(await screen.findByText("Tier Three Supplier")).toBeInTheDocument();

    expect(screen.getByText("Tier 3")).toBeInTheDocument();
    expect(screen.getByText("Tier Three Contact")).toBeInTheDocument();
    expect(screen.getByText("65%")).toBeInTheDocument();
    expect(screen.getByText("1.3 / 2")).toBeInTheDocument();
    expect(screen.getByText("Conditionally Qualified")).toBeInTheDocument();
    expect(screen.getByText("Medium Risk")).toBeInTheDocument();
    expect(screen.getByText("Corrective Action")).toBeInTheDocument();
  });

  test("uses Tier 1 internal fields and displays actual values", async (): Promise<void> => {
    configureResponses([
      { listTitle: TIER_1_LIST, items: [tier1Item] },
      { listTitle: TIER_2_LIST, items: [] },
      { listTitle: TIER_3_LIST, items: [] },
    ]);

    renderWidget();

    expect(await screen.findByText("Tier One Supplier")).toBeInTheDocument();

    expect(screen.getByText("Tier One Contact")).toBeInTheDocument();
    expect(screen.getByText("100%")).toBeInTheDocument();
    expect(screen.getByText("10 / 10")).toBeInTheDocument();
    expect(screen.getByText("13")).toBeInTheDocument();
    expect(screen.getByText("12")).toBeInTheDocument();
    expect(screen.getByText("7")).toBeInTheDocument();
    expect(screen.getByText("Qualified")).toBeInTheDocument();
    expect(screen.getByText("Low Risk")).toBeInTheDocument();
    expect(screen.getByText("Approve")).toBeInTheDocument();
  });

  test("uses Tier 2 environmental field and calculates relative score", async (): Promise<void> => {
    configureResponses([
      { listTitle: TIER_1_LIST, items: [] },
      { listTitle: TIER_2_LIST, items: [tier2Item] },
      { listTitle: TIER_3_LIST, items: [] },
    ]);

    renderWidget();

    expect(await screen.findByText("Tier Two Supplier")).toBeInTheDocument();

    expect(screen.getByText("10")).toBeInTheDocument();
    expect(screen.getByText("9")).toBeInTheDocument();
    expect(screen.getByText("8")).toBeInTheDocument();
    expect(screen.getByText("80%")).toBeInTheDocument();
    expect(screen.getByText("4 / 5")).toBeInTheDocument();
  });

  test("generates a system assessment note from the result", async (): Promise<void> => {
    configureResponses([
      { listTitle: TIER_1_LIST, items: [] },
      { listTitle: TIER_2_LIST, items: [tier2Item] },
      { listTitle: TIER_3_LIST, items: [] },
    ]);

    renderWidget();

    const assessmentNote: HTMLElement = await screen.findByText(
      /System-generated assessment for Tier 2:/i,
    );

    expect(assessmentNote).toHaveTextContent(
      "the SharePoint weighted score is 4 of 5 (80%)",
    );
    expect(assessmentNote).toHaveTextContent(
      "meets the configured qualification threshold",
    );
  });

  test("shows Modified By and Modified as update metadata", async (): Promise<void> => {
    configureResponses([
      { listTitle: TIER_1_LIST, items: [tier1Item] },
      { listTitle: TIER_2_LIST, items: [] },
      { listTitle: TIER_3_LIST, items: [] },
    ]);

    renderWidget();

    expect(await screen.findByText("Tier One Editor")).toBeInTheDocument();
    expect(screen.getByText("15 Aug 2026")).toBeInTheDocument();
    expect(screen.getByText("Last Updated By")).toBeInTheDocument();
    expect(screen.getByText("Last Updated")).toBeInTheDocument();
  });

  test("opens the matching list when View all is clicked", async (): Promise<void> => {
    const openSpy = jest
      .spyOn(window, "open")
      .mockImplementation((): Window | null => null);

    configureResponses([
      { listTitle: TIER_1_LIST, items: [] },
      { listTitle: TIER_2_LIST, items: [tier2Item] },
      { listTitle: TIER_3_LIST, items: [] },
    ]);

    renderWidget();

    const viewAllButton: HTMLElement = await screen.findByRole("button", {
      name: /view all/i,
    });

    fireEvent.click(viewAllButton);

    expect(openSpy).toHaveBeenCalledWith(
      "https://contoso.sharepoint.com/sites/esg/Lists/" +
        "Tier%202%20ESG%20Procurement%20Questionnaire/" +
        "AllItems.aspx",
      "_blank",
      "noopener,noreferrer",
    );
  });

  test("opens the exact item when View Full Assessment is clicked", async (): Promise<void> => {
    const openSpy = jest
      .spyOn(window, "open")
      .mockImplementation((): Window | null => null);

    configureResponses([
      { listTitle: TIER_1_LIST, items: [] },
      { listTitle: TIER_2_LIST, items: [] },
      { listTitle: TIER_3_LIST, items: [tier3Item] },
    ]);

    renderWidget();

    const assessmentButton: HTMLElement = await screen.findByRole("button", {
      name: /view full assessment/i,
    });

    fireEvent.click(assessmentButton);

    expect(openSpy).toHaveBeenCalledWith(
      "https://contoso.sharepoint.com/sites/esg/Lists/" +
        "Supplier%20Sustainability%20Questionnaires%20Tier%203/" +
        "DispForm.aspx?ID=33",
      "_blank",
      "noopener,noreferrer",
    );
  });

  test("shows the empty state when all lists are empty", async (): Promise<void> => {
    configureEmptyResponses();

    renderWidget();

    expect(
      await screen.findByText("No ESG questionnaire submissions were found."),
    ).toBeInTheDocument();

    expect(screen.getByRole("button", { name: /view all/i })).toBeDisabled();
  });

  test("shows an error when a SharePoint request fails", async (): Promise<void> => {
    configureResponses([
      { listTitle: TIER_1_LIST, items: [] },
      {
        listTitle: TIER_2_LIST,
        error: new Error("SharePoint API error"),
      },
      { listTitle: TIER_3_LIST, items: [] },
    ]);

    renderWidget();

    expect(
      await screen.findByText(
        "The latest ESG feedback could not be loaded from SharePoint.",
      ),
    ).toBeInTheDocument();
  });

  test("queries all three configured lists", async (): Promise<void> => {
    configureEmptyResponses();

    renderWidget();

    await screen.findByText("No ESG questionnaire submissions were found.");

    await waitFor((): void => {
      expect(mockGetByTitle).toHaveBeenCalledTimes(3);
    });

    expect(mockGetByTitle).toHaveBeenCalledWith(TIER_1_LIST);
    expect(mockGetByTitle).toHaveBeenCalledWith(TIER_2_LIST);
    expect(mockGetByTitle).toHaveBeenCalledWith(TIER_3_LIST);
  });
});
