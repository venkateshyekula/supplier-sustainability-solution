import * as React from "react";
import { Icon, Spinner, SpinnerSize } from "@fluentui/react";
import { KpiMetricCards } from "../../kpiMetricCards/components/KpiMetricCards";
import EsgFeedbackWidget from "../../esgFeedbackWidget/components/EsgFeedbackWidget";
import QuestionnaireQuickLinks from "../../questionnaireQuickLinks/components/QuestionnaireQuickLinks";
import SupplierEsgSearch from "../../supplierEsgSearch/components/SupplierEsgSearch";
import { ISupplierSustainabilityDashboardProps } from "./ISupplierSustainabilityDashboardProps";
import { ISiteSearchResult } from "../models/ISiteSearchResult";
import { SiteSearchService } from "../services/SiteSearchService";
import styles from "./SupplierSustainabilityDashboard.module.scss";

const SEARCH_DEBOUNCE_DELAY: number = 350;
const SEARCH_RESULT_LIMIT: number = 20;

type SiteSearchScope = "all" | "pages" | "documents" | "questionnaires";

const SCOPE_OPTIONS: { key: SiteSearchScope; label: string }[] = [
  { key: "all", label: "Site Search" },
  { key: "pages", label: "Site Pages" },
  { key: "questionnaires", label: "Questionnaires" },
  { key: "documents", label: "Documents" },
];

export function SupplierSustainabilityDashboard(
  props: ISupplierSustainabilityDashboardProps,
): React.ReactElement {
  const userDisplayName: string =
    props.context.pageContext.user.displayName?.trim() || "User";

  const [searchText, setSearchText] = React.useState<string>("");
  const [selectedSearchScope, setSelectedSearchScope] =
    React.useState<SiteSearchScope>("all");
  const [searchResults, setSearchResults] = React.useState<
    readonly ISiteSearchResult[]
  >([]);
  const [isSearchLoading, setIsSearchLoading] = React.useState<boolean>(false);
  const [searchError, setSearchError] = React.useState<string | undefined>(
    undefined,
  );
  const [isSearchFocused, setIsSearchFocused] = React.useState<boolean>(false);
  const [activeResultIndex, setActiveResultIndex] = React.useState<number>(-1);
  const [isScopeDropdownOpen, setIsScopeDropdownOpen] =
    React.useState<boolean>(false);

  const searchInputReference: React.RefObject<HTMLInputElement> =
    React.useRef<HTMLInputElement>(null);
  const scopeDropdownRef = React.useRef<HTMLDivElement>(null);

  const searchService: SiteSearchService =
    React.useMemo((): SiteSearchService => {
      return new SiteSearchService(props.context);
    }, [props.context]);

  const normalizedSearchText: string = searchText.trim();

  const searchShellRef = React.useRef<HTMLDivElement>(null);

  // Close scope dropdown and search results when clicking outside search container
  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent): void => {
      if (
        searchShellRef.current &&
        !searchShellRef.current.contains(event.target as Node)
      ) {
        setIsSearchFocused(false);
        setIsScopeDropdownOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  React.useEffect((): (() => void) => {
    if (normalizedSearchText.length < 2) {
      setSearchResults([]);
      setSearchError(undefined);
      setIsSearchLoading(false);
      setActiveResultIndex(-1);

      return (): void => {
        // No timeout to clear.
      };
    }

    let isDisposed: boolean = false;

    const timeoutId: number = window.setTimeout((): void => {
      setIsSearchLoading(true);
      setSearchError(undefined);

      searchService
        .searchSite(normalizedSearchText, SEARCH_RESULT_LIMIT)
        .then((results: ISiteSearchResult[]): void => {
          if (isDisposed) {
            return;
          }

          const filteredResults: ISiteSearchResult[] =
            filterSearchResultsByScope(results, selectedSearchScope);

          const limitedResults: ISiteSearchResult[] = filteredResults.slice(
            0,
            8,
          );

          setSearchResults(limitedResults);
          setActiveResultIndex(limitedResults.length > 0 ? 0 : -1);
        })
        .catch((error: unknown): void => {
          if (isDisposed) {
            return;
          }

          setSearchResults([]);
          setActiveResultIndex(-1);

          setSearchError(
            error instanceof Error
              ? error.message
              : "Unable to search the SharePoint site.",
          );
        })
        .finally((): void => {
          if (!isDisposed) {
            setIsSearchLoading(false);
          }
        });
    }, SEARCH_DEBOUNCE_DELAY);

    return (): void => {
      isDisposed = true;
      window.clearTimeout(timeoutId);
    };
  }, [normalizedSearchText, searchService, selectedSearchScope]);

  const shouldShowSearchSuggestions: boolean =
    isSearchFocused && normalizedSearchText.length >= 2;

  const resetSearchState = (): void => {
    setSearchText("");
    setSearchResults([]);
    setSearchError(undefined);
    setActiveResultIndex(-1);
    setIsSearchFocused(false);
    searchInputReference.current?.blur();
  };

  const openSearchResult = (result: ISiteSearchResult): void => {
    const targetPath: string = result.path.trim();

    if (!targetPath) {
      return;
    }

    window.open(targetPath, "_blank", "noopener,noreferrer");
    resetSearchState(); // ✅ Clears search input and hides suggestions
  };

  const openFullSearchResults = (): void => {
    if (!normalizedSearchText) {
      return;
    }

    const webAbsoluteUrl: string =
      props.context.pageContext.web.absoluteUrl.replace(/\/+$/, "");

    const scopedQuery: string = getFullSearchQuery(
      normalizedSearchText,
      selectedSearchScope,
    );

    const searchResultsUrl: string =
      `${webAbsoluteUrl}` +
      "/_layouts/15/" +
      "search.aspx/siteall" +
      `?q=${encodeURIComponent(scopedQuery)}`;

    // ✅ Changed from window.location.assign to window.open
    window.open(searchResultsUrl, "_blank", "noopener,noreferrer");
    resetSearchState(); // ✅ Clears search input on full search submit
  };

  const submitSearch = (): void => {
    const activeResult: ISiteSearchResult | undefined =
      activeResultIndex >= 0 ? searchResults[activeResultIndex] : undefined;

    if (activeResult) {
      openSearchResult(activeResult);
      return;
    }

    openFullSearchResults();
  };

  const handleSearchSubmit = (
    event: React.FormEvent<HTMLFormElement>,
  ): void => {
    event.preventDefault();
    submitSearch();
  };

  const handleScopeSelect = (scope: SiteSearchScope): void => {
    setSelectedSearchScope(scope);
    setIsScopeDropdownOpen(false);
    setSearchResults([]);
    setSearchError(undefined);
    setActiveResultIndex(-1);
    setIsSearchFocused(true);

    window.setTimeout((): void => {
      searchInputReference.current?.focus();
    }, 0);
  };

  const handleSearchTextChange = (
    event: React.ChangeEvent<HTMLInputElement>,
  ): void => {
    setSearchText(event.target.value);
    setSearchError(undefined);
    setActiveResultIndex(-1);
  };

  const handleSearchKeyDown = (
    event: React.KeyboardEvent<HTMLInputElement>,
  ): void => {
    if (event.key === "ArrowDown") {
      event.preventDefault();

      if (searchResults.length === 0) {
        return;
      }

      setActiveResultIndex((previousIndex: number): number => {
        if (previousIndex < 0) {
          return 0;
        }

        return Math.min(previousIndex + 1, searchResults.length - 1);
      });

      return;
    }

    if (event.key === "ArrowUp") {
      event.preventDefault();

      if (searchResults.length === 0) {
        return;
      }

      setActiveResultIndex((previousIndex: number): number => {
        if (previousIndex <= 0) {
          return 0;
        }

        return previousIndex - 1;
      });

      return;
    }

    if (event.key === "Home" && searchResults.length > 0) {
      event.preventDefault();
      setActiveResultIndex(0);
      return;
    }

    if (event.key === "End" && searchResults.length > 0) {
      event.preventDefault();
      setActiveResultIndex(searchResults.length - 1);
      return;
    }

    if (event.key === "Escape") {
      event.preventDefault();

      setIsSearchFocused(false);
      setActiveResultIndex(-1);

      searchInputReference.current?.blur();
    }
  };

  const clearSearch = (): void => {
    setSearchText("");
    setSearchResults([]);
    setSearchError(undefined);
    setActiveResultIndex(-1);
    setIsSearchFocused(true);

    window.setTimeout((): void => {
      searchInputReference.current?.focus();
    }, 0);
  };

  return (
    <section
      className={styles.dashboard}
      aria-label={"Supplier Sustainability dashboard"}
    >
      <header className={styles.welcomeHero}>
        <div className={styles.heroDecorationOne} aria-hidden="true" />
        <div className={styles.heroDecorationTwo} aria-hidden="true" />
        <div className={styles.heroDecorationThree} aria-hidden="true" />

        <div className={styles.welcomeContent}>
          <h1 className={styles.welcomeTitle}>
            <span className={styles.welcomePrefix}>Welcome,</span>
            <span className={styles.welcomeName}>{userDisplayName}!</span>
          </h1>

          <p className={styles.welcomeDescription}>
            Access Supplier Sustainability applications, questionnaires,
            supporting documents, analytics, and help resources.
          </p>

          <div className={styles.searchShell}>
            <form
              className={[
                styles.applicationSearch,
                isSearchFocused ? styles.applicationSearchFocused : "",
                shouldShowSearchSuggestions ? styles.applicationSearchOpen : "",
              ]
                .filter(Boolean)
                .join(" ")}
              role="search"
              aria-label="Search Supplier Sustainability applications"
              onSubmit={handleSearchSubmit}
            >
              <div className={styles.searchCategory} ref={scopeDropdownRef}>
                <button
                  type="button"
                  className={styles.customDropdownToggle}
                  onMouseDown={(
                    event: React.MouseEvent<HTMLButtonElement>,
                  ): void => {
                    event.preventDefault();
                  }}
                  onClick={(): void => setIsScopeDropdownOpen((prev) => !prev)}
                  aria-haspopup="listbox"
                  aria-expanded={isScopeDropdownOpen}
                  aria-label="Select search category"
                >
                  <span>
                    {
                      SCOPE_OPTIONS.find(
                        (opt) => opt.key === selectedSearchScope,
                      )?.label
                    }
                  </span>
                  <Icon
                    iconName={isScopeDropdownOpen ? "ChevronUp" : "ChevronDown"}
                    className={styles.categoryChevron}
                    aria-hidden="true"
                  />
                </button>

                {isScopeDropdownOpen && (
                  <ul className={styles.customDropdownMenu} role="listbox">
                    {SCOPE_OPTIONS.map((option) => {
                      const isSelected: boolean =
                        selectedSearchScope === option.key;
                      return (
                        <li
                          key={option.key}
                          role="option"
                          aria-selected={isSelected}
                          className={`${styles.customDropdownItem} ${
                            isSelected ? styles.customDropdownItemActive : ""
                          }`}
                          onClick={(): void => handleScopeSelect(option.key)}
                        >
                          {option.label}
                        </li>
                      );
                    })}
                  </ul>
                )}
              </div>

              <span className={styles.searchDivider} aria-hidden="true" />

              <div
                className={styles.searchInputContainer}
                style={{ position: "relative", overflow: "visible" }}
              >
                <button
                  type="submit"
                  className={styles.searchButton}
                  aria-label="Search"
                  title="Search"
                >
                  <Icon iconName="Search" aria-hidden="true" />
                </button>

                <label
                  htmlFor="supplier-application-search"
                  className={styles.visuallyHidden}
                >
                  Search applications and resources
                </label>

                <input
                  ref={searchInputReference}
                  id="supplier-application-search"
                  type="search"
                  className={styles.searchInput}
                  value={searchText}
                  placeholder={
                    "Type to search any application, page, " +
                    "document, or questionnaire"
                  }
                  autoComplete="off"
                  spellCheck={false}
                  aria-autocomplete="list"
                  aria-controls="supplier-application-search-results"
                  aria-expanded={shouldShowSearchSuggestions}
                  aria-activedescendant={
                    activeResultIndex >= 0
                      ? "supplier-application-search-result-" +
                        activeResultIndex.toString()
                      : undefined
                  }
                  onChange={handleSearchTextChange}
                  onFocus={(): void => {
                    setIsSearchFocused(true);
                  }}
                  onKeyDown={handleSearchKeyDown}
                />

                {searchText && (
                  <button
                    type="button"
                    className={styles.clearSearchButton}
                    aria-label="Clear search"
                    title="Clear search"
                    onMouseDown={(
                      event: React.MouseEvent<HTMLButtonElement>,
                    ): void => {
                      event.preventDefault();
                    }}
                    onClick={clearSearch}
                  >
                    <Icon iconName="Cancel" aria-hidden="true" />
                  </button>
                )}

                {/* ✅ KEEP THIS ONLY: Dropdown correctly anchored inside searchInputContainer */}
                {shouldShowSearchSuggestions && (
                  <div
                    id="supplier-application-search-results"
                    className={styles.searchResults}
                    role="listbox"
                    aria-label="Supplier Sustainability search results"
                  >
                    {isSearchLoading && (
                      <div
                        className={styles.searchStatus}
                        role="status"
                        aria-live="polite"
                      >
                        <Spinner size={SpinnerSize.small} />
                        <span>Searching this site...</span>
                      </div>
                    )}

                    {!isSearchLoading && searchError && (
                      <div className={styles.searchStatus} role="alert">
                        <Icon
                          iconName="ErrorBadge"
                          className={styles.searchResultsIcon}
                          aria-hidden="true"
                        />
                        <span>{searchError}</span>
                      </div>
                    )}

                    {!isSearchLoading &&
                      !searchError &&
                      searchResults.length === 0 && (
                        <div
                          className={styles.searchStatus}
                          role="status"
                          aria-live="polite"
                        >
                          <Icon
                            iconName="SearchIssue"
                            className={styles.searchResultsIcon}
                            aria-hidden="true"
                          />
                          <span>
                            No matching content was found in the selected
                            category.
                          </span>
                        </div>
                      )}

                    {!isSearchLoading &&
                      !searchError &&
                      searchResults.map(
                        (
                          result: ISiteSearchResult,
                          index: number,
                        ): React.ReactElement => {
                          const isActive: boolean = index === activeResultIndex;

                          return (
                            <a
                              id={`supplier-application-search-result-${index.toString()}`}
                              key={result.key}
                              href={result.path}
                              target="_blank"
                              rel="noopener noreferrer"
                              data-interception="off"
                              role="option"
                              aria-selected={isActive}
                              className={
                                isActive
                                  ? `${styles.searchResult} ${styles.searchResultActive}`
                                  : styles.searchResult
                              }
                              onMouseDown={(
                                event: React.MouseEvent<HTMLAnchorElement>,
                              ): void => {
                                event.preventDefault();
                              }}
                              onMouseEnter={(): void => {
                                setActiveResultIndex(index);
                              }}
                              onClick={(
                                event: React.MouseEvent<HTMLAnchorElement>,
                              ): void => {
                                // 1. Stop SharePoint Online SPA router from intercepting the link
                                event.nativeEvent.stopImmediatePropagation();
                                event.preventDefault();
                                event.stopPropagation();

                                // Open target path in new tab and reset search input
                                openSearchResult(result);

                                // 2. Explicitly force opening in a new tab
                                if (result.path) {
                                  window.open(
                                    result.path,
                                    "_blank",
                                    "noopener,noreferrer",
                                  );
                                }

                                // 3. Reset focus and hide suggestions UI
                                setIsSearchFocused(false);

                                if (searchInputReference.current) {
                                  searchInputReference.current.blur();
                                }
                              }}
                            >
                              <span
                                className={styles.searchResultIcon}
                                aria-hidden="true"
                              >
                                <Icon iconName={result.iconName} />
                              </span>

                              <span className={styles.searchResultContent}>
                                <span className={styles.searchResultTitle}>
                                  {result.title}
                                </span>

                                {result.summary && (
                                  <span className={styles.searchResultSummary}>
                                    {result.summary}
                                  </span>
                                )}

                                <span
                                  className={styles.searchResultPath}
                                  title={result.path}
                                >
                                  {getDisplayPath(result.path)}
                                </span>
                              </span>

                              <Icon
                                iconName="ChevronRight"
                                className={styles.searchResultChevron}
                                aria-hidden="true"
                              />
                            </a>
                          );
                        },
                      )}
                  </div>
                )}
              </div>
            </form>
          </div>
        </div>

        <div className={styles.heroVisual} aria-hidden="true">
          <span className={styles.heroVisualIcon}>
            <Icon iconName="BarChartVertical" />
          </span>
          <span className={styles.heroVisualIconSecondary}>
            <Icon iconName="ComplianceAudit" />
          </span>
          <span className={styles.heroVisualIconSmall}>
            <Icon iconName="CompletedSolid" />
          </span>
        </div>
      </header>

      <section
        className={styles.overlappingKpiSection}
        aria-labelledby={"supplier-esg-overview-title"}
      >
        <div className={styles.kpiSectionHeader}>
          <div>
            <h2
              id={"supplier-esg-overview-title"}
              className={styles.sectionTitle}
            >
              Supplier ESG Overview
            </h2>

            <p className={styles.sectionDescription}>
              Review current submission, qualification, and assessment metrics
              across all configured questionnaire tiers.
            </p>
          </div>
        </div>

        <KpiMetricCards
          context={props.context}
          tier1ListTitle={props.tier1ListTitle}
          tier2ListTitle={props.tier2ListTitle}
          tier3ListTitle={props.tier3ListTitle}
        />
      </section>

      <div className={styles.insightGrid}>
        <section
          className={styles.quickLinksColumn}
          aria-label={"Questionnaire and supporting document links"}
        >
          <QuestionnaireQuickLinks
            context={props.context}
            tier1QuestionnaireListTitle={props.tier1ListTitle}
            tier2QuestionnaireListTitle={props.tier2ListTitle}
            tier3QuestionnaireListTitle={props.tier3ListTitle}
            tier1DocumentLibraryTitle={props.tier1DocumentLibraryTitle}
            tier2DocumentLibraryTitle={props.tier2DocumentLibraryTitle}
            tier3DocumentLibraryTitle={props.tier3DocumentLibraryTitle}
          />
        </section>

        <aside
          className={styles.feedbackColumn}
          aria-label={"Latest ESG feedback"}
        >
          <EsgFeedbackWidget
            context={props.context}
            tier1ListTitle={props.tier1ListTitle}
            tier2ListTitle={props.tier2ListTitle}
            tier3ListTitle={props.tier3ListTitle}
          />
        </aside>
      </div>

      <section
        className={styles.dashboardSection}
        aria-label={"Supplier ESG submission search"}
      >
        <SupplierEsgSearch
          context={props.context}
          tier1ListTitle={props.tier1ListTitle}
          tier2ListTitle={props.tier2ListTitle}
          tier3ListTitle={props.tier3ListTitle}
          tier1SupplierNameField={props.tier1SupplierNameField}
          tier2SupplierNameField={props.tier2SupplierNameField}
          tier3SupplierNameField={props.tier3SupplierNameField}
        />
      </section>

      <div
        className={styles.supplierInformation}
        role="note"
        aria-label={"Information for suppliers"}
      >
        <span className={styles.supplierInformationIcon} aria-hidden="true">
          <Icon iconName="Info" />
        </span>

        <p className={styles.supplierInformationText}>
          <strong>Suppliers:</strong> Suppliers will receive direct links from
          Procurement to complete the appropriate Tier questionnaire. Please do
          not use this site to submit questionnaires.
        </p>
      </div>
    </section>
  );
}

function filterSearchResultsByScope(
  results: readonly ISiteSearchResult[],
  scope: SiteSearchScope,
): ISiteSearchResult[] {
  if (scope === "all") {
    return [...results];
  }

  return results.filter((result: ISiteSearchResult): boolean => {
    const normalizedPath: string = (result.path || "").trim().toLowerCase();
    const normalizedFileType: string = (result.fileType || "")
      .trim()
      .toLowerCase();
    const normalizedContentClass: string = (result.contentClass || "")
      .trim()
      .toLowerCase();
    const normalizedTitle: string = (result.title || "").trim().toLowerCase();

    if (scope === "pages") {
      return (
        normalizedPath.indexOf("/sitepages/") >= 0 ||
        normalizedFileType === "aspx"
      );
    }

    if (scope === "documents") {
      const documentFileTypes: readonly string[] = [
        "doc",
        "docx",
        "xls",
        "xlsx",
        "csv",
        "ppt",
        "pptx",
        "pdf",
        "txt",
        "zip",
        "rar",
        "7z",
        "jpg",
        "jpeg",
        "png",
        "gif",
        "svg",
      ];

      return (
        documentFileTypes.indexOf(normalizedFileType) >= 0 ||
        normalizedContentClass.indexOf("documentlibrary") >= 0
      );
    }

    if (scope === "questionnaires") {
      return (
        normalizedTitle.indexOf("questionnaire") >= 0 ||
        normalizedTitle.indexOf("assessment") >= 0 ||
        normalizedPath.indexOf("questionnaire") >= 0 ||
        normalizedPath.indexOf("casstech_ssq") >= 0 ||
        normalizedPath.indexOf("completed-questionnaires") >= 0
      );
    }

    return true;
  });
}

function getFullSearchQuery(
  searchText: string,
  scope: SiteSearchScope,
): string {
  switch (scope) {
    case "pages":
      return `${searchText} FileType:aspx`;

    case "documents":
      return (
        `${searchText} (` +
        "FileType:doc OR " +
        "FileType:docx OR " +
        "FileType:xls OR " +
        "FileType:xlsx OR " +
        "FileType:csv OR " +
        "FileType:ppt OR " +
        "FileType:pptx OR " +
        "FileType:pdf OR " +
        "FileType:txt" +
        ")"
      );

    case "questionnaires":
      return `${searchText} (questionnaire OR assessment)`;

    case "all":
    default:
      return searchText;
  }
}

function getDisplayPath(path: string): string {
  try {
    const url: URL = new URL(path);
    return decodeURIComponent(url.pathname) || path;
  } catch {
    return path;
  }
}

export default SupplierSustainabilityDashboard;
