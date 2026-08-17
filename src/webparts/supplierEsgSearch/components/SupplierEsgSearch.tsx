import * as React from 'react';

import {
  ActionButton,
  DatePicker,
  DefaultButton,
  DetailsListLayoutMode,
  Dropdown,
  IColumn,
  IDetailsHeaderProps,
  IDropdownOption,
  IconButton,
  IRenderFunction,
  Link,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  SelectionMode,
  ShimmeredDetailsList,
  Spinner,
  SpinnerSize,
  Stack,
  TextField
} from '@fluentui/react';

import styles from './SupplierEsgSearch.module.scss';

import { ISupplierEsgSearchProps } from './ISupplierEsgSearchProps';

import {
  QualificationStatus,
  Recommendation,
  RiskRating
} from '../models/IQualificationResult';

import {
  IListConfiguration,
  SupplierTier
} from '../models/IListConfiguration';

import { ISupplierSearchFilters } from '../models/ISupplierSearchFilters';
import { ISupplierSubmission } from '../models/ISupplierSubmission';

import { ISupplierSubmissionService } from '../services/ISupplierSubmissionService';
import { SupplierSubmissionService } from '../services/SupplierSubmissionService';

import {
  IQualificationThresholds,
  SupplierQualificationUtility
} from '../utilities/SupplierQualificationUtility';

export interface ISupplierEsgSearchState {
  allSubmissions: ISupplierSubmission[];
  displayedSubmissions: ISupplierSubmission[];

  filters: ISupplierSearchFilters;

  isLoading: boolean;
  errorMessage?: string;
  warningMessage?: string;

  sortedColumnKey: string;
  isSortedDescending: boolean;
}

export default class SupplierEsgSearch
  extends React.Component<
    ISupplierEsgSearchProps,
    ISupplierEsgSearchState
  > {

  private readonly submissionService:
    ISupplierSubmissionService;

  private columns: IColumn[];

  public constructor(props: ISupplierEsgSearchProps) {
    super(props);

    this.submissionService =
      new SupplierSubmissionService(props.context);

    this.state = {
      allSubmissions: [],
      displayedSubmissions: [],

      filters: {
        searchText: ''
      },

      isLoading: true,
      errorMessage: undefined,
      warningMessage: undefined,

      sortedColumnKey: 'created',
      isSortedDescending: true
    };

    this.columns = this.createColumns();
  }

  public componentDidMount(): void {
    this.loadSubmissions()
      .catch(
        (error: Error): void => {
          this.setState({
            isLoading: false,
            errorMessage:
              error.message ||
              'An unexpected error occurred.'
          });
        }
      );
  }

  public componentDidUpdate(
    previousProps: ISupplierEsgSearchProps
  ): void {
    const listConfigurationChanged: boolean =
      previousProps.tier1ListTitle !==
        this.props.tier1ListTitle ||
      previousProps.tier2ListTitle !==
        this.props.tier2ListTitle ||
      previousProps.tier3ListTitle !==
        this.props.tier3ListTitle ||
      previousProps.tier1SupplierNameField !==
        this.props.tier1SupplierNameField ||
      previousProps.tier2SupplierNameField !==
        this.props.tier2SupplierNameField ||
      previousProps.tier3SupplierNameField !==
        this.props.tier3SupplierNameField ||
      previousProps.qualifiedMinimum !==
        this.props.qualifiedMinimum ||
      previousProps.conditionalMinimum !==
        this.props.conditionalMinimum;

    if (listConfigurationChanged) {
      this.loadSubmissions()
        .catch(
          (error: Error): void => {
            this.setState({
              isLoading: false,
              errorMessage:
                error.message ||
                'An unexpected error occurred.'
            });
          }
        );
    }
  }

  public render(): React.ReactElement<ISupplierEsgSearchProps> {
    const tierOptions: IDropdownOption[] = [
      { key: '', text: 'All Tiers' },
      { key: 'Tier 1', text: 'Tier 1' },
      { key: 'Tier 2', text: 'Tier 2' },
      { key: 'Tier 3', text: 'Tier 3' }
    ];

    const qualificationOptions: IDropdownOption[] = [
      { key: '', text: 'All Qualifications' },
      { key: 'Qualified', text: 'Qualified' },
      {
        key: 'Conditionally Qualified',
        text: 'Conditionally Qualified'
      },
      {
        key: 'Not Qualified',
        text: 'Not Qualified'
      }
    ];

    return (
      <section className={styles.supplierEsgSearch}>
        <div className={styles.header}>
          <div>
            <h2 className={styles.title}>
              Search &amp; Filter Submissions
            </h2>

            <p className={styles.subtitle}>
              Search supplier sustainability questionnaire
              submissions across all tiers.
            </p>
          </div>

          <ActionButton
            iconProps={{ iconName: 'Refresh' }}
            text="Refresh"
            disabled={this.state.isLoading}
            onClick={this.handleRefresh}
          />
        </div>

        <div className={styles.filterCard}>
          <div className={styles.filterGrid}>
            <div className={styles.searchField}>
              <TextField
                label="Supplier"
                placeholder="Search supplier name, contact or email"
                value={this.state.filters.searchText}
                disabled={this.state.isLoading}
                onChange={this.handleSearchTextChange}
                onKeyDown={this.handleSearchKeyDown}
              />
            </div>

            <div className={styles.tierField}>
              <Dropdown
                label="Tier"
                options={tierOptions}
                selectedKey={this.state.filters.tier || ''}
                disabled={this.state.isLoading}
                onChange={this.handleTierChange}
              />
            </div>

            <div className={styles.qualificationField}>
              <Dropdown
                label="Qualification"
                options={qualificationOptions}
                selectedKey={this.state.filters.qualification || ''}
                disabled={this.state.isLoading}
                onChange={this.handleQualificationChange}
              />
            </div>

            <div className={styles.minField}>
              <TextField
                label="Minimum %"
                type="number"
                min={0}
                max={100}
                value={
                  this.state.filters.minimumPercentage !== undefined
                    ? this.state.filters.minimumPercentage.toString()
                    : ''
                }
                disabled={this.state.isLoading}
                onChange={this.handleMinimumPercentageChange}
              />
            </div>

            <div className={styles.maxField}>
              <TextField
                label="Maximum %"
                type="number"
                min={0}
                max={100}
                value={
                  this.state.filters.maximumPercentage !== undefined
                    ? this.state.filters.maximumPercentage.toString()
                    : ''
                }
                disabled={this.state.isLoading}
                onChange={this.handleMaximumPercentageChange}
              />
            </div>

            <div className={styles.dateField}>
              <DatePicker
                label="From date"
                placeholder="Select start date"
                value={this.state.filters.startDate}
                disabled={this.state.isLoading}
                onSelectDate={this.handleStartDateChange}
                allowTextInput={false}
              />
            </div>

            <div className={styles.dateField}>
              <DatePicker
                label="To date"
                placeholder="Select end date"
                value={this.state.filters.endDate}
                disabled={this.state.isLoading}
                onSelectDate={this.handleEndDateChange}
                allowTextInput={false}
              />
            </div>

            <div className={styles.filterActions}>
              <PrimaryButton
                text="Search"
                iconProps={{ iconName: 'Search' }}
                disabled={this.state.isLoading}
                onClick={this.applyFilters}
              />

              <DefaultButton
                text="Clear"
                iconProps={{ iconName: 'ClearFilter' }}
                disabled={this.state.isLoading}
                onClick={this.clearFilters}
              />
            </div>
          </div>
        </div>

        {this.renderWarning()}
        {this.renderError()}

        <div className={styles.summaryBar}>
          <div>
            <h3 className={styles.sectionTitle}>
              Recent Submissions
            </h3>

            <span className={styles.resultCount}>
              {this.state.displayedSubmissions.length}
              {' '}
              result
              {this.state.displayedSubmissions.length === 1
                ? ''
                : 's'}
            </span>
          </div>

          {this.state.isLoading && (
            <Spinner
              size={SpinnerSize.small}
              label="Loading submissions"
            />
          )}
        </div>

        <div className={styles.tableContainer}>
          <ShimmeredDetailsList
            items={this.state.displayedSubmissions}
            columns={this.columns}
            setKey="supplier-submissions"
            enableShimmer={this.state.isLoading}
            selectionMode={SelectionMode.none}
            layoutMode={DetailsListLayoutMode.justified}
            compact={false}
            onRenderDetailsHeader={this.onRenderDetailsHeader}
          />

          {!this.state.isLoading &&
            !this.state.errorMessage &&
            this.state.displayedSubmissions.length === 0 && (
              <div className={styles.emptyState}>
                <span className={styles.emptyIcon}>
                  &#128269;
                </span>

                <h3>No submissions found</h3>

                <p>
                  Change the filters or clear the search to
                  display all supplier submissions.
                </p>

                <DefaultButton
                  text="Clear filters"
                  onClick={this.clearFilters}
                />
              </div>
            )}
        </div>

        <div className={styles.ruleNote}>
          <strong>Qualification rules:</strong>{' '}
          SharePoint weighted scores are displayed unchanged. Qualification is
          evaluated relative to each tier maximum: Tier 1 = 10, Tier 2 = 5,
          and Tier 3 = 2. Qualified from {this.props.qualifiedMinimum}%;
          conditionally qualified from {this.props.conditionalMinimum}%; below{' '}
          {this.props.conditionalMinimum}% is not qualified.
        </div>
      </section>
    );
  }

  private createColumns(): IColumn[] {
    return [
      {
        key: 'supplierName',
        name: 'Supplier Name',
        fieldName: 'supplierName',
        minWidth: 170,
        maxWidth: 260,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this.handleColumnClick,
        onRender: (
          item: ISupplierSubmission
        ): React.ReactNode => {
          return (
            <div>
              <Link
                className={styles.supplierLink}
                onClick={(): void => {
                  this.openSubmission(item);
                }}
              >
                {item.supplierName}
              </Link>

              {item.email && (
                <div className={styles.secondaryText}>
                  {item.email}
                </div>
              )}
            </div>
          );
        }
      },
      {
        key: 'tier',
        name: 'Tier',
        fieldName: 'tier',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        onColumnClick: this.handleColumnClick,
        onRender: (
          item: ISupplierSubmission
        ): React.ReactNode => {
          return (
            <span
              className={
                styles.badge +
                ' ' +
                this.getTierClassName(item.tier)
              }
            >
              {item.tier}
            </span>
          );
        }
      },
      {
        key: 'overallPercentage',
        name: 'Overall %',
        fieldName: 'overallPercentage',
        minWidth: 80,
        maxWidth: 100,
        isResizable: true,
        onColumnClick: this.handleColumnClick,
        onRender: (
          item: ISupplierSubmission
        ): React.ReactNode => {
          return (
            <strong className={styles.percentage}>
              {SupplierQualificationUtility
                .formatPercentage(
                  item.overallPercentage
                )}
            </strong>
          );
        }
      },
      {
        key: 'qualification',
        name: 'Qualification',
        fieldName: 'qualification',
        minWidth: 150,
        maxWidth: 190,
        isResizable: true,
        onColumnClick: this.handleColumnClick,
        onRender: (
          item: ISupplierSubmission
        ): React.ReactNode => {
          return (
            <span
              className={
                styles.badge +
                ' ' +
                this.getQualificationClassName(
                  item.qualification
                )
              }
            >
              {item.qualification}
            </span>
          );
        }
      },
      {
        key: 'riskRating',
        name: 'Risk Rating',
        fieldName: 'riskRating',
        minWidth: 100,
        maxWidth: 130,
        isResizable: true,
        onColumnClick: this.handleColumnClick,
        onRender: (
          item: ISupplierSubmission
        ): React.ReactNode => {
          return (
            <span className={styles.riskText}>
              <span
                className={
                  styles.riskDot +
                  ' ' +
                  this.getRiskClassName(
                    item.riskRating
                  )
                }
              />
              {item.riskRating}
            </span>
          );
        }
      },
      {
        key: 'submittedByName',
        name: 'Submitted By',
        fieldName: 'submittedByName',
        minWidth: 110,
        maxWidth: 170,
        isResizable: true,
        onColumnClick: this.handleColumnClick,
        onRender: (
          item: ISupplierSubmission
        ): React.ReactNode => {
          return item.submittedByName || item.email || '-';
        }
      },
      {
        key: 'created',
        name: 'Submission Date',
        fieldName: 'created',
        minWidth: 115,
        maxWidth: 145,
        isResizable: true,
        isSorted: true,
        isSortedDescending: true,
        onColumnClick: this.handleColumnClick,
        onRender: (
          item: ISupplierSubmission
        ): React.ReactNode => {
          return this.formatDate(item.created);
        }
      },
      {
        key: 'recommendation',
        name: 'Recommendation',
        fieldName: 'recommendation',
        minWidth: 120,
        maxWidth: 160,
        isResizable: true,
        onColumnClick: this.handleColumnClick,
        onRender: (
          item: ISupplierSubmission
        ): React.ReactNode => {
          return (
            <span
              className={
                styles.badge +
                ' ' +
                this.getRecommendationClassName(
                  item.recommendation
                )
              }
            >
              {item.recommendation}
            </span>
          );
        }
      },
      {
        key: 'actions',
        name: 'Actions',
        minWidth: 100,
        maxWidth: 100,
        isResizable: false,
        onRender: (
          item: ISupplierSubmission
        ): React.ReactNode => {
          return (
            <IconButton
              iconProps={{ iconName: 'MoreVertical' }}
              title="Submission actions"
              ariaLabel={
                'Actions for ' + item.supplierName
              }
              menuProps={{
                items: [
                  {
                    key: 'open',
                    text: 'Open submission',
                    iconProps: {
                      iconName: 'OpenInNewWindow'
                    },
                    onClick: (): void => {
                      this.openSubmission(item);
                    }
                  },
                  {
                    key: 'copyEmail',
                    text: 'Copy supplier email',
                    iconProps: {
                      iconName: 'Copy'
                    },
                    disabled: !item.email,
                    onClick: (): void => {
                      this.copyEmail(item.email);
                    }
                  }
                ]
              }}
            />
          );
        }
      }
    ];
  }

  private loadSubmissions = async (): Promise<void> => {
    this.setState({
      isLoading: true,
      errorMessage: undefined,
      warningMessage: undefined
    });

    try {
      const listConfigurations:
        readonly IListConfiguration[] = [
          {
            listTitle: this.props.tier1ListTitle,
            tier: 'Tier 1',
            supplierNameDisplayName: 'Supplier Name',
            supplierNameInternalName:
              this.props.tier1SupplierNameField,
            emailInternalName: 'field_3',
            contactNameInternalName: 'field_4',
            overallPercentageInternalName:
              'OverallQuestionsPercentage',
            maximumWeightedScore: 10
          },
          {
            listTitle: this.props.tier2ListTitle,
            tier: 'Tier 2',
            supplierNameDisplayName: 'Supplier Name',
            supplierNameInternalName:
              this.props.tier2SupplierNameField,
            emailInternalName: 'Email',
            contactNameInternalName: 'Name',
            overallPercentageInternalName:
              'OverallQuestionsPercentage',
            maximumWeightedScore: 5
          },
          {
            listTitle: this.props.tier3ListTitle,
            tier: 'Tier 3',
            supplierNameDisplayName: 'Supplier Name',
            supplierNameInternalName:
              this.props.tier3SupplierNameField,
            emailInternalName: 'Email',
            contactNameInternalName: 'Name',
            overallPercentageInternalName:
              'OverallQuestionsPercentage',
            maximumWeightedScore: 2
          }
        ];

      const thresholds: IQualificationThresholds = {
        qualifiedMinimum:
          this.props.qualifiedMinimum,
        conditionalMinimum:
          this.props.conditionalMinimum
      };

      const submissions: ISupplierSubmission[] =
        await this.submissionService.getAllSubmissions(
          listConfigurations,
          thresholds
        );

      this.setState(
        {
          allSubmissions: submissions,
          displayedSubmissions: submissions,
          isLoading: false,
          errorMessage: undefined,
          warningMessage: undefined,
          filters: {
            searchText: ''
          }
        },
        (): void => {
          this.updateColumnSorting();
        }
      );
    } catch (error) {
      const message: string =
        error instanceof Error
          ? error.message
          : 'Unable to load supplier submissions.';

      this.setState({
        allSubmissions: [],
        displayedSubmissions: [],
        isLoading: false,
        errorMessage: message
      });
    }
  };

  private hasSelectedSearchCriteria(
    filters: ISupplierSearchFilters
  ): boolean {
    return Boolean(
      filters.searchText.trim() ||
      filters.tier ||
      filters.qualification ||
      filters.minimumPercentage !== undefined ||
      filters.maximumPercentage !== undefined ||
      filters.startDate ||
      filters.endDate
    );
  }

  private applyFilters = (): void => {
    const filters: ISupplierSearchFilters =
      this.state.filters;

    if (!this.hasSelectedSearchCriteria(filters)) {
      this.setState({
        warningMessage:
          'Please enter or select at least one search criterion before clicking Search.',
        errorMessage: undefined
      });

      return;
    }

    if (
      filters.minimumPercentage !== undefined &&
      filters.maximumPercentage !== undefined &&
      filters.minimumPercentage >
        filters.maximumPercentage
    ) {
      this.setState({
        warningMessage: undefined,
        errorMessage:
          'Minimum percentage cannot be greater than maximum percentage.'
      });

      return;
    }

    const searchText: string =
      filters.searchText.trim().toLowerCase();

    const startDate: Date | undefined =
      filters.startDate
        ? this.getStartOfDay(filters.startDate)
        : undefined;

    const endDate: Date | undefined =
      filters.endDate
        ? this.getEndOfDay(filters.endDate)
        : undefined;

    let filteredSubmissions: ISupplierSubmission[] =
      this.state.allSubmissions.filter(
        (
          submission: ISupplierSubmission
        ): boolean => {
          const supplierName: string =
            submission.supplierName.toLowerCase();

          const email: string =
            submission.email.toLowerCase();

          const contactName: string =
            submission.submittedByName.toLowerCase();

          const matchesSearch: boolean =
            !searchText ||
            supplierName.indexOf(searchText) !== -1 ||
            email.indexOf(searchText) !== -1 ||
            contactName.indexOf(searchText) !== -1;

          const matchesTier: boolean =
            !filters.tier ||
            submission.tier === filters.tier;

          const matchesQualification: boolean =
            !filters.qualification ||
            submission.qualification ===
              filters.qualification;

          const matchesMinimum: boolean =
            filters.minimumPercentage === undefined ||
            submission.overallPercentage >=
              filters.minimumPercentage;

          const matchesMaximum: boolean =
            filters.maximumPercentage === undefined ||
            submission.overallPercentage <=
              filters.maximumPercentage;

          const createdDate: Date =
            new Date(submission.created);

          const matchesStartDate: boolean =
            !startDate ||
            createdDate.getTime() >=
              startDate.getTime();

          const matchesEndDate: boolean =
            !endDate ||
            createdDate.getTime() <=
              endDate.getTime();

          return (
            matchesSearch &&
            matchesTier &&
            matchesQualification &&
            matchesMinimum &&
            matchesMaximum &&
            matchesStartDate &&
            matchesEndDate
          );
        }
      );

    filteredSubmissions =
      this.sortSubmissions(
        filteredSubmissions,
        this.state.sortedColumnKey,
        this.state.isSortedDescending
      );

    this.setState({
      displayedSubmissions:
        filteredSubmissions,
      warningMessage: undefined,
      errorMessage: undefined
    });
  };

  private clearFilters = (): void => {
    const clearedFilters: ISupplierSearchFilters = {
      searchText: ''
    };

    const sortedSubmissions: ISupplierSubmission[] =
      this.sortSubmissions(
        this.state.allSubmissions,
        'created',
        true
      );

    this.setState(
      {
        filters: clearedFilters,
        displayedSubmissions: sortedSubmissions,
        warningMessage: undefined,
        errorMessage: undefined,
        sortedColumnKey: 'created',
        isSortedDescending: true
      },
      (): void => {
        this.updateColumnSorting();
      }
    );
  };

  private handleRefresh = (): void => {
    this.loadSubmissions()
      .catch(
        (error: Error): void => {
          this.setState({
            isLoading: false,
            errorMessage: error.message
          });
        }
      );
  };

  private handleSearchTextChange = (
    event: React.FormEvent<
      HTMLInputElement | HTMLTextAreaElement
    >,
    value?: string
  ): void => {
    this.setState({
      filters: {
        ...this.state.filters,
        searchText: value || ''
      }
    });
  };

  private handleSearchKeyDown = (
    event: React.KeyboardEvent<HTMLInputElement>
  ): void => {
    if (event.key === 'Enter') {
      event.preventDefault();
      this.applyFilters();
    }
  };

  private handleTierChange = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    const tier: SupplierTier | undefined =
      option && option.key
        ? option.key as SupplierTier
        : undefined;

    this.setState({
      filters: {
        ...this.state.filters,
        tier: tier
      }
    });
  };

  private handleQualificationChange = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    const qualification:
      QualificationStatus | undefined =
      option && option.key
        ? option.key as QualificationStatus
        : undefined;

    this.setState({
      filters: {
        ...this.state.filters,
        qualification: qualification
      }
    });
  };

  private handleMinimumPercentageChange = (
    event: React.FormEvent<
      HTMLInputElement | HTMLTextAreaElement
    >,
    value?: string
  ): void => {
    this.setState({
      filters: {
        ...this.state.filters,
        minimumPercentage:
          this.parseOptionalPercentage(value)
      }
    });
  };

  private handleMaximumPercentageChange = (
    event: React.FormEvent<
      HTMLInputElement | HTMLTextAreaElement
    >,
    value?: string
  ): void => {
    this.setState({
      filters: {
        ...this.state.filters,
        maximumPercentage:
          this.parseOptionalPercentage(value)
      }
    });
  };

  private handleStartDateChange = (
    date: Date | null | undefined
  ): void => {
    this.setState({
      filters: {
        ...this.state.filters,
        startDate: date || undefined
      }
    });
  };

  private handleEndDateChange = (
    date: Date | null | undefined
  ): void => {
    this.setState({
      filters: {
        ...this.state.filters,
        endDate: date || undefined
      }
    });
  };

  private handleColumnClick = (
    event:
      React.MouseEvent<HTMLElement> | undefined,
    column: IColumn
  ): void => {
    const isSameColumn: boolean =
      this.state.sortedColumnKey === column.key;

    const isDescending: boolean =
      isSameColumn
        ? !this.state.isSortedDescending
        : false;

    const sortedItems: ISupplierSubmission[] =
      this.sortSubmissions(
        this.state.displayedSubmissions,
        column.key,
        isDescending
      );

    this.setState(
      {
        displayedSubmissions: sortedItems,
        sortedColumnKey: column.key,
        isSortedDescending: isDescending
      },
      (): void => {
        this.updateColumnSorting();
      }
    );
  };

  private sortSubmissions(
    submissions: readonly ISupplierSubmission[],
    columnKey: string,
    descending: boolean
  ): ISupplierSubmission[] {
    const copiedItems: ISupplierSubmission[] =
      submissions.slice();

    copiedItems.sort(
      (
        first: ISupplierSubmission,
        second: ISupplierSubmission
      ): number => {
        let result: number = 0;

        if (columnKey === 'overallPercentage') {
          result =
            first.overallPercentage -
            second.overallPercentage;
        } else if (columnKey === 'created') {
          result =
            new Date(first.created).getTime() -
            new Date(second.created).getTime();
        } else {
          const firstValue: string =
            this.getSortableString(
              first,
              columnKey
            );

          const secondValue: string =
            this.getSortableString(
              second,
              columnKey
            );

          result = firstValue.localeCompare(
            secondValue
          );
        }

        return descending ? -result : result;
      }
    );

    return copiedItems;
  }

  private getSortableString(
    item: ISupplierSubmission,
    columnKey: string
  ): string {
    switch (columnKey) {
      case 'supplierName':
        return item.supplierName.toLowerCase();

      case 'tier':
        return item.tier.toLowerCase();

      case 'qualification':
        return item.qualification.toLowerCase();

      case 'riskRating':
        return item.riskRating.toLowerCase();

      case 'submittedByName':
        return item.submittedByName.toLowerCase();

      case 'recommendation':
        return item.recommendation.toLowerCase();

      default:
        return '';
    }
  }

  private updateColumnSorting(): void {
    this.columns = this.columns.map(
      (column: IColumn): IColumn => {
        return {
          ...column,
          isSorted:
            column.key ===
            this.state.sortedColumnKey,
          isSortedDescending:
            column.key ===
            this.state.sortedColumnKey
              ? this.state.isSortedDescending
              : false
        };
      }
    );

    this.forceUpdate();
  }

  private parseOptionalPercentage(
    value?: string
  ): number | undefined {
    if (
      value === undefined ||
      value.trim() === ''
    ) {
      return undefined;
    }

    const parsedValue: number =
      parseFloat(value);

    if (!isFinite(parsedValue)) {
      return undefined;
    }

    return Math.max(
      0,
      Math.min(100, parsedValue)
    );
  }

  private getStartOfDay(date: Date): Date {
    const startDate: Date = new Date(date.getTime());
    startDate.setHours(0, 0, 0, 0);

    return startDate;
  }

  private getEndOfDay(date: Date): Date {
    const endDate: Date = new Date(date.getTime());
    endDate.setHours(23, 59, 59, 999);

    return endDate;
  }

  private formatDate(dateValue: string): string {
    if (!dateValue) {
      return '-';
    }

    const date: Date = new Date(dateValue);

    if (isNaN(date.getTime())) {
      return '-';
    }

    return date.toLocaleDateString(
      this.props.context.pageContext
        .cultureInfo.currentUICultureName ||
        'en-GB',
      {
        day: '2-digit',
        month: 'short',
        year: 'numeric'
      }
    );
  }

  private getTierClassName(
    tier: SupplierTier
  ): string {
    switch (tier) {
      case 'Tier 1':
        return styles.tier1;

      case 'Tier 2':
        return styles.tier2;

      case 'Tier 3':
        return styles.tier3;

      default:
        return '';
    }
  }

  private getQualificationClassName(
    qualification: QualificationStatus
  ): string {
    switch (qualification) {
      case 'Qualified':
        return styles.successBadge;

      case 'Conditionally Qualified':
        return styles.warningBadge;

      case 'Not Qualified':
        return styles.dangerBadge;

      default:
        return '';
    }
  }

  private getRiskClassName(
    riskRating: RiskRating
  ): string {
    switch (riskRating) {
      case 'Low Risk':
        return styles.lowRisk;

      case 'Medium Risk':
        return styles.mediumRisk;

      case 'High Risk':
        return styles.highRisk;

      default:
        return '';
    }
  }

  private getRecommendationClassName(
    recommendation: Recommendation
  ): string {
    switch (recommendation) {
      case 'Approve':
        return styles.successBadge;

      case 'Corrective Action':
        return styles.warningBadge;

      case 'Requires Review':
        return styles.dangerBadge;

      default:
        return '';
    }
  }

  private openSubmission(
    item: ISupplierSubmission
  ): void {
    window.open(
      item.sourceItemUrl,
      '_blank',
      'noopener,noreferrer'
    );
  }

  private copyEmail(email: string): void {
    if (!email) {
      return;
    }

    if (
      navigator.clipboard &&
      navigator.clipboard.writeText
    ) {
      navigator.clipboard
        .writeText(email)
        .catch(
          (): void => {
            this.copyUsingTextArea(email);
          }
        );

      return;
    }

    this.copyUsingTextArea(email);
  }

  private copyUsingTextArea(value: string): void {
    const textArea: HTMLTextAreaElement =
      document.createElement('textarea');

    textArea.value = value;
    textArea.style.position = 'fixed';
    textArea.style.opacity = '0';

    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();

    document.execCommand('copy');
    document.body.removeChild(textArea);
  }

  private renderWarning():
    React.ReactElement | undefined {
    if (!this.state.warningMessage) {
      return undefined;
    }

    return (
      <MessageBar
        messageBarType={MessageBarType.warning}
        isMultiline={false}
        onDismiss={(): void => {
          this.setState({
            warningMessage: undefined
          });
        }}
      >
        {this.state.warningMessage}
      </MessageBar>
    );
  }

  private renderError(): React.ReactNode {
    if (!this.state.errorMessage) {
      return undefined;
    }

    return (
      <MessageBar
        messageBarType={MessageBarType.error}
        isMultiline={true}
        onDismiss={(): void => {
          this.setState({
            errorMessage: undefined
          });
        }}
        actions={
          <Stack horizontal>
            <DefaultButton
              text="Retry"
              onClick={this.handleRefresh}
            />
          </Stack>
        }
      >
        {this.state.errorMessage}
      </MessageBar>
    );
  }

  private onRenderDetailsHeader:
    IRenderFunction<IDetailsHeaderProps> = (
      headerProps?: IDetailsHeaderProps,
      defaultRender?: IRenderFunction<IDetailsHeaderProps>
    ): JSX.Element | null => {
      if (!headerProps || !defaultRender) {
        return null;
      }

      return defaultRender(headerProps) || null;
    };
}