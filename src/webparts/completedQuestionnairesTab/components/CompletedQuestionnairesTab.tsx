import * as React from 'react';
import { Icon, MessageBar, MessageBarType } from '@fluentui/react';
import KpiMetricCards from '../../kpiMetricCards/components/KpiMetricCards';
import { IListConfiguration } from '../../supplierEsgSearch/models/IListConfiguration';
import { ISupplierSearchFilters } from '../../supplierEsgSearch/models/ISupplierSearchFilters';
import { ISupplierSubmission } from '../../supplierEsgSearch/models/ISupplierSubmission';
import { ISupplierSubmissionService } from '../../supplierEsgSearch/services/ISupplierSubmissionService';
import { SupplierSubmissionService } from '../../supplierEsgSearch/services/SupplierSubmissionService';
import { getSP } from '../../../pnpConfig';
import { EsgReviewQueueService } from '../services/EsgReviewQueueService';
import { IEsgReviewQueueService } from '../services/IEsgReviewQueueService';
import { IReviewQueueItem } from '../models/IReviewQueueItem';
import CompletedQuestionnaireFilters from './CompletedQuestionnaireFilters';
import EsgReviewQueue from './EsgReviewQueue';
import { ICompletedQuestionnairesTabProps } from './ICompletedQuestionnairesTabProps';
import { ITierListUrls } from './IRecentSubmissionsTableProps';
import RecentSubmissionsTable from './RecentSubmissionsTable';
import styles from './CompletedQuestionnairesTab.module.scss';

interface IState {
  allSubmissions: ISupplierSubmission[];
  filteredSubmissions: ISupplierSubmission[];
  reviewQueueItems: IReviewQueueItem[];
  filters: ISupplierSearchFilters;
  tierListUrls: ITierListUrls;
  recentItemsLimit: number;
  isLoading: boolean;
  errorMessage?: string;
}

interface IListMetadata {
  RootFolder?: { ServerRelativeUrl?: string };
}

export default class CompletedQuestionnairesTab extends React.Component<ICompletedQuestionnairesTabProps, IState> {
  private readonly submissionService: ISupplierSubmissionService;
  private readonly queueService: IEsgReviewQueueService;

  public constructor(props: ICompletedQuestionnairesTabProps) {
    super(props);
    this.submissionService = new SupplierSubmissionService(props.context);
    this.queueService = new EsgReviewQueueService();
    this.state = {
      allSubmissions: [],
      filteredSubmissions: [],
      reviewQueueItems: [],
      filters: { searchText: '' },
      tierListUrls: {},
      recentItemsLimit: 10,
      isLoading: true
    };
  }

  public componentDidMount(): void {
    this.loadData().catch((error: Error): void => {
      this.setState({ isLoading: false, errorMessage: error.message });
    });
  }

  public render(): React.ReactElement<ICompletedQuestionnairesTabProps> {
    return (
      <section className={styles.page} aria-label="Completed questionnaires">
        <CompletedQuestionnaireFilters
          filters={this.state.filters}
          isLoading={this.state.isLoading}
          onFiltersChange={this.handleFiltersChange}
          onSearch={this.applyFilters}
          onClear={this.clearFilters}
        />

        {this.state.errorMessage && (
          <MessageBar messageBarType={MessageBarType.error}>{this.state.errorMessage}</MessageBar>
        )}

        <section className={styles.kpis} aria-label="Submission metrics">
          <KpiMetricCards
            context={this.props.context}
            tier1ListTitle={this.props.tier1ListTitle}
            tier2ListTitle={this.props.tier2ListTitle}
            tier3ListTitle={this.props.tier3ListTitle}
          />
        </section>

        <div className={styles.contentGrid}>
          <RecentSubmissionsTable
            submissions={this.state.filteredSubmissions}
            isLoading={this.state.isLoading}
            recentItemsLimit={this.state.recentItemsLimit}
            tierListUrls={this.state.tierListUrls}
            onOpenSubmission={this.openSubmission}
            onOpenTierList={this.openUrl}
          />
          <EsgReviewQueue
            items={this.state.reviewQueueItems}
            isLoading={this.state.isLoading}
            errorMessage={this.state.errorMessage}
            onOpenItem={(item: IReviewQueueItem): void => this.openUrl(item.sourceItemUrl)}
          />
        </div>

        <footer className={styles.notice} role="note">
          <span className={styles.noticeIcon} aria-hidden="true"><Icon iconName="Info" /></span>
          <p><strong>Suppliers:</strong>{' '}Suppliers will receive direct links from Procurement to complete the appropriate Tier questionnaires. Please do not use this site to submit questionnaires.</p>
        </footer>
      </section>
    );
  }

  private getConfigurations(): readonly IListConfiguration[] {
    return [
      {
        listTitle: this.props.tier1ListTitle,
        tier: 'Tier 1',
        supplierNameDisplayName: 'Supplier Name',
        supplierNameInternalName: this.props.tier1SupplierNameField,
        emailInternalName: 'field_3',
        contactNameInternalName: 'field_4',
        overallPercentageInternalName: 'OverallQuestionsPercentage',
        qualifiedWeightedMinimum: 10,
        conditionalWeightedMinimum: 9.99
      },
      {
        listTitle: this.props.tier2ListTitle,
        tier: 'Tier 2',
        supplierNameDisplayName: 'Supplier Name',
        supplierNameInternalName: this.props.tier2SupplierNameField,
        emailInternalName: 'Email',
        contactNameInternalName: 'Name',
        overallPercentageInternalName: 'OverallQuestionsPercentage',
        qualifiedWeightedMinimum: 5,
        conditionalWeightedMinimum: 4.99
      },
      {
        listTitle: this.props.tier3ListTitle,
        tier: 'Tier 3',
        supplierNameDisplayName: 'Supplier Name',
        supplierNameInternalName: this.props.tier3SupplierNameField,
        emailInternalName: 'Email',
        contactNameInternalName: 'Name',
        overallPercentageInternalName: 'OverallQuestionsPercentage',
        qualifiedWeightedMinimum: 2,
        conditionalWeightedMinimum: 1.99
      }
    ];
  }

  private loadData = async (): Promise<void> => {
    this.setState({ isLoading: true, errorMessage: undefined });
    const results: [ISupplierSubmission[], ITierListUrls] = await Promise.all([
      this.submissionService.getAllSubmissions(this.getConfigurations()),
      this.loadTierListUrls()
    ]);
    const submissions: ISupplierSubmission[] = results[0];
    this.setState({
      allSubmissions: submissions,
      filteredSubmissions: submissions,
      reviewQueueItems: this.queueService.buildQueue(submissions),
      tierListUrls: results[1],
      isLoading: false
    });
  };

  private handleFiltersChange = (filters: ISupplierSearchFilters): void => {
    this.setState({ filters });
  };

  private applyFilters = (): void => {
    const f: ISupplierSearchFilters = this.state.filters;
    const term: string = f.searchText.trim().toLowerCase();
    const start: number | undefined = f.startDate ? this.startOfDay(f.startDate).getTime() : undefined;
    const end: number | undefined = f.endDate ? this.endOfDay(f.endDate).getTime() : undefined;
    const filtered: ISupplierSubmission[] = this.state.allSubmissions.filter((item: ISupplierSubmission): boolean => {
      const created: number = new Date(item.created).getTime();
      return (
        (!term || item.supplierName.toLowerCase().indexOf(term) >= 0 || item.email.toLowerCase().indexOf(term) >= 0 || item.submittedByName.toLowerCase().indexOf(term) >= 0) &&
        (!f.tier || item.tier === f.tier) &&
        (!f.qualification || item.qualification === f.qualification) &&
        (start === undefined || created >= start) &&
        (end === undefined || created <= end)
      );
    });
    this.setState({ filteredSubmissions: filtered });
  };

  private clearFilters = (): void => {
    this.setState({ filters: { searchText: '' }, filteredSubmissions: this.state.allSubmissions });
  };

  private async loadTierListUrls(): Promise<ITierListUrls> {
    const urls: Array<string | undefined> = await Promise.all([
      this.resolveListUrl(this.props.tier1ListTitle),
      this.resolveListUrl(this.props.tier2ListTitle),
      this.resolveListUrl(this.props.tier3ListTitle)
    ]);
    return { tier1: urls[0], tier2: urls[1], tier3: urls[2] };
  }

  private async resolveListUrl(title: string): Promise<string | undefined> {
    try {
      const metadata: IListMetadata = await getSP(this.props.context).web.lists
        .getByTitle(title).select('RootFolder/ServerRelativeUrl').expand('RootFolder')() as IListMetadata;
      const root: string = metadata.RootFolder?.ServerRelativeUrl || '';
      return root ? `${root.replace(/\/$/, '')}/AllItems.aspx` : undefined;
    } catch (error: unknown) {
      console.error(`Unable to resolve list URL for "${title}".`, error);
      return undefined;
    }
  }

  private openSubmission = (item: ISupplierSubmission): void => this.openUrl(item.sourceItemUrl);
  private openUrl = (url?: string): void => { if (url) window.open(url, '_blank', 'noopener,noreferrer'); };
  private startOfDay(date: Date): Date { const value: Date = new Date(date.getTime()); value.setHours(0, 0, 0, 0); return value; }
  private endOfDay(date: Date): Date { const value: Date = new Date(date.getTime()); value.setHours(23, 59, 59, 999); return value; }
}