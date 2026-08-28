import * as React from 'react';

import {
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize
} from '@fluentui/react';

import {
  IListConfiguration
} from '../../supplierEsgSearch/models/IListConfiguration';

import {
  IDashboardMetrics
} from '../models/IDashboardMetrics';

import {
  EsgDashboardService
} from '../services/EsgDashboardService';

import AverageScoreByTier
  from './AverageScoreByTier';

import DashboardKpiCards
  from './DashboardKpiCards';

import {
  IEsgDashboardProps
} from './IEsgDashboardProps';

import OverallEsgScore
  from './OverallEsgScore';

import PendingReviewAging
  from './PendingReviewAging';

import QualificationBreakdown
  from './QualificationBreakdown';

import QuickInsights
  from './QuickInsights';

import RecentSubmissionsSnapshot
  from './RecentSubmissionsSnapshot';

import RiskRatingAnalysis
  from './RiskRatingAnalysis';

import SubmissionsByTier
  from './SubmissionsByTier';

import SubmissionsTrend
  from './SubmissionsTrend';

import TopPendingActions
  from './TopPendingActions';

import styles from './EsgDashboard.module.scss';

interface IEsgDashboardState {
  metrics: IDashboardMetrics;
  loading: boolean;
  error?: string;
}

const EMPTY_DASHBOARD_METRICS:
  IDashboardMetrics = {
    total: 0,
    approved: 0,
    pending: 0,
    requiresAction: 0,
    highRisk: 0,

    averageScore: 0,

    totalChange: 0,
    approvedChange: 0,
    pendingChange: 0,
    actionChange: 0,
    highRiskChange: 0,
    scoreChange: 0,

    trend: [],
    byTier: [],
    qualification: [],
    risks: [],
    aging: [],
    actions: [],
    averageByTier: [],
    recent: [],
    insights: []
  };

export default class EsgDashboard
  extends React.Component<
    IEsgDashboardProps,
    IEsgDashboardState
  > {

  private readonly service:
    EsgDashboardService;

  public constructor(
    props: IEsgDashboardProps
  ) {
    super(props);

    this.service =
      new EsgDashboardService(
        props.context
      );

    this.state = {
      metrics:
        EMPTY_DASHBOARD_METRICS,

      loading:
        true,

      error:
        undefined
    };
  }

  public componentDidMount(): void {
    this.loadDashboard()
      .catch(
        (
          error: unknown
        ): void => {
          this.handleLoadingError(
            error
          );
        }
      );
  }

  public componentDidUpdate(
    previousProps:
      IEsgDashboardProps
  ): void {
    if (
      this.haveDataSourcePropertiesChanged(
        previousProps
      )
    ) {
      this.loadDashboard()
        .catch(
          (
            error: unknown
          ): void => {
            this.handleLoadingError(
              error
            );
          }
        );
    }
  }

  public render():
    React.ReactElement<IEsgDashboardProps> {
    const metrics:
      IDashboardMetrics =
        this.state.metrics;

    return (
      <section
        className={styles.dashboard}
        aria-label={
          'Supplier sustainability ESG dashboard'
        }
        aria-busy={
          this.state.loading
        }
      >
        {this.state.error && (
          <MessageBar
            messageBarType={
              MessageBarType.error
            }
            isMultiline={true}
            onDismiss={
              this.dismissError
            }
          >
            {this.state.error}
          </MessageBar>
        )}

        {this.state.loading && (
          <div
            className={
              styles.loadingContainer
            }
            role="status"
            aria-live="polite"
          >
            <Spinner
              size={SpinnerSize.large}
              label="Loading ESG dashboard"
            />
          </div>
        )}

        <div className={styles.kpiRow}>
          <div className={styles.kpiCards}>
            <DashboardKpiCards
              metrics={metrics}
              loading={
                this.state.loading
              }
            />
          </div>

          <div
            className={
              styles.overallScore
            }
          >
            <OverallEsgScore
              score={
                metrics.averageScore
              }
              change={
                metrics.scoreChange
              }
              target={
                this.props.targetScore
              }
            />
          </div>
        </div>

        <div className={styles.analyticsRow}>
          <SubmissionsTrend
            items={metrics.trend}
          />

          <SubmissionsByTier
            items={metrics.byTier}
          />

          <QualificationBreakdown
            items={
              metrics.qualification
            }
          />
        </div>

        <div className={styles.analysisRow}>
          <RiskRatingAnalysis
            items={metrics.risks}
          />

          <PendingReviewAging
            items={metrics.aging}
          />

          <TopPendingActions
            items={metrics.actions}
          />

          <AverageScoreByTier
            items={
              metrics.averageByTier
            }
          />
        </div>

        <div className={styles.bottomRow}>
          <RecentSubmissionsSnapshot
            items={metrics.recent}
            onOpen={
              this.openRecentSubmission
            }
          />

          <QuickInsights
            items={metrics.insights}
          />
        </div>
      </section>
    );
  }

  /**
   * Creates the SharePoint list configuration used by
   * EsgDashboardService.
   *
   * OverallQuestionsPercentage is read directly from the
   * configured SharePoint list.
   */
  private getConfigurations():
    readonly IListConfiguration[] {
    return [
      {
        listTitle:
          this.props.tier1ListTitle,

        tier:
          'Tier 1',

        supplierNameDisplayName:
          'Supplier Name',

        supplierNameInternalName:
          this.props
            .tier1SupplierNameField,

        emailInternalName:
          'field_3',

        contactNameInternalName:
          'field_4',

        overallPercentageInternalName:
          'OverallQuestionsPercentage',

        qualifiedWeightedMinimum:
          10,

        conditionalWeightedMinimum:
          9.99
      },
      {
        listTitle:
          this.props.tier2ListTitle,

        tier:
          'Tier 2',

        supplierNameDisplayName:
          'Supplier Name',

        supplierNameInternalName:
          this.props
            .tier2SupplierNameField,

        emailInternalName:
          'Email',

        contactNameInternalName:
          'Name',

        overallPercentageInternalName:
          'OverallQuestionsPercentage',

        qualifiedWeightedMinimum:
          5,

        conditionalWeightedMinimum:
          4.99
      },
      {
        listTitle:
          this.props.tier3ListTitle,

        tier:
          'Tier 3',

        supplierNameDisplayName:
          'Supplier Name',

        supplierNameInternalName:
          this.props
            .tier3SupplierNameField,

        emailInternalName:
          'Email',

        contactNameInternalName:
          'Name',

        overallPercentageInternalName:
          'OverallQuestionsPercentage',

        qualifiedWeightedMinimum:
          2,

        conditionalWeightedMinimum:
          1.99
      }
    ];
  }

  private loadDashboard =
    async (): Promise<void> => {
      this.setState({
        loading:
          true,

        error:
          undefined
      });

      try {
        const metrics:
          IDashboardMetrics =
            await this.service
              .getDashboard(
                this.getConfigurations()
              );

        this.setState({
          metrics,
          loading:
            false,
          error:
            undefined
        });
      } catch (
        error:
          unknown
      ) {
        this.handleLoadingError(
          error
        );
      }
    };

  private handleLoadingError(
    error: unknown
  ): void {
    const errorMessage: string =
      error instanceof Error
        ? error.message
        : 'Unable to load the ESG dashboard.';

    this.setState({
      loading:
        false,

      error:
        errorMessage
    });
  }

  private dismissError =
    (): void => {
      this.setState({
        error:
          undefined
      });
    };

  private openRecentSubmission =
    (
      item:
        IDashboardMetrics['recent'][number]
    ): void => {
      if (
        !item.sourceItemUrl
      ) {
        return;
      }

      window.open(
        item.sourceItemUrl,
        '_blank',
        'noopener,noreferrer'
      );
    };

  /**
   * Compares only the properties that affect the data source.
   *
   * The WebPartContext object is intentionally excluded from
   * JSON serialization and deep comparison.
   */
  private haveDataSourcePropertiesChanged(
    previousProps:
      IEsgDashboardProps
  ): boolean {
    return (
      previousProps
        .tier1ListTitle !==
        this.props
          .tier1ListTitle ||

      previousProps
        .tier2ListTitle !==
        this.props
          .tier2ListTitle ||

      previousProps
        .tier3ListTitle !==
        this.props
          .tier3ListTitle ||

      previousProps
        .tier1SupplierNameField !==
        this.props
          .tier1SupplierNameField ||

      previousProps
        .tier2SupplierNameField !==
        this.props
          .tier2SupplierNameField ||

      previousProps
        .tier3SupplierNameField !==
        this.props
          .tier3SupplierNameField ||

      previousProps
        .targetScore !==
        this.props
          .targetScore
    );
  }
}