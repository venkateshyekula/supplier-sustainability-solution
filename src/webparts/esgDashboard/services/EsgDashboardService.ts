import {
  WebPartContext
} from '@microsoft/sp-webpart-base';

import {
  IListConfiguration,
  SupplierTier
} from '../../supplierEsgSearch/models/IListConfiguration';

import {
  ISupplierSubmission
} from '../../supplierEsgSearch/models/ISupplierSubmission';

import {
  SupplierSubmissionService
} from '../../supplierEsgSearch/services/SupplierSubmissionService';

import {
  IActionDatum,
  IAgingDatum,
  IChartDatum,
  ITrendDatum
} from '../models/IDashboardChartData';

import {
  IDashboardInsight
} from '../models/IDashboardInsight';

import {
  IDashboardMetrics
} from '../models/IDashboardMetrics';

import {
  IDashboardSubmission
} from '../models/IDashboardSubmission';

import {
  IEsgDashboardService
} from './IEsgDashboardService';

export class EsgDashboardService
  implements IEsgDashboardService {

  private readonly submissionService:
    SupplierSubmissionService;

  public constructor(
    context: WebPartContext
  ) {
    this.submissionService =
      new SupplierSubmissionService(
        context
      );
  }

  public async getDashboard(
    configurations:
      readonly IListConfiguration[]
  ): Promise<IDashboardMetrics> {
    const rawSubmissions:
      ISupplierSubmission[] =
        await this.submissionService
          .getAllSubmissions(
            configurations
          );

    const allSubmissions:
      IDashboardSubmission[] =
        rawSubmissions.map(
          (
            submission:
              ISupplierSubmission
          ): IDashboardSubmission => {
            return this.mapSubmission(
              submission
            );
          }
        );

    const now: Date =
      new Date();

    const currentPeriodStart:
      Date =
        new Date(
          now.getFullYear(),
          now.getMonth(),
          now.getDate() - 30
        );

    const previousPeriodStart:
      Date =
        new Date(
          now.getFullYear(),
          now.getMonth(),
          now.getDate() - 60
        );

    const currentPeriodSubmissions:
      IDashboardSubmission[] =
        allSubmissions.filter(
          (
            submission:
              IDashboardSubmission
          ): boolean => {
            const submittedTime: number =
              this.getTimeValue(
                submission.submittedOn
              );

            return (
              submittedTime >=
              currentPeriodStart.getTime()
            );
          }
        );

    const previousPeriodSubmissions:
      IDashboardSubmission[] =
        allSubmissions.filter(
          (
            submission:
              IDashboardSubmission
          ): boolean => {
            const submittedTime: number =
              this.getTimeValue(
                submission.submittedOn
              );

            return (
              submittedTime >=
                previousPeriodStart.getTime() &&
              submittedTime <
                currentPeriodStart.getTime()
            );
          }
        );

    /*
     * Temporary mapping until dedicated ESG review fields
     * are available:
     *
     * Qualified               -> Approved by ESG
     * Conditionally Qualified -> Pending ESG Review
     * Not Qualified           -> Requires Action
     */
    const approvedCount: number =
      this.getQualificationCount(
        allSubmissions,
        'Qualified'
      );

    const pendingCount: number =
      this.getQualificationCount(
        allSubmissions,
        'Conditionally Qualified'
      );

    const requiresActionCount: number =
      this.getQualificationCount(
        allSubmissions,
        'Not Qualified'
      );

    const highRiskCount: number =
      allSubmissions.filter(
        (
          submission:
            IDashboardSubmission
        ): boolean => {
          return (
            submission.riskRating ===
            'High Risk'
          );
        }
      ).length;

    const averageScore: number =
      this.getAverage(
        allSubmissions.map(
          (
            submission:
              IDashboardSubmission
          ): number => {
            return submission.score;
          }
        )
      );

    const currentAverageScore: number =
      this.getAverage(
        currentPeriodSubmissions.map(
          (
            submission:
              IDashboardSubmission
          ): number => {
            return submission.score;
          }
        )
      );

    const previousAverageScore: number =
      this.getAverage(
        previousPeriodSubmissions.map(
          (
            submission:
              IDashboardSubmission
          ): number => {
            return submission.score;
          }
        )
      );

    const submissionsByTier:
      IChartDatum[] =
        this.buildCountChart(
          allSubmissions,
          [
            'Tier 1',
            'Tier 2',
            'Tier 3'
          ],
          (
            submission:
              IDashboardSubmission
          ): string => {
            return submission.tier;
          },
          [
            '#1675e0',
            '#ffb900',
            '#e84a6a'
          ]
        );

    const qualificationBreakdown:
      IChartDatum[] =
        this.buildCountChart(
          allSubmissions,
          [
            'Qualified',
            'Conditionally Qualified',
            'Not Qualified'
          ],
          (
            submission:
              IDashboardSubmission
          ): string => {
            return submission.qualification;
          },
          [
            '#36a852',
            '#f5b800',
            '#d83b32'
          ]
        );

    const riskBreakdown:
      IChartDatum[] =
        this.buildCountChart(
          allSubmissions,
          [
            'Low Risk',
            'Medium Risk',
            'High Risk'
          ],
          (
            submission:
              IDashboardSubmission
          ): string => {
            return submission.riskRating;
          },
          [
            '#36a852',
            '#f5b800',
            '#d83b32'
          ]
        );

    const recentSubmissions:
      IDashboardSubmission[] =
        allSubmissions
          .slice()
          .sort(
            (
              first:
                IDashboardSubmission,

              second:
                IDashboardSubmission
            ): number => {
              return (
                this.getTimeValue(
                  second.submittedOn
                ) -
                this.getTimeValue(
                  first.submittedOn
                )
              );
            }
          )
          .slice(
            0,
            5
          );

    const metrics:
      IDashboardMetrics = {
        total:
          allSubmissions.length,

        approved:
          approvedCount,

        pending:
          pendingCount,

        requiresAction:
          requiresActionCount,

        highRisk:
          highRiskCount,

        averageScore,

        totalChange:
          this.getPercentageChange(
            currentPeriodSubmissions.length,
            previousPeriodSubmissions.length
          ),

        approvedChange:
          this.getQualificationChange(
            currentPeriodSubmissions,
            previousPeriodSubmissions,
            'Qualified'
          ),

        pendingChange:
          this.getQualificationChange(
            currentPeriodSubmissions,
            previousPeriodSubmissions,
            'Conditionally Qualified'
          ),

        actionChange:
          this.getQualificationChange(
            currentPeriodSubmissions,
            previousPeriodSubmissions,
            'Not Qualified'
          ),

        highRiskChange:
          this.getHighRiskChange(
            currentPeriodSubmissions,
            previousPeriodSubmissions
          ),

        scoreChange:
          Math.round(
            currentAverageScore -
            previousAverageScore
          ),

        trend:
          this.buildSubmissionTrend(
            allSubmissions
          ),

        byTier:
          submissionsByTier,

        qualification:
          qualificationBreakdown,

        risks:
          riskBreakdown,

        aging:
          this.buildPendingReviewAging(
            allSubmissions
          ),

        actions:
          this.buildPendingActions(
            allSubmissions
          ),

        averageByTier:
          this.buildAverageScoreByTier(
            allSubmissions
          ),

        recent:
          recentSubmissions,

        insights: []
      };

    metrics.insights =
      this.buildInsights(
        metrics
      );

    return metrics;
  }

  private mapSubmission(
    submission:
      ISupplierSubmission
  ): IDashboardSubmission {
    return {
      key:
        submission.key,

      id:
        submission.id,

      supplierName:
        submission.supplierName,

      tier:
        submission.tier,

      score:
        this.normalizeScore(
          submission.overallPercentage,
          submission.tier
        ),

      qualification:
        submission.qualification,

      riskRating:
        submission.riskRating,

      recommendation:
        submission.recommendation,

      submittedOn:
        submission.created,

      submittedBy:
        submission.submittedByName,

      sourceItemUrl:
        submission.sourceItemUrl
    };
  }

  private normalizeScore(
    value: number,
    tier: SupplierTier
  ): number {
    const safeValue: number =
      isFinite(value)
        ? value
        : 0;

    /*
     * Values above 10 are treated as percentages.
     */
    if (safeValue > 10) {
      return this.clamp(
        safeValue,
        0,
        100
      );
    }

    let maximumScore: number;

    switch (tier) {
      case 'Tier 1':
        maximumScore = 10;
        break;

      case 'Tier 2':
        maximumScore = 5;
        break;

      case 'Tier 3':
        maximumScore = 2;
        break;

      default:
        maximumScore = 1;
        break;
    }

    const normalizedValue: number =
      (
        safeValue /
        maximumScore
      ) *
      100;

    return (
      Math.round(
        this.clamp(
          normalizedValue,
          0,
          100
        ) *
        100
      ) /
      100
    );
  }

  private clamp(
    value: number,
    minimum: number,
    maximum: number
  ): number {
    return Math.max(
      minimum,
      Math.min(
        maximum,
        value
      )
    );
  }

  private getAverage(
    values:
      readonly number[]
  ): number {
    if (values.length === 0) {
      return 0;
    }

    const total: number =
      values.reduce(
        (
          runningTotal: number,
          value: number
        ): number => {
          return (
            runningTotal +
            value
          );
        },
        0
      );

    return Math.round(
      total /
      values.length
    );
  }

  private getPercentageChange(
    currentValue: number,
    previousValue: number
  ): number {
    if (previousValue === 0) {
      return currentValue > 0
        ? 100
        : 0;
    }

    return Math.round(
      (
        (
          currentValue -
          previousValue
        ) /
        previousValue
      ) *
      100
    );
  }

  private getQualificationCount(
    submissions:
      readonly IDashboardSubmission[],
    qualification: string
  ): number {
    return submissions.filter(
      (
        submission:
          IDashboardSubmission
      ): boolean => {
        return (
          submission.qualification ===
          qualification
        );
      }
    ).length;
  }

  private getQualificationChange(
    currentSubmissions:
      readonly IDashboardSubmission[],

    previousSubmissions:
      readonly IDashboardSubmission[],

    qualification: string
  ): number {
    const currentCount: number =
      this.getQualificationCount(
        currentSubmissions,
        qualification
      );

    const previousCount: number =
      this.getQualificationCount(
        previousSubmissions,
        qualification
      );

    return this.getPercentageChange(
      currentCount,
      previousCount
    );
  }

  private getHighRiskChange(
    currentSubmissions:
      readonly IDashboardSubmission[],

    previousSubmissions:
      readonly IDashboardSubmission[]
  ): number {
    const currentCount: number =
      currentSubmissions.filter(
        (
          submission:
            IDashboardSubmission
        ): boolean => {
          return (
            submission.riskRating ===
            'High Risk'
          );
        }
      ).length;

    const previousCount: number =
      previousSubmissions.filter(
        (
          submission:
            IDashboardSubmission
        ): boolean => {
          return (
            submission.riskRating ===
            'High Risk'
          );
        }
      ).length;

    return this.getPercentageChange(
      currentCount,
      previousCount
    );
  }

  private buildCountChart(
    submissions:
      readonly IDashboardSubmission[],

    labels:
      readonly string[],

    getValue: (
      submission:
        IDashboardSubmission
    ) => string,

    colors:
      readonly string[]
  ): IChartDatum[] {
    return labels.map(
      (
        label: string,
        index: number
      ): IChartDatum => {
        const value: number =
          submissions.filter(
            (
              submission:
                IDashboardSubmission
            ): boolean => {
              return (
                getValue(
                  submission
                ) === label
              );
            }
          ).length;

        return {
          label,
          value,
          color:
            colors[index] ||
            '#64748b'
        };
      }
    );
  }

  private buildSubmissionTrend(
    submissions:
      readonly IDashboardSubmission[]
  ): ITrendDatum[] {
    const trend:
      ITrendDatum[] = [];

    const now: Date =
      new Date();

    for (
      let monthOffset: number = 5;
      monthOffset >= 0;
      monthOffset -= 1
    ) {
      const monthDate: Date =
        new Date(
          now.getFullYear(),
          now.getMonth() -
            monthOffset,
          1
        );

      const year: number =
        monthDate.getFullYear();

      const month: number =
        monthDate.getMonth();

      const monthCount: number =
        submissions.filter(
          (
            submission:
              IDashboardSubmission
          ): boolean => {
            const submittedDate:
              Date =
                new Date(
                  submission.submittedOn
                );

            return (
              !isNaN(
                submittedDate.getTime()
              ) &&
              submittedDate.getFullYear() ===
                year &&
              submittedDate.getMonth() ===
                month
            );
          }
        ).length;

      trend.push({
        label:
          monthDate
            .toLocaleDateString(
              'en-US',
              {
                month: 'short',
                year: 'numeric'
              }
            ),

        value:
          monthCount
      });
    }

    return trend;
  }

  private buildPendingReviewAging(
    submissions:
      readonly IDashboardSubmission[]
  ): IAgingDatum[] {
    const pendingSubmissions:
      IDashboardSubmission[] =
        submissions.filter(
          (
            submission:
              IDashboardSubmission
          ): boolean => {
            return (
              submission.qualification ===
              'Conditionally Qualified'
            );
          }
        );

    const getAgeInDays = (
      submission:
        IDashboardSubmission
    ): number => {
      const submissionTime: number =
        this.getTimeValue(
          submission.submittedOn
        );

      if (submissionTime <= 0) {
        return 0;
      }

      return Math.max(
        0,
        Math.floor(
          (
            Date.now() -
            submissionTime
          ) /
          86400000
        )
      );
    };

    return [
      {
        label: '0 - 7 days',
        value:
          pendingSubmissions.filter(
            (
              submission:
                IDashboardSubmission
            ): boolean => {
              return (
                getAgeInDays(
                  submission
                ) <= 7
              );
            }
          ).length,
        color: '#44a85b'
      },
      {
        label: '8 - 15 days',
        value:
          pendingSubmissions.filter(
            (
              submission:
                IDashboardSubmission
            ): boolean => {
              const age: number =
                getAgeInDays(
                  submission
                );

              return (
                age > 7 &&
                age <= 15
              );
            }
          ).length,
        color: '#f5b800'
      },
      {
        label: '16 - 30 days',
        value:
          pendingSubmissions.filter(
            (
              submission:
                IDashboardSubmission
            ): boolean => {
              const age: number =
                getAgeInDays(
                  submission
                );

              return (
                age > 15 &&
                age <= 30
              );
            }
          ).length,
        color: '#ff8c00'
      },
      {
        label: '31+ days',
        value:
          pendingSubmissions.filter(
            (
              submission:
                IDashboardSubmission
            ): boolean => {
              return (
                getAgeInDays(
                  submission
                ) > 30
              );
            }
          ).length,
        color: '#d83b32'
      }
    ];
  }

  private buildPendingActions(
    submissions:
      readonly IDashboardSubmission[]
  ): IActionDatum[] {
    const actionConfigurations:
      readonly {
        label: string;
        recommendationValue: string;
        icon: string;
        color: string;
      }[] = [
        {
          label:
            'Provide Additional Documentation',
          recommendationValue:
            'Request Additional Documentation',
          icon:
            'TextDocument',
          color:
            '#1675e0'
        },
        {
          label:
            'Address Corrective Action',
          recommendationValue:
            'Corrective Action',
          icon:
            'Warning',
          color:
            '#d83b32'
        },
        {
          label:
            'Reject Supplier',
          recommendationValue:
            'Reject',
          icon:
            'Blocked',
          color:
            '#8a5d00'
        }
      ];

    return actionConfigurations
      .map(
        (
          configuration
        ): IActionDatum => {
          const value: number =
            submissions.filter(
              (
                submission:
                  IDashboardSubmission
              ): boolean => {
                return (
                  String(
                    submission.recommendation
                  ) ===
                  configuration
                    .recommendationValue
                );
              }
            ).length;

          return {
            label:
              configuration.label,

            value,

            icon:
              configuration.icon,

            color:
              configuration.color
          };
        }
      )
      .filter(
        (
          action:
            IActionDatum
        ): boolean => {
          return action.value > 0;
        }
      );
  }

  private buildAverageScoreByTier(
    submissions:
      readonly IDashboardSubmission[]
  ): IChartDatum[] {
    const tiers:
      readonly SupplierTier[] = [
        'Tier 1',
        'Tier 2',
        'Tier 3'
      ];

    const colors:
      readonly string[] = [
        '#1675e0',
        '#ffb900',
        '#e84a6a'
      ];

    return tiers.map(
      (
        tier: SupplierTier,
        index: number
      ): IChartDatum => {
        const tierScores: number[] =
          submissions
            .filter(
              (
                submission:
                  IDashboardSubmission
              ): boolean => {
                return (
                  submission.tier ===
                  tier
                );
              }
            )
            .map(
              (
                submission:
                  IDashboardSubmission
              ): number => {
                return submission.score;
              }
            );

        return {
          label:
            tier,

          value:
            this.getAverage(
              tierScores
            ),

          color:
            colors[index]
        };
      }
    );
  }

  private buildInsights(
    metrics:
      IDashboardMetrics
  ): IDashboardInsight[] {
    const scoreInsight:
      IDashboardInsight = {
        key:
          'score',

        title:
          metrics.scoreChange >= 0
            ? 'ESG performance is improving'
            : 'ESG performance needs attention',

        description:
          `Overall average score changed by ` +
          `${Math.abs(
            metrics.scoreChange
          ).toString()}% in the last 30 days.`,

        icon:
          'Chart',

        tone:
          metrics.scoreChange >= 0
            ? 'positive'
            : 'warning'
      };

    const riskInsight:
      IDashboardInsight = {
        key:
          'risk',

        title:
          'Supplier risk overview',

        description:
          `${metrics.highRisk.toString()} ` +
          `suppliers are currently rated High Risk.`,

        icon:
          'Shield',

        tone:
          metrics.highRisk > 0
            ? 'warning'
            : 'positive'
      };

    const pendingInsight:
      IDashboardInsight = {
        key:
          'pending',

        title:
          'Focus on pending reviews',

        description:
          `${metrics.pending.toString()} ` +
          `submissions are awaiting ESG review.`,

        icon:
          'Clock',

        tone:
          metrics.pending > 0
            ? 'warning'
            : 'positive'
      };

    const actionInsight:
      IDashboardInsight = {
        key:
          'action',

        title:
          `${metrics.requiresAction.toString()} ` +
          `require immediate action`,

        description:
          'Suppliers may need procurement follow-up.',

        icon:
          'Warning',

        tone:
          metrics.requiresAction > 0
            ? 'danger'
            : 'positive'
      };

    return [
      scoreInsight,
      riskInsight,
      pendingInsight,
      actionInsight
    ];
  }

  private getTimeValue(
    value: string
  ): number {
    if (!value) {
      return 0;
    }

    const timestamp: number =
      new Date(
        value
      ).getTime();

    return isNaN(timestamp)
      ? 0
      : timestamp;
  }
}