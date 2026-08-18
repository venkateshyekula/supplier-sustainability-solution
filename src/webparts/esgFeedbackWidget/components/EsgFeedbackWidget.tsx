import * as React from 'react';
import { useEffect, useState } from 'react';

import {
  DefaultButton,
  IButtonStyles,
  Icon,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize
} from '@fluentui/react';

import { getSP } from '../../../pnpConfig';
import { IEsgFeedbackWidgetProps } from './IEsgFeedbackWidgetProps';

type SupplierTier = 'Tier 1' | 'Tier 2' | 'Tier 3';
type QualificationStatus =
  | 'Qualified'
  | 'Conditionally Qualified'
  | 'Not Qualified';
type RiskRating = 'Low Risk' | 'Medium Risk' | 'High Risk';
type Recommendation =
  | 'Approve'
  | 'Corrective Action'
  | 'Requires Review';

interface ITierConfiguration {
  tier: SupplierTier;
  listTitle: string;
  supplierNameInternalName: string;
  emailInternalName: string;
  contactNameInternalName: string;
  environmentalCountInternalName: string;
  socialCountInternalName: string;
  governanceCountInternalName: string;
  overallPercentageInternalName: string;
  maximumWeightedScore: number;
}

interface ISharePointUserValue {
  Title?: string;
  EMail?: string;
}

interface IRawSubmission {
  Id?: number;
  Created?: string;
  Modified?: string;
  Author?: ISharePointUserValue;
  Editor?: ISharePointUserValue;
  [key: string]: unknown;
}

interface ICategoryScore {
  name: string;
  value: number;
  iconName: string;
  color: string;
}

interface ILatestSubmission {
  id: number;
  tier: SupplierTier;
  listTitle: string;
  supplierName: string;
  supplierEmail: string;
  contactName: string;
  sharePointWeightedScore: number;
  maximumWeightedScore: number;
  overallScore: number;
  categoryScores: ICategoryScore[];
  qualification: QualificationStatus;
  riskRating: RiskRating;
  recommendation: Recommendation;
  lastUpdatedBy: string;
  lastUpdated: string;
  assessmentNote: string;
  created: string;
  itemUrl: string;
}

interface ITierFetchResult {
  configuration: ITierConfiguration;
  item?: IRawSubmission;
}

const QUALIFIED_MINIMUM = 70;
const CONDITIONAL_MINIMUM = 60;

const DEFAULT_TIER_CONFIGURATIONS: readonly ITierConfiguration[] = [
  {
    tier: 'Tier 1',
    listTitle: 'CASSTECH_SSQ',
    supplierNameInternalName: 'Supplier_x0020_Name',
    emailInternalName: 'field_3',
    contactNameInternalName: 'field_4',
    environmentalCountInternalName:
      'EnvironmentalQuestionsCount',
    socialCountInternalName: 'SocialQuestionsCount',
    governanceCountInternalName: 'GovernanceQuestionsCount',
    overallPercentageInternalName:
      'OverallQuestionsPercentage',
    maximumWeightedScore: 10
  },
  {
    tier: 'Tier 2',
    listTitle: 'Tier 2 ESG Procurement Questionnaire',
    supplierNameInternalName: 'Supplier_x0020_Name',
    emailInternalName: 'Email',
    contactNameInternalName: 'Name',
    environmentalCountInternalName: 'EnvQuestionsCount',
    socialCountInternalName: 'SocialQuestionsCount',
    governanceCountInternalName: 'GovernanceQuestionsCount',
    overallPercentageInternalName:
      'OverallQuestionsPercentage',
    maximumWeightedScore: 5
  },
  {
    tier: 'Tier 3',
    listTitle:
      'Supplier Sustainability Questionnaires Tier 3',
    supplierNameInternalName: 'Supplier_x0020_Name',
    emailInternalName: 'Email',
    contactNameInternalName: 'Name',
    environmentalCountInternalName: 'EnvQuestionsCount',
    socialCountInternalName: 'SocialQuestionsCount',
    governanceCountInternalName: 'GovernanceQuestionsCount',
    overallPercentageInternalName:
      'OverallQuestionsPercentage',
    maximumWeightedScore: 2
  }
];

const parseNumber = (value: unknown): number => {
  if (value === null || value === undefined || value === '') {
    return 0;
  }

  const normalizedValue: unknown =
    typeof value === 'string'
      ? value.replace('%', '').replace(',', '.').trim()
      : value;

  const numericValue: number =
    typeof normalizedValue === 'number'
      ? normalizedValue
      : Number(normalizedValue);

  if (
    isNaN(numericValue) ||
    !isFinite(numericValue) ||
    numericValue < 0
  ) {
    return 0;
  }

  return numericValue;
};

const getTextValue = (value: unknown): string => {
  if (typeof value === 'string') {
    return value.trim();
  }

  if (typeof value === 'number') {
    return value.toString();
  }

  return '';
};

const roundToTwoDecimals = (value: number): number => {
  return Math.round(value * 100) / 100;
};

const calculateRelativeScore = (
  sharePointWeightedScore: number,
  maximumWeightedScore: number
): number => {
  if (
    !isFinite(maximumWeightedScore) ||
    maximumWeightedScore <= 0
  ) {
    return 0;
  }

  const relativeScore: number =
    (sharePointWeightedScore / maximumWeightedScore) * 100;

  return roundToTwoDecimals(
    Math.min(Math.max(relativeScore, 0), 100)
  );
};

const evaluateQualification = (
  score: number
): {
  qualification: QualificationStatus;
  riskRating: RiskRating;
  recommendation: Recommendation;
} => {
  if (score >= QUALIFIED_MINIMUM) {
    return {
      qualification: 'Qualified',
      riskRating: 'Low Risk',
      recommendation: 'Approve'
    };
  }

  if (score >= CONDITIONAL_MINIMUM) {
    return {
      qualification: 'Conditionally Qualified',
      riskRating: 'Medium Risk',
      recommendation: 'Corrective Action'
    };
  }

  return {
    qualification: 'Not Qualified',
    riskRating: 'High Risk',
    recommendation: 'Requires Review'
  };
};

const generateAssessmentNote = (
  qualification: QualificationStatus,
  tier: SupplierTier,
  weightedScore: number,
  maximumWeightedScore: number,
  relativeScore: number
): string => {
  const scoreText: string =
    `${roundToTwoDecimals(weightedScore)} of ` +
    `${maximumWeightedScore} (${roundToTwoDecimals(relativeScore)}%)`;

  switch (qualification) {
    case 'Qualified':
      return (
        `System-generated assessment for ${tier}: ` +
        `the SharePoint weighted score is ${scoreText} and ` +
        'meets the configured qualification threshold.'
      );

    case 'Conditionally Qualified':
      return (
        `System-generated assessment for ${tier}: ` +
        `the SharePoint weighted score is ${scoreText} and ` +
        'meets the conditional threshold. Corrective action is recommended.'
      );

    default:
      return (
        `System-generated assessment for ${tier}: ` +
        `the SharePoint weighted score is ${scoreText} and ` +
        'is below the conditional threshold. Additional review is required.'
      );
  }
};

const formatDate = (
  value: string,
  localeName: string
): string => {
  if (!value) {
    return 'Not available';
  }

  const date: Date = new Date(value);

  if (isNaN(date.getTime())) {
    return 'Not available';
  }

  return date.toLocaleDateString(localeName || 'en-GB', {
    day: '2-digit',
    month: 'short',
    year: 'numeric'
  });
};

const getBadgeColors = (
  value: QualificationStatus | RiskRating | Recommendation
): { background: string; color: string } => {
  switch (value) {
    case 'Qualified':
    case 'Low Risk':
    case 'Approve':
      return {
        background: '#dff6dd',
        color: '#107c10'
      };

    case 'Conditionally Qualified':
    case 'Medium Risk':
    case 'Corrective Action':
      return {
        background: '#fff4ce',
        color: '#835c00'
      };

    default:
      return {
        background: '#fde7e9',
        color: '#a4262c'
      };
  }
};

const getCategoryScores = (
  item: IRawSubmission,
  configuration: ITierConfiguration
): ICategoryScore[] => {
  return [
    {
      name: 'Environmental',
      value: parseNumber(
        item[configuration.environmentalCountInternalName]
      ),
      iconName: 'Leaf',
      color: '#107c10'
    },
    {
      name: 'Social',
      value: parseNumber(
        item[configuration.socialCountInternalName]
      ),
      iconName: 'People',
      color: '#5c2d91'
    },
    {
      name: 'Governance',
      value: parseNumber(
        item[configuration.governanceCountInternalName]
      ),
      iconName: 'Bank',
      color: '#0078d4'
    }
  ];
};

const widgetStyle: React.CSSProperties = {
  boxSizing: 'border-box',
  width: '100%',
  maxWidth: 430,
  padding: 18,
  color: '#242424',
  background: '#ffffff',
  border: '1px solid #e1dfdd',
  borderRadius: 10,
  boxShadow: '0 2px 8px rgba(0, 0, 0, 0.08)',
  fontFamily: 'Segoe UI, sans-serif'
};

const headerStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  gap: 12,
  marginBottom: 14
};

const titleStyle: React.CSSProperties = {
  margin: 0,
  fontSize: 16,
  fontWeight: 700,
  lineHeight: '22px'
};

const viewAllButtonStyle: React.CSSProperties = {
  padding: 0,
  color: '#005a9e',
  background: 'transparent',
  border: 0,
  cursor: 'pointer',
  fontSize: 12,
  fontWeight: 600
};

const supplierRowStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'flex-start',
  justifyContent: 'space-between',
  gap: 8,
  marginBottom: 16
};

const supplierNameStyle: React.CSSProperties = {
  display: 'block',
  overflow: 'hidden',
  fontSize: 13,
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap'
};

const secondaryTextStyle: React.CSSProperties = {
  overflow: 'hidden',
  marginTop: 3,
  color: '#605e5c',
  fontSize: 11,
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap'
};

const tierBadgeStyle: React.CSSProperties = {
  flex: '0 0 auto',
  padding: '2px 8px',
  color: '#835c00',
  background: '#fff4ce',
  borderRadius: 10,
  fontSize: 10,
  fontWeight: 600
};

const scoreLayoutStyle: React.CSSProperties = {
  display: 'grid',
  gridTemplateColumns: '116px minmax(0, 1fr)',
  gap: 16,
  alignItems: 'center',
  marginBottom: 16
};

const scoreRingStyle: React.CSSProperties = {
  position: 'relative',
  width: 108,
  height: 108,
  borderRadius: '50%'
};

const scoreRingInnerStyle: React.CSSProperties = {
  position: 'absolute',
  inset: 9,
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  flexDirection: 'column',
  background: '#ffffff',
  borderRadius: '50%'
};

const scoreValueStyle: React.CSSProperties = {
  fontSize: 25,
  lineHeight: 1,
  color: '#201f1e'
};

const scoreCaptionStyle: React.CSSProperties = {
  marginTop: 5,
  color: '#605e5c',
  fontSize: 9,
  fontWeight: 600
};

const weightedScoreCaptionStyle: React.CSSProperties = {
  marginTop: 3,
  color: '#797775',
  fontSize: 9,
  fontWeight: 600
};

const categoryCardStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  gap: 8,
  marginBottom: 8,
  padding: '9px 10px',
  background: '#faf9f8',
  border: '1px solid #edebe9',
  borderRadius: 6
};

const categoryNameStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 7,
  fontSize: 11
};

const categoryScoreStyle: React.CSSProperties = {
  fontSize: 11
};

const detailsSectionStyle: React.CSSProperties = {
  paddingTop: 14,
  borderTop: '1px solid #edebe9'
};

const detailRowStyle: React.CSSProperties = {
  display: 'grid',
  gridTemplateColumns: '105px minmax(0, 1fr)',
  alignItems: 'center',
  gap: 10,
  minHeight: 28
};

const detailLabelStyle: React.CSSProperties = {
  color: '#605e5c',
  fontSize: 11
};

const detailValueStyle: React.CSSProperties = {
  color: '#323130',
  fontSize: 11,
  fontWeight: 500
};

const badgeStyle: React.CSSProperties = {
  justifySelf: 'start',
  padding: '2px 8px',
  borderRadius: 10,
  fontSize: 10,
  fontWeight: 600
};

const noteRowStyle: React.CSSProperties = {
  display: 'grid',
  gridTemplateColumns: '105px minmax(0, 1fr)',
  gap: 10,
  marginTop: 8
};

const noteTextStyle: React.CSSProperties = {
  color: '#323130',
  fontSize: 11,
  lineHeight: '16px'
};

const emptyStateStyle: React.CSSProperties = {
  padding: '24px 12px',
  color: '#605e5c',
  textAlign: 'center',
  fontSize: 12
};

const assessmentButtonStyles: IButtonStyles = {
  root: {
    marginTop: 16,
    minHeight: 34,
    color: '#005a9e',
    background: '#ffffff',
    borderColor: '#d2d0ce',
    borderRadius: 5
  },
  label: {
    fontSize: 12,
    fontWeight: 600
  },
  icon: {
    color: '#005a9e'
  }
};

export function EsgFeedbackWidget(
  props: IEsgFeedbackWidgetProps
): React.ReactElement {
  const [latestSubmission, setLatestSubmission] =
    useState<ILatestSubmission | undefined>(undefined);
  const [loading, setLoading] = useState<boolean>(true);
  const [errorMessage, setErrorMessage] = useState<string>('');

  /*
   * List titles come from the web-part properties. The defaults support
   * pre-production when an existing web-part instance has no saved values.
   */
  const tierConfigurations: readonly ITierConfiguration[] = [
    {
      ...DEFAULT_TIER_CONFIGURATIONS[0],
      listTitle:
        props.tier1ListTitle ||
        DEFAULT_TIER_CONFIGURATIONS[0].listTitle
    },
    {
      ...DEFAULT_TIER_CONFIGURATIONS[1],
      listTitle:
        props.tier2ListTitle ||
        DEFAULT_TIER_CONFIGURATIONS[1].listTitle
    },
    {
      ...DEFAULT_TIER_CONFIGURATIONS[2],
      listTitle:
        props.tier3ListTitle ||
        DEFAULT_TIER_CONFIGURATIONS[2].listTitle
    }
  ];

  useEffect((): (() => void) => {
    let isMounted: boolean = true;

    const fetchLatestItemForTier = async (
      configuration: ITierConfiguration
    ): Promise<ITierFetchResult> => {
      const sp = getSP(props.context);

      /*
       * Dynamic field selection allows each tier to provide its own
       * internal names while the mapping code stays common.
       */
      const selectFields: string[] = [
        'Id',
        'Created',
        'Modified',
        'Author/Title',
        'Author/EMail',
        'Editor/Title',
        'Editor/EMail',
        configuration.supplierNameInternalName,
        configuration.emailInternalName,
        configuration.contactNameInternalName,
        configuration.environmentalCountInternalName,
        configuration.socialCountInternalName,
        configuration.governanceCountInternalName,
        configuration.overallPercentageInternalName
      ];

      const items: IRawSubmission[] = await sp.web.lists
        .getByTitle(configuration.listTitle)
        .items.select(...selectFields)
        .expand('Author', 'Editor')
        .orderBy('Created', false)
        .top(1)();

      return {
        configuration,
        item: items[0]
      };
    };

    const loadLatestFeedback = async (): Promise<void> => {
      setLoading(true);
      setErrorMessage('');

      try {
        /*
         * Retrieve the newest item from each tier in parallel.
         * Each request returns at most one item.
         */
        const tierResults: ITierFetchResult[] =
          await Promise.all(
            tierConfigurations.map(
              (
                configuration: ITierConfiguration
              ): Promise<ITierFetchResult> => {
                return fetchLatestItemForTier(configuration);
              }
            )
          );

        if (!isMounted) {
          return;
        }

        /*
         * Remove tiers without submissions, then compare Created dates.
         * The first entry after sorting is the latest across all tiers.
         */
        const availableResults: ITierFetchResult[] =
          tierResults
            .filter(
              (result: ITierFetchResult): boolean => {
                return Boolean(result.item);
              }
            )
            .sort(
              (
                first: ITierFetchResult,
                second: ITierFetchResult
              ): number => {
                const firstTime: number = first.item?.Created
                  ? new Date(first.item.Created).getTime()
                  : 0;
                const secondTime: number = second.item?.Created
                  ? new Date(second.item.Created).getTime()
                  : 0;

                return secondTime - firstTime;
              }
            );

        const latestResult: ITierFetchResult | undefined =
          availableResults[0];

        if (!latestResult || !latestResult.item) {
          setLatestSubmission(undefined);
          return;
        }

        const item: IRawSubmission = latestResult.item;
        const configuration: ITierConfiguration =
          latestResult.configuration;
        const weightedScore: number = parseNumber(
          item[configuration.overallPercentageInternalName]
        );
        const overallScore: number = calculateRelativeScore(
          weightedScore,
          configuration.maximumWeightedScore
        );
        const evaluation = evaluateQualification(overallScore);
        const siteUrl: string =
          props.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
        const itemId: number =
          typeof item.Id === 'number' ? item.Id : 0;
        const created: string =
          typeof item.Created === 'string' ? item.Created : '';
        const modified: string =
          typeof item.Modified === 'string' ? item.Modified : '';

        setLatestSubmission({
          id: itemId,
          tier: configuration.tier,
          listTitle: configuration.listTitle,
          supplierName:
            getTextValue(
              item[configuration.supplierNameInternalName]
            ) ||
            item.Author?.Title?.trim() ||
            'Supplier name unavailable',
          supplierEmail:
            getTextValue(item[configuration.emailInternalName]) ||
            item.Author?.EMail?.trim() ||
            '',
          contactName:
            getTextValue(
              item[configuration.contactNameInternalName]
            ) ||
            item.Author?.Title?.trim() ||
            '',
          sharePointWeightedScore: weightedScore,
          maximumWeightedScore:
            configuration.maximumWeightedScore,
          overallScore,
          categoryScores: getCategoryScores(item, configuration),
          qualification: evaluation.qualification,
          riskRating: evaluation.riskRating,
          recommendation: evaluation.recommendation,
          lastUpdatedBy:
            item.Editor?.Title?.trim() ||
            item.Editor?.EMail?.trim() ||
            'Not available',
          lastUpdated: formatDate(
            modified,
            props.context.pageContext.cultureInfo
              .currentUICultureName
          ),
          assessmentNote: generateAssessmentNote(
            evaluation.qualification,
            configuration.tier,
            weightedScore,
            configuration.maximumWeightedScore,
            overallScore
          ),
          created,
          itemUrl:
            `${siteUrl}/Lists/` +
            `${encodeURIComponent(configuration.listTitle)}` +
            `/DispForm.aspx?ID=${itemId.toString()}`
        });
      } catch (error: unknown) {
        console.error('Error loading latest ESG feedback:', error);

        if (isMounted) {
          setLatestSubmission(undefined);
          setErrorMessage(
            'The latest ESG feedback could not be loaded from SharePoint.'
          );
        }
      } finally {
        if (isMounted) {
          setLoading(false);
        }
      }
    };

    loadLatestFeedback().catch(
      (error: unknown): void => {
        console.error(
          'Unhandled ESG feedback loading error:',
          error
        );

        if (isMounted) {
          setLatestSubmission(undefined);
          setErrorMessage(
            'An unexpected error occurred while loading ESG feedback.'
          );
          setLoading(false);
        }
      }
    );

    return (): void => {
      isMounted = false;
    };
  }, [
    props.context,
    props.tier1ListTitle,
    props.tier2ListTitle,
    props.tier3ListTitle
  ]);

  const openAssessment = (): void => {
    if (!latestSubmission) {
      return;
    }

    window.open(
      latestSubmission.itemUrl,
      '_blank',
      'noopener,noreferrer'
    );
  };

  const openAllAssessments = (): void => {
    if (!latestSubmission) {
      return;
    }

    const siteUrl: string =
      props.context.pageContext.web.absoluteUrl.replace(/\/$/, '');

    window.open(
      `${siteUrl}/Lists/${encodeURIComponent(
        latestSubmission.listTitle
      )}/AllItems.aspx`,
      '_blank',
      'noopener,noreferrer'
    );
  };

  if (loading) {
    return (
      <div role="status" aria-live="polite" style={{ padding: 16 }}>
        <Spinner
          size={SpinnerSize.small}
          label="Loading latest ESG feedback"
        />
      </div>
    );
  }

  const qualificationColors = latestSubmission
    ? getBadgeColors(latestSubmission.qualification)
    : getBadgeColors('Not Qualified');
  const riskColors = latestSubmission
    ? getBadgeColors(latestSubmission.riskRating)
    : getBadgeColors('High Risk');
  const recommendationColors = latestSubmission
    ? getBadgeColors(latestSubmission.recommendation)
    : getBadgeColors('Requires Review');

  return (
    <section aria-label="Latest ESG feedback" style={widgetStyle}>
      <div style={headerStyle}>
        <h3 style={titleStyle}>Latest ESG Feedback</h3>

        <button
          type="button"
          onClick={openAllAssessments}
          disabled={!latestSubmission}
          style={{
            ...viewAllButtonStyle,
            cursor: latestSubmission ? 'pointer' : 'default',
            opacity: latestSubmission ? 1 : 0.5
          }}
        >
          View all&nbsp;
          <Icon iconName="OpenInNewWindow" />
        </button>
      </div>

      {errorMessage && (
        <MessageBar messageBarType={MessageBarType.error}>
          {errorMessage}
        </MessageBar>
      )}

      {!latestSubmission ? (
        <div style={emptyStateStyle}>
          No ESG questionnaire submissions were found.
        </div>
      ) : (
        <>
          <div style={supplierRowStyle}>
            <div style={{ minWidth: 0 }}>
              <strong
                title={latestSubmission.supplierEmail}
                style={supplierNameStyle}
              >
                {latestSubmission.supplierName}
              </strong>
              {latestSubmission.contactName && (
                <div style={secondaryTextStyle}>
                  {latestSubmission.contactName}
                </div>
              )}
            </div>

            <span style={tierBadgeStyle}>
              {latestSubmission.tier}
            </span>
          </div>

          <div style={scoreLayoutStyle}>
            <div
              style={{
                ...scoreRingStyle,
                background:
                  `conic-gradient(#107c41 0 ` +
                  `${latestSubmission.overallScore}%, ` +
                  `#edebe9 ${latestSubmission.overallScore}% 100%)`
              }}
            >
              <div style={scoreRingInnerStyle}>
                <strong style={scoreValueStyle}>
                  {latestSubmission.overallScore}%
                </strong>
                <span style={scoreCaptionStyle}>
                  Overall ESG Score
                </span>
                <span style={weightedScoreCaptionStyle}>
                  {roundToTwoDecimals(
                    latestSubmission.sharePointWeightedScore
                  )}{' '}
                  / {latestSubmission.maximumWeightedScore}
                </span>
              </div>
            </div>

            <div>
              {latestSubmission.categoryScores.map(
                (
                  category: ICategoryScore
                ): React.ReactElement => (
                  <div key={category.name} style={categoryCardStyle}>
                    <span style={categoryNameStyle}>
                      <Icon
                        iconName={category.iconName}
                        styles={{
                          root: {
                            color: category.color,
                            fontSize: 14
                          }
                        }}
                      />
                      {category.name}
                    </span>

                    <strong style={categoryScoreStyle}>
                      {roundToTwoDecimals(category.value)}
                    </strong>
                  </div>
                )
              )}
            </div>
          </div>

          <div style={detailsSectionStyle}>
            <DetailBadgeRow
              label="Qualification"
              value={latestSubmission.qualification}
              colors={qualificationColors}
            />
            <DetailBadgeRow
              label="Risk Rating"
              value={latestSubmission.riskRating}
              colors={riskColors}
            />
            <DetailBadgeRow
              label="Recommendation"
              value={latestSubmission.recommendation}
              colors={recommendationColors}
            />
            <DetailTextRow
              label="Last Updated By"
              value={latestSubmission.lastUpdatedBy}
            />
            <DetailTextRow
              label="Last Updated"
              value={latestSubmission.lastUpdated}
            />

            <div style={noteRowStyle}>
              <span style={detailLabelStyle}>Assessment Note</span>
              <span style={noteTextStyle}>
                {latestSubmission.assessmentNote}
              </span>
            </div>
          </div>

          <DefaultButton
            text="View Full Assessment"
            iconProps={{ iconName: 'Forward' }}
            onClick={openAssessment}
            styles={assessmentButtonStyles}
          />
        </>
      )}
    </section>
  );
}

interface IDetailBadgeRowProps {
  label: string;
  value: string;
  colors: {
    background: string;
    color: string;
  };
}

function DetailBadgeRow(
  props: IDetailBadgeRowProps
): React.ReactElement {
  return (
    <div style={detailRowStyle}>
      <span style={detailLabelStyle}>{props.label}</span>
      <span
        style={{
          ...badgeStyle,
          color: props.colors.color,
          background: props.colors.background
        }}
      >
        {props.value}
      </span>
    </div>
  );
}

interface IDetailTextRowProps {
  label: string;
  value: string;
}

function DetailTextRow(
  props: IDetailTextRowProps
): React.ReactElement {
  return (
    <div style={detailRowStyle}>
      <span style={detailLabelStyle}>{props.label}</span>
      <span style={detailValueStyle}>{props.value}</span>
    </div>
  );
}

export default EsgFeedbackWidget;
