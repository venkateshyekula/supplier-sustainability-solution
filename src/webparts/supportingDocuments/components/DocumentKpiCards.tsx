import * as React from 'react';

import {
  Icon,
  Shimmer,
  ShimmerElementType
} from '@fluentui/react';

import {
  SupplierTier
} from '../../supplierEsgSearch/models/IListConfiguration';

import {
  ISupplierAttachmentStatus,
  SupplierAttachmentStatus
} from '../models/ISupplierAttachmentStatus';

import {
  ISupplierDocument
} from '../models/ISupplierDocument';

import styles from './DocumentKpiCards.module.scss';
import { IMissingAttachment } from '../models/IMissingAttachment';

export interface IDocumentKpiCardsProps {
  documents: readonly ISupplierDocument[];

  attachmentStatuses:
    readonly ISupplierAttachmentStatus[];
  
  missingAttachments: 
   readonly IMissingAttachment[];  

  isLoading: boolean;
}

type DocumentMetricKey =
  | 'total'
  | 'submitted'
  | 'pending'
  | 'missing';

interface ITierCounts {
  tier1: number;
  tier2: number;
  tier3: number;
}

interface IDocumentKpiCard {
  key: DocumentMetricKey;
  title: string;
  value: number;
  subtitle: string;
  iconName: string;
  cardClassName: string;
  iconClassName: string;
  tierCounts: ITierCounts;
}

const SHIMMER_ELEMENTS = [
  {
    type: ShimmerElementType.circle,
    height: 52
  },
  {
    type: ShimmerElementType.gap,
    width: 12
  },
  {
    type: ShimmerElementType.line,
    width: '55%',
    height: 20
  }
];

export function DocumentKpiCards(
  props: IDocumentKpiCardsProps
): React.ReactElement<IDocumentKpiCardsProps> {
  const getStatusCount = (
    status: SupplierAttachmentStatus
  ): number => {
    return props.attachmentStatuses.filter(
      (
        item: ISupplierAttachmentStatus
      ): boolean => {
        return item.status === status;
      }
    ).length;
  };

  const getTierStatusCount = (
    tier: SupplierTier,
    status?: SupplierAttachmentStatus
  ): number => {
    return props.attachmentStatuses.filter(
      (
        item: ISupplierAttachmentStatus
      ): boolean => {
        if (item.tier !== tier) {
          return false;
        }

        return status
          ? item.status === status
          : true;
      }
    ).length;
  };

  const getStatusTierCounts = (
    status: SupplierAttachmentStatus
  ): ITierCounts => {
    return {
      tier1:
        getTierStatusCount(
          'Tier 1',
          status
        ),

      tier2:
        getTierStatusCount(
          'Tier 2',
          status
        ),

      tier3:
        getTierStatusCount(
          'Tier 3',
          status
        )
    };
  };

  const getDocumentTierCounts =
    (): ITierCounts => {
      return {
        tier1:
          props.documents.filter(
            (
              document: ISupplierDocument
            ): boolean => {
              return document.tier === 'Tier 1';
            }
          ).length,

        tier2:
          props.documents.filter(
            (
              document: ISupplierDocument
            ): boolean => {
              return document.tier === 'Tier 2';
            }
          ).length,

        tier3:
          props.documents.filter(
            (
              document: ISupplierDocument
            ): boolean => {
              return document.tier === 'Tier 3';
            }
          ).length
      };
    };

  const getPercentageText = (
    value: number,
    total: number
  ): string => {
    if (total <= 0) {
      return '0% of submissions';
    }

    const percentage: number =
      Math.round(
        (
          value /
          total
        ) *
        100
      );

    return `${percentage.toString()}% of submissions`;
  };

  const totalDocuments: number =
    props.documents.length;

  const totalStatuses: number =
    props.attachmentStatuses.length;

  const submittedCount: number =
    getStatusCount(
      'Submitted'
    );

  const pendingCount: number =
    getStatusCount(
      'Pending'
    );

  const missingCount: number =
    getStatusCount(
      'Missing'
    );

  const cards:
    readonly IDocumentKpiCard[] = [
      {
        key: 'total',
        title: 'Total Documents',
        value: totalDocuments,
        subtitle: 'All tiers',
        iconName: 'FabricFolder',
        cardClassName: styles.totalCard,
        iconClassName: styles.totalIcon,
        tierCounts: getDocumentTierCounts()
      },
      {
        key: 'submitted',
        title: 'Submitted',
        value: submittedCount,
        subtitle:
          getPercentageText(
            submittedCount,
            totalStatuses
          ),
        iconName: 'Completed',
        cardClassName: styles.submittedCard,
        iconClassName: styles.submittedIcon,
        tierCounts:
          getStatusTierCounts(
            'Submitted'
          )
      },
      {
        key: 'pending',
        title: 'Pending',
        value: pendingCount,
        subtitle:
          getPercentageText(
            pendingCount,
            totalStatuses
          ),
        iconName: 'Clock',
        cardClassName: styles.pendingCard,
        iconClassName: styles.pendingIcon,
        tierCounts:
          getStatusTierCounts(
            'Pending'
          )
      },
      {
        key: 'missing',
        title: 'Missing Attachments',
        value: missingCount,
        subtitle:
          getPercentageText(
            missingCount,
            totalStatuses
          ),
        iconName: 'Warning',
        cardClassName: styles.missingCard,
        iconClassName: styles.missingIcon,
        tierCounts:
          getStatusTierCounts(
            'Missing'
          )
      }
    ];

  return (
  <section
    className={styles.metricsLayout}
    aria-label="Supporting document metrics"
    aria-busy={props.isLoading}
  >
    <div className={styles.cardsGrid}>
      {cards.map(
        (
          card: IDocumentKpiCard
        ): React.ReactElement => {
          return (
            <Shimmer
              key={card.key}
              className={styles.shimmer}
              isDataLoaded={!props.isLoading}
              shimmerElements={
                SHIMMER_ELEMENTS
              }
            >
              <article
                className={
                  `${styles.card} ` +
                  `${card.cardClassName}`
                }
                aria-label={
                  `${card.title}: ` +
                  `${card.value.toString()}`
                }
              >
                <span
                  className={
                    `${styles.iconContainer} ` +
                    `${card.iconClassName}`
                  }
                  aria-hidden="true"
                >
                  <Icon
                    iconName={card.iconName}
                    className={styles.icon}
                  />
                </span>

                <div className={styles.cardContent}>
                  <span className={styles.cardTitle}>
                    {card.title}
                  </span>

                  <strong className={styles.cardValue}>
                    {card.value}
                  </strong>

                  <span className={styles.cardSubtitle}>
                    {card.subtitle}
                  </span>
                </div>

                <div className={styles.tierBreakdown}>
                  <div>
                    <span>Tier 1</span>

                    <strong>
                      {card.tierCounts.tier1}
                    </strong>
                  </div>

                  <div>
                    <span>Tier 2</span>

                    <strong>
                      {card.tierCounts.tier2}
                    </strong>
                  </div>

                  <div>
                    <span>Tier 3</span>

                    <strong>
                      {card.tierCounts.tier3}
                    </strong>
                  </div>
                </div>
              </article>
            </Shimmer>
          );
        }
      )}
    </div>
  </section>
);
}

export default DocumentKpiCards;