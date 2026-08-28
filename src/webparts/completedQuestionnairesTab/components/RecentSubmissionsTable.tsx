import * as React from "react";
import {
  DetailsListLayoutMode,
  IColumn,
  IconButton,
  Link,
  SelectionMode,
  ShimmeredDetailsList,
} from "@fluentui/react";
import {
  QualificationStatus,
  Recommendation,
  RiskRating,
} from "../../supplierEsgSearch/models/IQualificationResult";
import { SupplierTier } from "../../supplierEsgSearch/models/IListConfiguration";
import { ISupplierSubmission } from "../../supplierEsgSearch/models/ISupplierSubmission";
import { IRecentSubmissionsTableProps } from "./IRecentSubmissionsTableProps";
import styles from "./RecentSubmissionsTable.module.scss";

export function RecentSubmissionsTable(
  props: IRecentSubmissionsTableProps,
): React.ReactElement<IRecentSubmissionsTableProps> {
  const RECENT_ITEMS_LIMIT: number = 5;

  const visible: readonly ISupplierSubmission[] = props.submissions.slice(
    0,
    RECENT_ITEMS_LIMIT,
  );

  const tierClass = (tier: SupplierTier): string =>
    tier === "Tier 1"
      ? styles.tier1
      : tier === "Tier 2"
        ? styles.tier2
        : styles.tier3;
  const statusClass = (value: QualificationStatus): string =>
    value === "Qualified"
      ? styles.success
      : value === "Conditionally Qualified"
        ? styles.warning
        : styles.danger;
  const riskClass = (value: RiskRating): string =>
    value === "Low Risk"
      ? styles.low
      : value === "Medium Risk"
        ? styles.medium
        : styles.high;
  const recommendationClass = (value: Recommendation): string =>
    value === "Approve"
      ? styles.success
      : value === "Corrective Action"
        ? styles.warning
        : styles.danger;
  const formatDate = (value: string): string => {
    const date: Date = new Date(value);
    return isNaN(date.getTime())
      ? "-"
      : date.toLocaleDateString("en-GB", {
          day: "2-digit",
          month: "short",
          year: "numeric",
        });
  };
  const formatScore = (value: number): string =>
    `${Math.round((isFinite(value) ? value : 0) * 100) / 100}%`;

  const columns: IColumn[] = [
    {
      key: "supplierName",
      name: "Supplier Name",
      fieldName: "supplierName",
      minWidth: 175,
      maxWidth: 230,
      onRender: (item: ISupplierSubmission): React.ReactNode => (
        <Link onClick={(): void => props.onOpenSubmission(item)}>
          {item.supplierName}
        </Link>
      ),
    },
    {
      key: "tier",
      name: "Tier",
      fieldName: "tier",
      minWidth: 65,
      maxWidth: 80,
      onRender: (item: ISupplierSubmission): React.ReactNode => (
        <span className={`${styles.badge} ${tierClass(item.tier)}`}>
          {item.tier}
        </span>
      ),
    },
    {
      key: "score",
      name: "Score",
      fieldName: "overallPercentage",
      minWidth: 58,
      maxWidth: 72,
      onRender: (item: ISupplierSubmission): React.ReactNode =>
        formatScore(item.overallPercentage),
    },
    {
      key: "qualification",
      name: "ESG Status",
      fieldName: "qualification",
      minWidth: 135,
      maxWidth: 170,
      onRender: (item: ISupplierSubmission): React.ReactNode => (
        <span className={`${styles.badge} ${statusClass(item.qualification)}`}>
          {item.qualification}
        </span>
      ),
    },
    {
      key: "risk",
      name: "Risk Rating",
      fieldName: "riskRating",
      minWidth: 95,
      maxWidth: 115,
      onRender: (item: ISupplierSubmission): React.ReactNode => (
        <span className={styles.risk}>
          <i className={riskClass(item.riskRating)} />
          {item.riskRating}
        </span>
      ),
    },
    {
      key: "submittedBy",
      name: "Reviewer",
      fieldName: "submittedByName",
      minWidth: 95,
      maxWidth: 125,
      onRender: (item: ISupplierSubmission): React.ReactNode =>
        item.submittedByName || "Unassigned",
    },
    {
      key: "created",
      name: "Submission Date",
      fieldName: "created",
      minWidth: 105,
      maxWidth: 130,
      onRender: (item: ISupplierSubmission): React.ReactNode =>
        formatDate(item.created),
    },
    {
      key: "recommendation",
      name: "Recommendation",
      fieldName: "recommendation",
      minWidth: 120,
      maxWidth: 155,
      onRender: (item: ISupplierSubmission): React.ReactNode => (
        <span
          className={`${styles.badge} ${recommendationClass(item.recommendation)}`}
        >
          {item.recommendation}
        </span>
      ),
    },
    {
      key: "actions",
      name: "Actions",
      minWidth: 60,
      maxWidth: 60,
      onRender: (item: ISupplierSubmission): React.ReactNode => (
        <IconButton
          iconProps={{ iconName: "MoreVertical" }}
          menuProps={{
            items: [
              {
                key: "open",
                text: "Open submission",
                iconProps: { iconName: "OpenInNewWindow" },
                onClick: (): void => props.onOpenSubmission(item),
              },
            ],
          }}
        />
      ),
    },
  ];

  return (
    <section
      className={styles.panel}
      aria-labelledby="recent-submissions-title"
    >
      <header className={styles.header}>
        <div>
          <h2 id="recent-submissions-title">Recent Submissions</h2>
          <span>
            Showing {visible.length} of {props.submissions.length}
          </span>
        </div>
      </header>
      <div className={styles.table}>
        <div className={styles.scroll}>
          <ShimmeredDetailsList
            items={visible as ISupplierSubmission[]}
            columns={columns}
            enableShimmer={props.isLoading}
            selectionMode={SelectionMode.none}
            layoutMode={DetailsListLayoutMode.justified}
            compact={true}
          />
        </div>
        {!props.isLoading && visible.length === 0 && (
          <div className={styles.empty}>No matching submissions found.</div>
        )}
      </div>
      <footer className={styles.footer}>
        <span>Go to Completed Questionnaires:</span>
        <Link
          disabled={!props.tierListUrls.tier1}
          onClick={(): void => props.onOpenTierList(props.tierListUrls.tier1)}
        >
          Tier 1
        </Link>
        <span>|</span>
        <Link
          disabled={!props.tierListUrls.tier2}
          onClick={(): void => props.onOpenTierList(props.tierListUrls.tier2)}
        >
          Tier 2
        </Link>
        <span>|</span>
        <Link
          disabled={!props.tierListUrls.tier3}
          onClick={(): void => props.onOpenTierList(props.tierListUrls.tier3)}
        >
          Tier 3
        </Link>
        <span aria-hidden="true">→</span>
      </footer>
    </section>
  );
}

export default RecentSubmissionsTable;
