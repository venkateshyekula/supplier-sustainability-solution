import * as React from "react";

import { Icon } from "@fluentui/react";

import { KpiMetricCards } from "../../kpiMetricCards/components/KpiMetricCards";

import EsgFeedbackWidget from "../../esgFeedbackWidget/components/EsgFeedbackWidget";

import QuestionnaireQuickLinks from "../../questionnaireQuickLinks/components/QuestionnaireQuickLinks";

import SupplierEsgSearch from "../../supplierEsgSearch/components/SupplierEsgSearch";

import { ISupplierSustainabilityDashboardProps } from "./ISupplierSustainabilityDashboardProps";

import styles from "./SupplierSustainabilityDashboard.module.scss";

export function SupplierSustainabilityDashboard(
  props: ISupplierSustainabilityDashboardProps,
): React.ReactElement {
  const userDisplayName: string =
    props.context.pageContext.user.displayName?.trim() || "User";

  /*const userDisplayName: string =
  props.context.pageContext.user.displayName || "User";

  const firstName: string =
  getFirstName(userDisplayName);*/

  return (
    <section
      className={styles.dashboard}
      aria-label="Supplier Sustainability dashboard"
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
        aria-labelledby="supplier-esg-overview-title"
      >
        <div className={styles.kpiSectionHeader}>
          <div>
            <h2
              id="supplier-esg-overview-title"
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
          aria-label="Latest ESG feedback"
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
        aria-labelledby="supplier-search-title"
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
        aria-label="Information for suppliers"
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

/*function getFirstName(displayName: string): string {
  const normalizedDisplayName: string = displayName.trim().replace(/\s+/g, " ");

  if (!normalizedDisplayName) {
    return "User";
  }

  const commaIndex: number = normalizedDisplayName.indexOf(",");

  if (commaIndex >= 0) {
    const nameAfterComma: string = normalizedDisplayName
      .substring(commaIndex + 1)
      .trim();

    if (nameAfterComma) {
      return nameAfterComma.split(" ")[0] || "User";
    }
  }

  return normalizedDisplayName.split(" ")[0] || "User";
}*/

export default SupplierSustainabilityDashboard;
