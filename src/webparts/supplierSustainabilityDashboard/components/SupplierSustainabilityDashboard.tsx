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
  return (
    <section
      className={styles.dashboard}
      aria-label="Supplier sustainability dashboard"
    >
      <header className={styles.hero}>
        <div className={styles.heroIcon} aria-hidden="true">
          <Icon iconName="Globe" />
        </div>

        <div className={styles.heroContent}>
          <h1 className={styles.heroTitle}>
            Supplier Sustainability Dashboard
          </h1>
          <p className={styles.heroDescription}>
            Monitor supplier ESG submissions, qualification outcomes, supporting
            documents, and recent assessment activity across all configured
            questionnaire tiers.
          </p>
        </div>
      </header>

      <section
        className={styles.dashboardSection}
        aria-labelledby="supplier-esg-overview-title"
      >
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
          aria-label="Questionnaire and supporting document links"
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

export default SupplierSustainabilityDashboard;
