import * as React from "react";

import { Icon, MessageBar, MessageBarType } from "@fluentui/react";

import {
  IListConfiguration,
  SupplierTier,
} from "../../supplierEsgSearch/models/IListConfiguration";

import { ISupplierSubmission } from "../../supplierEsgSearch/models/ISupplierSubmission";

import { SupplierSubmissionService } from "../../supplierEsgSearch/services/SupplierSubmissionService";

import { IDocumentLibraryConfiguration } from "../models/IDocumentLibraryConfiguration";

import { IMissingAttachment } from "../models/IMissingAttachment";

import { ISupplierDocument } from "../models/ISupplierDocument";

import { ISupplierDocumentFilters } from "../models/ISupplierDocumentFilters";

import { ISupplierFolderSummary } from "../models/ISupplierFolderSummary";

import { ITierAttachmentOverview } from "../models/ITierAttachmentOverview";

import { SupplierDocumentService } from "../services/SupplierDocumentService";

import DocumentKpiCards from "./DocumentKpiCards";

import MissingAttachments from "./MissingAttachments";

import RecentDocumentsTable from "./RecentDocumentsTable";

import SupportingDocumentFilters from "./SupportingDocumentFilters";

import SupportingDocumentQuickActions from "./SupportingDocumentQuickActions";

import { ISupportingDocumentsProps } from "./ISupportingDocumentsProps";

import SupplierFolderTable from "./SupplierFolderTable";

import TierAttachmentOverview from "./TierAttachmentOverview";

import styles from "./SupportingDocuments.module.scss";
import { ISupplierAttachmentStatus } from "../models/ISupplierAttachmentStatus";

interface ISupportingDocumentsState {
  allDocuments: ISupplierDocument[];
  displayedDocuments: ISupplierDocument[];

  allFolders: ISupplierFolderSummary[];
  displayedFolders: ISupplierFolderSummary[];

  submissions: ISupplierSubmission[];
  attachmentStatuses: ISupplierAttachmentStatus[];
  missingAttachments: IMissingAttachment[];
  tierOverview: ITierAttachmentOverview[];

  filters: ISupplierDocumentFilters;

  libraryUrls: Record<string, string>;

  isLoading: boolean;
  errorMessage?: string;
}

export default class SupportingDocuments extends React.Component<
  ISupportingDocumentsProps,
  ISupportingDocumentsState
> {
  private readonly documentService: SupplierDocumentService;

  private readonly submissionService: SupplierSubmissionService;

  public constructor(props: ISupportingDocumentsProps) {
    super(props);

    this.documentService = new SupplierDocumentService(props.context);

    this.submissionService = new SupplierSubmissionService(props.context);

    this.state = {
      allDocuments: [],
      displayedDocuments: [],

      allFolders: [],
      displayedFolders: [],

      submissions: [],
      missingAttachments: [],
      attachmentStatuses: [],
      tierOverview: [],

      filters: {
        searchText: "",
      },

      libraryUrls: {},

      isLoading: true,
      errorMessage: undefined,
    };
  }

  public componentDidMount(): void {
    this.loadData().catch((error: Error): void => {
      this.setState({
        isLoading: false,
        errorMessage: error.message || "An unexpected error occurred.",
      });
    });
  }

  public componentDidUpdate(previousProps: ISupportingDocumentsProps): void {
    const configurationChanged: boolean =
      previousProps.tier1DocumentLibraryTitle !==
        this.props.tier1DocumentLibraryTitle ||
      previousProps.tier2DocumentLibraryTitle !==
        this.props.tier2DocumentLibraryTitle ||
      previousProps.tier3DocumentLibraryTitle !==
        this.props.tier3DocumentLibraryTitle ||
      previousProps.tier1QuestionnaireListTitle !==
        this.props.tier1QuestionnaireListTitle ||
      previousProps.tier2QuestionnaireListTitle !==
        this.props.tier2QuestionnaireListTitle ||
      previousProps.tier3QuestionnaireListTitle !==
        this.props.tier3QuestionnaireListTitle ||
      previousProps.tier1SupplierNameField !==
        this.props.tier1SupplierNameField ||
      previousProps.tier2SupplierNameField !==
        this.props.tier2SupplierNameField ||
      previousProps.tier3SupplierNameField !==
        this.props.tier3SupplierNameField;

    if (configurationChanged) {
      this.loadData().catch((error: Error): void => {
        this.setState({
          isLoading: false,
          errorMessage: error.message || "An unexpected error occurred.",
        });
      });
    }
  }

  public render(): React.ReactElement {
    return (
      <section className={styles.page} aria-label="Supporting documents">
        <header className={styles.pageHeader}>
          <div>
            <h1>Supporting Documents</h1>

            <p>
              Track questionnaire submissions, supplier folders, and supporting
              files across all tiers.
            </p>
          </div>
        </header>

        <section
          className={styles.filters}
          aria-label="Supporting document filters"
        >
          <SupportingDocumentFilters
            filters={this.state.filters}
            isLoading={this.state.isLoading}
            onFiltersChange={(filters: ISupplierDocumentFilters): void => {
              this.setState({
                filters,
              });
            }}
            onSearch={this.applyFilters}
            onClear={this.clearFilters}
          />
        </section>

        {this.state.errorMessage && (
          <div className={styles.errorMessage}>
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={true}
              onDismiss={(): void => {
                this.setState({
                  errorMessage: undefined,
                });
              }}
            >
              {this.state.errorMessage}
            </MessageBar>
          </div>
        )}

        <section
          className={styles.metricsSection}
          aria-label="Supporting document metrics"
        >
          <div className={styles.kpiCardsSection}>
            <DocumentKpiCards
              documents={this.state.allDocuments}
              attachmentStatuses={this.state.attachmentStatuses}
              missingAttachments={this.state.missingAttachments}
              isLoading={this.state.isLoading}
            />
          </div>

          <div className={styles.tierOverviewSection}>
            <TierAttachmentOverview
              items={this.state.tierOverview}
              isLoading={this.state.isLoading}
              onOpenSubmission={this.openUrl}
            />
          </div>
        </section>

        <div className={styles.contentGrid}>
          <main className={styles.mainColumn}>
            <section
              className={styles.tableSection}
              aria-label="Supplier document folders"
            >
              <SupplierFolderTable
                items={this.state.displayedFolders}
                isLoading={this.state.isLoading}
                onOpenFolder={this.openFolder}
              />
            </section>

            <section
              className={styles.tableSection}
              aria-label="Recent supporting documents"
            >
              <RecentDocumentsTable
                items={this.state.displayedDocuments}
                isLoading={this.state.isLoading}
                onOpenDocument={this.openDocument}
              />
            </section>
          </main>

          <aside
            className={styles.sideColumn}
            aria-label="Document alerts and actions"
          >
            <section
              className={styles.sidePanel}
              aria-label="Missing attachments"
            >
              <MissingAttachments
                items={this.state.missingAttachments}
                isLoading={this.state.isLoading}
                onOpenSubmission={(item: IMissingAttachment): void => {
                  this.openUrl(item.submissionItemUrl);
                }}
              />
            </section>

            <section
              className={styles.sidePanel}
              aria-label="Supporting document quick actions"
            >
              <SupportingDocumentQuickActions
                tier1Url={this.state.libraryUrls["Tier 1"]}
                tier2Url={this.state.libraryUrls["Tier 2"]}
                tier3Url={this.state.libraryUrls["Tier 3"]}
                guidanceUrl={this.props.documentGuidanceUrl}
                onOpen={this.openUrl}
              />
            </section>
          </aside>
        </div>

        <footer
          className={styles.notice}
          aria-label="Supporting document information"
        >
          <span className={styles.noticeIcon} aria-hidden="true">
            <Icon iconName="Info" />
          </span>

          <p>
            <strong>Important:</strong> Please Upload All Supporting Documents for the Appropriate Tiers Questionnaries.
            Ensure that all documents are uploaded to the correct tier document library and that they are properly named and organized. This will help us maintain accurate records and facilitate efficient review and approval processes.
          </p>
        </footer>
      </section>
    );
  }

  private getDocumentLibraryConfigurations(): readonly IDocumentLibraryConfiguration[] {
    return [
      {
        tier: "Tier 1",
        libraryTitle: this.props.tier1DocumentLibraryTitle,
      },
      {
        tier: "Tier 2",
        libraryTitle: this.props.tier2DocumentLibraryTitle,
      },
      {
        tier: "Tier 3",
        libraryTitle: this.props.tier3DocumentLibraryTitle,
      },
    ];
  }

  private getQuestionnaireConfigurations(): readonly IListConfiguration[] {
    return [
      {
        listTitle: this.props.tier1QuestionnaireListTitle,

        tier: "Tier 1",

        supplierNameDisplayName: "Supplier Name",

        supplierNameInternalName: this.props.tier1SupplierNameField,

        emailInternalName: "field_3",

        contactNameInternalName: "field_4",

        overallPercentageInternalName: "OverallQuestionsPercentage",

        qualifiedWeightedMinimum: 10,

        conditionalWeightedMinimum: 9.99,
      },
      {
        listTitle: this.props.tier2QuestionnaireListTitle,

        tier: "Tier 2",

        supplierNameDisplayName: "Supplier Name",

        supplierNameInternalName: this.props.tier2SupplierNameField,

        emailInternalName: "Email",

        contactNameInternalName: "Name",

        overallPercentageInternalName: "OverallQuestionsPercentage",

        qualifiedWeightedMinimum: 5,

        conditionalWeightedMinimum: 4.99,
      },
      {
        listTitle: this.props.tier3QuestionnaireListTitle,

        tier: "Tier 3",

        supplierNameDisplayName: "Supplier Name",

        supplierNameInternalName: this.props.tier3SupplierNameField,

        emailInternalName: "Email",

        contactNameInternalName: "Name",

        overallPercentageInternalName: "OverallQuestionsPercentage",

        qualifiedWeightedMinimum: 2,

        conditionalWeightedMinimum: 1.99,
      },
    ];
  }

  private loadData = async (): Promise<void> => {
    this.setState({
      isLoading: true,
      errorMessage: undefined,
    });

    try {
      const libraryConfigurations: readonly IDocumentLibraryConfiguration[] =
        this.getDocumentLibraryConfigurations();

      const questionnaireConfigurations: readonly IListConfiguration[] =
        this.getQuestionnaireConfigurations();

      const results: [
        ISupplierDocument[],
        ISupplierSubmission[],
        Record<string, string>,
      ] = await Promise.all([
        this.documentService.getAllDocuments(libraryConfigurations),

        this.submissionService.getAllSubmissions(questionnaireConfigurations),

        this.documentService.getLibraryUrls(libraryConfigurations),
      ]);

      const documents: ISupplierDocument[] = results[0];

      const submissions: ISupplierSubmission[] = results[1];

      const libraryUrls: Record<string, string> = results[2];

      const folders: ISupplierFolderSummary[] =
        this.buildFolderSummaries(documents);

      const missingAttachments: IMissingAttachment[] =
        this.buildMissingAttachments(submissions, documents);

      const tierOverview: ITierAttachmentOverview[] = this.buildTierOverview(
        submissions,
        documents,
        missingAttachments,
      );

      this.setState({
        allDocuments: documents,

        displayedDocuments: documents,

        allFolders: folders,

        displayedFolders: folders,

        submissions,

        missingAttachments,

        tierOverview,

        filters: {
          searchText: "",
        },

        libraryUrls,

        isLoading: false,

        errorMessage: undefined,
      });
    } catch (error: unknown) {
      this.setState({
        allDocuments: [],
        displayedDocuments: [],

        allFolders: [],
        displayedFolders: [],

        submissions: [],
        missingAttachments: [],
        tierOverview: [],

        libraryUrls: {},

        isLoading: false,

        errorMessage:
          error instanceof Error
            ? error.message
            : "Unable to load supporting document data.",
      });
    }
  };

  private buildMissingAttachments(
    submissions: readonly ISupplierSubmission[],

    documents: readonly ISupplierDocument[],
  ): IMissingAttachment[] {
    const documentKeys: Set<string> = new Set<string>();

    documents.forEach((document: ISupplierDocument): void => {
      documentKeys.add(
        this.createSupplierTierKey(document.supplierName, document.tier),
      );
    });

    const processedKeys: Set<string> = new Set<string>();

    const missingAttachments: IMissingAttachment[] = [];

    submissions.forEach((submission: ISupplierSubmission): void => {
      const supplierTierKey: string = this.createSupplierTierKey(
        submission.supplierName,
        submission.tier,
      );

      if (processedKeys.has(supplierTierKey)) {
        return;
      }

      processedKeys.add(supplierTierKey);

      if (documentKeys.has(supplierTierKey)) {
        return;
      }

      missingAttachments.push({
        key: `missing-${submission.key}`,

        supplierName: submission.supplierName,

        tier: submission.tier,

        submissionId: submission.id,

        submissionDate: submission.created,

        qualification: submission.qualification,

        riskRating: submission.riskRating,

        recommendation: submission.recommendation,

        submissionItemUrl: submission.sourceItemUrl,
      });
    });

    return missingAttachments.sort(
      (
        first: IMissingAttachment,

        second: IMissingAttachment,
      ): number => {
        return (
          new Date(second.submissionDate).getTime() -
          new Date(first.submissionDate).getTime()
        );
      },
    );
  }

  private buildTierOverview(
    submissions: readonly ISupplierSubmission[],

    documents: readonly ISupplierDocument[],

    missingAttachments: readonly IMissingAttachment[],
  ): ITierAttachmentOverview[] {
    const tiers: readonly SupplierTier[] = ["Tier 1", "Tier 2", "Tier 3"];

    return tiers.map((tier: SupplierTier): ITierAttachmentOverview => {
      const tierSubmissions: ISupplierSubmission[] = submissions
        .filter((submission: ISupplierSubmission): boolean => {
          return submission.tier === tier;
        })
        .sort(
          (
            first: ISupplierSubmission,

            second: ISupplierSubmission,
          ): number => {
            return (
              new Date(second.created).getTime() -
              new Date(first.created).getTime()
            );
          },
        );

      const uniqueSubmissionSuppliers: Set<string> = new Set<string>();

      tierSubmissions.forEach((submission: ISupplierSubmission): void => {
        uniqueSubmissionSuppliers.add(
          this.createSupplierTierKey(submission.supplierName, tier),
        );
      });

      const suppliersWithDocuments: Set<string> = new Set<string>();

      documents
        .filter((document: ISupplierDocument): boolean => {
          return document.tier === tier;
        })
        .forEach((document: ISupplierDocument): void => {
          suppliersWithDocuments.add(
            this.createSupplierTierKey(document.supplierName, tier),
          );
        });

      const latestSubmission: ISupplierSubmission | undefined =
        tierSubmissions[0];

      return {
        tier,

        submissions: uniqueSubmissionSuppliers.size,

        suppliersWithDocuments: suppliersWithDocuments.size,

        missingAttachments: missingAttachments.filter(
          (item: IMissingAttachment): boolean => {
            return item.tier === tier;
          },
        ).length,

        latestSupplier: latestSubmission
          ? latestSubmission.supplierName
          : undefined,

        latestSubmissionDate: latestSubmission
          ? latestSubmission.created
          : undefined,

        latestQualification: latestSubmission
          ? latestSubmission.qualification
          : undefined,

        latestRiskRating: latestSubmission
          ? latestSubmission.riskRating
          : undefined,

        latestRecommendation: latestSubmission
          ? latestSubmission.recommendation
          : undefined,

        latestSubmissionUrl: latestSubmission
          ? latestSubmission.sourceItemUrl
          : undefined,
      };
    });
  }

  private buildFolderSummaries(
    documents: readonly ISupplierDocument[],
  ): ISupplierFolderSummary[] {
    const folderMap: Record<string, ISupplierFolderSummary> = {};

    documents.forEach((document: ISupplierDocument): void => {
      const key: string = `${document.tier}|${document.folderServerRelativeUrl}`;

      if (!folderMap[key]) {
        folderMap[key] = {
          key,

          supplierName: document.supplierName,

          tier: document.tier,

          libraryTitle: document.libraryTitle,

          documentCount: 0,

          totalFileSize: 0,

          latestUploadDate: undefined,

          latestModifiedBy: undefined,

          folderServerRelativeUrl: document.folderServerRelativeUrl,
        };
      }

      const folder: ISupplierFolderSummary = folderMap[key];

      folder.documentCount += 1;

      folder.totalFileSize += this.getSafeFileSize(document.fileSize);

      const existingLatestTime: number = folder.latestUploadDate
        ? this.getDateTime(folder.latestUploadDate)
        : 0;

      const documentTime: number = this.getDateTime(document.created);

      if (documentTime > existingLatestTime) {
        folder.latestUploadDate = document.created;

        folder.latestModifiedBy = document.modifiedBy;
      }
    });

    return Object.keys(folderMap)
      .map((key: string): ISupplierFolderSummary => {
        return folderMap[key];
      })
      .sort(
        (
          first: ISupplierFolderSummary,

          second: ISupplierFolderSummary,
        ): number => {
          const supplierComparison: number = first.supplierName.localeCompare(
            second.supplierName,
          );

          if (supplierComparison !== 0) {
            return supplierComparison;
          }

          return first.tier.localeCompare(second.tier);
        },
      );
  }

  private applyFilters = (): void => {
    const filters: ISupplierDocumentFilters = this.state.filters;

    const searchText: string = filters.searchText.trim().toLowerCase();

    const uploadedBy: string = (filters.uploadedBy || "").trim().toLowerCase();

    const startTime: number | undefined = filters.startDate
      ? this.getStartOfDay(filters.startDate).getTime()
      : undefined;

    const endTime: number | undefined = filters.endDate
      ? this.getEndOfDay(filters.endDate).getTime()
      : undefined;

    const filteredDocuments: ISupplierDocument[] =
      this.state.allDocuments.filter((document: ISupplierDocument): boolean => {
        const createdTime: number = this.getDateTime(document.created);

        const supplierName: string = (
          document.supplierName || ""
        ).toLowerCase();

        const fileName: string = (document.fileName || "").toLowerCase();

        const createdBy: string = (document.createdBy || "").toLowerCase();

        const matchesSearch: boolean =
          !searchText ||
          supplierName.indexOf(searchText) >= 0 ||
          fileName.indexOf(searchText) >= 0;

        const matchesTier: boolean =
          !filters.tier || document.tier === filters.tier;

        const matchesUploader: boolean =
          !uploadedBy || createdBy.indexOf(uploadedBy) >= 0;

        const matchesStartDate: boolean =
          startTime === undefined || createdTime >= startTime;

        const matchesEndDate: boolean =
          endTime === undefined || createdTime <= endTime;

        return (
          matchesSearch &&
          matchesTier &&
          matchesUploader &&
          matchesStartDate &&
          matchesEndDate
        );
      });

    this.setState({
      displayedDocuments: filteredDocuments,

      displayedFolders: this.buildFolderSummaries(filteredDocuments),
    });
  };

  private clearFilters = (): void => {
    this.setState({
      filters: {
        searchText: "",
      },

      displayedDocuments: this.state.allDocuments,

      displayedFolders: this.state.allFolders,
    });
  };

  private createSupplierTierKey(supplierName: string, tier: string): string {
    return `${tier}|${this.normalizeSupplierName(supplierName)}`;
  }

  private normalizeSupplierName(value: string): string {
    return (value || "")
      .trim()
      .toLowerCase()
      .replace(/&/g, "and")
      .replace(/[^a-z0-9]+/g, "");
  }

  private getSafeFileSize(value: number): number {
    return typeof value === "number" && isFinite(value) && value >= 0
      ? value
      : 0;
  }

  private getDateTime(value?: string): number {
    if (!value) {
      return 0;
    }

    const time: number = new Date(value).getTime();

    return isNaN(time) ? 0 : time;
  }

  private openFolder = (item: ISupplierFolderSummary): void => {
    this.openUrl(item.folderServerRelativeUrl);
  };

  private openDocument = (item: ISupplierDocument): void => {
    this.openUrl(item.fileServerRelativeUrl);
  };

  private openUrl = (url?: string): void => {
    if (!url) {
      return;
    }

    const normalizedUrl: string = url.trim();

    if (!normalizedUrl) {
      return;
    }

    const isAbsoluteUrl: boolean =
      normalizedUrl.toLowerCase().indexOf("http://") === 0 ||
      normalizedUrl.toLowerCase().indexOf("https://") === 0;

    const relativePrefix: string = normalizedUrl.charAt(0) === "/" ? "" : "/";

    const targetUrl: string = isAbsoluteUrl
      ? normalizedUrl
      : `${window.location.origin}${relativePrefix}${normalizedUrl}`;

    window.open(targetUrl, "_blank", "noopener,noreferrer");
  };

  private getStartOfDay(date: Date): Date {
    const value: Date = new Date(date.getTime());

    value.setHours(0, 0, 0, 0);

    return value;
  }

  private getEndOfDay(date: Date): Date {
    const value: Date = new Date(date.getTime());

    value.setHours(23, 59, 59, 999);

    return value;
  }
}
