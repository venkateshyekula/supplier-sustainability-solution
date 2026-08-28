import * as React from 'react';
import { Icon } from '@fluentui/react';
import { IHelpPageProps } from './IHelpPageProps';
import styles from './HelpPage.module.scss';

interface IGettingStartedCard {
  key: string;
  iconName: string;
  title: string;
  description: string;
  color: string;
  backgroundColor: string;
}

interface IResourceCard {
  key: string;
  iconName: string;
  title: string;
  items: readonly string[];
  color: string;
  backgroundColor: string;
  url?: string;
}

interface IQuickLink {
  key: string;
  iconName: string;
  title: string;
  description: string;
  color: string;
  backgroundColor: string;
  url?: string;
}

interface IFrequentlyAskedQuestion {
  key: string;
  question: string;
  answer: string;
}

interface IHelpPageState {
  expandedFaqKey?: string;
}

// eslint-disable-next-line @typescript-eslint/no-var-requires
const supportBanner: string = require('../assets/support-banner.png');

const GETTING_STARTED_CARDS: readonly IGettingStartedCard[] = [
  {
    key: 'receiveInvitation',
    iconName: 'Mail',
    title: 'Receive Invitation',
    description:
      'Suppliers receive a secure questionnaire link directly from Procurement.',
    color: '#3568d4',
    backgroundColor: '#edf3ff'
  },
  {
    key: 'selectQuestionnaire',
    iconName: 'Questionnaire',
    title: 'Select Questionnaire',
    description:
      'Open the questionnaire Tier assigned to the supplier by Procurement.',
    color: '#6750c9',
    backgroundColor: '#f1eeff'
  },
  {
    key: 'completeSurvey',
    iconName: 'EditNote',
    title: 'Complete Survey',
    description:
      'Answer all questions accurately and provide the requested information.',
    color: '#287bb5',
    backgroundColor: '#eaf6ff'
  },
  {
    key: 'submitResponses',
    iconName: 'CloudUpload',
    title: 'Submit Responses',
    description:
      'Review the responses and submit the completed questionnaire.',
    color: '#5b56c8',
    backgroundColor: '#f0efff'
  },
  {
    key: 'trackStatus',
    iconName: 'Completed',
    title: 'Track Status',
    description:
      'Procurement and ESG reviewers monitor the submitted assessment.',
    color: '#4771c9',
    backgroundColor: '#eef3ff'
  }
];

const FAQ_ITEMS: readonly IFrequentlyAskedQuestion[] = [
  {
    key: 'invitation',
    question:
      'How do I know which Tier questionnaire applies to my company?',
    answer:
      'Procurement will send the supplier a direct link to the appropriate Tier questionnaire. Suppliers should use only the link provided by Procurement.'
  },
  {
    key: 'progress',
    question:
      'Can I save my progress and return to the survey later?',
    answer:
      'The ability to save and return depends on the questionnaire configuration. Review the form instructions before closing the questionnaire and contact support if progress is not retained.'
  },
  {
    key: 'documents',
    question:
      'What documents are required to complete the questionnaire?',
    answer:
      'Required documents depend on the supplier Tier and questionnaire responses. Typical supporting evidence may include policies, certifications, compliance records, environmental reports, and governance documentation.'
  },
  {
    key: 'submitted',
    question:
      'How will I be notified about my survey status?',
    answer:
      'Procurement or the ESG review team will communicate follow-up actions. Internal users can monitor qualification, risk rating, recommendations, and supporting-document status through the dashboard.'
  },
  {
    key: 'changes',
    question:
      'Can a submitted questionnaire be changed?',
    answer:
      'Contact Procurement or the support team if submitted information must be corrected. Do not submit a duplicate questionnaire unless instructed to do so.'
  }
];

export default class HelpPage extends React.Component<
  IHelpPageProps,
  IHelpPageState
> {
  public constructor(props: IHelpPageProps) {
    super(props);

    this.state = {
      expandedFaqKey: undefined
    };
  }

  public render(): React.ReactElement<IHelpPageProps> {
    const resourceCards: readonly IResourceCard[] =
      this.getResourceCards();

    const quickLinks: readonly IQuickLink[] =
      this.getQuickLinks();

    const supportEmail: string | undefined =
      this.getTrimmedValue(this.props.supportEmail);

    const procurementEmail: string | undefined =
      this.getTrimmedValue(this.props.procurementEmail);

    const supportPhone: string | undefined =
      this.getTrimmedValue(this.props.supportPhone);

    const supportHours: string | undefined =
      this.getTrimmedValue(this.props.supportHours);

    return (
      <section
        className={styles.helpPage}
        aria-label="Supplier Sustainability help and support"
      >
        <div className={styles.contentLayout}>
          <main className={styles.mainColumn}>
            <header
              className={styles.heroBanner}
              style={{ backgroundImage: `url(${supportBanner})` }}
            >
              <div className={styles.heroContent}>
                <h1 className={styles.heroTitle}>
                  Help &amp; Support
                </h1>

                <p className={styles.heroDescription}>
                  Find guidance, instructions, and support for submitting and managing
                  Supplier Sustainability Surveys.
                </p>
              </div>
            </header>

            <section
              className={styles.section}
              aria-labelledby="getting-started-title"
            >
              <div className={styles.sectionHeader}>
                <h2 id="getting-started-title">
                  Getting Started
                </h2>

                <p>
                  New to the Supplier Sustainability
                  Survey? Follow these steps to get
                  started.
                </p>
              </div>

              <div
                className={styles.stepProgress}
                aria-label="Getting started steps"
              >
                {GETTING_STARTED_CARDS.map(
                  (
                    card: IGettingStartedCard,
                    index: number
                  ): React.ReactElement => {
                    const stepNumber: number = index + 1;

                    return (
                      <React.Fragment key={card.key}>
                        <span
                          className={styles.stepProgressNumber}
                          aria-label={`Step ${stepNumber.toString()}`}
                        >
                          {stepNumber}
                        </span>

                        {index < GETTING_STARTED_CARDS.length - 1 && (
                          <span
                            className={styles.stepProgressLine}
                            aria-hidden="true"
                          />
                        )}
                      </React.Fragment>
                    );
                  }
                )}
              </div>

              <div className={styles.gettingStartedGrid}>
                {GETTING_STARTED_CARDS.map(
                  (card: IGettingStartedCard): React.ReactElement => {
                    return (
                      <article
                        key={card.key}
                        className={styles.gettingStartedCard}
                      >
                        <span
                          className={styles.gettingStartedIcon}
                          style={{
                            color: card.color,
                            backgroundColor: card.backgroundColor
                          }}
                          aria-hidden="true"
                        >
                          <Icon iconName={card.iconName} />
                        </span>

                        <h3>{card.title}</h3>

                        <p>{card.description}</p>
                      </article>
                    );
                  }
                )}
              </div>
            </section>

            <section
              className={styles.section}
              aria-labelledby="guides-resources-title"
            >
              <div className={styles.sectionHeader}>
                <h2 id="guides-resources-title">
                  Guides &amp; Resources
                </h2>

                <p>
                  Access detailed guides and resources
                  to help complete the survey.
                </p>
              </div>

              <div className={styles.resourceGrid}>
                {resourceCards.map(
                  (resource: IResourceCard): React.ReactElement => {
                    const hasResourceUrl: boolean = this.hasValue(
                      resource.url
                    );

                    return (
                      <article
                        key={resource.key}
                        className={styles.resourceCard}
                        style={{
                          backgroundColor: resource.backgroundColor
                        }}
                      >
                        <div className={styles.resourceHeader}>
                          <span
                            className={styles.resourceIcon}
                            style={{
                              color: resource.color
                            }}
                            aria-hidden="true"
                          >
                            <Icon iconName={resource.iconName} />
                          </span>

                          <h3>{resource.title}</h3>
                        </div>

                        <ul>
                          {resource.items.map(
                            (item: string): React.ReactElement => {
                              return <li key={item}>{item}</li>;
                            }
                          )}
                        </ul>

                        <button
                          type="button"
                          className={styles.resourceAction}
                          disabled={!hasResourceUrl}
                          title={
                            hasResourceUrl
                              ? `Open ${resource.title}`
                              : `${resource.title} is not configured`
                          }
                          onClick={(): void => {
                            this.openUrl(resource.url);
                          }}
                        >
                          <span>View Guide</span>

                          <Icon
                            iconName="ChevronRight"
                            aria-hidden="true"
                          />
                        </button>
                      </article>
                    );
                  }
                )}
              </div>
            </section>

            <section
              className={styles.section}
              aria-labelledby="faq-title"
            >
              <div className={styles.faqHeading}>
                <div className={styles.sectionHeader}>
                  <h2 id="faq-title">
                    Frequently Asked Questions
                  </h2>

                  <p>
                    Common questions about supplier
                    sustainability questionnaires.
                  </p>
                </div>
              </div>

              <div className={styles.faqList}>
                {FAQ_ITEMS.map(
                  (
                    faq: IFrequentlyAskedQuestion
                  ): React.ReactElement => {
                    const isExpanded: boolean =
                      this.state.expandedFaqKey === faq.key;

                    const answerId: string = `faq-answer-${faq.key}`;

                    return (
                      <article key={faq.key} className={styles.faqItem}>
                        <button
                          type="button"
                          className={styles.faqQuestion}
                          aria-expanded={isExpanded}
                          aria-controls={answerId}
                          onClick={(): void => {
                            this.toggleFaq(faq.key);
                          }}
                        >
                          <span>{faq.question}</span>

                          <Icon
                            iconName={
                              isExpanded ? 'ChevronUp' : 'ChevronDown'
                            }
                            aria-hidden="true"
                          />
                        </button>

                        {isExpanded && (
                          <div
                            id={answerId}
                            className={styles.faqAnswer}
                          >
                            {faq.answer}
                          </div>
                        )}
                      </article>
                    );
                  }
                )}
              </div>

              <div className={styles.faqFooter}>
                <button
                  type="button"
                  className={styles.viewAllFaq}
                  disabled={!this.hasValue(this.props.faqPageUrl)}
                  onClick={(): void => {
                    this.openUrl(this.props.faqPageUrl);
                  }}
                >
                  <span>View All FAQs</span>

                  <Icon iconName="Forward" aria-hidden="true" />
                </button>
              </div>
            </section>
          </main>

          <aside
            className={styles.sideColumn}
            aria-label="Support options"
          >
            <section className={styles.supportCard}>
              <h2>Need Additional Help?</h2>

              <p className={styles.supportIntroduction}>
                Our support team is here to help with
                any questions or issues.
              </p>

              <div className={styles.supportItem}>
                <span
                  className={`${styles.supportIcon} ${styles.emailIcon}`}
                  aria-hidden="true"
                >
                  <Icon iconName="Mail" />
                </span>

                <div>
                  <h3>Email Support</h3>

                  {supportEmail && (
                    <a href={`mailto:${supportEmail}`}>
                      {supportEmail}
                    </a>
                  )}

                  {procurementEmail && (
                    <a href={`mailto:${procurementEmail}`}>
                      {procurementEmail}
                    </a>
                  )}

                  {!supportEmail && !procurementEmail && (
                    <p className={styles.supportUnavailable}>
                      Email support is not configured.
                    </p>
                  )}
                </div>
              </div>

              <div className={styles.supportItem}>
                <span
                  className={`${styles.supportIcon} ${styles.phoneIcon}`}
                  aria-hidden="true"
                >
                  <Icon iconName="Phone" />
                </span>

                <div>
                  <h3>Phone Support</h3>

                  {supportPhone ? (
                    <a href={`tel:${this.getTelephoneValue(supportPhone)}`}>
                      {supportPhone}
                    </a>
                  ) : (
                    <p className={styles.supportUnavailable}>
                      Phone support is not configured.
                    </p>
                  )}
                </div>
              </div>

              <div className={styles.supportItem}>
                <span
                  className={`${styles.supportIcon} ${styles.hoursIcon}`}
                  aria-hidden="true"
                >
                  <Icon iconName="Clock" />
                </span>

                <div>
                  <h3>Support Hours</h3>

                  <p>
                    {supportHours || 'Support hours are not configured.'}
                  </p>
                </div>
              </div>
            </section>

            <section className={styles.quickLinksCard}>
              <h2>Quick Links</h2>

              <div className={styles.quickLinks}>
                {quickLinks.map(
                  (link: IQuickLink): React.ReactElement => {
                    const isEnabled: boolean = this.hasValue(link.url);

                    return (
                      <button
                        key={link.key}
                        type="button"
                        className={styles.quickLinkItem}
                        disabled={!isEnabled}
                        title={
                          isEnabled
                            ? link.title
                            : `${link.title} is not configured`
                        }
                        onClick={(): void => {
                          this.openUrl(link.url);
                        }}
                      >
                        <span
                          className={styles.quickLinkIcon}
                          style={{
                            color: link.color,
                            backgroundColor: link.backgroundColor
                          }}
                          aria-hidden="true"
                        >
                          <Icon iconName={link.iconName} />
                        </span>

                        <span className={styles.quickLinkContent}>
                          <strong>{link.title}</strong>
                          <small>{link.description}</small>
                        </span>

                        <Icon
                          iconName="ChevronRight"
                          className={styles.quickLinkArrow}
                          aria-hidden="true"
                        />
                      </button>
                    );
                  }
                )}
              </div>
            </section>

            <section className={styles.reportIssueCard}>
              <div className={styles.reportIssueContent}>
                <h2>Report an Issue</h2>

                <p>
                  Experiencing a problem? Contact
                  the support team through the
                  service portal.
                </p>

                <button
                  type="button"
                  className={styles.reportIssueButton}
                  disabled={!this.hasValue(this.props.serviceNowUrl)}
                  onClick={(): void => {
                    this.openUrl(this.props.serviceNowUrl);
                  }}
                >
                  <Icon iconName="Ticket" aria-hidden="true" />
                  <span>Create Support Ticket</span>
                </button>
              </div>

              <div
                className={styles.reportIssueIllustration}
                aria-hidden="true"
              >
                <Icon iconName="ClipboardList" />
              </div>
            </section>

            <section className={styles.importantNote} role="note">
              <span
                className={styles.importantNoteIcon}
                aria-hidden="true"
              >
                <Icon iconName="Info" />
              </span>

              <div>
                <h2>Important Note</h2>

                <p>
                  Please ensure all information
                  provided is accurate and up to
                  date. Incomplete submissions may
                  delay review and approval.
                </p>
              </div>
            </section>
          </aside>
        </div>
      </section>
    );
  }

  private getResourceCards(): readonly IResourceCard[] {
    return [
      {
        key: 'supplierGuide',
        iconName: 'Contact',
        title: 'For Suppliers',
        items: [
          'Survey overview',
          'Required information',
          'Best practices',
          'Help and support'
        ],
        color: '#376bc7',
        backgroundColor: '#eef5ff',
        url: this.props.supplierGuideUrl
      },
      {
        key: 'procurementGuide',
        iconName: 'ClipboardList',
        title: 'For Procurement',
        items: [
          'Supplier onboarding',
          'Tier classification',
          'Evaluation process',
          'Corrective actions'
        ],
        color: '#248253',
        backgroundColor: '#eefbf4',
        url: this.props.procurementGuideUrl
      },
      {
        key: 'compliance',
        iconName: 'ComplianceAudit',
        title: 'Compliance & Policies',
        items: [
          'ESG policies',
          'Data privacy',
          'Code of conduct',
          'Reporting standards'
        ],
        color: '#c98419',
        backgroundColor: '#fff8e9',
        url: this.props.policyComplianceUrl
      },
      {
        key: 'training',
        iconName: 'Education',
        title: 'Training Videos',
        items: [
          'How to complete a survey',
          'Upload documents',
          'Track submission status'
        ],
        color: '#6c4ac7',
        backgroundColor: '#f5f0ff',
        url: this.props.userGuideUrl
      }
    ];
  }

  private getQuickLinks(): readonly IQuickLink[] {
    return [
      {
        key: 'guidelines',
        iconName: 'OpenBook',
        title: 'View ESG Guidelines',
        description: 'Download the latest assessment guidelines',
        color: '#4b61c2',
        backgroundColor: '#eef1ff',
        url: this.props.esgGuidelinesUrl
      },
      {
        key: 'userGuide',
        iconName: 'DownloadDocument',
        title: 'Download User Guide',
        description: 'Step-by-step instructions',
        color: '#3878c8',
        backgroundColor: '#edf6ff',
        url: this.props.userGuideUrl
      },
      {
        key: 'supplierPortal',
        iconName: 'NavigateExternalInline',
        title: 'Supplier Portal Access',
        description: 'Go to the supplier portal',
        color: '#258254',
        backgroundColor: '#edf9f2',
        url: this.props.supplierPortalUrl
      },
      {
        key: 'policies',
        iconName: 'Shield',
        title: 'Policy & Compliance',
        description: 'Review policies and compliance guidance',
        color: '#c57a14',
        backgroundColor: '#fff7e8',
        url: this.props.policyComplianceUrl
      }
    ];
  }

  private toggleFaq(faqKey: string): void {
    this.setState((previousState: IHelpPageState): IHelpPageState => {
      return {
        expandedFaqKey:
          previousState.expandedFaqKey === faqKey ? undefined : faqKey
      };
    });
  }

  private hasValue(value?: string): boolean {
    return Boolean(value && value.trim());
  }

  private getTrimmedValue(value?: string): string | undefined {
    if (!this.hasValue(value)) {
      return undefined;
    }

    return (value || '').trim();
  }

  private openUrl(url?: string): void {
    const normalizedUrl: string | undefined = this.getTrimmedValue(url);

    if (!normalizedUrl) {
      return;
    }

    const lowerUrl: string = normalizedUrl.toLowerCase();

    const isHttpUrl: boolean =
      lowerUrl.indexOf('http://') === 0 ||
      lowerUrl.indexOf('https://') === 0;

    const isMailUrl: boolean = lowerUrl.indexOf('mailto:') === 0;

    const isTelephoneUrl: boolean = lowerUrl.indexOf('tel:') === 0;

    if (isMailUrl || isTelephoneUrl) {
      window.location.href = normalizedUrl;
      return;
    }

    let targetUrl: string;

    if (isHttpUrl) {
      targetUrl = normalizedUrl;
    } else {
      const relativePath: string =
        normalizedUrl.charAt(0) === '/'
          ? normalizedUrl
          : `/${normalizedUrl}`;

      targetUrl = `${window.location.origin}${relativePath}`;
    }

    window.open(targetUrl, '_blank', 'noopener,noreferrer');
  }

  private getTelephoneValue(value: string): string {
    return value.replace(/[^+\d]/g, '');
  }
}