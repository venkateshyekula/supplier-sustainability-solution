import * as React from 'react';
import * as ReactDom from 'react-dom';

import {
  Version
} from '@microsoft/sp-core-library';

import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';

import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import HelpPage from './components/HelpPage';

import {
  IHelpPageProps
} from './components/IHelpPageProps';

import {
  HelpDocumentService
} from './services/HelpDocumentService';

import {
  IHelpDocumentLinks
} from './models/IHelpDocumentLinks';

export interface IHelpPageWebPartProps {
  helpDocumentLibraryTitle: string;

  supportEmail: string;
  procurementEmail: string;
  supportPhone: string;
  supportHours: string;

  serviceNowUrl: string;

  /*
   * These properties are optional fallback values.
   * PnPConfig/service values take precedence.
   */
  esgGuidelinesUrl: string;
  userGuideUrl: string;
  supplierPortalUrl: string;
  policyComplianceUrl: string;

  supplierGuideUrl: string;
  procurementGuideUrl: string;

  faqPageUrl: string;
}

const DEFAULT_HELP_DOCUMENT_LIBRARY_TITLE:
  string =
    'Supplier Sustainability Resources';

const DEFAULT_SUPPORT_EMAIL:
  string =
    'support@liquid.tech';

const DEFAULT_PROCUREMENT_EMAIL:
  string =
    'GroupProcurement@cassava.tech';

const DEFAULT_SUPPORT_PHONE:
  string =
    '+27 11 585 0000';

const DEFAULT_SUPPORT_HOURS:
  string =
    'Mon - Fri, 8:30 AM - 5:30 PM (GMT/SAST)';

const DEFAULT_SERVICE_NOW_URL:
  string =
    'https://oneliquidsupport.service-now.com/' +
    'staffportal?id=itsm_index';

const EMPTY_DOCUMENT_LINKS:
  IHelpDocumentLinks = {
    esgGuidelinesUrl: '',
    userGuideUrl: '',
    supplierPortalUrl: '',
    policyComplianceUrl: '',
    supplierGuideUrl: '',
    procurementGuideUrl: '',
    faqPageUrl: ''
  };

export default class HelpPageWebPart
  extends BaseClientSideWebPart<
    IHelpPageWebPartProps
  > {

  private documentLinks:
    IHelpDocumentLinks =
      EMPTY_DOCUMENT_LINKS;

  private isDocumentLinksLoading:
    boolean = true;

  private documentLinksError?:
    string;

  protected async onInit():
    Promise<void> {
    await super.onInit();

    this.setDefaultProperties();

    await this.loadDocumentLinks();
  }

  public render(): void {
    const componentProperties:
      IHelpPageProps = {
        context:
          this.context,

        supportEmail:
          this.getValue(
            this.properties.supportEmail,
            DEFAULT_SUPPORT_EMAIL
          ),

        procurementEmail:
          this.getValue(
            this.properties.procurementEmail,
            DEFAULT_PROCUREMENT_EMAIL
          ),

        supportPhone:
          this.getValue(
            this.properties.supportPhone,
            DEFAULT_SUPPORT_PHONE
          ),

        supportHours:
          this.getValue(
            this.properties.supportHours,
            DEFAULT_SUPPORT_HOURS
          ),

        serviceNowUrl:
          this.getValue(
            this.properties.serviceNowUrl,
            DEFAULT_SERVICE_NOW_URL
          ),

        /*
         * Dynamically resolved links take precedence.
         * Property-pane values are fallback values.
         */
        esgGuidelinesUrl:
          this.getResolvedUrl(
            this.documentLinks
              .esgGuidelinesUrl,
            this.properties
              .esgGuidelinesUrl
          ),

        userGuideUrl:
          this.getResolvedUrl(
            this.documentLinks
              .userGuideUrl,
            this.properties
              .userGuideUrl
          ),

        supplierPortalUrl:
          this.getResolvedUrl(
            this.documentLinks
              .supplierPortalUrl,
            this.properties
              .supplierPortalUrl
          ),

        policyComplianceUrl:
          this.getResolvedUrl(
            this.documentLinks
              .policyComplianceUrl,
            this.properties
              .policyComplianceUrl
          ),

        supplierGuideUrl:
          this.getResolvedUrl(
            this.documentLinks
              .supplierGuideUrl,
            this.properties
              .supplierGuideUrl
          ),

        procurementGuideUrl:
          this.getResolvedUrl(
            this.documentLinks
              .procurementGuideUrl,
            this.properties
              .procurementGuideUrl
          ),

        faqPageUrl:
          this.getResolvedUrl(
            this.documentLinks
              .faqPageUrl,
            this.properties
              .faqPageUrl
          ),

        documentLinksLoading:
          this.isDocumentLinksLoading,

        documentLinksError:
          this.documentLinksError
      };

    const element:
      React.ReactElement<
        IHelpPageProps
      > =
        React.createElement(
          HelpPage,
          componentProperties
        );

    ReactDom.render(
      element,
      this.domElement
    );
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(
      this.domElement
    );
  }

  protected get dataVersion():
    Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration():
    IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description:
              'Configure Help page contacts, ' +
              'resource library, fallback URLs, ' +
              'and support options.'
          },

          groups: [
            {
              groupName:
                'Dynamic document configuration',

              groupFields: [
                PropertyPaneTextField(
                  'helpDocumentLibraryTitle',
                  {
                    label:
                      'Help document library title',

                    description:
                      'SharePoint library used by the ' +
                      'dynamic Help document service.'
                  }
                )
              ]
            },

            {
              groupName:
                'Support contacts',

              groupFields: [
                PropertyPaneTextField(
                  'supportEmail',
                  {
                    label:
                      'Support email'
                  }
                ),

                PropertyPaneTextField(
                  'procurementEmail',
                  {
                    label:
                      'Procurement email'
                  }
                ),

                PropertyPaneTextField(
                  'supportPhone',
                  {
                    label:
                      'Support phone'
                  }
                ),

                PropertyPaneTextField(
                  'supportHours',
                  {
                    label:
                      'Support hours'
                  }
                ),

                PropertyPaneTextField(
                  'serviceNowUrl',
                  {
                    label:
                      'Support ticket URL'
                  }
                )
              ]
            },

            {
              groupName:
                'Fallback quick links',

              groupFields: [
                PropertyPaneTextField(
                  'esgGuidelinesUrl',
                  {
                    label:
                      'ESG Guidelines fallback URL'
                  }
                ),

                PropertyPaneTextField(
                  'userGuideUrl',
                  {
                    label:
                      'User Guide fallback URL'
                  }
                ),

                PropertyPaneTextField(
                  'supplierPortalUrl',
                  {
                    label:
                      'Supplier Portal fallback URL'
                  }
                ),

                PropertyPaneTextField(
                  'policyComplianceUrl',
                  {
                    label:
                      'Policy and Compliance fallback URL'
                  }
                )
              ]
            },

            {
              groupName:
                'Fallback guides and FAQ',

              groupFields: [
                PropertyPaneTextField(
                  'supplierGuideUrl',
                  {
                    label:
                      'Supplier Guide fallback URL'
                  }
                ),

                PropertyPaneTextField(
                  'procurementGuideUrl',
                  {
                    label:
                      'Procurement Guide fallback URL'
                  }
                ),

                PropertyPaneTextField(
                  'faqPageUrl',
                  {
                    label:
                      'FAQ page fallback URL'
                  }
                )
              ]
            }
          ]
        }
      ]
    };
  }

  private async loadDocumentLinks():
    Promise<void> {
    this.isDocumentLinksLoading =
      true;

    this.documentLinksError =
      undefined;

    /*
     * Render the loading state immediately.
     */
    this.render();

    try {
      const service:
        HelpDocumentService =
          new HelpDocumentService(
            this.context
          );

      const libraryTitle:
        string =
          this.getValue(
            this.properties
              .helpDocumentLibraryTitle,
            DEFAULT_HELP_DOCUMENT_LIBRARY_TITLE
          );

      const resolvedLinks:
        IHelpDocumentLinks =
          await service
            .getDocumentLinks(
              libraryTitle
            );

      this.documentLinks = {
        esgGuidelinesUrl:
          this.getValue(
            resolvedLinks
              .esgGuidelinesUrl,
            ''
          ),

        userGuideUrl:
          this.getValue(
            resolvedLinks
              .userGuideUrl,
            ''
          ),

        supplierPortalUrl:
          this.getValue(
            resolvedLinks
              .supplierPortalUrl,
            ''
          ),

        policyComplianceUrl:
          this.getValue(
            resolvedLinks
              .policyComplianceUrl,
            ''
          ),

        supplierGuideUrl:
          this.getValue(
            resolvedLinks
              .supplierGuideUrl,
            ''
          ),

        procurementGuideUrl:
          this.getValue(
            resolvedLinks
              .procurementGuideUrl,
            ''
          ),

        faqPageUrl:
          this.getValue(
            resolvedLinks
              .faqPageUrl,
            ''
          )
      };
    } catch (
      error: unknown
    ) {
      this.documentLinks =
        EMPTY_DOCUMENT_LINKS;

      this.documentLinksError =
        error instanceof Error
          ? error.message
          : 'Unable to load Help documents.';
    } finally {
      this.isDocumentLinksLoading =
        false;

      this.render();
    }
  }

  private setDefaultProperties(): void {
    if (
      !this.hasValue(
        this.properties
          .helpDocumentLibraryTitle
      )
    ) {
      this.properties
        .helpDocumentLibraryTitle =
          DEFAULT_HELP_DOCUMENT_LIBRARY_TITLE;
    }

    if (
      !this.hasValue(
        this.properties.supportEmail
      )
    ) {
      this.properties.supportEmail =
        DEFAULT_SUPPORT_EMAIL;
    }

    if (
      !this.hasValue(
        this.properties
          .procurementEmail
      )
    ) {
      this.properties
        .procurementEmail =
          DEFAULT_PROCUREMENT_EMAIL;
    }

    if (
      !this.hasValue(
        this.properties.supportPhone
      )
    ) {
      this.properties.supportPhone =
        DEFAULT_SUPPORT_PHONE;
    }

    if (
      !this.hasValue(
        this.properties.supportHours
      )
    ) {
      this.properties.supportHours =
        DEFAULT_SUPPORT_HOURS;
    }

    if (
      !this.hasValue(
        this.properties.serviceNowUrl
      )
    ) {
      this.properties.serviceNowUrl =
        DEFAULT_SERVICE_NOW_URL;
    }

    this.initializeOptionalProperty(
      'esgGuidelinesUrl'
    );

    this.initializeOptionalProperty(
      'userGuideUrl'
    );

    this.initializeOptionalProperty(
      'supplierPortalUrl'
    );

    this.initializeOptionalProperty(
      'policyComplianceUrl'
    );

    this.initializeOptionalProperty(
      'supplierGuideUrl'
    );

    this.initializeOptionalProperty(
      'procurementGuideUrl'
    );

    this.initializeOptionalProperty(
      'faqPageUrl'
    );
  }

  private initializeOptionalProperty(
    propertyName:
      keyof IHelpPageWebPartProps
  ): void {
    if (
      this.properties[
        propertyName
      ] === undefined
    ) {
      this.properties[
        propertyName
      ] = '';
    }
  }

  private getResolvedUrl(
    dynamicUrl:
      string | undefined,

    fallbackUrl:
      string | undefined
  ): string {
    if (
      this.hasValue(
        dynamicUrl
      )
    ) {
      return (
        dynamicUrl ||
        ''
      ).trim();
    }

    return this.getValue(
      fallbackUrl,
      ''
    );
  }

  private getValue(
    value:
      string | undefined,

    fallbackValue:
      string
  ): string {
    const normalizedValue:
      string =
        value
          ? value.trim()
          : '';

    return (
      normalizedValue ||
      fallbackValue
    );
  }

  private hasValue(
    value:
      string | undefined
  ): boolean {
    return Boolean(
      value &&
      value.trim()
    );
  }
}