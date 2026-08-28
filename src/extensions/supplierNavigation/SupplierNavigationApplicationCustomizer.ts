import * as React from 'react';
import * as ReactDom from 'react-dom';

import {
  Log
} from '@microsoft/sp-core-library';

import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import SupplierNavigation
  from './components/SupplierNavigation';

import {
  ISupplierNavigationProps
} from './components/ISupplierNavigationProps';

import SupplierFooter
  from './components/SupplierFooter';

import {
  ISupplierFooterProps
} from './components/SupplierFooter';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';

const LOG_SOURCE: string =
  'SupplierNavigationApplicationCustomizer';

export interface ISupplierNavigationApplicationCustomizerProperties {
  organizationName?: string;

  homeUrl?: string;
  completedQuestionnairesUrl?: string;
  supportingDocumentsUrl?: string;
  dashboardUrl?: string;
  helpUrl?: string;

  footerText?: string;
  privacyUrl?: string;
  accessibilityUrl?: string;
  confidentialityText: string;
  allowedSiteUrls?: string;
}

export default class SupplierNavigationApplicationCustomizer
  extends BaseApplicationCustomizer<
    ISupplierNavigationApplicationCustomizerProperties
  > {

  private topPlaceholder:
    PlaceholderContent | undefined;

  private bottomPlaceholder:
    PlaceholderContent | undefined;

  /**
   * Initializes the navigation and footer and listens for
   * SharePoint placeholder and client-side page changes.
   */
  public onInit(): Promise<void> {
    if (!this.isCurrentSiteAllowed()) {
      Log.warn(
        LOG_SOURCE,
        'Supplier Navigation is not enabled for this site.'
      );

      // Display visual error/warning banner in SharePoint UI
      this.renderUnauthorizedWarning();
      return Promise.resolve();
    }

    this.context
      .placeholderProvider
      .changedEvent
      .add(
        this,
        this.renderPlaceholders
      );

    this.context
      .application
      .navigatedEvent
      .add(
        this,
        this.handlePageNavigation
      );

    this.renderPlaceholders();

    return Promise.resolve();
  }

  /**
 * Renders an error banner if the extension is added to an unauthorized site.
 */
  private renderUnauthorizedWarning(): void {
    if (!this.topPlaceholder) {
      this.topPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top
        );
    }

    if (this.topPlaceholder && this.topPlaceholder.domElement) {
      const warningElement = React.createElement(
        MessageBar,
        {
          messageBarType: MessageBarType.error,
          isMultiline: false
        },
        'Supplier Navigation Extension is installed on an unauthorized site. Please contact your administrator.'
      );

      ReactDom.unmountComponentAtNode(this.topPlaceholder.domElement);
      ReactDom.render(warningElement, this.topPlaceholder.domElement);
    }
  }

  /**
   * Removes event subscriptions and unmounts the React
   * components when the extension is disposed.
   */
  public onDispose(): void {
    this.context
      .placeholderProvider
      .changedEvent
      .remove(
        this,
        this.renderPlaceholders
      );

    this.context
      .application
      .navigatedEvent
      .remove(
        this,
        this.handlePageNavigation
      );

    this.disposeTopPlaceholder();
    this.disposeBottomPlaceholder();
  }

  private isCurrentSiteAllowed(): boolean {
    const configuredUrls: string =
      this.properties.allowedSiteUrls
        ? this.properties.allowedSiteUrls.trim()
        : '';

    /*
     * No allowlist means site-level app installation
     * controls where the extension runs.
     */
    if (!configuredUrls) {
      return true;
    }

    const currentSiteUrl: string =
      this.normalizeSiteUrl(
        this.context.pageContext.web.absoluteUrl
      );

    const allowedSiteUrls: string[] =
      configuredUrls
        .split(';')
        .map(
          (siteUrl: string): string => {
            return this.normalizeSiteUrl(
              siteUrl
            );
          }
        )
        .filter(
          (siteUrl: string): boolean => {
            return Boolean(siteUrl);
          }
        );

    return allowedSiteUrls.indexOf(
      currentSiteUrl
    ) >= 0;
  }

  private normalizeSiteUrl(
    value: string
  ): string {
    return (
      value ||
      ''
    )
      .trim()
      .replace(
        /\/+$/,
        ''
      )
      .toLowerCase();
  }

  /**
   * Re-renders the navigation after modern SharePoint
   * client-side page navigation.
   */
  private handlePageNavigation =
    (): void => {
      this.renderPlaceholders();
    };

  /**
   * Renders both the Top navigation and Bottom footer.
   */
  private renderPlaceholders =
    (): void => {
      this.renderNavigation();
      this.renderFooter();
    };

  /**
   * Creates the Top placeholder and renders the
   * SupplierNavigation component.
   */

  private getOptionalPropertyValue(
    value: string | undefined
  ): string | undefined {
    const normalizedValue: string =
      value
        ? value.trim()
        : '';

    return normalizedValue
      ? normalizedValue
      : undefined;
  }

  private renderNavigation(): void {
    if (!this.topPlaceholder) {
      this.topPlaceholder =
        this.context
          .placeholderProvider
          .tryCreateContent(
            PlaceholderName.Top,
            {
              onDispose:
                this.handleTopPlaceholderDispose
            }
          );
    }

    if (
      !this.topPlaceholder ||
      !this.topPlaceholder.domElement
    ) {
      Log.warn(
        LOG_SOURCE,
        'The SharePoint Top placeholder is not available.'
      );

      return;
    }

    const webAbsoluteUrl: string =
      this.context
        .pageContext
        .web
        .absoluteUrl
        .replace(
          /\/$/,
          ''
        );

    /*const homeUrl: string =
      this.getPropertyValue(
        this.properties.homeUrl,
        `${webAbsoluteUrl}/SitePages/Home.aspx`
      );

    const completedQuestionnairesUrl: string =
      this.getPropertyValue(
        this.properties.completedQuestionnairesUrl,
        `${webAbsoluteUrl}` +
        '/SitePages/' +
        'Completed-Questionnaires.aspx'
      );

    const supportingDocumentsUrl: string =
      this.getPropertyValue(
        this.properties.supportingDocumentsUrl,
        `${webAbsoluteUrl}` +
        '/SitePages/' +
        'Supporting-Documents.aspx'
      );

    const dashboardUrl: string =
      this.getPropertyValue(
        this.properties.dashboardUrl,
        `${webAbsoluteUrl}/SitePages/Dashboard.aspx`
      );

    const helpUrl: string =
      this.getPropertyValue(
        this.properties.helpUrl,
        `${webAbsoluteUrl}/SitePages/Help.aspx`
      );*/

    /*
     * This object matches the SupplierNavigation component
     * currently used by the extension.
     */
    const navigationProps:
      ISupplierNavigationProps = {
      renderMode:
        'navigation',

      organizationName:
        this.getPropertyValue(
          this.properties.organizationName,
          'Cassava Technologies'
        ),

      homeUrl:
        this.getPropertyValue(
          this.properties.homeUrl,
          `${webAbsoluteUrl}/SitePages/Home.aspx`
        ),

      completedQuestionnairesUrl:
        this.getPropertyValue(
          this.properties.completedQuestionnairesUrl,
          `${webAbsoluteUrl}/SitePages/Completed-Questionnaires.aspx`
        ),

      supportingDocumentsUrl:
        this.getPropertyValue(
          this.properties.supportingDocumentsUrl,
          `${webAbsoluteUrl}/SitePages/Supporting-Documents.aspx`
        ),

      dashboardUrl:
        this.getPropertyValue(
          this.properties.dashboardUrl,
          `${webAbsoluteUrl}/SitePages/Dashboard.aspx`
        ),

      helpUrl:
        this.getPropertyValue(
          this.properties.helpUrl,
          `${webAbsoluteUrl}/SitePages/Help.aspx`
        ),

      currentPath:
        window.location.href,

      footerText:
        this.getPropertyValue(
          this.properties.footerText,
          'Supplier ESG questionnaire and sustainability portal'
        ),

      privacyUrl:
        this.getOptionalPropertyValue(
          this.properties.privacyUrl
        ),

      accessibilityUrl:
        this.getOptionalPropertyValue(
          this.properties.accessibilityUrl
        )
    };

    const navigationElement:
      React.ReactElement<
        ISupplierNavigationProps
      > = React.createElement(
        SupplierNavigation,
        navigationProps
      );

    ReactDom.render(
      navigationElement,
      this.topPlaceholder.domElement
    );
  }

  /**
   * Creates the Bottom placeholder and renders the
   * SupplierFooter component.
   */
  private renderFooter(): void {
    if (!this.bottomPlaceholder) {
      this.bottomPlaceholder =
        this.context
          .placeholderProvider
          .tryCreateContent(
            PlaceholderName.Bottom,
            {
              onDispose:
                this.handleBottomPlaceholderDispose
            }
          );
    }

    if (
      !this.bottomPlaceholder ||
      !this.bottomPlaceholder.domElement
    ) {
      Log.warn(
        LOG_SOURCE,
        'The SharePoint Bottom placeholder is not available.'
      );

      return;
    }

    const footerProps:
      ISupplierFooterProps = {
      organizationName:
        this.getPropertyValue(
          this.properties.organizationName,
          'Cassava Technologies'
        ),

      confidentialityText:
        this.getPropertyValue(
          this.properties.confidentialityText,
          'All information is confidential and must be used ' +
          'in accordance with our privacy and security policies.'
        )
    };

    const footerElement:
      React.ReactElement<
        ISupplierFooterProps
      > = React.createElement(
        SupplierFooter,
        footerProps
      );

    ReactDom.render(
      footerElement,
      this.bottomPlaceholder.domElement
    );
  }

  /**
   * Called when SharePoint disposes the Top placeholder.
   */
  private handleTopPlaceholderDispose =
    (): void => {
      this.disposeTopPlaceholder();
    };

  /**
   * Called when SharePoint disposes the Bottom placeholder.
   */
  private handleBottomPlaceholderDispose =
    (): void => {
      this.disposeBottomPlaceholder();
    };

  /**
   * Unmounts the navigation from the Top placeholder.
   */
  private disposeTopPlaceholder(): void {
    if (
      this.topPlaceholder &&
      this.topPlaceholder.domElement
    ) {
      ReactDom.unmountComponentAtNode(
        this.topPlaceholder.domElement
      );
    }

    this.topPlaceholder =
      undefined;
  }

  /**
   * Unmounts the footer from the Bottom placeholder.
   */
  private disposeBottomPlaceholder(): void {
    if (
      this.bottomPlaceholder &&
      this.bottomPlaceholder.domElement
    ) {
      ReactDom.unmountComponentAtNode(
        this.bottomPlaceholder.domElement
      );
    }

    this.bottomPlaceholder =
      undefined;
  }

  /**
   * Returns a trimmed extension property or its fallback.
   */
  private getPropertyValue(
    value: string | undefined,
    fallbackValue: string
  ): string {
    const normalizedValue: string =
      value
        ? value.trim()
        : '';

    return (
      normalizedValue ||
      fallbackValue
    );
  }
}