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
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import EsgDashboard
  from './components/EsgDashboard';

import {
  IEsgDashboardProps
} from './components/IEsgDashboardProps';

export interface IEsgDashboardWebPartProps {
  tier1ListTitle: string;
  tier2ListTitle: string;
  tier3ListTitle: string;

  tier1SupplierNameField: string;
  tier2SupplierNameField: string;
  tier3SupplierNameField: string;

  tier1DocumentLibraryTitle: string;
  tier2DocumentLibraryTitle: string;
  tier3DocumentLibraryTitle: string;

  completedQuestionnairesUrl: string;
  supportingDocumentsUrl: string;

  targetScore: number;
}

export default class EsgDashboardWebPart
  extends BaseClientSideWebPart<
    IEsgDashboardWebPartProps
  > {

  public render(): void {
    const componentProps:
      IEsgDashboardProps = {
        context:
          this.context,

        tier1ListTitle:
          this.getRequiredTextProperty(
            this.properties.tier1ListTitle,
            'CASSTECH_SSQ'
          ),

        tier2ListTitle:
          this.getRequiredTextProperty(
            this.properties.tier2ListTitle,
            'Tier 2 ESG Procurement Questionnaire'
          ),

        tier3ListTitle:
          this.getRequiredTextProperty(
            this.properties.tier3ListTitle,
            'Supplier Sustainability Questionnaires Tier 3'
          ),

        tier1SupplierNameField:
          this.getRequiredTextProperty(
            this.properties.tier1SupplierNameField,
            'Supplier_x0020_Name'
          ),

        tier2SupplierNameField:
          this.getRequiredTextProperty(
            this.properties.tier2SupplierNameField,
            'Supplier_x0020_Name'
          ),

        tier3SupplierNameField:
          this.getRequiredTextProperty(
            this.properties.tier3SupplierNameField,
            'Supplier_x0020_Name'
          ),

        tier1DocumentLibraryTitle:
          this.getRequiredTextProperty(
            this.properties.tier1DocumentLibraryTitle,
            'TestSupply'
          ),

        tier2DocumentLibraryTitle:
          this.getRequiredTextProperty(
            this.properties.tier2DocumentLibraryTitle,
            'SSQ-Tier2 Attachments'
          ),

        tier3DocumentLibraryTitle:
          this.getRequiredTextProperty(
            this.properties.tier3DocumentLibraryTitle,
            'SSQ-Tier3 Attachments'
          ),

        completedQuestionnairesUrl:
          this.getOptionalTextProperty(
            this.properties.completedQuestionnairesUrl
          ),

        supportingDocumentsUrl:
          this.getOptionalTextProperty(
            this.properties.supportingDocumentsUrl
          ),

        targetScore:
          this.getValidTargetScore(
            this.properties.targetScore
          )
      };

    const element:
      React.ReactElement<IEsgDashboardProps> =
        React.createElement(
          EsgDashboard,
          componentProps
        );

    ReactDom.render(
      element,
      this.domElement
    );
  }

  protected async onInit(): Promise<void> {
    await super.onInit();

    this.initializeDefaultProperties();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(
      this.domElement
    );
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration():
    IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description:
              'Configure the ESG dashboard data sources, ' +
              'navigation links, and target score.'
          },

          groups: [
            {
              groupName:
                'Questionnaire lists',

              groupFields: [
                PropertyPaneTextField(
                  'tier1ListTitle',
                  {
                    label:
                      'Tier 1 questionnaire list title',

                    description:
                      'Enter the exact SharePoint display ' +
                      'name of the Tier 1 questionnaire list.'
                  }
                ),

                PropertyPaneTextField(
                  'tier2ListTitle',
                  {
                    label:
                      'Tier 2 questionnaire list title',

                    description:
                      'Enter the exact SharePoint display ' +
                      'name of the Tier 2 questionnaire list.'
                  }
                ),

                PropertyPaneTextField(
                  'tier3ListTitle',
                  {
                    label:
                      'Tier 3 questionnaire list title',

                    description:
                      'Enter the exact SharePoint display ' +
                      'name of the Tier 3 questionnaire list.'
                  }
                )
              ]
            },

            {
              groupName:
                'Supplier Name fields',

              groupFields: [
                PropertyPaneTextField(
                  'tier1SupplierNameField',
                  {
                    label:
                      'Tier 1 Supplier Name internal field',

                    description:
                      'Enter the internal SharePoint column ' +
                      'name containing the Tier 1 supplier name.'
                  }
                ),

                PropertyPaneTextField(
                  'tier2SupplierNameField',
                  {
                    label:
                      'Tier 2 Supplier Name internal field',

                    description:
                      'Enter the internal SharePoint column ' +
                      'name containing the Tier 2 supplier name.'
                  }
                ),

                PropertyPaneTextField(
                  'tier3SupplierNameField',
                  {
                    label:
                      'Tier 3 Supplier Name internal field',

                    description:
                      'Enter the internal SharePoint column ' +
                      'name containing the Tier 3 supplier name.'
                  }
                )
              ]
            },

            {
              groupName:
                'Supporting document libraries',

              groupFields: [
                PropertyPaneTextField(
                  'tier1DocumentLibraryTitle',
                  {
                    label:
                      'Tier 1 document library title'
                  }
                ),

                PropertyPaneTextField(
                  'tier2DocumentLibraryTitle',
                  {
                    label:
                      'Tier 2 document library title'
                  }
                ),

                PropertyPaneTextField(
                  'tier3DocumentLibraryTitle',
                  {
                    label:
                      'Tier 3 document library title'
                  }
                )
              ]
            },

            {
              groupName:
                'Dashboard settings',

              groupFields: [
                PropertyPaneSlider(
                  'targetScore',
                  {
                    label:
                      'Target score',

                    min:
                      0,

                    max:
                      100,

                    step:
                      1,

                    showValue:
                      true
                  }
                )
              ]
            },

            {
              groupName:
                'Navigation links',

              groupFields: [
                PropertyPaneTextField(
                  'completedQuestionnairesUrl',
                  {
                    label:
                      'Completed Questionnaires URL',

                    description:
                      'Optional URL used by dashboard links ' +
                      'for completed questionnaire records.'
                  }
                ),

                PropertyPaneTextField(
                  'supportingDocumentsUrl',
                  {
                    label:
                      'Supporting Documents URL',

                    description:
                      'Optional URL used to open the ' +
                      'Supporting Documents page.'
                  }
                )
              ]
            }
          ]
        }
      ]
    };
  }

  /**
   * Initializes missing configuration values for new and
   * previously added instances of the web part.
   */
  private initializeDefaultProperties(): void {
    if (
      !this.properties
        .tier1ListTitle ||
      !this.properties
        .tier1ListTitle
        .trim()
    ) {
      this.properties.tier1ListTitle =
        'CASSTECH_SSQ';
    }

    if (
      !this.properties
        .tier2ListTitle ||
      !this.properties
        .tier2ListTitle
        .trim()
    ) {
      this.properties.tier2ListTitle =
        'Tier 2 ESG Procurement Questionnaire';
    }

    if (
      !this.properties
        .tier3ListTitle ||
      !this.properties
        .tier3ListTitle
        .trim()
    ) {
      this.properties.tier3ListTitle =
        'Supplier Sustainability Questionnaires Tier 3';
    }

    if (
      !this.properties
        .tier1SupplierNameField ||
      !this.properties
        .tier1SupplierNameField
        .trim()
    ) {
      this.properties.tier1SupplierNameField =
        'Supplier_x0020_Name';
    }

    if (
      !this.properties
        .tier2SupplierNameField ||
      !this.properties
        .tier2SupplierNameField
        .trim()
    ) {
      this.properties.tier2SupplierNameField =
        'Supplier_x0020_Name';
    }

    if (
      !this.properties
        .tier3SupplierNameField ||
      !this.properties
        .tier3SupplierNameField
        .trim()
    ) {
      this.properties.tier3SupplierNameField =
        'Supplier_x0020_Name';
    }

    if (
      !this.properties
        .tier1DocumentLibraryTitle ||
      !this.properties
        .tier1DocumentLibraryTitle
        .trim()
    ) {
      this.properties
        .tier1DocumentLibraryTitle =
          'TestSupply';
    }

    if (
      !this.properties
        .tier2DocumentLibraryTitle ||
      !this.properties
        .tier2DocumentLibraryTitle
        .trim()
    ) {
      this.properties
        .tier2DocumentLibraryTitle =
          'SSQ-Tier2 Attachments';
    }

    if (
      !this.properties
        .tier3DocumentLibraryTitle ||
      !this.properties
        .tier3DocumentLibraryTitle
        .trim()
    ) {
      this.properties
        .tier3DocumentLibraryTitle =
          'SSQ-Tier3 Attachments';
    }

    if (
      !this.properties
        .completedQuestionnairesUrl
    ) {
      this.properties
        .completedQuestionnairesUrl = '';
    }

    if (
      !this.properties
        .supportingDocumentsUrl
    ) {
      this.properties
        .supportingDocumentsUrl = '';
    }

    this.properties.targetScore =
      this.getValidTargetScore(
        this.properties.targetScore
      );
  }

  /**
   * Returns a trimmed required text-property value.
   */
  private getRequiredTextProperty(
    value: string | undefined,
    defaultValue: string
  ): string {
    if (!value) {
      return defaultValue;
    }

    const trimmedValue: string =
      value.trim();

    return trimmedValue
      ? trimmedValue
      : defaultValue;
  }

  /**
   * Returns a trimmed optional text-property value.
   */
  private getOptionalTextProperty(
    value: string | undefined
  ): string {
    return value
      ? value.trim()
      : '';
  }

  /**
   * Ensures targetScore remains between zero and one hundred.
   */
  private getValidTargetScore(
    value: number | undefined
  ): number {
    if (
      value === undefined ||
      !isFinite(value)
    ) {
      return 80;
    }

    return Math.min(
      Math.max(
        Math.round(value),
        0
      ),
      100
    );
  }
}