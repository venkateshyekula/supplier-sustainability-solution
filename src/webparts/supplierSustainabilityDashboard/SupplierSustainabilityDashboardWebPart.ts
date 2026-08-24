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

import SupplierSustainabilityDashboard
  from './components/SupplierSustainabilityDashboard';

import {
  ISupplierSustainabilityDashboardProps
} from './components/ISupplierSustainabilityDashboardProps';

export interface ISupplierSustainabilityDashboardWebPartProps {
  tier1ListTitle: string;
  tier2ListTitle: string;
  tier3ListTitle: string;

  tier1SupplierNameField: string;
  tier2SupplierNameField: string;
  tier3SupplierNameField: string;

  tier1DocumentLibraryTitle: string;
  tier2DocumentLibraryTitle: string;
  tier3DocumentLibraryTitle: string;

  qualifiedMinimum: number;
  conditionalMinimum: number;
}

export default class SupplierSustainabilityDashboardWebPart
  extends BaseClientSideWebPart<
    ISupplierSustainabilityDashboardWebPartProps
  > {
  public render(): void {
    const componentProps:
      ISupplierSustainabilityDashboardProps = {
        context: this.context,

        tier1ListTitle:
          this.properties.tier1ListTitle ||
          'CASSTECH_SSQ',

        tier2ListTitle:
          this.properties.tier2ListTitle ||
          'Tier 2 ESG Procurement Questionnaire',

        tier3ListTitle:
          this.properties.tier3ListTitle ||
          'Supplier Sustainability Questionnaires Tier 3',

        tier1SupplierNameField:
          this.properties.tier1SupplierNameField ||
          'Supplier_x0020_Name',

        tier2SupplierNameField:
          this.properties.tier2SupplierNameField ||
          'Supplier_x0020_Name',

        tier3SupplierNameField:
          this.properties.tier3SupplierNameField ||
          'Supplier_x0020_Name',

        tier1DocumentLibraryTitle:
          this.properties.tier1DocumentLibraryTitle ||
          'TestSupply',

        tier2DocumentLibraryTitle:
          this.properties.tier2DocumentLibraryTitle ||
          'SSQ-Tier2 Attachments',

        tier3DocumentLibraryTitle:
          this.properties.tier3DocumentLibraryTitle ||
          'SSQ-Tier3 Attachments',

        qualifiedMinimum:
          this.properties.qualifiedMinimum !== undefined
            ? this.properties.qualifiedMinimum
            : 70,

        conditionalMinimum:
          this.properties.conditionalMinimum !== undefined
            ? this.properties.conditionalMinimum
            : 52
      };

    const element: React.ReactElement<
      ISupplierSustainabilityDashboardProps
    > = React.createElement(
      SupplierSustainabilityDashboard,
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
              'Configure Supplier Sustainability Dashboard ' +
              'data sources and qualification values.'
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
                      'Tier 1 questionnaire list title'
                  }
                ),

                PropertyPaneTextField(
                  'tier2ListTitle',
                  {
                    label:
                      'Tier 2 questionnaire list title'
                  }
                ),

                PropertyPaneTextField(
                  'tier3ListTitle',
                  {
                    label:
                      'Tier 3 questionnaire list title'
                  }
                )
              ]
            },

            {
              groupName:
                'Supplier Name internal fields',

              groupFields: [
                PropertyPaneTextField(
                  'tier1SupplierNameField',
                  {
                    label:
                      'Tier 1 Supplier Name internal field'
                  }
                ),

                PropertyPaneTextField(
                  'tier2SupplierNameField',
                  {
                    label:
                      'Tier 2 Supplier Name internal field'
                  }
                ),

                PropertyPaneTextField(
                  'tier3SupplierNameField',
                  {
                    label:
                      'Tier 3 Supplier Name internal field'
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
                'Qualification values',

              groupFields: [
                PropertyPaneSlider(
                  'qualifiedMinimum',
                  {
                    label:
                      'Qualified minimum value',
                    min: 0,
                    max: 100,
                    step: 1,
                    showValue: true
                  }
                ),

                PropertyPaneSlider(
                  'conditionalMinimum',
                  {
                    label:
                      'Conditionally qualified minimum value',
                    min: 0,
                    max: 100,
                    step: 1,
                    showValue: true
                  }
                )
              ]
            }
          ]
        }
      ]
    };
  }

  private initializeDefaultProperties(): void {
    if (!this.properties.tier1ListTitle) {
      this.properties.tier1ListTitle =
        'CASSTECH_SSQ';
    }

    if (!this.properties.tier2ListTitle) {
      this.properties.tier2ListTitle =
        'Tier 2 ESG Procurement Questionnaire';
    }

    if (!this.properties.tier3ListTitle) {
      this.properties.tier3ListTitle =
        'Supplier Sustainability Questionnaires Tier 3';
    }

    if (!this.properties.tier1SupplierNameField) {
      this.properties.tier1SupplierNameField =
        'Supplier_x0020_Name';
    }

    if (!this.properties.tier2SupplierNameField) {
      this.properties.tier2SupplierNameField =
        'Supplier_x0020_Name';
    }

    if (!this.properties.tier3SupplierNameField) {
      this.properties.tier3SupplierNameField =
        'Supplier_x0020_Name';
    }

    if (!this.properties.tier1DocumentLibraryTitle) {
      this.properties.tier1DocumentLibraryTitle =
        'TestSupply';
    }

    if (!this.properties.tier2DocumentLibraryTitle) {
      this.properties.tier2DocumentLibraryTitle =
        'SSQ-Tier2 Attachments';
    }

    if (!this.properties.tier3DocumentLibraryTitle) {
      this.properties.tier3DocumentLibraryTitle =
        'SSQ-Tier3 Attachments';
    }

    if (
      this.properties.qualifiedMinimum === undefined
    ) {
      this.properties.qualifiedMinimum = 70;
    }

    if (
      this.properties.conditionalMinimum === undefined
    ) {
      this.properties.conditionalMinimum = 52;
    }
  }
}