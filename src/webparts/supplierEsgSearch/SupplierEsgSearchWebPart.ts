import * as React from 'react';
import * as ReactDom from 'react-dom';

import { Version } from '@microsoft/sp-core-library';

import {
  IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';

import SupplierEsgSearch from './components/SupplierEsgSearch';

import {
  ISupplierEsgSearchProps
} from './components/ISupplierEsgSearchProps';

export interface ISupplierEsgSearchWebPartProps {
  tier1ListTitle: string;
  tier2ListTitle: string;
  tier3ListTitle: string;

  tier1SupplierNameField: string;
  tier2SupplierNameField: string;
  tier3SupplierNameField: string;

  tier1MaxPoints: number;
  tier2MaxPoints: number;
  tier3MaxPoints: number;

  tier1WeightagePercentage: number;
  tier2WeightagePercentage: number;
  tier3WeightagePercentage: number;
}

export default class SupplierEsgSearchWebPart
  extends BaseClientSideWebPart<
    ISupplierEsgSearchWebPartProps
  > {
  public render(): void {
    const element: React.ReactElement<
      ISupplierEsgSearchProps
    > = React.createElement(
      SupplierEsgSearch,
      {
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
          'Supplier_x0020_Name'
      }
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
              'Configure the questionnaire lists and qualification thresholds.'
          },
          groups: [
            {
              groupName: 'Questionnaire lists',
              groupFields: [
                PropertyPaneTextField(
                  'tier1ListTitle',
                  {
                    label: 'Tier 1 list title'
                  }
                ),

                PropertyPaneTextField(
                  'tier2ListTitle',
                  {
                    label: 'Tier 2 list title'
                  }
                ),

                PropertyPaneTextField(
                  'tier3ListTitle',
                  {
                    label: 'Tier 3 list title'
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
                      'Tier 1 Supplier Name internal name'
                  }
                ),

                PropertyPaneTextField(
                  'tier2SupplierNameField',
                  {
                    label:
                      'Tier 2 Supplier Name internal name'
                  }
                ),

                PropertyPaneTextField(
                  'tier3SupplierNameField',
                  {
                    label:
                      'Tier 3 Supplier Name internal name'
                  }
                )
              ]
            },
            {
              groupName:
                'Qualification thresholds',
              groupFields: [
                PropertyPaneSlider(
                  'qualifiedMinimum',
                  {
                    label:
                      'Qualified minimum percentage',
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
                      'Conditionally qualified minimum percentage',
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

  /**
   * Initializes properties that have not yet been configured
   * through the web-part property pane.
   */
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
  }
}