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

import CompletedQuestionnairesTab
  from './components/CompletedQuestionnairesTab';

import {
  ICompletedQuestionnairesTabProps
} from './components/ICompletedQuestionnairesTabProps';

export interface ICompletedQuestionnairesTabWebPartProps {
  tier1ListTitle: string;
  tier2ListTitle: string;
  tier3ListTitle: string;

  tier1SupplierNameField: string;
  tier2SupplierNameField: string;
  tier3SupplierNameField: string;

  tier1DocumentLibraryTitle: string;
  tier2DocumentLibraryTitle: string;
  tier3DocumentLibraryTitle: string;
}

export default class CompletedQuestionnairesTabWebPart
  extends BaseClientSideWebPart<
    ICompletedQuestionnairesTabWebPartProps
  > {
  protected async onInit(): Promise<void> {
    await super.onInit();

    this.initializeDefaultProperties();
  }

  public render(): void {
    const componentProps:
      ICompletedQuestionnairesTabProps = {
        context: this.context,

        tier1ListTitle:
          this.getPropertyValue(
            this.properties.tier1ListTitle,
            'CASSTECH_SSQ'
          ),

        tier2ListTitle:
          this.getPropertyValue(
            this.properties.tier2ListTitle,
            'Tier 2 ESG Procurement Questionnaire'
          ),

        tier3ListTitle:
          this.getPropertyValue(
            this.properties.tier3ListTitle,
            'Supplier Sustainability Questionnaires Tier 3'
          ),

        tier1SupplierNameField:
          this.getPropertyValue(
            this.properties.tier1SupplierNameField,
            'Supplier_x0020_Name'
          ),

        tier2SupplierNameField:
          this.getPropertyValue(
            this.properties.tier2SupplierNameField,
            'Supplier_x0020_Name'
          ),

        tier3SupplierNameField:
          this.getPropertyValue(
            this.properties.tier3SupplierNameField,
            'Supplier_x0020_Name'
          ),

        tier1DocumentLibraryTitle:
          this.getPropertyValue(
            this.properties.tier1DocumentLibraryTitle,
            'TestSupply'
          ),

        tier2DocumentLibraryTitle:
          this.getPropertyValue(
            this.properties.tier2DocumentLibraryTitle,
            'SSQ-Tier2 Attachments'
          ),

        tier3DocumentLibraryTitle:
          this.getPropertyValue(
            this.properties.tier3DocumentLibraryTitle,
            'SSQ-Tier3 Attachments'
          )
      };

    const element: React.ReactElement<
      ICompletedQuestionnairesTabProps
    > = React.createElement(
      CompletedQuestionnairesTab,
      componentProps
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
              'Configure the completed questionnaire lists, ' +
              'supplier fields, and supporting document libraries.'
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
            }
          ]
        }
      ]
    };
  }

  private initializeDefaultProperties(): void {
    if (!this.properties.tier1ListTitle?.trim()) {
      this.properties.tier1ListTitle =
        'CASSTECH_SSQ';
    }

    if (!this.properties.tier2ListTitle?.trim()) {
      this.properties.tier2ListTitle =
        'Tier 2 ESG Procurement Questionnaire';
    }

    if (!this.properties.tier3ListTitle?.trim()) {
      this.properties.tier3ListTitle =
        'Supplier Sustainability Questionnaires Tier 3';
    }

    if (!this.properties.tier1SupplierNameField?.trim()) {
      this.properties.tier1SupplierNameField =
        'Supplier_x0020_Name';
    }

    if (!this.properties.tier2SupplierNameField?.trim()) {
      this.properties.tier2SupplierNameField =
        'Supplier_x0020_Name';
    }

    if (!this.properties.tier3SupplierNameField?.trim()) {
      this.properties.tier3SupplierNameField =
        'Supplier_x0020_Name';
    }

    if (!this.properties.tier1DocumentLibraryTitle?.trim()) {
      this.properties.tier1DocumentLibraryTitle =
        'TestSupply';
    }

    if (!this.properties.tier2DocumentLibraryTitle?.trim()) {
      this.properties.tier2DocumentLibraryTitle =
        'SSQ-Tier2 Attachments';
    }

    if (!this.properties.tier3DocumentLibraryTitle?.trim()) {
      this.properties.tier3DocumentLibraryTitle =
        'SSQ-Tier3 Attachments';
    }
  }

  private getPropertyValue(
    value: string,
    defaultValue: string
  ): string {
    const normalizedValue: string =
      value ? value.trim() : '';

    return normalizedValue || defaultValue;
  }
}