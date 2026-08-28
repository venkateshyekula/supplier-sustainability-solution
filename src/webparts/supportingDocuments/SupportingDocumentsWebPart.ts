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

import SupportingDocuments
  from './components/SupportingDocuments';

import {
  ISupportingDocumentsProps
} from './components/ISupportingDocumentsProps';

export interface ISupportingDocumentsWebPartProps {
  tier1QuestionnaireListTitle: string;
  tier2QuestionnaireListTitle: string;
  tier3QuestionnaireListTitle: string;

  tier1SupplierNameField: string;
  tier2SupplierNameField: string;
  tier3SupplierNameField: string;

  tier1DocumentLibraryTitle: string;
  tier2DocumentLibraryTitle: string;
  tier3DocumentLibraryTitle: string;

  documentGuidanceUrl: string;
}

export default class SupportingDocumentsWebPart
  extends BaseClientSideWebPart<
    ISupportingDocumentsWebPartProps
  > {

  public render(): void {
    const componentProps:
      ISupportingDocumentsProps = {
        context:
          this.context,

        tier1QuestionnaireListTitle:
          this.getPropertyValue(
            this.properties
              .tier1QuestionnaireListTitle,
            'CASSTECH_SSQ'
          ),

        tier2QuestionnaireListTitle:
          this.getPropertyValue(
            this.properties
              .tier2QuestionnaireListTitle,
            'Tier 2 ESG Procurement Questionnaire'
          ),

        tier3QuestionnaireListTitle:
          this.getPropertyValue(
            this.properties
              .tier3QuestionnaireListTitle,
            'Supplier Sustainability Questionnaires Tier 3'
          ),

        tier1SupplierNameField:
          this.getPropertyValue(
            this.properties
              .tier1SupplierNameField,
            'Supplier_x0020_Name'
          ),

        tier2SupplierNameField:
          this.getPropertyValue(
            this.properties
              .tier2SupplierNameField,
            'Supplier_x0020_Name'
          ),

        tier3SupplierNameField:
          this.getPropertyValue(
            this.properties
              .tier3SupplierNameField,
            'Supplier_x0020_Name'
          ),

        tier1DocumentLibraryTitle:
          this.getPropertyValue(
            this.properties
              .tier1DocumentLibraryTitle,
            'TestSupply'
          ),

        tier2DocumentLibraryTitle:
          this.getPropertyValue(
            this.properties
              .tier2DocumentLibraryTitle,
            'SSQ-Tier2 Attachments'
          ),

        tier3DocumentLibraryTitle:
          this.getPropertyValue(
            this.properties
              .tier3DocumentLibraryTitle,
            'SSQ-Tier3 Attachments'
          ),

        documentGuidanceUrl:
          this.getOptionalPropertyValue(
            this.properties
              .documentGuidanceUrl
          )
      };

    const element: React.ReactElement<
      ISupportingDocumentsProps
    > = React.createElement(
      SupportingDocuments,
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
              'Configure the SharePoint questionnaire ' +
              'lists, supplier-name fields, supporting ' +
              'document libraries, and guidance link.'
          },

          groups: [
            {
              groupName:
                'Questionnaire lists',

              groupFields: [
                PropertyPaneTextField(
                  'tier1QuestionnaireListTitle',
                  {
                    label:
                      'Tier 1 questionnaire list title',

                    description:
                      'Enter the exact SharePoint display ' +
                      'name of the Tier 1 questionnaire list.'
                  }
                ),

                PropertyPaneTextField(
                  'tier2QuestionnaireListTitle',
                  {
                    label:
                      'Tier 2 questionnaire list title',

                    description:
                      'Enter the exact SharePoint display ' +
                      'name of the Tier 2 questionnaire list.'
                  }
                ),

                PropertyPaneTextField(
                  'tier3QuestionnaireListTitle',
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
                'Supplier Name internal fields',

              groupFields: [
                PropertyPaneTextField(
                  'tier1SupplierNameField',
                  {
                    label:
                      'Tier 1 Supplier Name internal field',

                    description:
                      'Enter the internal SharePoint field ' +
                      'name containing the Tier 1 supplier name.'
                  }
                ),

                PropertyPaneTextField(
                  'tier2SupplierNameField',
                  {
                    label:
                      'Tier 2 Supplier Name internal field',

                    description:
                      'Enter the internal SharePoint field ' +
                      'name containing the Tier 2 supplier name.'
                  }
                ),

                PropertyPaneTextField(
                  'tier3SupplierNameField',
                  {
                    label:
                      'Tier 3 Supplier Name internal field',

                    description:
                      'Enter the internal SharePoint field ' +
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
                      'Tier 1 document library title',

                    description:
                      'Enter the exact SharePoint display ' +
                      'name of the Tier 1 document library.'
                  }
                ),

                PropertyPaneTextField(
                  'tier2DocumentLibraryTitle',
                  {
                    label:
                      'Tier 2 document library title',

                    description:
                      'Enter the exact SharePoint display ' +
                      'name of the Tier 2 document library.'
                  }
                ),

                PropertyPaneTextField(
                  'tier3DocumentLibraryTitle',
                  {
                    label:
                      'Tier 3 document library title',

                    description:
                      'Enter the exact SharePoint display ' +
                      'name of the Tier 3 document library.'
                  }
                )
              ]
            },

            {
              groupName:
                'Document guidance',

              groupFields: [
                PropertyPaneTextField(
                  'documentGuidanceUrl',
                  {
                    label:
                      'Document guidance URL',

                    description:
                      'Optional URL for supporting-document ' +
                      'instructions or guidance.'
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
   * Initializes missing values for existing web-part instances.
   */
  private initializeDefaultProperties(): void {
    if (
      !this.properties
        .tier1QuestionnaireListTitle?.trim()
    ) {
      this.properties
        .tier1QuestionnaireListTitle =
          'CASSTECH_SSQ';
    }

    if (
      !this.properties
        .tier2QuestionnaireListTitle?.trim()
    ) {
      this.properties
        .tier2QuestionnaireListTitle =
          'Tier 2 ESG Procurement Questionnaire';
    }

    if (
      !this.properties
        .tier3QuestionnaireListTitle?.trim()
    ) {
      this.properties
        .tier3QuestionnaireListTitle =
          'Supplier Sustainability Questionnaires Tier 3';
    }

    if (
      !this.properties
        .tier1SupplierNameField?.trim()
    ) {
      this.properties
        .tier1SupplierNameField =
          'Supplier_x0020_Name';
    }

    if (
      !this.properties
        .tier2SupplierNameField?.trim()
    ) {
      this.properties
        .tier2SupplierNameField =
          'Supplier_x0020_Name';
    }

    if (
      !this.properties
        .tier3SupplierNameField?.trim()
    ) {
      this.properties
        .tier3SupplierNameField =
          'Supplier_x0020_Name';
    }

    if (
      !this.properties
        .tier1DocumentLibraryTitle?.trim()
    ) {
      this.properties
        .tier1DocumentLibraryTitle =
          'TestSupply';
    }

    if (
      !this.properties
        .tier2DocumentLibraryTitle?.trim()
    ) {
      this.properties
        .tier2DocumentLibraryTitle =
          'SSQ-Tier2 Attachments';
    }

    if (
      !this.properties
        .tier3DocumentLibraryTitle?.trim()
    ) {
      this.properties
        .tier3DocumentLibraryTitle =
          'SSQ-Tier3 Attachments';
    }

    if (
      this.properties
        .documentGuidanceUrl === undefined
    ) {
      this.properties.documentGuidanceUrl = '';
    }
  }

  /**
   * Returns a trimmed property value or its required fallback.
   */
  private getPropertyValue(
    value: string | undefined,
    fallbackValue: string
  ): string {
    const trimmedValue: string =
      value?.trim() || '';

    return trimmedValue || fallbackValue;
  }

  /**
   * Returns a trimmed optional property value.
   */
  private getOptionalPropertyValue(
    value: string | undefined
  ): string {
    return value?.trim() || '';
  }
}