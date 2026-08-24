import * as React from 'react';
import * as ReactDom from 'react-dom';

import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
} from '@microsoft/sp-webpart-base';
import {
  PropertyPaneTextField,
  IPropertyPaneConfiguration
} from '@microsoft/sp-property-pane';

import QuestionnaireQuickLinks from './components/QuestionnaireQuickLinks';
import { IQuestionnaireQuickLinksProps } from './components/IQuestionnaireQuickLinksProps';

export interface IQuestionnaireQuickLinksWebPartProps {
  tier1QuestionnaireListTitle: string;
  tier2QuestionnaireListTitle: string;
  tier3QuestionnaireListTitle: string;

  tier1DocumentLibraryTitle: string;
  tier2DocumentLibraryTitle: string;
  tier3DocumentLibraryTitle: string;
}

export default class QuestionnaireQuickLinksWebPart
  extends BaseClientSideWebPart<IQuestionnaireQuickLinksWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IQuestionnaireQuickLinksProps> =
      React.createElement(
        QuestionnaireQuickLinks,
        {
          context: this.context,
          tier1QuestionnaireListTitle:
            this.properties
              .tier1QuestionnaireListTitle
              ?.trim() ||
            'CASSTECH_SSQ',

          tier2QuestionnaireListTitle:
            this.properties
              .tier2QuestionnaireListTitle
              ?.trim() ||
            'Tier 2 ESG Procurement Questionnaire',

          tier3QuestionnaireListTitle:
            this.properties
              .tier3QuestionnaireListTitle
              ?.trim() ||
            'Supplier Sustainability Questionnaires Tier 3',

          tier1DocumentLibraryTitle:
            this.properties
              .tier1DocumentLibraryTitle
              ?.trim() ||
            'TestSupply',

          tier2DocumentLibraryTitle:
            this.properties
              .tier2DocumentLibraryTitle
              ?.trim() ||
            'SSQ-Tier2 Attachments',

          tier3DocumentLibraryTitle:
            this.properties
              .tier3DocumentLibraryTitle
              ?.trim() ||
            'SSQ-Tier3 Attachments'
        }
      );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description:
              'Configure questionnaire lists and supporting document libraries for this SharePoint environment.'
          },
          groups: [
            {
              groupName: 'Completed questionnaire lists',
              groupFields: [
                PropertyPaneTextField('tier1QuestionnaireListTitle', {
                  label: 'Tier 1 questionnaire list title'
                }),
                PropertyPaneTextField('tier2QuestionnaireListTitle', {
                  label: 'Tier 2 questionnaire list title'
                }),
                PropertyPaneTextField('tier3QuestionnaireListTitle', {
                  label: 'Tier 3 questionnaire list title'
                })
              ]
            },
            {
              groupName: 'Supporting document libraries',
              groupFields: [
                PropertyPaneTextField('tier1DocumentLibraryTitle', {
                  label: 'Tier 1 document library title'
                }),
                PropertyPaneTextField('tier2DocumentLibraryTitle', {
                  label: 'Tier 2 document library title'
                }),
                PropertyPaneTextField('tier3DocumentLibraryTitle', {
                  label: 'Tier 3 document library title'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}