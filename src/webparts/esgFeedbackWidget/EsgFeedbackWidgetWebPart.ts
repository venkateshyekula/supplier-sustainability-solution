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

import EsgFeedbackWidget from './components/EsgFeedbackWidget';

import {
  IEsgFeedbackWidgetProps
} from './components/IEsgFeedbackWidgetProps';

export interface IEsgFeedbackWidgetWebPartProps {
  tier1ListTitle: string;
  tier2ListTitle: string;
  tier3ListTitle: string;
}

export default class EsgFeedbackWidgetWebPart
  extends BaseClientSideWebPart<
    IEsgFeedbackWidgetWebPartProps
  > {
  public render(): void {
    const element: React.ReactElement<
      IEsgFeedbackWidgetProps
    > = React.createElement(
      EsgFeedbackWidget,
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
          'Supplier Sustainability Questionnaires Tier 3'
      }
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
              'Configure the SharePoint questionnaire lists used by the ESG Feedback widget.'
          },
          groups: [
            {
              groupName:
                'Questionnaire list configuration',

              groupFields: [
                PropertyPaneTextField(
                  'tier1ListTitle',
                  {
                    label:
                      'Tier 1 questionnaire list title',

                    description:
                      'Enter the exact SharePoint display name of the Tier 1 questionnaire list.'
                  }
                ),

                PropertyPaneTextField(
                  'tier2ListTitle',
                  {
                    label:
                      'Tier 2 questionnaire list title',

                    description:
                      'Enter the exact SharePoint display name of the Tier 2 questionnaire list.'
                  }
                ),

                PropertyPaneTextField(
                  'tier3ListTitle',
                  {
                    label:
                      'Tier 3 questionnaire list title',

                    description:
                      'Enter the exact SharePoint display name of the Tier 3 questionnaire list.'
                  }
                )
              ]
            }
          ]
        }
      ]
    };
  }
}