import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import EsgFeedbackWidget from './components/EsgFeedbackWidget';
import { IEsgFeedbackWidgetProps } from './components/IEsgFeedbackWidgetProps';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

export interface IEsgFeedbackWidgetWebPartProps {
  description: string;
  tier1ListTitle: string;
tier2ListTitle: string;
tier3ListTitle: string;
}

export default class EsgFeedbackWidgetWebPart extends BaseClientSideWebPart<IEsgFeedbackWidgetWebPartProps> {

  public render(): void {
  const element: React.ReactElement<
    IEsgFeedbackWidgetProps
  > = React.createElement(
    EsgFeedbackWidget,
    {
      context: this.context,

      tier1ListTitle:
        this.properties.tier1ListTitle,

      tier2ListTitle:
        this.properties.tier2ListTitle,

      tier3ListTitle:
        this.properties.tier3ListTitle
    }
  );

  ReactDom.render(
    element,
    this.domElement
  );
}

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
            'Configure the Supplier ESG questionnaire lists.'
        },
        groups: [
          {
            groupName:
              'SharePoint list configuration',

            groupFields: [
              PropertyPaneTextField(
                'tier1ListTitle',
                {
                  label:
                    'Tier 1 list title'
                }
              ),

              PropertyPaneTextField(
                'tier2ListTitle',
                {
                  label:
                    'Tier 2 list title'
                }
              ),

              PropertyPaneTextField(
                'tier3ListTitle',
                {
                  label:
                    'Tier 3 list title'
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
