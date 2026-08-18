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

import KpiMetricCards from './components/KpiMetricCards';

import {
  IKpiMetricCardsProps
} from './components/IKpiMetricCardsProps';

export interface IKpiMetricCardsWebPartProps {
  tier1ListTitle: string;
  tier2ListTitle: string;
  tier3ListTitle: string;
}

export default class KpiMetricCardsWebPart
  extends BaseClientSideWebPart<
    IKpiMetricCardsWebPartProps
  > {
  public render(): void {
    const element: React.ReactElement<
      IKpiMetricCardsProps
    > = React.createElement(
      KpiMetricCards,
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
              'Configure the SharePoint lists used by the KPI cards.'
          },
          groups: [
            {
              groupName:
                'Supplier ESG list configuration',

              groupFields: [
                PropertyPaneTextField(
                  'tier1ListTitle',
                  {
                    label:
                      'Tier 1 list title',

                    description:
                      'Enter the exact SharePoint display name of the Tier 1 list.'
                  }
                ),

                PropertyPaneTextField(
                  'tier2ListTitle',
                  {
                    label:
                      'Tier 2 list title',

                    description:
                      'Enter the exact SharePoint display name of the Tier 2 list.'
                  }
                ),

                PropertyPaneTextField(
                  'tier3ListTitle',
                  {
                    label:
                      'Tier 3 list title',

                    description:
                      'Enter the exact SharePoint display name of the Tier 3 list.'
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