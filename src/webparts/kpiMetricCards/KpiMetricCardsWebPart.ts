import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import KpiMetricCards from './components/KpiMetricCards';
import { IKpiMetricCardsProps } from './components/IKpiMetricCardsProps';

export interface IKpiMetricCardsWebPartProps {
  description: string;
}

export default class KpiMetricCardsWebPart extends BaseClientSideWebPart<IKpiMetricCardsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IKpiMetricCardsProps> = React.createElement(
      KpiMetricCards,
      {
        context: this.context
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
}
