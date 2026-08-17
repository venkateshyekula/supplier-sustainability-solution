import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import EsgFeedbackWidget from './components/EsgFeedbackWidget';
import { IEsgFeedbackWidgetProps } from './components/IEsgFeedbackWidgetProps';

export interface IEsgFeedbackWidgetWebPartProps {
  description: string;
}

export default class EsgFeedbackWidgetWebPart extends BaseClientSideWebPart<IEsgFeedbackWidgetWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IEsgFeedbackWidgetProps> = React.createElement(
      EsgFeedbackWidget,
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
