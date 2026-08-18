import {
  WebPartContext
} from '@microsoft/sp-webpart-base';

export interface IKpiMetricCardsProps {
  context: WebPartContext;

  tier1ListTitle: string;
  tier2ListTitle: string;
  tier3ListTitle: string;
}