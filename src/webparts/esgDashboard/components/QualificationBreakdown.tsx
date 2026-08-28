import * as React from 'react';

import {
  IChartDatum
} from '../models/IDashboardChartData';

import {
  DonutPanel
} from './DonutPanel';

export interface IQualificationBreakdownProps {
  items: readonly IChartDatum[];

  onViewDetails?(): void;
}

export function QualificationBreakdown(
  props: IQualificationBreakdownProps
): React.ReactElement<IQualificationBreakdownProps> {
  return (
    <DonutPanel
      title="Qualification Breakdown"
      subtitle="Based on Tier-specific weighted thresholds"
      items={props.items}
      onViewDetails={props.onViewDetails}
    />
  );
}

export default QualificationBreakdown;