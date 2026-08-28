import * as React from 'react';

import {
  IChartDatum
} from '../models/IDashboardChartData';

import {
  DonutPanel
} from './DonutPanel';

export interface IRiskRatingAnalysisProps {
  items: readonly IChartDatum[];

  onViewDetails?(): void;
}

export function RiskRatingAnalysis(
  props: IRiskRatingAnalysisProps
): React.ReactElement<IRiskRatingAnalysisProps> {
  return (
    <DonutPanel
      title="Risk Rating Analysis"
      subtitle="Risk distribution across all submissions"
      items={props.items}
      onViewDetails={props.onViewDetails}
    />
  );
}

export default RiskRatingAnalysis;