import * as React from 'react';

import {
  IChartDatum
} from '../models/IDashboardChartData';

import {
  DonutPanel
} from './DonutPanel';

export interface ISubmissionsByTierProps {
  items: readonly IChartDatum[];

  onViewDetails?(): void;
}

export function SubmissionsByTier(
  props: ISubmissionsByTierProps
): React.ReactElement<ISubmissionsByTierProps> {
  return (
    <DonutPanel
      title="Submissions by Tier"
      subtitle="Distribution of submissions across Tiers"
      items={props.items}
      onViewDetails={props.onViewDetails}
    />
  );
}

export default SubmissionsByTier;