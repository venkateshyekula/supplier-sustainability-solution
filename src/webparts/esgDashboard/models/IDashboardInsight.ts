export type DashboardInsightTone =
  | 'positive'
  | 'warning'
  | 'danger'
  | 'info';

export interface IDashboardInsight {
  key: string;
  title: string;
  description: string;
  icon: string;
  tone: DashboardInsightTone;
}