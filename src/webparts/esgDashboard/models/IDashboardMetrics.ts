import { IActionDatum, IAgingDatum, IChartDatum, ITrendDatum } from './IDashboardChartData';
import { IDashboardInsight } from './IDashboardInsight';
import { IDashboardSubmission } from './IDashboardSubmission';
export interface IDashboardMetrics {
    total: number; approved: number; pending: number; requiresAction: number; highRisk: number; averageScore: number;
    totalChange: number; approvedChange: number; pendingChange: number; actionChange: number; highRiskChange: number; scoreChange: number;
    trend: ITrendDatum[]; byTier: IChartDatum[]; qualification: IChartDatum[]; risks: IChartDatum[];
    aging: IAgingDatum[]; actions: IActionDatum[]; averageByTier: IChartDatum[];
    recent: IDashboardSubmission[]; insights: IDashboardInsight[];
}