import { IListConfiguration } from '../../supplierEsgSearch/models/IListConfiguration';
import { IDashboardMetrics } from '../models/IDashboardMetrics';
export interface IEsgDashboardService { 
    getDashboard(
        configurations: readonly IListConfiguration[]
    ): 
    Promise<IDashboardMetrics>; 
}