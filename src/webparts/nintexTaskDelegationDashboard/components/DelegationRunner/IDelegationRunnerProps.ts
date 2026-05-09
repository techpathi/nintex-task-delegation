import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDelegationRunnerProps {
  context: WebPartContext;
  nintexApiBaseUrl: string;
  tokenListUrl: string;
  tokenColumnName: string;
  tokenTitleValue: string;
  recentDaysLimit: number;
  bufferMinutes?: number;
}
