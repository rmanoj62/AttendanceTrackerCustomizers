import { SPHttpClient } from '@microsoft/sp-http';

export interface IAttendanceTrackerProps {
  siteUrl: string;
  userEmail: string;
  userName: string;
  spHttpClient: SPHttpClient;
}