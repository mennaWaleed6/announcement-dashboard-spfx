import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";

export interface IAnnouncementDashboardProps {
  title: string;
  description: string;
  IsFiltering: boolean;
  Items: number;
  Layout: string;
  ListName: string;
  Language: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  websiteUrl: string;
  color: string;

  context: WebPartContext;
}
