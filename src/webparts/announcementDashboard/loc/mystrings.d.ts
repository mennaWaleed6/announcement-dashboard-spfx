declare interface IAnnouncementDashboardWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;

  HeaderTitle: string;
  HeaderDescription: string;
  HeaderCategory: string;
  HeaderPriority: string;
  HeaderDueDate: string;
  HeaderAssignedTo: string;
}

declare module "AnnouncementDashboardWebPartStrings" {
  const strings: IAnnouncementDashboardWebPartStrings;
  export = strings;
}
