declare interface IHtmlCssWebPartStrings {
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
  PropertyPaneDescription: string;
  ContentFieldsGroupName: string;
  BehaviorGroupName: string;
  EditButtonLabel: string;
  EnableCssFieldLabel: string;
  ShowPlaceholderFieldLabel: string;
}

declare module "HtmlCssWebPartStrings" {
  const strings: IHtmlCssWebPartStrings;
  export = strings;
}
