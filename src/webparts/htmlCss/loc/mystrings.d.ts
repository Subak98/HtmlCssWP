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

  SourceGroupName: string;
  LibraryFieldLabel: string;
  FileFieldLabel: string;
  RefreshButtonLabel: string;
  LoadingLibraries: string;
  LoadingFiles: string;
  NoFilesFound: string;
  FilesFound: string;
  LibrariesLoadError: string;
  FilesLoadError: string;

  SelectFilePlaceholderTitle: string;
  SelectFilePlaceholderText: string;
  LoadingContent: string;
  FrameTitle: string;
  FormatErrorTitle: string;
  FormatErrorHint: string;
  NotFoundError: string;
  ForbiddenError: string;
  TooLargeError: string;
  GenericLoadError: string;
}

declare module "HtmlCssWebPartStrings" {
  const strings: IHtmlCssWebPartStrings;
  export = strings;
}
