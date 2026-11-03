declare interface IHtmlCssWebPartStrings {
  PropertyPaneDescription: string;
  ContentFieldsGroupName: string;
  LibraryFieldLabel: string;
  FileFieldLabel: string;
  OpenFileButtonLabel: string;
  Loading: string;
  NoFilesFound: string;
  LoadError: string;
}

declare module "HtmlCssWebPartStrings" {
  const strings: IHtmlCssWebPartStrings;
  export = strings;
}
