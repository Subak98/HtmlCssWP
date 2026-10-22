define([], function () {
  return {
    PropertyPaneDescription:
      "Choose a library and a style file (.txt / .html with markup, or .json with html, css and js properties) to render on the page.",
    BasicGroupName: "Group Name",
    DescriptionFieldLabel: "Description Field",
    AppLocalEnvironmentSharePoint:
      "The app is running on your local environment as SharePoint web part",
    AppLocalEnvironmentTeams:
      "The app is running on your local environment as Microsoft Teams app",
    AppLocalEnvironmentOffice:
      "The app is running on your local environment in office.com",
    AppLocalEnvironmentOutlook:
      "The app is running on your local environment in Outlook",
    AppSharePointEnvironment: "The app is running on SharePoint page",
    AppTeamsTabEnvironment: "The app is running in Microsoft Teams",
    AppOfficeEnvironment: "The app is running in office.com",
    AppOutlookEnvironment: "The app is running in Outlook",
    UnknownEnvironment: "The app is running in an unknown environment",

    // Property pane
    SourceGroupName: "Style source",
    LibraryFieldLabel: "Library",
    FileFieldLabel: "Style file",
    RefreshButtonLabel: "Refresh file list",
    LoadingLibraries: "Loading libraries...",
    LoadingFiles: "Loading files...",
    NoFilesFound: "No files found in this library.",
    FilesFound: "{0} file(s) available.",
    LibrariesLoadError: "Could not load the libraries of this site.",
    FilesLoadError: "Could not load the files of this library.",

    // Rendering
    SelectFilePlaceholderTitle: "No style file selected",
    SelectFilePlaceholderText:
      "Open the property pane, pick a library and select a style file.",
    LoadingContent: "Loading content...",
    FrameTitle: "Custom content",
    FormatErrorTitle:
      "The selected file is not in the expected property format.",
    FormatErrorHint:
      "Use a .json file with the properties html, css and js (each a string), or a .txt / .html file containing HTML with optional style and script blocks.",
    NotFoundError:
      "The selected file could not be found. It may have been moved or deleted.",
    ForbiddenError: "You do not have permission to read the selected file.",
    TooLargeError: "The selected file is too large (the limit is 1 MB).",
    GenericLoadError: "The selected file could not be loaded.",
  };
});
