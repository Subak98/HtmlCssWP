import { SPHttpClient } from "@microsoft/sp-http";

export interface IHtmlCssProps {
  spHttpClient: SPHttpClient;
  /** Absolute URL of the current web. */
  webUrl: string;
  /** Server-relative URL of the .html file picked in the property pane. */
  fileUrl?: string;
  /** Show a friendly placeholder in edit mode when no file is selected. */
  showPlaceholderWhenEmpty: boolean;
  /** Unique instance id used to scope the injected <style> block. */
  instanceId: string;
  isEditMode: boolean;
}

export interface IHtmlCssWebPartWebPartProps {
  /** Id (GUID) of the document library the file is picked from. Defaults to the Style Library. */
  libraryId?: string;
  /** Server-relative URL of the selected .html file. */
  fileUrl?: string;
  showPlaceholderWhenEmpty?: boolean;
}
