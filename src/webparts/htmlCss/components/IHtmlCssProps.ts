import type { StyleLibraryService } from "../services/StyleLibraryService";

export interface IHtmlCssProps {
  /** Server-relative URL of the selected style file, e.g. /sites/hr/Style Library/banner.json */
  fileUrl: string;
  service: StyleLibraryService;
  /** Unique instance id; also used to match the iframe's height messages to this web part. */
  instanceId: string;
  /** SharePoint display mode: 1 = Read, 2 = Edit. */
  displayMode: number;
  /** SharePoint theme colours, passed into the sandbox as CSS custom properties. */
  themeVars: { [name: string]: string };
}

/** Properties persisted on the web part. Note: NO html/css/js is stored here any more. */
export interface IHtmlCssWebPartWebPartProps {
  /** Id (GUID) of the library chosen in the property pane. Empty = default to the Style Library. */
  libraryId: string;
  /** Server-relative URL of the selected style file. */
  fileUrl: string;
}
