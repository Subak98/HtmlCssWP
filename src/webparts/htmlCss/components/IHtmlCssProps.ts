export interface IHtmlCssProps {
  html: string;
  css: string;
  enableCss: boolean;
  showPlaceholderWhenEmpty: boolean;
  /** Unique instance id used to scope the injected <style> block. */
  instanceId: string;
  /** True when the page/web part is in edit mode (Pages edit mode in SharePoint). */
  displayMode: number;
}
export interface IHtmlCssWebPartWebPartProps {
  /** Raw HTML markup entered by the editor. */
  html: string;
  /** Raw CSS entered by the editor. Scoped automatically to this instance. */
  css: string;
  /** Master switch to enable/disable the custom CSS injection. */
  enableCss: boolean;
  /** Show a friendly placeholder in edit mode when no HTML has been entered yet. */
  showPlaceholderWhenEmpty: boolean;
}
