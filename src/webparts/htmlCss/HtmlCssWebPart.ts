import * as React from "react";
import * as ReactDom from "react-dom";
import { DisplayMode, Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneDropdownOption,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneDropdown,
  PropertyPaneLabel,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "HtmlCssWebPartStrings";
import HtmlCss from "./components/HtmlCss";
import {
  IHtmlCssProps,
  IHtmlCssWebPartWebPartProps,
} from "./components/IHtmlCssProps";
import {
  ILibraryInfo,
  StyleLibraryService,
} from "./services/StyleLibraryService";

export default class HtmlCssWebPart extends BaseClientSideWebPart<IHtmlCssWebPartWebPartProps> {
  private _instanceId: string = "";
  private _service!: StyleLibraryService;
  private _themeVars: { [name: string]: string } = {};
  private _hasRendered: boolean = false;

  // ---- property pane state (loaded lazily when the pane opens) ----
  private _libraries: ILibraryInfo[] = [];
  private _libraryOptions: IPropertyPaneDropdownOption[] = [];
  private _fileOptions: IPropertyPaneDropdownOption[] = [];
  private _styleLibraryId: string = "";
  /** Library whose files are currently in _fileOptions. */
  private _activeLibraryId: string = "";
  private _librariesLoaded: boolean = false;
  private _loadingLibraries: boolean = false;
  private _loadingFiles: boolean = false;
  private _pickerMessage: string = "";

  public render(): void {
    this._hasRendered = true;

    const element: React.ReactElement<IHtmlCssProps> = React.createElement(
      HtmlCss,
      {
        fileUrl: this.properties.fileUrl,
        service: this._service,
        instanceId: this._instanceId,
        displayMode: this.displayMode as unknown as number,
        themeVars: this._themeVars,
      },
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._instanceId = `html-css-wp-${this.context.instanceId}`;
    this._service = new StyleLibraryService(this.context);
    return super.onInit();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors } = currentTheme;
    const vars: { [name: string]: string } = {};

    if (semanticColors) {
      if (semanticColors.bodyText) {
        vars["--bodyText"] = semanticColors.bodyText;
      }
      if (semanticColors.link) {
        vars["--link"] = semanticColors.link;
      }
      if (semanticColors.linkHovered) {
        vars["--linkHovered"] = semanticColors.linkHovered;
      }
    }

    // The content lives in a sandboxed iframe, so the theme colours are passed into it
    // (CSS variables set on this.domElement would not cross the iframe boundary).
    this._themeVars = vars;
    if (this._hasRendered) {
      this.render();
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  // -------------------------------------------------------------------------
  // Property pane
  // -------------------------------------------------------------------------

  /** Library to show as selected: the saved one, otherwise the Style Library by default. */
  private _effectiveLibraryId(): string {
    return this.properties.libraryId || this._styleLibraryId || "";
  }

  protected onPropertyPaneConfigurationStart(): void {
    if (this._librariesLoaded) {
      return;
    }

    this._loadingLibraries = true;
    this._pickerMessage = strings.LoadingLibraries;
    this.context.propertyPane.refresh();

    this._service
      .getLibraries()
      .then((libraries) => {
        this._libraries = libraries;
        this._libraryOptions = libraries.map((l) => ({
          key: l.id,
          text: l.isStyleLibrary ? `${l.title} (default)` : l.title,
        }));
        const styleLibrary = libraries.filter((l) => l.isStyleLibrary)[0];
        this._styleLibraryId = styleLibrary ? styleLibrary.id : "";
        this._librariesLoaded = true;
        this._pickerMessage = "";
      })
      .catch(() => {
        this._pickerMessage = strings.LibrariesLoadError;
      })
      .then(() => {
        this._loadingLibraries = false;
        return this._loadFiles(this._effectiveLibraryId());
      })
      .then(() => this.context.propertyPane.refresh())
      .catch(() => undefined);
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any,
  ): void {
    if (propertyPath === "libraryId" && newValue !== this._activeLibraryId) {
      // A different library was picked: forget the previous file and list the new library's files.
      this.properties.fileUrl = "";
      this._fileOptions = [];
      this._loadFiles(newValue as string)
        .then(() => this.context.propertyPane.refresh())
        .catch(() => undefined);
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  private async _loadFiles(libraryId: string): Promise<void> {
    this._activeLibraryId = libraryId;
    this._fileOptions = [];

    if (!libraryId) {
      return;
    }

    const library: ILibraryInfo | undefined = this._libraries.filter(
      (l) => l.id === libraryId,
    )[0];

    this._loadingFiles = true;
    this._pickerMessage = strings.LoadingFiles;
    this.context.propertyPane.refresh();

    try {
      const files = await this._service.getFiles(
        libraryId,
        library ? library.rootUrl : "",
      );
      if (this._activeLibraryId !== libraryId) {
        return; // user already switched to another library
      }
      this._fileOptions = files.map((f) => ({ key: f.url, text: f.path }));
      this._pickerMessage =
        files.length === 0
          ? strings.NoFilesFound
          : strings.FilesFound.replace("{0}", String(files.length));
    } catch (e) {
      if (this._activeLibraryId === libraryId) {
        this._pickerMessage = strings.FilesLoadError;
      }
    } finally {
      if (this._activeLibraryId === libraryId) {
        this._loadingFiles = false;
      }
    }
  }

  private _onRefreshClicked = (): void => {
    this._loadFiles(this._effectiveLibraryId())
      .then(() => {
        this.context.propertyPane.refresh();
        this.render(); // also re-read the selected file, e.g. after it was edited
      })
      .catch(() => undefined);
  };

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.SourceGroupName,
              groupFields: [
                PropertyPaneDropdown("libraryId", {
                  label: strings.LibraryFieldLabel,
                  options: this._libraryOptions,
                  selectedKey: this._effectiveLibraryId(),
                  disabled:
                    this._loadingLibraries || this._libraryOptions.length === 0,
                }),
                PropertyPaneDropdown("fileUrl", {
                  label: strings.FileFieldLabel,
                  options: this._fileOptions,
                  selectedKey: this.properties.fileUrl,
                  disabled:
                    this._loadingLibraries ||
                    this._loadingFiles ||
                    this._fileOptions.length === 0,
                }),
                PropertyPaneLabel("pickerStatus", {
                  text: this._pickerMessage,
                }),
                PropertyPaneButton("refreshFiles", {
                  text: strings.RefreshButtonLabel,
                  icon: "Refresh",
                  buttonType: PropertyPaneButtonType.Normal,
                  disabled: this._loadingLibraries || this._loadingFiles,
                  onClick: this._onRefreshClicked,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
// Re-exported so DisplayMode stays a single import point if consumers
// need it; not required for the web part itself.
export { DisplayMode };
