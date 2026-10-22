import * as React from "react";
import * as ReactDom from "react-dom";
import { DisplayMode, Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneField,
  type IPropertyPaneDropdownOption,
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

  // Property pane state
  private _libraries: ILibraryInfo[] = [];
  private _libraryOptions: IPropertyPaneDropdownOption[] = [];
  private _fileOptions: IPropertyPaneDropdownOption[] = [];
  private _paneLoading: boolean = false;
  private _paneError: string = "";

  public render(): void {
    const element: React.ReactElement<IHtmlCssProps> = React.createElement(
      HtmlCss,
      {
        spHttpClient: this.context.spHttpClient,
        webUrl: this.context.pageContext.web.absoluteUrl,
        fileUrl: this.properties.fileUrl,
        showPlaceholderWhenEmpty:
          this.properties.showPlaceholderWhenEmpty !== false,
        instanceId: this._instanceId,
        isEditMode: this.displayMode === DisplayMode.Edit,
      },
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._instanceId = `html-css-wp-${this.context.instanceId}`;
    this._service = new StyleLibraryService(
      this.context.spHttpClient,
      this.context.pageContext.web.absoluteUrl,
    );
    return super.onInit();
  }

  // ---- Property pane: library picker -> file dropdown -------------------

  protected onPropertyPaneConfigurationStart(): void {
    this._refreshPane().catch(() => {
      /* handled inside */
    });
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown,
  ): void {
    if (propertyPath === "libraryId" && oldValue !== newValue) {
      // A different library was picked: drop the old file and list the new library's files.
      this.properties.fileUrl = undefined;
      this._fileOptions = [];
      this._refreshPane().catch(() => {
        /* handled inside */
      });
    }
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  /** Loads libraries (once), defaults to the Style Library, then loads that library's files. */
  private async _refreshPane(): Promise<void> {
    this._paneLoading = true;
    this._paneError = "";
    this.context.propertyPane.refresh();

    try {
      if (this._libraries.length === 0) {
        this._libraries = await this._service.getLibraries();
        this._libraryOptions = this._libraries.map((l) => ({
          key: l.id,
          text: l.title,
        }));
      }

      if (!this.properties.libraryId) {
        const styleLibrary = this._libraries.filter((l) => l.isStyleLibrary)[0];
        if (styleLibrary) {
          this.properties.libraryId = styleLibrary.id;
        }
      }

      const library = this._libraries.filter(
        (l) => l.id === this.properties.libraryId,
      )[0];
      this._fileOptions = library
        ? (await this._service.getFiles(library)).map((f) => ({
            key: f.serverRelativeUrl,
            text: f.displayName,
          }))
        : [];
    } catch (err) {
      this._paneError = `${strings.LoadError} ${(err as Error).message}`;
    } finally {
      this._paneLoading = false;
      this.context.propertyPane.refresh();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const fields: IPropertyPaneField<unknown>[] = [
      PropertyPaneDropdown("libraryId", {
        label: strings.LibraryFieldLabel,
        options: this._libraryOptions,
        selectedKey: this.properties.libraryId,
        disabled: this._paneLoading || this._libraryOptions.length === 0,
      }),
      PropertyPaneDropdown("fileUrl", {
        label: strings.FileFieldLabel,
        options: this._fileOptions,
        selectedKey: this.properties.fileUrl,
        disabled: this._paneLoading || this._fileOptions.length === 0,
      }),
    ];

    if (this._paneLoading) {
      fields.push(PropertyPaneLabel("", { text: strings.Loading }));
    } else if (this._paneError) {
      fields.push(PropertyPaneLabel("", { text: this._paneError }));
    } else if (this.properties.libraryId && this._fileOptions.length === 0) {
      fields.push(PropertyPaneLabel("", { text: strings.NoFilesFound }));
    }

    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.ContentFieldsGroupName,
              groupFields: fields,
            },
          ],
        },
      ],
    };
  }

  // ---- Unchanged plumbing ----------------------------------------------

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null,
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null,
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }
}
