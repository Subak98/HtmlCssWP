import * as React from "react";
import * as ReactDom from "react-dom";
import { DisplayMode, Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneLabel,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "HtmlCssWebPartStrings";
import HtmlCss from "./components/HtmlCss";
import {
  IHtmlCssProps,
  IHtmlCssWebPartWebPartProps,
} from "./components/IHtmlCssProps";
import { PropertyFieldCodeEditor } from "./propertyPane/PropertyFieldCodeEditor";

export interface IHtmlCssWebPartProps {
  description: string;
}

export default class HtmlCssWebPart extends BaseClientSideWebPart<IHtmlCssWebPartWebPartProps> {
  private _instanceId: any;

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  public render(): void {
    const element: React.ReactElement<IHtmlCssProps> = React.createElement(
      HtmlCss,
      {
        html: this.properties.html,
        css: this.properties.css,
        enableCss: this.properties.enableCss !== false,
        showPlaceholderWhenEmpty:
          this.properties.showPlaceholderWhenEmpty !== false,
        instanceId: this._instanceId,
        displayMode: this.displayMode as unknown as number,
      },
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._instanceId = `html-css-wp-${this.context.instanceId}`;
    return super.onInit();
  }

  /**
   * Called by the custom property pane field when the user hits "Save"
   * inside the HTML/CSS editor panel. Persists both properties in one
   * go and triggers a re-render.
   */
  private _onCodeSaved = (html: string, css: string): void => {
    this.properties.html = html;
    this.properties.css = css;
    this.render();
    // Refresh the property pane so the field reflects the saved values
    // (e.g. if the pane is re-opened without a full page reload).
    this.context.propertyPane.refresh();
  };

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
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

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.ContentFieldsGroupName,
              groupFields: [
                PropertyFieldCodeEditor("html", {
                  label: strings.EditButtonLabel,
                  html: this.properties.html,
                  css: this.properties.css,
                  onSave: this._onCodeSaved,
                }),
                PropertyPaneLabel("html", {
                  text:
                    this.properties.html && this.properties.html.trim()
                      ? `${this.properties.html.trim().length} characters of HTML saved.`
                      : "No HTML saved yet.",
                }),
              ],
            },
            // {
            //   groupName: strings.BehaviorGroupName,
            //   groupFields: [
            //     PropertyPaneToggle("enableCss", {
            //       label: strings.EnableCssFieldLabel,
            //       checked: this.properties.enableCss !== false,
            //     }),
            //     PropertyPaneToggle("showPlaceholderWhenEmpty", {
            //       label: strings.ShowPlaceholderFieldLabel,
            //       checked: this.properties.showPlaceholderWhenEmpty !== false,
            //     }),
            //   ],
            // },
          ],
        },
      ],
    };
  }
}
// Re-exported so DisplayMode stays a single import point if consumers
// need it; not required for the web part itself.
export { DisplayMode };
