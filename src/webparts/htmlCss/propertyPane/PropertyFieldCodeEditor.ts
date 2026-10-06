import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType,
  IPropertyPaneCustomFieldProps
} from '@microsoft/sp-property-pane';
import CodeEditorHost, { ICodeEditorHostProps } from './CodeEditorHost';

export interface IPropertyFieldCodeEditorProps {
  /** Text shown on the launcher button, e.g. "Edit HTML / CSS". */
  label: string;
  html: string;
  css: string;
  disabled?: boolean;
  /** Called with the new html/css when the user clicks Save in the panel. */
  onSave: (html: string, css: string) => void;
}

export interface IPropertyFieldCodeEditorInternalProps
  extends IPropertyFieldCodeEditorProps,
    IPropertyPaneCustomFieldProps {
  key: string;
}

class PropertyFieldCodeEditorBuilder implements IPropertyPaneField<IPropertyFieldCodeEditorInternalProps> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldCodeEditorInternalProps;

  constructor(targetProperty: string, properties: IPropertyFieldCodeEditorProps) {
    this.targetProperty = targetProperty;
    this.properties = {
      key: `${targetProperty}-code-editor-field`,
      label: properties.label,
      html: properties.html,
      css: properties.css,
      disabled: properties.disabled,
      onSave: properties.onSave,
      onRender: this._onRender.bind(this),
      onDispose: this._onDispose.bind(this)
    };
  }

  // Required by IPropertyPaneField; the property pane calls onRender/onDispose
  // directly, this hook is unused for custom fields.
  public render(): void {
    /* no-op */
  }

  private _onDispose(element: HTMLElement): void {
    ReactDom.unmountComponentAtNode(element);
  }

  private _onRender(elem: HTMLElement): void {
    const element: React.ReactElement<ICodeEditorHostProps> = React.createElement(CodeEditorHost, {
      label: this.properties.label,
      html: this.properties.html,
      css: this.properties.css,
      disabled: this.properties.disabled,
      onSave: (html: string, css: string) => {
        // Keep the field's own copy in sync in case the pane re-renders
        // before the parent web part's onPropertyPaneFieldChanged fires.
        this.properties.html = html;
        this.properties.css = css;
        this.properties.onSave(html, css);
      }
    });
    ReactDom.render(element, elem);
  }
}

/**
 * Factory function following the same pattern as PropertyPaneTextField etc.
 * Usage:
 *   PropertyFieldCodeEditor('html', {
 *     label: 'Edit HTML / CSS',
 *     html: this.properties.html,
 *     css: this.properties.css,
 *     onSave: (html, css) => { ... }
 *   })
 */
export function PropertyFieldCodeEditor(
  targetProperty: string,
  properties: IPropertyFieldCodeEditorProps
): IPropertyPaneField<IPropertyFieldCodeEditorInternalProps> {
  return new PropertyFieldCodeEditorBuilder(targetProperty, properties);
}
