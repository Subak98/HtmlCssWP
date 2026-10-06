import * as React from 'react';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import styles from './CodeEditorHost.module.scss';

export interface ICodeEditorHostProps {
  label: string;
  html: string;
  css: string;
  disabled?: boolean;
  onSave: (html: string, css: string) => void;
}

type TabKey = 'html' | 'css' | 'preview';

interface ICodeEditorHostState {
  isOpen: boolean;
  activeTab: TabKey;
  draftHtml: string;
  draftCss: string;
  previewSrcDoc: string;
}

/**
 * Renders inside the property pane. Shows a single "Edit HTML / CSS" button;
 * clicking it opens a full-height Panel with three tabs: HTML source,
 * CSS source, and a sandboxed iframe live preview that combines both.
 */
export default class CodeEditorHost extends React.Component<ICodeEditorHostProps, ICodeEditorHostState> {
  constructor(props: ICodeEditorHostProps) {
    super(props);
    this.state = {
      isOpen: false,
      activeTab: 'html',
      draftHtml: props.html || '',
      draftCss: props.css || '',
      previewSrcDoc: this._buildPreviewDoc(props.html || '', props.css || '')
    };
  }

  public render(): React.ReactElement<ICodeEditorHostProps> {
    const { label, disabled } = this.props;
    const { isOpen, activeTab, draftHtml, draftCss, previewSrcDoc } = this.state;

    return (
      <div>
        <DefaultButton
          className={styles.launcher}
          iconProps={{ iconName: 'Code' }}
          text={label || 'Edit HTML / CSS'}
          disabled={disabled}
          onClick={this._openPanel}
        />

        <Panel
          className={styles.editorPanel}
          isOpen={isOpen}
          onDismiss={this._closePanel}
          type={PanelType.large}
          headerText="Edit HTML / CSS"
          closeButtonAriaLabel="Close"
          isFooterAtBottom
        >
          <div className={styles.panelBody}>
            <MessageBar messageBarType={MessageBarType.warning}>
              Custom HTML/CSS runs directly on the page. Only paste markup you trust.
            </MessageBar>

            <div className={styles.tabBar}>
              <button
                className={activeTab === 'html' ? styles.tabButtonActive : styles.tabButton}
                onClick={() => this._setTab('html')}
              >
                HTML
              </button>
              <button
                className={activeTab === 'css' ? styles.tabButtonActive : styles.tabButton}
                onClick={() => this._setTab('css')}
              >
                CSS
              </button>
              <button
                className={activeTab === 'preview' ? styles.tabButtonActive : styles.tabButton}
                onClick={() => this._setTab('preview', true)}
              >
                Preview
              </button>
            </div>

            <div className={styles.editorArea}>
              {activeTab === 'html' && (
                <textarea
                  className={styles.codeTextarea}
                  spellCheck={false}
                  value={draftHtml}
                  onChange={this._onHtmlChange}
                  placeholder="<div class=&quot;my-widget&quot;>Hello world</div>"
                />
              )}

              {activeTab === 'css' && (
                <textarea
                  className={styles.codeTextarea}
                  spellCheck={false}
                  value={draftCss}
                  onChange={this._onCssChange}
                  placeholder=".my-widget { color: #0078d4; }"
                />
              )}

              {activeTab === 'preview' && (
                <iframe
                  className={styles.previewFrame}
                  title="HTML/CSS preview"
                  sandbox="allow-same-origin"
                  srcDoc={previewSrcDoc}
                />
              )}
            </div>

            <p className={styles.hint}>
              CSS is automatically scoped to this web part instance when rendered on the page, so it
              will not leak onto the rest of the site. The preview tab approximates the final look.
            </p>

            <div className={styles.footer}>
              <PrimaryButton text="Save" onClick={this._save} />
              <DefaultButton text="Cancel" onClick={this._closePanel} />
            </div>
          </div>
        </Panel>
      </div>
    );
  }

  private _openPanel = (): void => {
    this.setState({
      isOpen: true,
      draftHtml: this.props.html || '',
      draftCss: this.props.css || ''
    });
  };

  private _closePanel = (): void => {
    this.setState({ isOpen: false });
  };

  private _setTab = (tab: TabKey, refreshPreview?: boolean): void => {
    this.setState((prev) => ({
      activeTab: tab,
      previewSrcDoc: refreshPreview
        ? this._buildPreviewDoc(prev.draftHtml, prev.draftCss)
        : prev.previewSrcDoc
    }));
  };

  private _onHtmlChange = (ev: React.ChangeEvent<HTMLTextAreaElement>): void => {
    this.setState({ draftHtml: ev.target.value });
  };

  private _onCssChange = (ev: React.ChangeEvent<HTMLTextAreaElement>): void => {
    this.setState({ draftCss: ev.target.value });
  };

  private _save = (): void => {
    this.props.onSave(this.state.draftHtml, this.state.draftCss);
    this.setState({ isOpen: false });
  };

  private _buildPreviewDoc(html: string, css: string): string {
    return `<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8" />
    <style>
      body { font-family: 'Segoe UI', Arial, sans-serif; margin: 12px; }
      ${css || ''}
    </style>
  </head>
  <body>
    ${html || '<p style="color:#a19f9d;">Nothing to preview yet.</p>'}
  </body>
</html>`;
  }
}
