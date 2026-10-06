import * as React from "react";
import styles from "./HtmlCss.module.scss";
import type { IHtmlCssProps } from "./IHtmlCssProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { Icon } from "@fluentui/react";
interface IHtmlCssWebPartState {
  hasRenderError: boolean;
  errorMessage: string;
}

/**
 * Wraps a CSS string so every rule is scoped under the instance's
 * unique root selector. This is a light-weight scoping approach
 * (no shadow DOM) that keeps custom CSS from leaking onto the rest
 * of the page, while still allowing normal cascading CSS authoring.
 */

function scopeCss(css: string, scopeSelector: string): string {
  if (!css || !css.trim()) {
    return "";
  }

  // Strip HTML comments/script-ish content defensively; this is a
  // property-pane-authored value (site owner/editor controlled), not
  // arbitrary anonymous user input, but we still avoid obvious injection
  // vectors like </style> breaking out of the tag.
  const sanitized = css.replace(/<\/style>/gi, "");

  // Very small "scoper": prefix each top-level selector group with the
  // instance root. Handles the common cases (selectors separated by
  // commas, rules separated by `}`) without pulling in a full CSS parser.
  const ruleRegex = /([^{}]+)\{([^{}]*)\}/g;
  let scoped = "";
  let match: RegExpExecArray | null;

  while ((match = ruleRegex.exec(sanitized)) !== null) {
    const selectors = match[1]
      .split(",")
      .map((s) => s.trim())
      .filter((s) => s.length > 0)
      .map((s) => {
        // Leave @-rules (media queries, keyframes, font-face) untouched;
        // they are handled at a coarser level below.
        if (s.indexOf("@") === 0) {
          return s;
        }
        return `${scopeSelector} ${s}`;
      })
      .join(", ");
    scoped += `${selectors} {${match[2]}}\n`;
  }

  // Fallback: if the regex-based pass produced nothing (e.g. the author
  // used @media / @keyframes blocks the simple regex can't split), just
  // wrap the whole thing in a scoping @layer-less block via :where so the
  // browser applies specificity sanely without us re-implementing a CSS parser.
  if (!scoped.trim()) {
    scoped = `${scopeSelector} { all: revert; }\n${sanitized}`;
  }

  return scoped;
}

export default class HtmlCss extends React.Component<
  IHtmlCssProps,
  IHtmlCssWebPartState
> {
  private _styleTag: HTMLStyleElement | undefined;

  constructor(props: IHtmlCssProps) {
    super(props);
    this.state = { hasRenderError: false, errorMessage: "" };
  }

  public componentDidMount(): void {
    this._applyCss();
  }

  public componentDidUpdate(prevProps: IHtmlCssProps): void {
    if (
      prevProps.css !== this.props.css ||
      prevProps.enableCss !== this.props.enableCss
    ) {
      this._applyCss();
    }
  }
  public componentWillUnmount(): void {
    this._removeCss();
  }

  public render(): React.ReactElement<IHtmlCssProps> {
    const { html, showPlaceholderWhenEmpty, displayMode, instanceId } =
      this.props;
    const isEditMode = displayMode === 2;
    const isEmpty = !html || !html.trim();

    if (isEmpty && showPlaceholderWhenEmpty && isEditMode) {
      return (
        <div className={styles.htmlCssWebPart} id={instanceId}>
          <div className={styles.emptyPlaceholder}>
            <Icon iconName="Code" className={styles.icon} />
            <div className={styles.title}>No content yet</div>
            <div className={styles.subtitle}>
              Open the property pane and use &quot;Edit HTML / CSS&quot; to add
              your markup and styles.
            </div>
          </div>
        </div>
      );
    }

    return (
      <div className={styles.htmlCssWebPart} id={instanceId}>
        {this.state.hasRenderError && (
          <div className={styles.errorBanner}>
            There was a problem rendering the custom HTML:{" "}
            {this.state.errorMessage}
          </div>
        )}
        <div
          className={styles.renderRoot}
          // eslint-disable-next-line react/no-danger
          dangerouslySetInnerHTML={{ __html: html || "" }}
        />
      </div>
    );
  }
  public componentDidCatch(error: Error): void {
    this.setState({ hasRenderError: true, errorMessage: error.message });
  }

  private _applyCss(): void {
    this._removeCss();

    if (!this.props.enableCss || !this.props.css || !this.props.css.trim()) {
      return;
    }

    const scopeSelector = `#${this.props.instanceId}`;
    const cssText = scopeCss(this.props.css, scopeSelector);

    const styleTag = document.createElement("style");
    styleTag.setAttribute("data-html-css-webpart", this.props.instanceId);
    styleTag.innerHTML = cssText;
    document.head.appendChild(styleTag);
    this._styleTag = styleTag;
  }
  private _removeCss(): void {
    if (this._styleTag && this._styleTag.parentNode) {
      this._styleTag.parentNode.removeChild(this._styleTag);
    }
    this._styleTag = undefined;
  }
}
