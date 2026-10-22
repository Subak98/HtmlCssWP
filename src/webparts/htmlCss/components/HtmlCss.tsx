import * as React from "react";
import styles from "./HtmlCss.module.scss";
import type { IHtmlCssProps } from "./IHtmlCssProps";
import {
  Icon,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
} from "@fluentui/react";
import { StyleLibraryService } from "../services/StyleLibraryService";
import { IParsedHtmlFile, parseHtmlFile } from "../utils/htmlFileParser";
import { scopeCss } from "../utils/scopeCss";

type ViewState =
  | { status: "idle" }
  | { status: "loading" }
  | { status: "error"; message: string }
  | { status: "ready"; parsed: IParsedHtmlFile };

async function fetchText(url: string): Promise<string> {
  const res = await fetch(url, { credentials: "same-origin" });
  if (!res.ok) {
    throw new Error(`${res.status} ${res.statusText}`);
  }
  return res.text();
}

function loadExternalScript(host: HTMLElement, src: string): Promise<void> {
  return new Promise<void>((resolve) => {
    const el = document.createElement("script");
    el.src = src;
    el.onload = (): void => resolve();
    el.onerror = (): void => resolve(); // a broken CDN link should not block the rest
    host.appendChild(el);
  });
}

/**
 * Files authored as standalone pages often wait for DOMContentLoaded / load,
 * which already fired for the SharePoint page. While an inline script runs we
 * capture those listeners and call them right after it has executed.
 */
function runWithReadyShim(run: () => void): void {
  const queued: Array<{
    target: Document | Window;
    type: string;
    listener: EventListenerOrEventListenerObject;
  }> = [];
  const restore: Array<() => void> = [];

  const patch = (target: Document | Window, type: string): void => {
    const original = target.addEventListener;
    target.addEventListener = function (
      t: string,
      l: EventListenerOrEventListenerObject,
      o?: boolean | AddEventListenerOptions,
    ): void {
      if (t === type) {
        queued.push({ target, type, listener: l });
        return;
      }
      original.call(target, t, l, o);
    } as typeof target.addEventListener;
    restore.push(() => {
      target.addEventListener = original;
    });
  };

  patch(document, "DOMContentLoaded");
  patch(window, "load");
  try {
    run();
  } finally {
    restore.forEach((r) => r());
  }

  queued.forEach(({ target, type, listener }) => {
    const evt = new Event(type);
    if (typeof listener === "function") {
      listener.call(target, evt);
    } else {
      listener.handleEvent(evt);
    }
  });
}

function runInlineScript(code: string): void {
  // SharePoint's CSP blocks inline <script> elements but allows 'unsafe-eval',
  // so inline code is run with an indirect (global-scope) eval.
  // Wrapped in a block so top-level let/const don't clash when the same file
  // runs twice on a page; `var` and function declarations still become global,
  // so inline onclick="fn()" handlers keep working.
  const globalEval = eval;
  runWithReadyShim(() => {
    try {
      globalEval(`{\n${code}\n}`);
    } catch (err) {
      // eslint-disable-next-line no-console
      console.error("Error while running script from the selected file:", err);
    }
  });
}

const HtmlCss: React.FC<IHtmlCssProps> = (props) => {
  const {
    spHttpClient,
    webUrl,
    fileUrl,
    showPlaceholderWhenEmpty,
    instanceId,
    isEditMode,
  } = props;

  const [state, setState] = React.useState<ViewState>({ status: "idle" });
  const scriptHostRef = React.useRef<HTMLDivElement>(null);

  // 1) Fetch + validate whenever the selected file changes.
  React.useEffect(() => {
    if (!fileUrl) {
      setState({ status: "idle" });
      return undefined;
    }

    let cancelled = false;
    const fileName = fileUrl.substring(fileUrl.lastIndexOf("/") + 1);
    setState({ status: "loading" });

    new StyleLibraryService(spHttpClient, webUrl)
      .getFileContent(fileUrl)
      .then((text) => {
        if (cancelled) {
          return;
        }
        const result = parseHtmlFile(text, window.location.origin + fileUrl);
        setState(
          result.ok
            ? { status: "ready", parsed: result.file }
            : {
                status: "error",
                message: `The selected file "${fileName}" is not in property format. It must contain HTML and CSS (JavaScript is optional). ${result.reason}`,
              },
        );
      })
      .catch((err: Error) => {
        if (!cancelled) {
          setState({
            status: "error",
            message: `Could not read "${fileName}": ${err.message}`,
          });
        }
      });

    return () => {
      cancelled = true;
    };
  }, [fileUrl, webUrl, spHttpClient]);

  // 2) Once the markup is in the DOM, inject the scoped CSS and run the scripts.
  React.useEffect(() => {
    const host = scriptHostRef.current;
    if (state.status !== "ready" || !host) {
      return undefined;
    }

    const { styles: sheets, scripts } = state.parsed;
    const added: HTMLElement[] = [];
    let cancelled = false;

    const apply = async (): Promise<void> => {
      // CSS: inline blocks as-is, linked sheets fetched so they can be scoped too.
      const cssParts: string[] = [];
      const unscopedLinks: string[] = [];
      for (const s of sheets) {
        if (s.text) {
          cssParts.push(s.text);
        } else if (s.href) {
          try {
            cssParts.push(await fetchText(s.href));
          } catch {
            unscopedLinks.push(s.href); // e.g. cross-origin without CORS
          }
        }
      }
      if (cancelled) {
        return;
      }

      const styleTag = document.createElement("style");
      styleTag.setAttribute("data-html-css-webpart", instanceId);
      styleTag.textContent = cssParts
        .map((c) => scopeCss(c, `#${instanceId}`))
        .join("\n");
      document.head.appendChild(styleTag);
      added.push(styleTag);

      unscopedLinks.forEach((href) => {
        const link = document.createElement("link");
        link.rel = "stylesheet";
        link.href = href;
        document.head.appendChild(link);
        added.push(link);
      });

      // JS: in document order; external scripts finish loading before the next one runs.
      for (const s of scripts) {
        if (cancelled) {
          return;
        }
        if (s.src) {
          await loadExternalScript(host, s.src);
        } else if (s.text) {
          runInlineScript(s.text);
        }
      }
    };

    apply().catch((err: Error) => {
      if (!cancelled) {
        setState({
          status: "error",
          message: `There was a problem rendering the file: ${err.message}`,
        });
      }
    });

    return () => {
      cancelled = true;
      added.forEach((el) => el.parentNode && el.parentNode.removeChild(el));
      host.innerHTML = ""; // removes the <script> elements we added
    };
  }, [state, instanceId]);

  return (
    <div className={styles.htmlCssWebPart} id={instanceId}>
      {state.status === "loading" && (
        <Spinner size={SpinnerSize.small} label="Loading..." />
      )}

      {state.status === "error" && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          {state.message}
        </MessageBar>
      )}

      {state.status === "idle" && isEditMode && showPlaceholderWhenEmpty && (
        <div className={styles.emptyPlaceholder}>
          <Icon iconName="Code" className={styles.icon} />
          <div className={styles.title}>No file selected</div>
          <div className={styles.subtitle}>
            Open the property pane, pick a library and choose the .html file to
            render.
          </div>
        </div>
      )}

      {state.status === "ready" && (
        <div
          className={styles.renderRoot}
          // eslint-disable-next-line react/no-danger
          dangerouslySetInnerHTML={{ __html: state.parsed.bodyHtml }}
        />
      )}

      {/* Scripts are appended here, outside React-managed markup. */}
      <div ref={scriptHostRef} style={{ display: "none" }} />
    </div>
  );
};

export default HtmlCss;
