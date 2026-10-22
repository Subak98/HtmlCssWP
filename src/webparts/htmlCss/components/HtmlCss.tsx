import * as React from "react";
import styles from "./HtmlCss.module.scss";
import type { IHtmlCssProps } from "./IHtmlCssProps";
import { Icon, Spinner, SpinnerSize } from "@fluentui/react";
import * as strings from "HtmlCssWebPartStrings";
import type {
  IStyleFileContent,
  ReadFailureKind,
} from "../services/StyleLibraryService";
import {
  buildSandboxDocument,
  HEIGHT_MESSAGE_TYPE,
} from "../utils/buildSandboxDocument";

type LoadStatus = "idle" | "loading" | "ready" | "error";

interface IHtmlCssState {
  status: LoadStatus;
  content?: IStyleFileContent;
  errorKind?: ReadFailureKind;
  errorDetail?: string;
  frameHeight: number;
  /** True once the iframe has reported its real content height. */
  measured: boolean;
}

/** Height used until the iframe reports its real height (or forever, if it can't). */
const FALLBACK_HEIGHT: number = 300;
const MAX_FRAME_HEIGHT: number = 10000;

function errorMessageFor(kind: ReadFailureKind | undefined): string {
  switch (kind) {
    case "format":
      return strings.FormatErrorTitle;
    case "notFound":
      return strings.NotFoundError;
    case "forbidden":
      return strings.ForbiddenError;
    case "tooLarge":
      return strings.TooLargeError;
    default:
      return strings.GenericLoadError;
  }
}

export default class HtmlCss extends React.Component<
  IHtmlCssProps,
  IHtmlCssState
> {
  private _frameRef: React.RefObject<HTMLIFrameElement> =
    React.createRef<HTMLIFrameElement>();
  private _loadToken: number = 0;
  private _unmounted: boolean = false;
  private _srcDocKey: string = "";
  private _srcDoc: string = "";

  constructor(props: IHtmlCssProps) {
    super(props);
    this.state = {
      status: "idle",
      frameHeight: FALLBACK_HEIGHT,
      measured: false,
    };
  }

  public componentDidMount(): void {
    window.addEventListener("message", this._onMessage);
    this._load().catch(() => undefined);
  }

  public componentDidUpdate(prevProps: IHtmlCssProps): void {
    if (prevProps.fileUrl !== this.props.fileUrl) {
      this._load().catch(() => undefined);
    }
  }

  public componentWillUnmount(): void {
    this._unmounted = true;
    window.removeEventListener("message", this._onMessage);
  }

  public render(): React.ReactElement<IHtmlCssProps> {
    const { displayMode, instanceId } = this.props;
    const { status, content, errorKind, errorDetail, frameHeight, measured } =
      this.state;
    const isEditMode: boolean = displayMode === 2;

    let body: React.ReactNode = null;

    if (status === "idle") {
      // Nothing selected yet: guide the author in edit mode, show nothing to readers.
      if (isEditMode) {
        body = (
          <div className={styles.emptyPlaceholder}>
            <Icon iconName="Code" className={styles.icon} />
            <div className={styles.title}>
              {strings.SelectFilePlaceholderTitle}
            </div>
            <div className={styles.subtitle}>
              {strings.SelectFilePlaceholderText}
            </div>
          </div>
        );
      }
    } else if (status === "loading") {
      body = (
        <Spinner size={SpinnerSize.medium} label={strings.LoadingContent} />
      );
    } else if (status === "error") {
      body = (
        <div className={styles.errorBanner} role="alert">
          <div className={styles.errorTitle}>{errorMessageFor(errorKind)}</div>
          {errorKind === "format" && (
            <div className={styles.errorDetail}>{strings.FormatErrorHint}</div>
          )}
          {errorDetail && (
            <div className={styles.errorDetail}>{errorDetail}</div>
          )}
        </div>
      );
    } else if (status === "ready" && content) {
      body = (
        <iframe
          ref={this._frameRef}
          className={styles.frame}
          title={strings.FrameTitle}
          // allow-scripts WITHOUT allow-same-origin => opaque origin: the file's JS cannot touch the
          // SharePoint page, cookies or tokens. Popups are allowed (and escape the sandbox) so that
          // links open as normal pages in a new tab.
          sandbox="allow-scripts allow-popups allow-popups-to-escape-sandbox"
          referrerPolicy="no-referrer"
          srcDoc={this._getSrcDoc(content)}
          scrolling={measured ? "no" : "yes"}
          style={{ height: frameHeight }}
        />
      );
    }

    return (
      <div className={styles.htmlCssWebPart} id={instanceId}>
        {body}
      </div>
    );
  }

  private _getSrcDoc(content: IStyleFileContent): string {
    const key: string = JSON.stringify([
      content,
      this.props.themeVars,
      this.props.instanceId,
    ]);
    if (key !== this._srcDocKey) {
      this._srcDoc = buildSandboxDocument({
        content,
        instanceId: this.props.instanceId,
        themeVars: this.props.themeVars,
      });
      this._srcDocKey = key;
    }
    return this._srcDoc;
  }

  private async _load(): Promise<void> {
    const { fileUrl, service } = this.props;
    const token: number = ++this._loadToken;

    if (!fileUrl) {
      this.setState({
        status: "idle",
        content: undefined,
        errorKind: undefined,
        errorDetail: undefined,
      });
      return;
    }

    this.setState({
      status: "loading",
      measured: false,
      frameHeight: FALLBACK_HEIGHT,
    });

    const result = await service.readFile(fileUrl);
    if (this._unmounted || token !== this._loadToken) {
      return; // a newer selection superseded this request
    }

    if (result.ok) {
      this.setState({
        status: "ready",
        content: result.content,
        errorKind: undefined,
        errorDetail: undefined,
        measured: false,
        frameHeight: FALLBACK_HEIGHT,
      });
    } else {
      this.setState({
        status: "error",
        content: undefined,
        errorKind: result.kind,
        errorDetail: result.detail,
      });
    }
  }

  /** Receives the content height from the sandboxed iframe (it can't be measured from outside). */
  private _onMessage = (ev: MessageEvent): void => {
    const frame: HTMLIFrameElement | null = this._frameRef.current;
    if (!frame || ev.source !== frame.contentWindow) {
      return; // not from our iframe
    }

    const data: any = ev.data;
    if (
      !data ||
      data.type !== HEIGHT_MESSAGE_TYPE ||
      data.id !== this.props.instanceId
    ) {
      return;
    }

    const height: number = Number(data.height);
    if (!isFinite(height) || height < 0) {
      return;
    }

    const clamped: number = Math.min(Math.ceil(height), MAX_FRAME_HEIGHT);
    if (clamped !== this.state.frameHeight || !this.state.measured) {
      this.setState({ frameHeight: clamped, measured: true });
    }
  };
}
