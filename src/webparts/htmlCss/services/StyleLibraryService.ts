import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import type { WebPartContext } from "@microsoft/sp-webpart-base";

/** Properties a style file MUST contain (each a string; empty strings are allowed). */
export const REQUIRED_PROPERTIES: string[] = ["html", "css", "js"];

/** Largest style file the web part will download and render. */
export const MAX_FILE_BYTES: number = 1024 * 1024; // 1 MB

/** How many library items are listed in the file dropdown. */
export const MAX_LISTED_FILES: number = 2000;

/** Only these extensions are read. Anything else is reported as "not in the expected format". */
const SUPPORTED_EXTENSION: RegExp = /\.(json|txt|html?)$/i;

export interface ILibraryInfo {
  id: string;
  title: string;
  /** Server-relative URL of the library root, e.g. /sites/hr/Style Library */
  rootUrl: string;
  isStyleLibrary: boolean;
}

export interface IStyleFileInfo {
  /** Server-relative URL of the file (used as the dropdown key and stored in the web part properties). */
  url: string;
  name: string;
  /** Path relative to the library root, e.g. "banners/welcome.json". Shown in the dropdown. */
  path: string;
}

export interface IStyleFileContent {
  html: string;
  css: string;
  js: string;
}

export type ReadFailureKind =
  | "format"
  | "notFound"
  | "forbidden"
  | "tooLarge"
  | "error";

export type ReadResult =
  | { ok: true; content: IStyleFileContent }
  | { ok: false; kind: ReadFailureKind; detail?: string };

/**
 * Pure function (no network): validates the text of a style file.
 * Valid = a JSON object that has string properties "html", "css" and "js".
 */
export function parseStyleFile(text: string): ReadResult {
  let data: any;
  try {
    data = JSON.parse((text || "").replace(/^\uFEFF/, ""));
  } catch (e) {
    return { ok: false, kind: "format", detail: "The file is not valid JSON." };
  }

  if (!data || typeof data !== "object" || Array.isArray(data)) {
    return {
      ok: false,
      kind: "format",
      detail: "The JSON root must be an object.",
    };
  }

  const missing: string[] = REQUIRED_PROPERTIES.filter(
    (key) => typeof data[key] !== "string",
  );
  if (missing.length > 0) {
    return {
      ok: false,
      kind: "format",
      detail: `Missing or invalid propert${missing.length > 1 ? "ies" : "y"}: ${missing.join(", ")}.`,
    };
  }

  return { ok: true, content: { html: data.html, css: data.css, js: data.js } };
}

/**
 * Pure function (no network): parses a .txt/.html file that contains plain markup, e.g.
 *
 *   <style> .box { color: red } </style>
 *   <div class="box">Hello</div>
 *   <script> ... </script>        (optional)
 *
 * <style> blocks become "css", inline <script> blocks become "js" and everything else is "html".
 * A full HTML document (<html>/<head>/<body>) is also accepted. The file must contain at least one
 * HTML element, otherwise it is reported as not being in the expected format.
 */
export function parseMarkupFile(text: string): ReadResult {
  let source: string = (text || "").replace(/^\uFEFF/, "");

  const cssParts: string[] = [];
  const jsParts: string[] = [];

  source = source.replace(
    /<style\b[^>]*>([\s\S]*?)<\/style\s*>/gi,
    (_m: string, body: string) => {
      cssParts.push(body.trim());
      return "";
    },
  );

  source = source.replace(
    /<script\b([^>]*)>([\s\S]*?)<\/script\s*>/gi,
    (_m: string, attrs: string, body: string) => {
      // External scripts (src=...) can't load inside the sandbox (its CSP only allows inline code).
      if (!/\bsrc\s*=/i.test(attrs)) {
        jsParts.push(body.trim());
      }
      return "";
    },
  );

  // If a whole document was supplied, keep only what belongs in the body.
  const bodyMatch: RegExpMatchArray | null = source.match(
    /<body\b[^>]*>([\s\S]*)<\/body\s*>/i,
  );
  if (bodyMatch) {
    source = bodyMatch[1];
  } else {
    source = source
      .replace(/<!doctype[^>]*>/gi, "")
      .replace(/<head\b[^>]*>[\s\S]*?<\/head\s*>/gi, "")
      .replace(/<\/?(html|head|body)\b[^>]*>/gi, "");
  }

  const html: string = source.trim();
  const withoutComments: string = html.replace(/<!--[\s\S]*?-->/g, "");
  if (!/<[a-z][a-z0-9-]*[\s>/]/i.test(withoutComments)) {
    return {
      ok: false,
      kind: "format",
      detail: "No HTML elements were found in the file.",
    };
  }

  return {
    ok: true,
    content: { html, css: cssParts.join("\n\n"), js: jsParts.join("\n\n") },
  };
}

/**
 * Pure function: picks the right parser for the file's content.
 *  - Text that starts with "{" is treated as the JSON property format ({ html, css, js }).
 *  - .json files are always validated as JSON (so a broken JSON file reports a JSON error).
 *  - Anything else (.txt / .html / .htm) is parsed as plain markup with optional <style>/<script>.
 */
export function parseStyleContent(text: string, fileUrl: string): ReadResult {
  const trimmed: string = (text || "").replace(/^\uFEFF/, "").trim();
  const isJsonFile: boolean = /\.json$/i.test(fileUrl || "");

  if (isJsonFile || trimmed.charAt(0) === "{") {
    return parseStyleFile(trimmed);
  }
  return parseMarkupFile(trimmed);
}

export class StyleLibraryService {
  private _sp: SPFI;

  constructor(context: WebPartContext) {
    this._sp = spfi().using(SPFx(context));
  }

  /**
   * Document libraries of the current site. The Style Library is returned even though SharePoint
   * marks it hidden, and is always sorted first so it can be used as the default.
   */
  public async getLibraries(): Promise<ILibraryInfo[]> {
    const lists: any[] = await this._sp.web.lists
      .filter("BaseTemplate eq 101")
      .select("Id", "Title", "Hidden", "RootFolder/ServerRelativeUrl")
      .expand("RootFolder")();
    console.log(lists);

    return lists
      .map((l) => {
        const rootUrl: string = l.RootFolder
          ? l.RootFolder.ServerRelativeUrl
          : "";
        return {
          id: String(l.Id),
          title: String(l.Title),
          rootUrl,
          isStyleLibrary: /\/style library$/i.test(rootUrl),
          hidden: !!l.Hidden,
        };
      })
      .filter((l) => !l.hidden || l.isStyleLibrary)
      .sort((a, b) => {
        if (a.isStyleLibrary !== b.isStyleLibrary) {
          return a.isStyleLibrary ? -1 : 1;
        }
        return a.title.localeCompare(b.title);
      })
      .map((l) => ({
        id: l.id,
        title: l.title,
        rootUrl: l.rootUrl,
        isStyleLibrary: l.isStyleLibrary,
      }));
  }

  /**
   * All files in the library, including files in sub-folders (folders themselves are excluded).
   * Filtering/sorting is done client-side so this also works on large libraries without hitting
   * the list view threshold; only the first MAX_LISTED_FILES items are considered.
   */
  public async getFiles(
    libraryId: string,
    libraryRootUrl: string,
  ): Promise<IStyleFileInfo[]> {
    const items: any[] = await this._sp.web.lists
      .getById(libraryId)
      .items.select("Id", "FileLeafRef", "FileRef", "FSObjType")
      .top(MAX_LISTED_FILES)();

    const prefix: string = (libraryRootUrl || "").toLowerCase() + "/";

    return items
      .filter((i) => Number(i.FSObjType) === 0 && !!i.FileRef)
      .map((i) => {
        const url: string = String(i.FileRef);
        const path: string =
          url.toLowerCase().indexOf(prefix) === 0
            ? url.substring(prefix.length)
            : String(i.FileLeafRef);
        return { url, name: String(i.FileLeafRef), path };
      })
      .sort((a, b) => a.path.localeCompare(b.path));
  }

  /** Downloads and validates the selected style file. Never throws; failures are returned. */
  public async readFile(fileUrl: string): Promise<ReadResult> {
    if (!fileUrl) {
      return { ok: false, kind: "error", detail: "No file selected." };
    }

    // Don't download images, PDFs, etc. - they can't be in the property format anyway.
    if (!SUPPORTED_EXTENSION.test(fileUrl)) {
      return {
        ok: false,
        kind: "format",
        detail: "Only .json, .txt and .html files are supported.",
      };
    }

    try {
      const info: any = await this._sp.web
        .getFileByServerRelativePath(fileUrl)
        .select("Length")();
      const length: number = parseInt(info && info.Length, 10);
      if (isFinite(length) && length > MAX_FILE_BYTES) {
        return { ok: false, kind: "tooLarge" };
      }

      const text: string = await this._sp.web
        .getFileByServerRelativePath(fileUrl)
        .getText();
      return parseStyleContent(text, fileUrl);
    } catch (e) {
      const status: number | undefined = e && (e as any).status;
      if (status === 404) {
        return { ok: false, kind: "notFound" };
      }
      if (status === 401 || status === 403) {
        return { ok: false, kind: "forbidden" };
      }
      return {
        ok: false,
        kind: "error",
        detail: e instanceof Error ? e.message : undefined,
      };
    }
  }
}
