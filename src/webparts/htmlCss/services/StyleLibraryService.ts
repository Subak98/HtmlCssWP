import { SPHttpClient } from '@microsoft/sp-http';

export interface ILibraryInfo {
  id: string;
  title: string;
  /** Server-relative URL of the library root, e.g. /sites/team/Style Library */
  rootUrl: string;
  isStyleLibrary: boolean;
}

export interface IFileInfo {
  serverRelativeUrl: string;
  /** Path relative to the library root, e.g. "cards/card.html" */
  displayName: string;
}

/** Set to true to list only .html / .htm files in the file dropdown. */
const ONLY_HTML_FILES: boolean = false;

interface IListRow {
  Id: string;
  Title: string;
  Hidden: boolean;
  RootFolder: { ServerRelativeUrl: string };
}

interface IItemRow {
  FileRef: string;
  FSObjType: number;
}

export class StyleLibraryService {
  constructor(
    private _http: SPHttpClient,
    private _webUrl: string,
  ) {}

  /** All document libraries of the current web (Style Library is always included, even if hidden). */
  public async getLibraries(): Promise<ILibraryInfo[]> {
    const url =
      `${this._webUrl}/_api/web/lists?$select=Id,Title,Hidden,RootFolder/ServerRelativeUrl` +
      `&$expand=RootFolder&$filter=BaseTemplate eq 101`;
    const json = await this._getJson<{ value: IListRow[] }>(url);

    return json.value
      .map((l) => ({
        id: l.Id,
        title: l.Title,
        rootUrl: l.RootFolder.ServerRelativeUrl,
        isStyleLibrary: /\/style library$/i.test(l.RootFolder.ServerRelativeUrl),
        hidden: l.Hidden,
      }))
      .filter((l) => !l.hidden || l.isStyleLibrary)
      .map(({ id, title, rootUrl, isStyleLibrary }) => ({
        id,
        title,
        rootUrl,
        isStyleLibrary,
      }))
      .sort((a, b) =>
        a.isStyleLibrary !== b.isStyleLibrary
          ? a.isStyleLibrary
            ? -1
            : 1
          : a.title.localeCompare(b.title),
      );
  }

  /** Every file in the library, including files inside sub-folders. */
  public async getFiles(library: ILibraryInfo): Promise<IFileInfo[]> {
    const files: IFileInfo[] = [];
    const prefix = `${library.rootUrl}/`;
    let next: string | undefined =
      `${this._webUrl}/_api/web/lists(guid'${library.id}')/items` +
      `?$select=FileRef,FSObjType&$top=1000`;

    while (next) {
      const json: { value: IItemRow[]; 'odata.nextLink'?: string } =
        await this._getJson(next);

      json.value.forEach((item) => {
        if (item.FSObjType !== 0) {
          return; // folder
        }
        if (ONLY_HTML_FILES && !/\.html?$/i.test(item.FileRef)) {
          return;
        }
        files.push({
          serverRelativeUrl: item.FileRef,
          displayName:
            item.FileRef.indexOf(prefix) === 0
              ? item.FileRef.substring(prefix.length)
              : item.FileRef,
        });
      });
      next = json['odata.nextLink'];
    }

    return files.sort((a, b) => a.displayName.localeCompare(b.displayName));
  }

  /** Raw text of a file, given its server-relative URL. */
  public async getFileContent(serverRelativeUrl: string): Promise<string> {
    const path = encodeURIComponent(serverRelativeUrl.replace(/'/g, "''"));
    const url = `${this._webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${path}')/$value`;
    const res = await this._http.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) {
      throw new Error(`${res.status} ${res.statusText}`);
    }
    return res.text();
  }

  private async _getJson<T>(url: string): Promise<T> {
    const res = await this._http.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) {
      throw new Error(`${res.status} ${res.statusText}`);
    }
    return (await res.json()) as T;
  }
}
