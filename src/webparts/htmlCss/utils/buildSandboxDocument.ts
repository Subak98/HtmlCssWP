import type { IStyleFileContent } from "../services/StyleLibraryService";

/** postMessage type the iframe uses to tell the web part how tall its content is. */
export const HEIGHT_MESSAGE_TYPE: string = "dv-htmlcss-height";

/**
 * Extra HTTPS origins (besides the current SharePoint tenant origin) that images/media inside the
 * sandbox may load from. Keep this list as small as possible. Example: ["https://cdn.contoso.com"]
 */
export const TRUSTED_RESOURCE_ORIGINS: string[] = [];

export interface ISandboxOptions {
  content: IStyleFileContent;
  instanceId: string;
  /** CSS custom properties from the SharePoint theme, e.g. { "--bodyText": "#323130" } */
  themeVars: { [name: string]: string };
}

// Only plain colour/length-like values from the theme are written into the document.
const SAFE_THEME_VALUE: RegExp = /^[#a-z0-9(),.\s%-]+$/i;
const SAFE_THEME_NAME: RegExp = /^--[a-z0-9-]+$/i;

function buildThemeCss(themeVars: { [name: string]: string }): string {
  const lines: string[] = [];
  Object.keys(themeVars || {}).forEach((name) => {
    const value: string = themeVars[name];
    if (SAFE_THEME_NAME.test(name) && value && SAFE_THEME_VALUE.test(value)) {
      lines.push(`${name}: ${value};`);
    }
  });
  return lines.length > 0 ? `:root { ${lines.join(" ")} }` : "";
}

/**
 * Builds the full HTML document rendered inside <iframe sandbox="allow-scripts ...">.
 *
 * Why an iframe: the file's JavaScript is authored content. The iframe is sandboxed WITHOUT
 * "allow-same-origin", so the document gets an opaque origin: its script cannot read the
 * SharePoint DOM, cookies, local storage or tokens, and cannot call SharePoint/Graph as the viewer.
 * The CSP below additionally blocks network calls (connect-src 'none'), forms, frames and plugins,
 * so the script can only manipulate its own document.
 */
export function buildSandboxDocument(options: ISandboxOptions): string {
  const origin: string = window.location.origin;
  const imgOrigins: string = [origin].concat(TRUSTED_RESOURCE_ORIGINS).join(" ");

  const csp: string = [
    "default-src 'none'",
    "script-src 'unsafe-inline'",
    "style-src 'unsafe-inline'",
    `img-src data: ${imgOrigins}`,
    `media-src data: ${imgOrigins}`,
    "font-src data:",
    "connect-src 'none'",
    "form-action 'none'",
    "frame-src 'none'",
    "object-src 'none'",
    `base-uri ${origin}`
  ].join("; ");

  // Prevent "</script>" / "</style>" inside the author's code from ending the element early.
  const safeJs: string = (options.content.js || "").replace(/<\/script/gi, "<\\/script");
  const safeCss: string = (options.content.css || "").replace(/<\/style/gi, "<\\/style");

  // Instance id is embedded in a script: JSON-encode it and neutralise "<".
  const idLiteral: string = JSON.stringify(options.instanceId).replace(/</g, "\\u003c");

  // Measures the content wrapper (not the viewport, which would never shrink) and reports it.
  const heightReporter: string = `(function () {
  var root = document.getElementById("__dv_root");
  function post() {
    var h = Math.ceil(root.getBoundingClientRect().height);
    parent.postMessage({ type: "${HEIGHT_MESSAGE_TYPE}", id: ${idLiteral}, height: h }, "*");
  }
  if (window.ResizeObserver) { new ResizeObserver(post).observe(root); }
  window.addEventListener("load", post);
  post();
})();`;

  return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="${csp}">
<base href="${window.location.href.split("#")[0].replace(/"/g, "%22")}" target="_blank">
<style>
${buildThemeCss(options.themeVars)}
html, body { margin: 0; padding: 0; }
body { font-family: "Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif; color: var(--bodyText, #323130); }
a { color: var(--link, #0078d4); }
a:hover { color: var(--linkHovered, #004578); }
#__dv_root { display: flow-root; }
</style>
<style>
${safeCss}
</style>
</head>
<body>
<div id="__dv_root">${options.content.html || ""}</div>
<script>${heightReporter}</script>
<script>
${safeJs}
</script>
</body>
</html>`;
}
