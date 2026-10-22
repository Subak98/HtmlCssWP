export interface IParsedStyle {
  /** Inline <style> content. */
  text?: string;
  /** Absolute URL of a <link rel="stylesheet">. */
  href?: string;
}

export interface IParsedScript {
  /** Inline <script> content. */
  text?: string;
  /** Absolute URL of a <script src>. */
  src?: string;
}

export interface IParsedHtmlFile {
  /** Markup that remains after <style>, <link rel=stylesheet> and <script> are taken out. */
  bodyHtml: string;
  styles: IParsedStyle[];
  scripts: IParsedScript[];
}

export type ParseResult =
  | { ok: true; file: IParsedHtmlFile }
  | { ok: false; reason: string };

/**
 * "Property format" = the .html file contains:
 *   - HTML markup (anything left in <body> besides style/script)   [required]
 *   - CSS  (a <style> block or <link rel="stylesheet">)            [required]
 *   - JS   (a <script> block or <script src>)                      [optional]
 * To make CSS optional too, remove the CSS check below.
 */
export function parseHtmlFile(source: string, fileUrl: string): ParseResult {
  if (!/\.html?$/i.test(fileUrl)) {
    return { ok: false, reason: "It is not an .html file." };
  }

  // DOMParser never executes scripts or loads resources.
  const doc = new DOMParser().parseFromString(source, "text/html");
  const abs = (u: string): string => new URL(u, fileUrl).href;

  const styles: IParsedStyle[] = [];
  const scripts: IParsedScript[] = [];

  doc.querySelectorAll('style, link[rel~="stylesheet"]').forEach((el) => {
    if (el.tagName === "STYLE") {
      if (el.textContent && el.textContent.trim()) {
        styles.push({ text: el.textContent });
      }
    } else {
      const href = el.getAttribute("href");
      if (href) {
        styles.push({ href: abs(href) });
      }
    }
    if (el.parentNode) {
      el.parentNode.removeChild(el);
    }
  });

  doc.querySelectorAll("script").forEach((el) => {
    const src = el.getAttribute("src");
    if (src) {
      scripts.push({ src: abs(src) });
    } else if (el.textContent && el.textContent.trim()) {
      scripts.push({ text: el.textContent });
    }
    if (el.parentNode) {
      el.parentNode.removeChild(el);
    }
  });

  const bodyHtml = doc.body ? doc.body.innerHTML.trim() : "";

  const missing: string[] = [];
  if (!bodyHtml) {
    missing.push("HTML");
  }
  if (styles.length === 0) {
    missing.push("CSS");
  }

  if (missing.length > 0) {
    return { ok: false, reason: `Missing: ${missing.join(", ")}.` };
  }
  return { ok: true, file: { bodyHtml, styles, scripts } };
}
