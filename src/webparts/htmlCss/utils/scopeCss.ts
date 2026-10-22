/**
 * Prefixes every selector with the web part's root selector so the file's CSS
 * cannot leak onto the rest of the SharePoint page. Uses the browser's own CSS
 * parser, so @media / @supports blocks are handled; @keyframes, @font-face etc.
 * are passed through unchanged.
 *
 * `html`, `body` and `:root` selectors are mapped to the web part root itself.
 */
export function scopeCss(css: string, scope: string): string {
  const sheet = new CSSStyleSheet();
  sheet.replaceSync(css);
  return scopeRules(sheet.cssRules, scope);
}

function scopeRules(rules: CSSRuleList, scope: string): string {
  const out: string[] = [];
  for (let i = 0; i < rules.length; i++) {
    out.push(scopeRule(rules[i], scope));
  }
  return out.join('\n');
}

function scopeRule(rule: CSSRule, scope: string): string {
  if (rule instanceof CSSStyleRule) {
    const selectors = rule.selectorText
      .split(/,(?![^(]*\))/) // commas inside :is(...) / :not(...) stay put
      .map((s) => scopeSelector(s.trim(), scope))
      .join(', ');
    return `${selectors} { ${rule.style.cssText} }`;
  }
  if (rule instanceof CSSMediaRule) {
    return `@media ${rule.media.mediaText} { ${scopeRules(rule.cssRules, scope)} }`;
  }
  if (rule instanceof CSSSupportsRule) {
    return `@supports ${rule.conditionText} { ${scopeRules(rule.cssRules, scope)} }`;
  }
  return rule.cssText;
}

function scopeSelector(selector: string, scope: string): string {
  if (/^(html|body|:root)$/i.test(selector)) {
    return scope;
  }
  return `${scope} ${selector.replace(/^(html|body|:root)\s+/i, '')}`;
}
