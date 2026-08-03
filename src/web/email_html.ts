import type { MailInlineImage } from "../mail/types.ts";

function normalizeContentId(input: string | undefined): string {
  return (input ?? "")
    .trim()
    .replace(/^cid:/i, "")
    .replace(/^<|>$/g, "")
    .toLowerCase();
}

function rewriteCidImages(
  html: string,
  inlineImages: MailInlineImage[] | undefined,
): string {
  const contentIdMap = new Map<string, string>();
  for (const image of inlineImages ?? []) {
    const dataUrl = `data:${image.contentType};base64,${image.dataBase64}`;
    const contentId = normalizeContentId(image.contentId);
    if (contentId) contentIdMap.set(contentId, dataUrl);
    const normalizedName = normalizeContentId(image.name);
    if (normalizedName) contentIdMap.set(normalizedName, dataUrl);
  }

  return html.replace(
    /(<img\b[^>]*\bsrc\s*=\s*)(["'])(cid:[^"']+)\2/gi,
    (_full, prefix, quote, src) => {
      const resolved = contentIdMap.get(normalizeContentId(String(src)));
      return resolved
        ? `${prefix}${quote}${resolved}${quote}`
        : `${prefix}${quote}${src}${quote}`;
    },
  );
}

/**
 * 邮件 HTML 来自外部发送者，因此只能进入无脚本、无同源权限的 iframe。
 * 这里再移除主动内容和事件属性，既减少无效资源，也避免旧邮件模板影响工作台。
 */
export function sanitizeEmailHtml(
  html: string,
  inlineImages: MailInlineImage[] | undefined,
): string {
  return rewriteCidImages(html, inlineImages)
    .replace(/<!doctype[^>]*>/gi, "")
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<iframe[\s\S]*?<\/iframe>/gi, "")
    .replace(/<object[\s\S]*?<\/object>/gi, "")
    .replace(/<embed[\s\S]*?>/gi, "")
    .replace(/<form[\s\S]*?<\/form>/gi, "")
    .replace(/<base[\s\S]*?>/gi, "")
    .replace(/<meta[\s\S]*?>/gi, "")
    .replace(/<link[\s\S]*?>/gi, "")
    .replace(/<\/?(html|body|head)[^>]*>/gi, "")
    .replace(/\son\w+\s*=\s*(".*?"|'.*?'|[^\s>]+)/gi, "")
    .replace(/\s(href|src)\s*=\s*(['"])\s*javascript:[^'"]*\2/gi, ' $1="#"')
    .replace(/<a\b/gi, '<a target="_blank" rel="noopener noreferrer"');
}

export function buildReaderDocumentHtml(
  html: string,
  inlineImages: MailInlineImage[] | undefined,
): string {
  const body = sanitizeEmailHtml(html, inlineImages);
  return `<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data: https: http:; style-src 'unsafe-inline'; font-src data: https:; media-src data: https: http:" />
    <base target="_blank" />
    <style>
      :root { color-scheme: light; --text: #182230; --muted: #536173; --line: #dfe4ea; --accent: #1769aa; }
      * { box-sizing: border-box; max-width: 100%; }
      html, body { margin: 0; padding: 0; background: #fff; color: var(--text); }
      body { padding: 2px 0 36px; font: 14px/1.72 Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; overflow-wrap: anywhere; }
      img { max-width: 100% !important; height: auto !important; }
      table { max-width: 100% !important; table-layout: auto; }
      pre, code { white-space: pre-wrap; overflow-wrap: anywhere; }
      a { color: var(--accent); }
      blockquote { margin: 1rem 0; padding-left: 1rem; border-left: 3px solid var(--line); color: var(--muted); }
    </style>
  </head>
  <body>${body}</body>
</html>`;
}
