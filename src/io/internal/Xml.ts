export function xmlEscape(value: string): string {
  return value
    .split("&").join("&amp;")
    .split("<").join("&lt;")
    .split(">").join("&gt;")
    .split('"').join("&quot;")
    .split("'").join("&apos;");
}

export function xmlUnescape(value: string): string {
  return value
    .split("&lt;").join("<")
    .split("&gt;").join(">")
    .split("&quot;").join('"')
    .split("&apos;").join("'")
    .split("&amp;").join("&");
}

export function readTagText(xml: string, tagName: string): string | undefined {
  const regex = new RegExp(`<${tagName}[^>]*>([\\s\\S]*?)<\\/${tagName}>`, "i");
  const match = regex.exec(xml);
  return match ? xmlUnescape(match[1]) : undefined;
}

export function getAttribute(tag: string, attributeName: string): string | undefined {
  const regex = new RegExp(`${attributeName}="([^"]*)"`, "i");
  const match = regex.exec(tag);
  return match ? xmlUnescape(match[1]) : undefined;
}
