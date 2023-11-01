export default function formatTag(variableName: string, items: any[]) {
  let tag = sanitizeAndUppercaseTag(variableName);
  let number = "";
  let sameVariableName = 0;
  for (let i = 0; i < items.length; i++) {
    let item = items[i];
    if ((item.tag.substring(0, (item.tag as string).indexOf(":")) || item.tag) === tag) {
      sameVariableName++;
    }
  }
  if (sameVariableName) {
    number = (sameVariableName + 1).toString();
  }
  return [tag + (number ? ":" + number : ""), tag, number];
}

function sanitizeAndUppercaseTag(tag: string) {
  tag = tag
    .toUpperCase()
    .replace(/[^A-Z0-9_]/g, "_")
    .replace(/[_]+/g, "_");
  if (tag[0] === "_") {
    tag = tag.slice(1);
  }
  if (tag[tag.length - 1] === "_") {
    tag = tag.slice(0, -1);
  }
  return tag;
}
