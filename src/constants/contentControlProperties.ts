/*
 * IMPORTANT! TITLE AND TAG ARE REVERSED IN THE UI:
 * - "id" -- is unique, automatically generated when insert control into document. Unfortunately, when user clicks a control in the document, we can't get the "id" of the clicked control, so we must use the "text" and "tag" to identify which component was clicked
 * - "title" -- is displayed as the label in the document UI, on left/right of cc text in a pill box, or on top as a tooltip. IT MAY BE DUPLICATED, so it's not good UI to show the variable name as the title. Instead, lets use it to show the "type" of variable.
 * - "tag" -- NOT displayed. It is internal. So, we must use this to identify the component, lets use this as the variable name.
 * - "text" -- We can NOT display the variable name as the "title" or "tag". So, lets show the variable name as the "text". This is the only place we can consistently display it.
 * - "tagAndText" -- When user clicks inside the Word document, WE DO NOT KNOW exactly what content control they clicked on. We only know the "text", not the "id". So, because the "text" and "tag" must be the same value, that means it must also be unique. So, lets auto-increment the tag/text value with suffix "_2" "_3" "_4" so each instance is unique.
 */

export const TYPES = {
  numberElement: "NUMBER",
  textElement: "TEXT",
  dataElement: "DATA",
  component: "COMPONENT",
  conditional: "CONDITIONAL",
  scenario: "SCENARIO",
};

export const createNumberElement = function (variableName: string, allDataElements: any[]) {
  return {
    type: TYPES.numberElement,
    name: (function () {
      variableName = sanitizeAndUppercaseTag(variableName) + "_" + (allDataElements.length + 1);
      return variableName;
    })(),
    appearance: "Tags" as const,
    color: "#666666",
  };
};
export const createTextElement = function (variableName: string, allDataElements: any[]) {
  return {
    type: TYPES.textElement,
    name: (function () {
      variableName = sanitizeAndUppercaseTag(variableName) + "_" + (allDataElements.length + 1);
      return variableName;
    })(),
    appearance: "Tags" as const,
    color: "#666666",
  };
};
export const createDataElement = function (variableName: string, allDataElements: any[]) {
  return {
    type: TYPES.dataElement,
    name: (function () {
      variableName = sanitizeAndUppercaseTag(variableName) + "_" + (allDataElements.length + 1);
      return variableName;
    })(),
    appearance: "BoundingBox" as const,
    color: "#666666",
  };
};
export const createComponent = function (variableName: string, allComponents: any[]) {
  return {
    type: TYPES.component,
    name: (function () {
      variableName = sanitizeAndUppercaseTag(variableName) + "_" + (allComponents.length + 1);
      return variableName;
    })(),
    appearance: "BoundingBox" as const,
    color: "#666666",
    htmlAfter: " <br /> ",
  };
};

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

/*
 * Deprecated?
 */
export const TAGNAMES = {
  component: "COMPONENT",
  conditional: "CONDITIONAL",
  scenario: "SCENARIO",
};
