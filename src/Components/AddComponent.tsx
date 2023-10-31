import React, { useState } from "react";
import { DefaultButton, Stack } from "@fluentui/react";
import { ComponentTestData } from "@src/testData";

/* global setTimeout, console, Word, require */

/**
 * Inserts a content control, then inserts formatted content into the new content control,
 * as base64, from a couple pre-defined strings.
 */
const AddComponent = () => {
  const [document, set_document] = useState("NO_PICK");

  return (
    <div style={{ margin: "0 5px 10px" }}>
      <Stack
        horizontal
        style={{ justifyContent: "space-between", alignItems: "center", margin: "0 0 10px", padding: "0" }}
      >
        <h3 style={{ margin: "0", padding: "0" }}>Add predefined content:</h3>
      </Stack>
      <Stack horizontal className="faf-fieldgroup">
        <select
          value={document}
          className="faf-fieldgroup-input"
          style={{ height: "32px" }}
          onChange={(value) => {
            set_document(value.target.value);
          }}
        >
          <option value="NO_PICK">Choose content:</option>
          <option value="comp_simple_word">Simple text</option>
          <option value="comp_with_table">With table</option>
        </select>
        <DefaultButton
          className="faf-fieldgroup-button"
          iconProps={{ iconName: "ChevronRight" }}
          onClick={() => {
            insert(document);
            set_document("");
          }}
        >
          Add
        </DefaultButton>
      </Stack>
    </div>
  );
};

export default AddComponent;

/**
 * Add a component to the template, into the current cursor selection
 */
function insert(documentName: string) {
  console.log("ADD", documentName);
  return Word.run(async (context) => {
    // 0. Get base64 data content

    let base64DataContent = "";
    switch (documentName) {
      case "comp_with_table":
        base64DataContent = ComponentTestData.comp_with_table.data;
        break;

      case "comp_simple_word":
        base64DataContent = ComponentTestData.comp_simple_word.data;
        break;

      default:
        Promise.reject("ERROR - Document does not exist");
        return;
    }
    documentName += "_" + Math.round(Math.random() * 1000);
    console.log("Insert base64DataContent", documentName, base64DataContent);

    // 1. Insert into document

    const contentRange = context.document.getSelection().getRange();
    const contentControl = contentRange.insertContentControl();
    contentControl.set({
      tag: "COMPONENT",
      title: documentName.toUpperCase(),
      appearance: "Hidden",
    });
    contentControl.insertHtml("<div>Adding selected content...</div>", "Start");
    await context.sync();
    contentControl.load("insertFileFromBase64");
    await context.sync();
    contentControl.insertFileFromBase64(base64DataContent, "Replace");
    await context.sync();

    // 2. Debug

    // Attempt to trigger reload of document content by adding HTML before/after
    // This is also useful for better UI, even if the component gets added successfully.

    // insert space after
    console.warn("insertHtml After");
    const rangeAfter = contentControl.getRange("After");
    rangeAfter.load(["insertHtml", "html"]);
    await context.sync();
    rangeAfter.insertHtml("&nbsp;<br />&nbsp;", "Start");
    await context.sync();
    rangeAfter.select("End");
    console.warn("insertHtml After done");
    // insert space before
    console.warn("insertText Before");
    const rangeBefore = contentControl.getRange("Before");
    rangeBefore.load("text");
    await context.sync();
    console.log("insertText Before done", rangeBefore.text);
    // insert html before
    rangeBefore.load("insertHtml Before");
    await context.sync();
    rangeBefore.insertHtml("&nbsp;", "End");
    await context.sync();
    rangeBefore.select("End");
    console.warn("insertHtml Before done");

    // 5000 ms is as long as the user will wait for a component to be inserted
    // Try to force update the document
    setTimeout(() => {
      console.log("context.document.load() after 5000 ms");
      context.document.load();
      return context.sync();
    }, 5000);
  });
}
