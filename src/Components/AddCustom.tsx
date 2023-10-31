import React, { useState } from "react";
import { DefaultButton, Stack } from "@fluentui/react";

/* global Office, console, Word, require */

/**
 * This uses the add() helper function below to insert a content control into the document,
 * then insert formatted content into the new content control, as base64, XML, or text.
 */
const AddCustom = () => {
  const [string, set_string] = useState("");
  return (
    <div style={{ margin: "0 5px 10px" }}>
      <Stack
        horizontal
        style={{ justifyContent: "space-between", alignItems: "center", margin: "0 0 10px", padding: "0" }}
      >
        <h3 style={{ margin: "0", padding: "0" }}>Add custom string:</h3>
      </Stack>
      <Stack className="faf-fieldgroup" style={{ margin: "0" }}>
        <textarea
          defaultValue=""
          onChange={(e) => {
            set_string(e.target.value);
          }}
          placeholder="Insert string with formatting"
        ></textarea>
        <Stack horizontal style={{ justifyContent: "space-between", margin: "0 15px 0 5px" }}>
          <DefaultButton
            className="faf-fieldgroup-button"
            style={{ whiteSpace: "nowrap", border: "none" }}
            onClick={async () => {
              await add(string, "xml");
            }}
          >
            XML
          </DefaultButton>
          <DefaultButton
            className="faf-fieldgroup-button"
            style={{ whiteSpace: "nowrap", border: "none" }}
            onClick={() => {
              add(string, "base64");
            }}
          >
            Base64
          </DefaultButton>
          <DefaultButton
            className="faf-fieldgroup-button"
            style={{ whiteSpace: "nowrap", border: "none" }}
            onClick={() => {
              add(string, "data");
            }}
          >
            dataAsync
          </DefaultButton>
        </Stack>
      </Stack>
    </div>
  );
};

export default AddCustom;

function add(contentToInsert, type: "base64" | "xml" | "data" = "base64") {
  return new Promise((resolve, reject) => {
    const documentName = "COMP_" + Date.now();
    Word.run(async (context) => {
      try {
        const selection = context.document.getSelection();
        const contentRange = selection.getRange("Content");
        const contentControl = contentRange.insertContentControl();
        contentControl.tag = "COMPONENT";
        contentControl.title = documentName.toUpperCase();
        contentControl.insertHtml("<div>Adding custom content...</div>", "Start");
        contentControl.load("cannotEdit");
        await context.sync();
        contentControl.appearance = "BoundingBox";
        contentControl.cannotEdit = false;
        if (type === "data") {
          contentControl.select();
          await asyncFunctionWithCallback(Office.context.document.setSelectedDataAsync, [contentToInsert]);
        } else if (type === "xml") {
          contentControl.load("insertOoxml");
          await context.sync();
          contentControl.insertOoxml(contentToInsert, "Replace");
        } else {
          contentControl.load("insertFileFromBase64");
          await context.sync();
          contentControl.insertFileFromBase64(contentToInsert, "Replace");
        }
        await contentControl.context.sync();
        await context.sync();
        // insert line break if there is no text before
        const rangeBefore = contentControl.getRange("Before");
        const textBefore = rangeBefore.getTextRanges([" "], true).load();
        textBefore.load("items");
        await context.sync();
        if (textBefore.items.length === 0) {
          contentControl.insertBreak("Line", "Before");
          await context.sync();
        }
        // insert line break if there is no text after
        const rangeAfter = contentControl.getRange("After");
        const textAfter = rangeAfter.getTextRanges([" "], true).load();
        textAfter.load("items");
        await context.sync();
        if (textAfter.items.length === 0) {
          contentControl.insertBreak("Line", "After");
          await context.sync();
        }
        // sync
        await context.sync();
        context.document.body.load();
        context.document.load();
        resolve(true);
      } catch (error) {
        reject(error);
      }
    });
  });
}

function asyncFunctionWithCallback(func, args) {
  return new Promise((resolve, reject) => {
    func.apply(null, [...args, (err, result) => (err ? reject(err) : resolve(result))]);
  });
}
