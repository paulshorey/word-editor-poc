import React, { useState } from "react";
import { DefaultButton, Stack, TextField } from "@fluentui/react";
import componentsState, { componentsStateType } from "@src/state/componentsState";

/* global console, Word, require */

const AddComponent = () => {
  const components: componentsStateType = componentsState((state) => state as componentsStateType);
  const [documentContent, set_documentContent] = useState("");
  return (
    <Stack className="faf-fieldgroup">
      <textarea defaultValue="" onChange={console.log} placeholder="Insert OOXML or BASE64"></textarea>
      <Stack horizontal style={{ justifyContent: "space-between", margin: "0 15px 0 5px" }}>
        <DefaultButton
          className="faf-fieldgroup-button"
          style={{ whiteSpace: "nowrap", border: "none" }}
          iconProps={{ iconName: "ChevronRight" }}
        >
          Insert XML
        </DefaultButton>
        <DefaultButton
          className="faf-fieldgroup-button"
          style={{ whiteSpace: "nowrap", border: "none" }}
          iconProps={{ iconName: "ChevronRight" }}
        >
          Insert Base64
        </DefaultButton>
      </Stack>
    </Stack>
  );
};

export default AddComponent;

import { TAGNAMES } from "@src/constants/constants";
function insertXML(contentToInsert, isXML = false) {
  const documentName = "COMP_" + Date.now();
  Word.run(async (context) => {
    const contentRange = context.document.getSelection().getRange("Content");
    const contentControl = contentRange.insertContentControl();
    contentControl.tag = TAGNAMES.component; // `COMPONENT#${loadDocument}#${timeStamp}`
    contentControl.title = documentName.toUpperCase();
    contentControl.insertHtml("<div>Loading component content...</div>", "Start");
    contentControl.load("cannotEdit");
    await context.sync();
    contentControl.appearance = "BoundingBox";
    contentControl.cannotEdit = false;
    if (isXML) {
      contentControl.load("insertOoxml");
      await context.sync();
      contentControl.insertOoxml(contentToInsert, "Replace");
    } else {
      contentControl.load("insertFileFromBase64");
      await context.sync();
      contentControl.insertFileFromBase64(contentToInsert, "Replace");
    }
    // await range.context.sync();
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
  });
}
