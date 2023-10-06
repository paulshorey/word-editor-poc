import React, { useState } from "react";
import { DefaultButton } from "@fluentui/react";
import { contextLoad } from "@src/lib/commandUtils";
import { ComponentTestData } from "@src/testdata/TestData";

/* global Word, require */

const handleClick = (loadDocument: string) => {
  let base64DataContent;
  switch (loadDocument) {
    case "comp_with_table":
      base64DataContent = ComponentTestData.comp_with_table.data;
      break;

    case "comp_simple_word":
      base64DataContent = ComponentTestData.comp_simple_word.data;
      break;

    default:
      return Promise.reject("ERROR - Document does not exist");
  }

  const timeStamp = new Date().getTime();
  return Word.run(async (context) => {
    const contentRange = context.document.getSelection();
    const contentControl = contentRange.insertContentControl();
    contentControl.title = "COMPONENT";
    contentControl.tag = `COMPONENT#${loadDocument}#${timeStamp}`;
    contentControl.color = "purple";
    contentControl.cannotDelete = false;
    contentControl.cannotEdit = false;
    contentControl.appearance = "Tags";

    contextLoad(context, contentControl);
    contentControl.insertFileFromBase64(base64DataContent, "Replace").load();
    return context.sync();
  });
};

const AddComponent = () => {
  const [loadDocument, setLoadDocument] = useState("NO_PICK");

  const fns = () => {
    handleClick(loadDocument);
  };

  return (
    <div className="faf-fieldset" style={{ margin: "10px" }}>
      <select style={{ height: "32px" }} onChange={(value) => setLoadDocument(value.target.value)}>
        <option value="NO_PICK">Pick a Document</option>
        <option value="comp_simple_word">Simple Document</option>
        <option value="comp_with_table">Complex Document</option>
      </select>
      <DefaultButton className="faf-button" iconProps={{ iconName: "ChevronRight" }} onClick={fns}>
        Add Component
      </DefaultButton>
    </div>
  );
};

export default AddComponent;
