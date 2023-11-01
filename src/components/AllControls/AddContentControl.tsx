import React from "react";
import { DefaultButton } from "@fluentui/react";

/* global Word, require */

const handleClick = (tagName: string) => {
  return Word.run(async (context) => {
    const contentRange = context.document.getSelection();
    const contentControl = contentRange.insertContentControl();
    contentControl.title = "Content Control Title ***";
    contentControl.tag = tagName;
    contentControl.appearance = "Tags";
    contentControl.color = "green";
    contentControl.cannotDelete = false;
    contentControl.cannotEdit = false;

    return context.sync();
  });
};

interface AddContentControlInterface {
  tagName?: string;
}
const AddContentControl = ({ tagName }: AddContentControlInterface) => {
  tagName = tagName || "CC_TAG";
  return (
    <DefaultButton className="faf-button" iconProps={{ iconName: "ChevronRight" }} onClick={() => handleClick(tagName)}>
      Add Content Control
    </DefaultButton>
  );
};

export default AddContentControl;
