import React from "react";
import { DefaultButton } from "@fluentui/react";

/* global Word, require */

const handleClick = () => {
  return Word.run(async (context) => {
    const contentRange = context.document.getSelection();
    const contentControl = contentRange.insertContentControl();
    contentControl.title = "Content Control Title ***";
    contentControl.tag = "CC_TAG";
    contentControl.appearance = "Tags";
    contentControl.color = "Red";
    contentControl.cannotDelete = false;

    // contentControl.cannotEdit = true;
    // contentControl.appearance = "BoundingBox";

    return context.sync();
  });
};

const AddContentControl = () => {
  return (
    <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={handleClick}>
      Add Content Control
    </DefaultButton>
  );
};

export default AddContentControl;
