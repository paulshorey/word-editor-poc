import React from "react";
import { DefaultButton } from "@fluentui/react";

/* global Word, require */

const handleClick = () => {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello New Component", Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "green";

    await context.sync();
  });
};

const AddComponent = () => {
  return (
    <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={handleClick}>
      Add Component
    </DefaultButton>
  );
};

export default AddComponent;
