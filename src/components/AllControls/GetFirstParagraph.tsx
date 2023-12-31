import React from "react";
import { DefaultButton } from "@fluentui/react";

/* global Word, require */

const handleClick = () => {
  return Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst(); // .getNext();
    firstParagraph.font.set({
      name: "Courier New",
      bold: true,
      size: 18,
      color: "red",
    });
    await context.sync();

    const doc = context.document;
    const originalRange = doc.getSelection();

    originalRange.load("text");
    context.load(firstParagraph);
    await context.sync();

    // eslint-disable-next-line office-addins/load-object-before-read
    doc.body.insertParagraph(" " + firstParagraph.text, Word.InsertLocation.end);
    await context.sync();
  });
};

const GetFirstParagraph = () => {
  return (
    <DefaultButton className="faf-button" iconProps={{ iconName: "ChevronRight" }} onClick={handleClick}>
      Get First Para
    </DefaultButton>
  );
};

export default GetFirstParagraph;
