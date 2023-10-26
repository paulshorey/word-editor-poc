import React from "react";
import { DefaultButton } from "@fluentui/react";
import { TAGNAMES } from "@src/constants/constants";
import componentsState, { componentsStateType } from "@src/state/componentsState";

/* global Word, require */

const handleClick = (components) => {
  const data = [
    { id: 2083969571, tag: "COMPONENT", title: "__PHRASE_1807828__" },
    { id: 675455105, tag: "COMPONENT", title: "__PHRASE_1806233__" },
    { id: 122621298, tag: "COMPONENT", title: "__PHRASE_1831982__" },
    { id: 489286671, tag: "COMPONENT", title: "__PHRASE_1839685__" },
    { id: 710833076, tag: "COMPONENT", title: "__PHRASE_1863429__" },
    { id: 1852679990, tag: "COMPONENT", title: "__PHRASE_1843516__" },
    { id: 1269873108, tag: "COMPONENT", title: "__PHRASE_1870564__" },
    { id: 475076653, tag: "COMPONENT", title: "__PHRASE_1883381__" },
    { id: 263214859, tag: "COMPONENT", title: "__PHRASE_1865299__" },
    { id: 506761104, tag: "COMPONENT", title: "__PHRASE_1865294__" },
    { id: 1961460814, tag: "COMPONENT", title: "__PHRASE_1883398__" },
    { id: 1815657127, tag: "COMPONENT", title: "__PHRASE_1806238__" },
    { id: 2052523545, tag: "COMPONENT", title: "__PHRASE_1806753__" },
    { id: 1660952709, tag: "COMPONENT", title: "__PHRASE_1804976__" },
    { id: 335636801, tag: "COMPONENT", title: "__PHRASE_1883396__" },
    { id: 519070901, tag: "COMPONENT", title: "__PHRASE_1863429__" },
    { id: 1976928106, tag: "COMPONENT", title: "__PHRASE_1808834__" },
    { id: 1995487196, tag: "COMPONENT", title: "__PHRASE_1808837__" },
    { id: 800090292, tag: "COMPONENT", title: "__PHRASE_1803417__" },
    { id: 1123128691, tag: "COMPONENT", title: "__PHRASE_1759074__" },
    { id: 2010321275, tag: "COMPONENT", title: "__PHRASE_658639__" },
    { id: 1776689281, tag: "COMPONENT", title: "__PHRASE_716206__" },
    { id: 2025484991, tag: "COMPONENT", title: "__PHRASE_1844205__" },
    { id: 1263717135, tag: "COMPONENT", title: "__PHRASE_25019__" },
  ];
  return Word.run(async (context) => {
    let body = context.document.body;
    for (const d of data) {
      const title = d.title.toUpperCase().replace("/_/g", " ");
      let paragraph = body.insertParagraph(title, Word.InsertLocation.end);

      let contentControl = paragraph.insertContentControl();
      contentControl.set({
        appearance: "BoundingBox",
        cannotEdit: false,
        cannotDelete: false,
        tag: TAGNAMES.component, // `COMPONENT#${loadDocument}#${timeStamp}`
        title: d.title.toUpperCase().replace("/_/g", " "),
      });
      // paragraph.insertBreak("Line", "After");
      // await context.sync();
      // contentControl.track(); // This is important. If not specified, click will trigger everything
      // contentControl.onEntered.add(async (event) => {
      //   // eslint-disable-next-line no-undef
      //   console.log("Hi4", event.ids, title);
      // });
    }

    await context.sync().catch((err) => {
      // eslint-disable-next-line no-undef
      console.log("===> ERROR", err);
    });
    // eslint-disable-next-line no-undef
    console.log("===> Here we go...");
    components.loadAll();
  });
};

const GetDon = () => {
  const components: componentsStateType = componentsState((state) => state as componentsStateType);
  return (
    <DefaultButton
      className="faf-button"
      iconProps={{ iconName: "ChevronRight" }}
      onClick={() => handleClick(components)}
    >
      Get Don
    </DefaultButton>
  );
};

export default GetDon;
