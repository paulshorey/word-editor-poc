import React, { useState } from "react";
import { DefaultButton, TextField } from "@fluentui/react";

/* global Word, require */

const defaultCondition = (context: Word.RequestContext, condition: string | null) => {
  const contentRange = context.document.getSelection().getRange("Whole");
  const contentControl = contentRange.insertContentControl();
  contentControl.set({
    appearance: "Tags",
    cannotEdit: false,
    cannotDelete: false,
    color: "cyan",
    tag: "Default",
    title: condition || "Un-Conditional",
  });
  return context.sync();
};

const handleClick = (tagName: string, condition: string) => {
  return Word.run(async (context) => {
    const contentRange = context.document.getSelection().getRange();
    const contentControl = contentRange.insertContentControl();
    contentControl.set({
      appearance: "Tags",
      cannotEdit: false,
      cannotDelete: false,
      color: "cyan",
      tag: tagName,
      title: "CONDITION",
    });

    context.load(contentControl);
    // eslint-disable-next-line no-undef
    console.log("===>", condition);
    await context.sync();
    defaultCondition(context, condition);
    context.load(contentControl);
    await context.sync();
    contentControl.select("End");
    return context.sync();
  });
};

interface AddConditionalComponentInterface {
  tagName?: string;
}
const AddConditionalComponent = ({ tagName = "CONDITIONAL" }: AddConditionalComponentInterface) => {
  const [value, setValue] = useState<string | null>();

  return (
    <div>
      <h2 style={{ margin: "0", padding: "0" }}>Conditional Components:</h2>

      <div className="faf-fieldset" style={{ margin: "10px" }}>
        <TextField
          style={{
            width: "100%",
            minWidth: "200px",
            flexGrow: "1",
          }}
          onChange={(_e, value) => {
            setValue(value);
          }}
          placeholder="CONDITION"
        />

        <DefaultButton
          className="faf-button"
          iconProps={{ iconName: "ChevronRight" }}
          onClick={() => handleClick(tagName, value)}
        >
          Add
        </DefaultButton>
      </div>
    </div>
  );
};

export default AddConditionalComponent;
