import React from "react";
import { DefaultButton, TextField } from "@fluentui/react";

/* global Word, require */

const addDataElement = (varname) => {
  // convert varname to uppercase, remove all spaces and special characters
  varname = varname
    .toUpperCase()
    .replace(/[^A-Z0-9_]/g, "_")
    .replace(/[_]+/g, "_");
  if (varname[0] === "_") {
    varname = varname.slice(1);
  }
  if (varname[varname.length - 1] === "_") {
    varname = varname.slice(0, -1);
  }
  // insert into document
  return Word.run(async (context) => {
    const contentRange = context.document.getSelection();
    const contentControl = contentRange.insertContentControl();
    contentControl.title = "";
    contentControl.tag = varname;
    contentControl.color = "#666666";
    contentControl.cannotDelete = false;
    contentControl.cannotEdit = false;
    contentControl.appearance = "Tags";
    contentControl.insertText(varname, "Replace");
    contentControl.cannotEdit = true;
    context.sync().then(() => {
      contentControl.cannotEdit = true;
    });
  });
};

const AddComponent = () => {
  const [varname, set_varname] = React.useState("");
  return (
    <div className="faf-fieldset" style={{ margin: "10px" }}>
      <TextField
        style={{
          width: "100%",
          minWidth: "200px",
          flexGrow: "1",
        }}
        onChange={(_e, value) => {
          set_varname(value);
        }}
        placeholder="DATA_ELEMENT_NAME"
      />
      <DefaultButton
        className="faf-button"
        iconProps={{ iconName: "ChevronRight" }}
        onClick={() => {
          addDataElement(varname);
        }}
        style={{
          width: "40px",
          flexGrow: "0",
        }}
      >
        Add
      </DefaultButton>
    </div>
  );
};

export default AddComponent;
