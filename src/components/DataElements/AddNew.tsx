import React from "react";
import { DefaultButton, TextField } from "@fluentui/react";
import dataElementsState, { dataElementsStateType } from "@src/state/dataElements";

/* global Word, require */

const AddDataElement = () => {
  const dataElements = dataElementsState((state) => state as dataElementsStateType);
  const [tag, set_tag] = React.useState("");
  return (
    <div className="faf-fieldset" style={{ margin: "10px" }}>
      <TextField
        style={{
          width: "100%",
          minWidth: "200px",
          flexGrow: "1",
        }}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            dataElements.insertTag(tag);
          }
          if (e.key.length > 1) return;
        }}
        onChange={(_e, value) => {
          set_tag(value);
        }}
        placeholder="ELEMENT_NAME"
      />
      <DefaultButton
        iconProps={{ iconName: "ChevronRight" }}
        onClick={() => {
          dataElements.insertTag(tag);
        }}
        style={{
          width: "40px",
          flexGrow: "0",
          background: "none",
        }}
      >
        Add
      </DefaultButton>
    </div>
  );
};

export default AddDataElement;
