import React from "react";
import { DefaultButton, TextField, Stack } from "@fluentui/react";
import dataElementsState, { dataElementsStateType } from "@src/state/dataElements";

/* global Word, require */

const AddDataElement = () => {
  const dataElements = dataElementsState((state) => state as dataElementsStateType);
  const [tag, set_tag] = React.useState("");
  return (
    <Stack horizontal className="faf-fieldgroup">
      <TextField
        className="faf-fieldgroup-input"
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
        className="faf-fieldgroup-button"
        iconProps={{ iconName: "ChevronRight" }}
        onClick={() => {
          dataElements.insertTag(tag);
        }}
      >
        Add
      </DefaultButton>
    </Stack>
  );
};

export default AddDataElement;
