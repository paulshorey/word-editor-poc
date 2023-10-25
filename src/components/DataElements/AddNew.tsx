import React from "react";
import { DefaultButton, TextField, Stack } from "@fluentui/react";
import dataElementsState, { dataElementsStateType } from "@src/state/dataElementsState";

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
            set_tag("");
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
          set_tag("");
        }}
      >
        Add
      </DefaultButton>
    </Stack>
  );
};

export default AddDataElement;
