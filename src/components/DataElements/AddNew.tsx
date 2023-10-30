import React from "react";
import { DefaultButton, TextField, Stack } from "@fluentui/react";
import dataElementsState, { dataElementsStateType } from "@src/state/dataElementsState";

/* global setTimeout, Word, require */

const AddDataElement = () => {
  const dataElements = dataElementsState((state) => state as dataElementsStateType);
  const [tag, set_tag] = React.useState("");
  const [loading, set_loading] = React.useState(false);
  const handleInsertTag = (value: string) => {
    dataElements.insertTag(value);
    set_tag("");
    set_loading(true);
    setTimeout(() => {
      set_loading(false);
    }, 500);
  };
  return (
    <Stack horizontal className="faf-fieldgroup">
      <TextField
        value={tag}
        disabled={loading}
        className="faf-fieldgroup-input"
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            handleInsertTag(tag);
          }
          if (e.key.length > 1) return;
        }}
        onChange={(_e, value) => {
          set_tag(value);
        }}
        placeholder={loading ? "Loading..." : "ELEMENT_NAME"}
      />
      <DefaultButton
        disabled={loading}
        className="faf-fieldgroup-button"
        iconProps={{ iconName: "ChevronRight" }}
        onClick={() => {
          handleInsertTag(tag);
        }}
      >
        Add
      </DefaultButton>
    </Stack>
  );
};

export default AddDataElement;
