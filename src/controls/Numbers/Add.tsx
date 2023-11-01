import React, { useEffect } from "react";
import { DefaultButton, TextField, Stack } from "@fluentui/react";
import controlsState, { controlsStateType } from "@src/state/controls";

/* global setTimeout, Word, require */

const AddDataElement = () => {
  const controls = controlsState((state) => state as controlsStateType);
  const [number, set_number] = React.useState("");
  const [loading, set_loading] = React.useState(false);
  const handleInsertNumber = (value: string) => {
    // controls.setLabel("Number");
    controls.insertTag("NUMBER", "Number", value);
    set_number("");
    set_loading(true);
    setTimeout(() => {
      set_loading(false);
    }, 500);
  };
  useEffect(() => {
    controls.loadAll();
  }, []);
  return (
    <Stack horizontal style={{ marginBottom: "6px" }}>
      <TextField
        type="number"
        value={number}
        disabled={loading}
        className="faf-fieldgroup-input"
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            handleInsertNumber(number);
          }
          if (e.key.length > 1) return;
        }}
        onChange={(_e, value) => {
          set_number(value);
        }}
        placeholder={loading ? "Loading..." : "Enter number"}
      />
      <DefaultButton
        disabled={loading}
        className="faf-fieldgroup-button"
        iconProps={{ iconName: "ChevronRight" }}
        onClick={() => {
          handleInsertNumber(number);
        }}
      >
        Add
      </DefaultButton>
    </Stack>
  );
};

export default AddDataElement;
