import React, { useEffect } from "react";
import { DefaultButton, TextField, Stack } from "@fluentui/react";
import controlsState, { controlsStateType } from "@src/state/controls";

/* global setTimeout, Word, require */

const AddDataElement = () => {
  const controls = controlsState((state) => state as controlsStateType);
  const [text, set_text] = React.useState("");
  const [loading, set_loading] = React.useState(false);
  const handleInsertText = (value: string) => {
    // controls.setLabel("Text");
    controls.insertTag("Text", "Text", value);
    set_text("");
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
        value={text}
        disabled={loading}
        className="faf-fieldgroup-input"
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            handleInsertText(text);
          }
          if (e.key.length > 1) return;
        }}
        onChange={(_e, value) => {
          set_text(value);
        }}
        placeholder={loading ? "Loading..." : "Enter text"}
      />
      <DefaultButton
        disabled={loading}
        className="faf-fieldgroup-button"
        iconProps={{ iconName: "ChevronRight" }}
        onClick={() => {
          handleInsertText(text);
        }}
      >
        Add
      </DefaultButton>
    </Stack>
  );
};

export default AddDataElement;
