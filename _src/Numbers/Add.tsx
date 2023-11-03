import React, { useEffect } from "react";
import { DefaultButton, TextField, Stack } from "@fluentui/react";
import controlsState, { controlsStateType } from "@src/state/controls";

/* global setTimeout, Word, require */

const AddDataElement = () => {
  const controls = controlsState("Number")((state) => state as controlsStateType);
  const [tag, set_tag] = React.useState("");
  const [loading, set_loading] = React.useState(false);
  const handleInsertTag = (value: string) => {
    controls.insertTag(value);
    set_tag("");
    set_loading(true);
    setTimeout(() => {
      set_loading(false);
    }, 500);
  };
  useEffect(() => {
    controls.loadAll();
  }, []);
  return (
    <div>
      {/* <TextField
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
        placeholder={loading ? "Loading..." : "Enter text"}
      /> */}
      <DefaultButton
        disabled={loading}
        iconProps={{ iconName: "ChevronRight" }}
        onClick={() => {
          handleInsertTag("NUMBER");
        }}
        style={{ whiteSpace: "nowrap", margin: "0 0 12px" }}
      >
        Insert number variable
      </DefaultButton>
    </div>
  );
};

export default AddDataElement;
