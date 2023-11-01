/* global console, setTimeout */
import React, { useEffect } from "react";
import { TextField, Stack, IconButton } from "@fluentui/react";
import controlsState, { controlsStateType, control } from "@src/state/controls";

type Props = {
  control: control;
  isSelected: boolean;
};

const Edit = ({ control, isSelected: isSelected }: Props) => {
  const controls = controlsState((state) => state as controlsStateType);
  const [text, set_text] = React.useState("");
  const [loading, set_loading] = React.useState(false);
  const handleEditText = (value: string) => {
    // controls.setLabel("Text");
    // controls.insertTag("Text", value);
    controls.editValue(control.id, value);
    set_text("");
    set_loading(true);
    setTimeout(() => {
      set_loading(false);
    }, 500);
  };
  useEffect(() => {
    set_text(control.value);
  }, [control]);
  let selectedStyles = {};
  if (isSelected) {
    selectedStyles = {
      border: "solid 1px #4aaaff",
      borderRadius: "5px",
    };
  }
  return (
    <Stack
      horizontal
      style={{
        ...selectedStyles,
        width: "100%",
        justifyContent: "space-between",
      }}
    >
      <a
        href={`#${control.tag}`}
        onClick={() => {
          controls.selectId(control.id);
        }}
        style={{ color: isSelected ? "rgb(0, 120, 212)" : "#4aaaff", fontWeight: "bold", margin: "6px 3px" }}
      >
        {control.tag}
      </a>
      <TextField
        value={text}
        disabled={loading}
        className="faf-fieldgroup-input"
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            handleEditText(text);
          }
          if (e.key.length > 1) return;
        }}
        onChange={(_e, value) => {
          set_text(value);
        }}
        placeholder={loading ? "Loading..." : "Enter text"}
      />
      <IconButton
        iconProps={{ iconName: "ChromeClose" }}
        title="Emoji"
        ariaLabel="Emoji"
        onClick={() => {
          controls.deleteId(control.id);
        }}
      />
    </Stack>
  );
};

export default Edit;
