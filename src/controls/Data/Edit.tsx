/* global console, setTimeout */
import React, { useEffect } from "react";
import { TextField, Stack, IconButton } from "@fluentui/react";
import controlsState, { controlsStateType, control } from "@src/state/dataElements";

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
        style={{ color: isSelected ? "rgb(0, 120, 212)" : "#4aaaff", fontWeight: "bold", margin: "6px 0 6px 3px" }}
      >
        {control.tag}
      </a>
      <a
        href={`#${control.tag}`}
        onClick={() => {
          controls.selectId(control.id);
        }}
        style={{ color: "inherit", margin: "6px 3px 6px 0" }}
      >
        {text}
      </a>
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
