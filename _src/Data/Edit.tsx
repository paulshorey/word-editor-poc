/* global console */
import React, { useEffect } from "react";
import { TextField, Stack, IconButton } from "@fluentui/react";
import controlsState, { controlsStateType, control } from "@src/state/controls";

type Props = {
  control: control;
  isSelected: boolean;
};

const Edit = ({ control, isSelected: isSelected }: Props) => {
  const controls = controlsState("Data")((state) => state as controlsStateType);
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
        style={{ color: isSelected ? "rgb(0, 120, 212)" : "#4aaaff", fontWeight: "bold", margin: "7px" }}
      >
        {control.tag}
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

function cleanTag(tag: string): string {
  return tag.substring(0, tag.indexOf(":")) || tag;
}
