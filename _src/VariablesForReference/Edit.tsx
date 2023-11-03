/* global console */
import React, { useEffect } from "react";
import { TextField, Stack, IconButton } from "@fluentui/react";
import dataElementsState, { dataElementsStateType, dataElement } from "./state";

type Props = {
  control: dataElement;
  isSelected: boolean;
};

const Edit = ({ control, isSelected: isSelected }: Props) => {
  const dataElements = dataElementsState((state) => state as dataElementsStateType);
  const [tag, set_tag] = React.useState(cleanTag(control.tag));
  let selectedStyles = {};
  if (isSelected) {
    selectedStyles = {
      background: "#90b4d1",
      borderRadius: "5px",
    };
  }
  useEffect(() => {
    if (control.tag !== tag) {
      set_tag(cleanTag(control.tag));
    }
  }, [control]);
  return (
    <Stack
      horizontal
      style={{
        ...selectedStyles,
        width: "100%",
        justifyContent: "space-between",
      }}
    >
      <TextField
        onFocus={() => {
          dataElements.selectId(control.id);
        }}
        onBlur={() => {
          dataElements.renameId(control.id, cleanTag(tag));
        }}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            dataElements.renameId(control.id, cleanTag(tag));
          }
          if (e.key.length > 1) return;
        }}
        onChange={(_e, value) => {
          set_tag(cleanTag(value));
        }}
        value={tag}
      />
      <IconButton
        iconProps={{ iconName: "Accept" }}
        title="Emoji"
        ariaLabel="Emoji"
        onClick={() => {
          dataElements.renameId(control.id, tag);
        }}
      />
      <IconButton
        iconProps={{ iconName: "ChromeClose" }}
        title="Emoji"
        ariaLabel="Emoji"
        onClick={() => {
          dataElements.deleteId(control.id);
        }}
      />
    </Stack>
  );
};

export default Edit;

function cleanTag(tag: string): string {
  return tag.substring(0, tag.indexOf(":")) || tag;
}
