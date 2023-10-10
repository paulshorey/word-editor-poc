/* global console */
import React from "react";
import { TextField, Stack, IconButton } from "@fluentui/react";
import dataElementsState, { dataElementsStateType, dataElement } from "@src/state/dataElementsState";
import * as wordDocument from "@src/state/wordDocument";

type Props = {
  control: dataElement;
  selectedTag: string;
};

const Fieldset = ({ control, selectedTag }: Props) => {
  const dataElements = dataElementsState((state) => state as dataElementsStateType);
  const [tag, set_tag] = React.useState(control.tag);
  let selectedStyles = {};
  if (control.tag === selectedTag) {
    selectedStyles = {
      background: "#90b4d1",
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
      <TextField
        onFocus={() => {
          wordDocument.scrollToId(control.id);
        }}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            dataElements.renameId(control.id, tag);
          }
          if (e.key.length > 1) return;
        }}
        onChange={(_e, value) => {
          set_tag(value);
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
      {/* <IconButton
        iconProps={{ iconName: "BullseyeTarget" }}
        title="Emoji"
        ariaLabel="Emoji"
        onClick={() => {
          wordDocument.scrollToId(control.id);
        }}
      /> */}
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

export default Fieldset;
