/* global console */
import React from "react";
import { TextField, Stack, IconButton, DefaultButton } from "@fluentui/react";
import dataElementsState, { dataElementsStateType, dataElement } from "@src/state/dataElements";

type Props = {
  control: dataElement;
};

const Fieldset = ({ control }: Props) => {
  const dataElements = dataElementsState((state) => state as dataElementsStateType);
  const [tag, set_tag] = React.useState(control.tag);

  return (
    <Stack horizontal style={{ width: "100%", justifyContent: "space-between" }}>
      <Stack horizontal wrap>
        <TextField
          style={{
            flexGrow: "1",
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
      </Stack>
      <IconButton
        iconProps={{ iconName: "BullseyeTarget" }}
        title="Emoji"
        ariaLabel="Emoji"
        onClick={() => {
          dataElements.selectId(control.id);
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

export default Fieldset;
