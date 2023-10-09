/* global console */
import React from "react";
import { TextField, Stack, IconButton } from "@fluentui/react";
import * as wordDocument from "@src/state/wordDocument";
import { dataElement } from "@src/state/conditionalComponentsState";
import componentsState, { componentsStateType } from "@src/state/componentsState";

type Props = {
  control: dataElement;
};

const ComponentFieldset = ({ control }: Props) => {
  const components: componentsStateType = componentsState((state) => state as componentsStateType);
  const [tag, setTag] = React.useState(control.title);
  return (
    <Stack horizontal style={{ width: "100%", justifyContent: "space-between" }}>
      <Stack horizontal wrap>
        <TextField
          style={{
            flexGrow: "1",
          }}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              // conditionalComponents.renameId(control.id, tag);
            }
            if (e.key.length > 1) return;
          }}
          onChange={(_e, value) => {
            setTag(value);
          }}
          value={tag}
        />
        <IconButton
          iconProps={{ iconName: "Accept" }}
          title="Emoji"
          ariaLabel="Emoji"
          onClick={() => {
            // conditionalComponents.renameId(control.id, tag);
          }}
        />
      </Stack>
      <IconButton
        iconProps={{ iconName: "BullseyeTarget" }}
        title="Emoji"
        ariaLabel="Emoji"
        onClick={() => {
          wordDocument.scrollToId(control.id);
        }}
      />
      <IconButton
        iconProps={{ iconName: "ChromeClose" }}
        title="Emoji"
        ariaLabel="Emoji"
        onClick={() => {
          components.deleteId(control.id);
        }}
      />
    </Stack>
  );
};

export default ComponentFieldset;
