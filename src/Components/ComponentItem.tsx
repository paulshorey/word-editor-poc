/* global console */
import React from "react";
import { Stack, IconButton } from "@fluentui/react";
import componentsState, { componentType, componentsStateType } from "@src/Components/state";

type Props = {
  control: componentType;
};

const ComponentItem = ({ control }: Props) => {
  const components: componentsStateType = componentsState((state) => state as componentsStateType);
  let selectedStyles = {};

  return (
    <Stack horizontal style={{ ...selectedStyles, width: "100%", justifyContent: "space-between" }}>
      <b>{control.title}</b>
      <IconButton
        iconProps={{ iconName: "ChromeClose" }}
        title="Emoji"
        ariaLabel="Emoji"
        onClick={() => {
          components.delete(control.id);
        }}
      />
    </Stack>
  );
};

export default ComponentItem;
