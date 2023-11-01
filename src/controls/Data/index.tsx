/* global console */
import React from "react";
import controlsState, { controlsStateType } from "@src/state/controls";
import Item from "./Edit";
import AddNew from "./Add";
import { Stack } from "@fluentui/react";

const ViewDataElements = () => {
  const controls = controlsState((state) => state as controlsStateType);
  return (
    <div style={{ margin: "0 0 10px 0" }}>
      <Stack
        horizontal
        style={{ justifyContent: "space-between", alignItems: "center", margin: "0 0 12px", padding: "0" }}
      >
        <h3 style={{ margin: "0", padding: "0" }}>Data Elements:</h3>
        <button onClick={controls.loadAll}>sync</button>
      </Stack>
      <AddNew />
      {controls.items
        ?.filter((control) => control.title === "DATA")
        .map((control) => (
          <Item key={control.id} control={control} isSelected={control.id === controls.selectedId} />
        ))}
      {/* <code>{JSON.stringify(controls.items)}</code> */}
    </div>
  );
};

export default ViewDataElements;
