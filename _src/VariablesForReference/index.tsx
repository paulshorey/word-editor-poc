import React from "react";
import dataElementsState, { dataElementsStateType } from "./state";
import Item from "./Edit";
import AddNew from "./Add";
import { Stack } from "@fluentui/react";

const ViewDataElements = () => {
  const dataElements = dataElementsState((state) => state as dataElementsStateType);
  return (
    <div style={{ margin: "0 0 10px 0" }}>
      <Stack
        horizontal
        style={{ justifyContent: "space-between", alignItems: "center", margin: "0 0 10px", padding: "0" }}
      >
        <h3 style={{ margin: "0", padding: "0" }}>Data Elements:</h3>
        <button onClick={dataElements.loadAll}>sync</button>
      </Stack>
      <AddNew />
      {dataElements.items?.map((control) => (
        <Item key={control.id} control={control} isSelected={control.id === dataElements.selectedId} />
      ))}
      {/* <code>{JSON.stringify(dataElements.items)}</code> */}
    </div>
  );
};

export default ViewDataElements;
