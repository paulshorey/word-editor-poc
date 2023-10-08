import React from "react";
import ComponentFieldset from "@src/components/Components/ComponentFieldset";
import AddComponent from "@src/components/Components/AddComponent";
import { Stack } from "@fluentui/react";
import componentsState, { componentsStateType } from "@src/state/componentsState";

const ViewComponents = () => {
  const components: componentsStateType = componentsState((state) => state as componentsStateType);
  return (
    <div style={{ margin: "0 5px 10px" }}>
      <Stack
        horizontal
        style={{ justifyContent: "space-between", alignItems: "center", margin: "0 0 10px", padding: "0" }}
      >
        <h3 style={{ margin: "0", padding: "0" }}>Components:</h3>
        <button onClick={components?.loadAll}>reload</button>
      </Stack>
      <AddComponent />
      {(components.items.length > 0 &&
        components.items.map((control) => <ComponentFieldset key={control.id} control={control} />)) || (
        <div>
          <code>None</code>
        </div>
      )}
    </div>
  );
};

export default ViewComponents;
