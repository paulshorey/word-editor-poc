import React from "react";
import CCFieldset from "@src/controls/Conditionals/CCFieldset";
import AddConditionalComponent from "./AddConditionalComponent";
import { Stack } from "@fluentui/react";
import conditionalComponentsState, { conditionalComponentsStateType } from "@src/state/conditionals";
import * as wordDocument from "@src/state/wordDocument";

const ViewConditionalComponents = () => {
  const conditionalComponents: conditionalComponentsStateType = conditionalComponentsState(
    (state) => state as conditionalComponentsStateType
  );
  const [selectedTag, setSelectedTag] = React.useState("");
  wordDocument.state.subscribe((state) => {
    setSelectedTag(state.selectedTag);
  });

  return (
    <div style={{ margin: "0 5px 10px" }}>
      <Stack
        horizontal
        style={{ justifyContent: "space-between", alignItems: "center", margin: "0 0 10px", padding: "0" }}
      >
        <h3 style={{ margin: "0", padding: "0" }}>Conditional Components:</h3>
        <button onClick={conditionalComponents?.loadAll}>reload</button>
      </Stack>
      <AddConditionalComponent />
      {conditionalComponents.items.map((control) => (
        <CCFieldset key={control.id} control={control} selectedTag={selectedTag} />
      )) || (
        <div>
          <code>None</code>
        </div>
      )}
    </div>
  );
};

export default ViewConditionalComponents;
