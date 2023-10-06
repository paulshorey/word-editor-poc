import React from "react";
import dataElementsState, { dataElementsStateType } from "@src/state/dataElements";
import Fieldset from "@src/components/DataElements/Fieldset";
import AddNew from "./AddNew";
import { Stack } from "@fluentui/react";
import wordDocumentState from "@src/state/wordDocument";
import { wordDocumentStateType } from "@src/state/wordDocument";

const ViewDataElements = () => {
  const dataElements = dataElementsState((state) => state as dataElementsStateType);
  const wordDocument = wordDocumentState((state) => state as wordDocumentStateType);
  return (
    <div style={{ margin: "0 0 10px 0" }}>
      <Stack
        horizontal
        style={{ justifyContent: "space-between", alignItems: "center", margin: "0 0 10px", padding: "0" }}
      >
        <h3 style={{ margin: "0", padding: "0" }}>Data Elements:</h3>
        <button onClick={dataElements.loadAll}>reload</button>
      </Stack>
      <AddNew />
      {dataElements.items?.map((control) => (
        <Fieldset key={control.id} control={control} />
      ))}
      <Stack horizontal>
        <span>Selected tag:</span>
        <span>{wordDocument.selectedTag}</span>
      </Stack>
    </div>
  );
};

export default ViewDataElements;
