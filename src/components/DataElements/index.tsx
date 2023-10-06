import React from "react";
import dataElementsState, { dataElementsStateType } from "@src/state/dataElements";
import Fieldset from "@src/components/DataElements/Fieldset";
import AddNew from "./AddNew";

const ViewDataElements = () => {
  const dataElements = dataElementsState((state) => state as dataElementsStateType);
  return (
    <div style={{ margin: "0 5px 10px" }}>
      <div className="faf-fieldset" style={{ justifyContent: "space-between", margin: "0 0 10px" }}>
        <h3>Data Elements:</h3>
        <button onClick={dataElements.loadAll}>reload</button>
      </div>
      <AddNew />
      {dataElements.items?.map((control) => <Fieldset key={control.id} control={control} />) || (
        <div>
          <code>none</code>
        </div>
      )}
    </div>
  );
};

export default ViewDataElements;
