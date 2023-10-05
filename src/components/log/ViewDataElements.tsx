import React from "react";
import dataElementsState, { dataElementsStateType } from "@src/state/dataElements";

const ViewDataElements = () => {
  const dataElements = dataElementsState((state) => state as dataElementsStateType);
  return (
    <div style={{ margin: "30px 0 0 10px" }}>
      <h3 style={{ margin: "0", padding: "0" }}>RECENTLY-ADDED DATA ELEMENTS:</h3>
      <sup style={{ display: "block", margin: "0", padding: "5px 0 0 0", lineHeight: "1" }}>
        (This does not include data elements that were added previously, before page reload.)
      </sup>
      <pre style={{ margin: "10px" }}>
        <code>{JSON.stringify(dataElements.usedInDocument || null, null, 2)}</code>
      </pre>
    </div>
  );
};

export default ViewDataElements;
