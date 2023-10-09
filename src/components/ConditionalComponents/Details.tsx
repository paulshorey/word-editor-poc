import React from "react";
import { Popup } from "@fluentui/react";
import conditionalComponentsState, {
  conditionalComponentsStateType,
  dataElement,
  outputOption,
} from "@src/state/conditionalComponentsState";

type DetailLinesType = {
  lines: outputOption[];
};
const DetailLines = ({ lines }: DetailLinesType) => {
  if (lines.length === 0) {
    return null;
  }
  return (
    <div>
      {lines.map((line) => (
        <div
          key={line.id}
          style={{
            display: "flex",
            flexDirection: "column",
            backgroundColor: "bisque",
            padding: "4px",
            marginBottom: "8px",
          }}
        >
          <div style={{ display: "flex" }}>
            <div style={{ width: "80px" }}>Scenario:</div>
            <div>{line.title}</div>
          </div>
          <div style={{ display: "flex" }}>
            <div style={{ width: "80px" }}>Rule:</div>
            <div>TRUE</div>
          </div>
        </div>
      ))}
    </div>
  );
};

type DetailProps = {
  control: dataElement;
};

const Details = ({ control: { id, title } }: DetailProps) => {
  const conditionalComponents: conditionalComponentsStateType = conditionalComponentsState(
    (state) => state as conditionalComponentsStateType
  );
  const myObj = conditionalComponents.getItemById(id);

  return (
    <div
      style={{
        border: "1px dotted black",
        padding: "0 8px",
        margin: "4px 0 8px",
        backgroundColor: "whitesmoke",
        width: "100%",
      }}
    >
      <Popup>
        <h2>{title}</h2>

        {myObj?.outputOptions?.length > 0 && <DetailLines lines={myObj?.outputOptions} />}
      </Popup>
    </div>
  );
};

export default Details;
