import React, { useEffect, useState } from "react";
import { DefaultButton, Popup, TextField } from "@fluentui/react";
import conditionalComponentsState, {
  conditionalComponentsStateType,
  dataElement,
  outputOption,
} from "@src/state/conditionals";

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
  const [newRule, setNewRule] = useState("FALSE");
  const [newScenarioName, setNewScenarioName] = useState("");
  const conditionalComponents: conditionalComponentsStateType = conditionalComponentsState(
    (state) => state as conditionalComponentsStateType
  );
  const currentConditionalComponent = conditionalComponents.getItemById(id);
  const SCENARIO_PLACEHOLDER = `Scenario_${currentConditionalComponent?.outputOptions?.length}`;

  useEffect(() => {
    setNewScenarioName(SCENARIO_PLACEHOLDER);
  }, [currentConditionalComponent?.outputOptions?.length]);

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

        {currentConditionalComponent?.outputOptions?.length > 0 && (
          <DetailLines lines={currentConditionalComponent?.outputOptions} />
        )}

        <hr />

        <div style={{ display: "flex", gap: "4px", margin: "8px 0", flexDirection: "column" }}>
          <TextField
            value={newScenarioName}
            className="faf-fieldgroup-input"
            onChange={(_e, value) => {
              setNewScenarioName(value);
            }}
            placeholder={SCENARIO_PLACEHOLDER}
          ></TextField>

          <TextField
            value={newRule}
            className="faf-fieldgroup-input"
            onKeyDown={(e) => {
              if (e.key.length > 1) return;
            }}
            onChange={(_e, value) => {
              setNewRule(value);
            }}
            placeholder="RULE"
          />

          <div style={{ display: "flex", justifyContent: "space-around", margin: "8px 0" }}>
            <DefaultButton
              className="faf-fieldgroup-button"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={() => {
                conditionalComponents.insertScenario(id, newScenarioName, newRule);
              }}
            >
              Add
            </DefaultButton>
          </div>
        </div>
      </Popup>
    </div>
  );
};

export default Details;
