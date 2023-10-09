import React from "react";
import { DefaultButton, Stack, TextField } from "@fluentui/react";
import { TAGNAMES } from "@src/constants/constants";
import { conditionalComponentsStateType } from "@src/state/conditionalComponentsState";
import conditionalComponentsState from "@src/state/conditionalComponentsState";

/* global Word, require */

interface AddConditionalComponentInterface {
  tagName?: string;
}
const AddConditionalComponent = ({ tagName = TAGNAMES.conditional }: AddConditionalComponentInterface) => {
  const conditionalComponents: conditionalComponentsStateType = conditionalComponentsState(
    (state) => state as conditionalComponentsStateType
  );
  const [title, setTitle] = React.useState("");
  return (
    <Stack horizontal style={{ margin: "10px 0", width: "100%" }}>
      <TextField
        style={{
          width: "100%",
          minWidth: "200px",
          flexGrow: "1",
        }}
        onKeyDown={(e) => {
          // if (e.key === "Enter") {
          //   conditionalComponents.insertTag(tagName, title);
          // }
          if (e.key.length > 1) return;
        }}
        onChange={(_e, value) => {
          setTitle(value);
        }}
        placeholder="Name"
      />

      <DefaultButton
        className="faf-button"
        iconProps={{ iconName: "ChevronRight" }}
        onClick={() => conditionalComponents.insertTag(tagName, title)}
      >
        Add
      </DefaultButton>
    </Stack>
  );
};

export default AddConditionalComponent;
