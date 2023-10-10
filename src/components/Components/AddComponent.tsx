import React, { useState } from "react";
import { DefaultButton, Stack } from "@fluentui/react";
import componentsState, { componentsStateType } from "@src/state/componentsState";

/* global Word, require */

const AddComponent = () => {
  const components: componentsStateType = componentsState((state) => state as componentsStateType);
  const [loadDocument, setLoadDocument] = useState("NO_PICK");

  return (
    <Stack horizontal className="faf-fieldgroup">
      <select
        className="faf-fieldgroup-input"
        style={{ height: "32px" }}
        onChange={(value) => setLoadDocument(value.target.value)}
      >
        <option value="NO_PICK">Pick a Document</option>
        <option value="comp_simple_word">Simple Document</option>
        <option value="comp_with_table">Complex Document</option>
      </select>
      <DefaultButton
        className="faf-fieldgroup-button"
        iconProps={{ iconName: "ChevronRight" }}
        onClick={() => components.insertTag(loadDocument)}
      >
        Add
      </DefaultButton>
    </Stack>
  );
};

export default AddComponent;