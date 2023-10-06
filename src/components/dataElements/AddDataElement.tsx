import React from "react";
import { DefaultButton, TextField } from "@fluentui/react";
import dataElementsState, { dataElementsStateType } from "@src/state/dataElements";

/* global Word, require */

const AddDataElement = () => {
  const dataElements = dataElementsState((state) => state as dataElementsStateType);
  const [varname, set_varname] = React.useState("");
  return (
    <div className="faf-fieldset" style={{ margin: "10px" }}>
      <TextField
        style={{
          width: "100%",
          minWidth: "200px",
          flexGrow: "1",
        }}
        onChange={(_e, value) => {
          set_varname(value);
        }}
        placeholder="DATA_ELEMENT_NAME"
      />
      <DefaultButton
        className="faf-button"
        iconProps={{ iconName: "ChevronRight" }}
        onClick={() => {
          dataElements.insertToDocumentByName(varname);
        }}
        style={{
          width: "40px",
          flexGrow: "0",
        }}
      >
        Add
      </DefaultButton>
    </div>
  );
};

export default AddDataElement;
