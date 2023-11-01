import React, { useEffect } from "react";
import { DefaultButton, TextField, Stack } from "@fluentui/react";
import controlsState, { controlsStateType } from "@src/state/controls";

/* global setTimeout, Word, require */

const AddDataElement = () => {
  const controls = controlsState((state) => state as controlsStateType);
  // const [text, set_text] = React.useState("");
  // const [tag, set_tag] = React.useState("");
  const [loading, set_loading] = React.useState(false);

  const handleInsertData = (tag, text) => {
    console.log(["handleInsertData", tag, text]);
    controls.insertTag("DATA", tag, text);
    // set_text("");
    set_loading(true);
    setTimeout(() => {
      set_loading(false);
    }, 500);
  };

  useEffect(() => {
    controls.loadAll();
  }, []);

  const onChange: any = (e: any) => {
    console.log(["onChange e.target.value", e.target.value]);
    let arr = e.target.value.split("::");
    let tag = arr[0];
    let text = arr[1];
    console.log(["onChange arr", arr, tag, text]);
    // set_text(text);
    // set_tag(tag);
    handleInsertData(tag, text);
  };
  return (
    <Stack horizontal style={{ marginBottom: "6px" }}>
      <select
        disabled={loading}
        className="faf-fieldgroup-input"
        // onKeyDown={(e) => {
        //   if (e.key === "Enter") {
        //     handleInsertData();
        //   }
        //   if (e.key.length > 1) return;
        // }}
        onChange={onChange}
        // placeholder={loading ? "Loading..." : "Choose data field"}
        style={{ height: "36px", padding: "6px", color: "#666" }}
      >
        <option value={""}>Add a data element</option>
        <option value={"key1::value1"}>key1</option>
        <option value={"key2::value2"}>key2</option>
        <option value={"key3::value3"}>key3</option>
        <option value={"key4::value4"}>key4</option>
        <option value={"key5::value5"}>key5</option>
        <option value={"key6::value6"}>key6</option>
      </select>
      {/* <DefaultButton
        disabled={loading}
        className="faf-fieldgroup-button"
        iconProps={{ iconName: "ChevronRight" }}
        onClick={() => {
          handleInsertData();
        }}
      >
        Add
      </DefaultButton> */}
    </Stack>
  );
};

export default AddDataElement;
