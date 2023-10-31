import React from "react";
import Components from "@src/Components";
import CustomComponent from "@src/Components/AddCustom";

/* global window, document, Office, Word, require */

export interface Props {
  title: string;
  isOfficeInitialized: boolean;
}

export default function Taskpane({ title, isOfficeInitialized }: Props) {
  if (!isOfficeInitialized) {
    return (
      <div>
        <h3>{title}</h3>
        <p>Please sideload your addin to see app body.</p>
      </div>
    );
  }

  return (
    <div className="faf-taskpane">
      <hr />
      <Components />
      <hr />
      <CustomComponent />
    </div>
  );
}
