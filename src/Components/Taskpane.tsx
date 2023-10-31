import React from "react";
import AddComponent from "@src/components/AddComponent";
import AddCustom from "@src/components/AddCustom";

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
      <AddComponent />
      <hr />
      <AddCustom />
    </div>
  );
}
