import React from "react";
import DataElements from "@src/controls/Data";
import Texts from "@src/controls/Texts";
import Numbers from "@src/controls/Numbers";
import AllControls from "../AllControls";

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
      <DataElements />
      <hr />
      <Texts />
      <hr />
      <Numbers />
      <hr />
      <AllControls />
    </div>
  );
}
