import React from "react";
import AddComponent from "@src/components/commands/AddComponent";
import GetFirstParagraph from "@src/components/commands/GetFirstParagraph";
import AddContentControl from "@src/components/commands/AddContentControl";
import ToggleCCDeletable from "@src/components/commands/ToggleCCDeletable";
import AddDataElement from "./dataElements/AddDataElement";
import ViewVariables from "./dataElements/ViewDataElements";
import PrepareCC4Save from "@src/components/commands/PrepareCC4Save";
import Scroll2LastComponent from "@src/components/commands/Scroll2LastComponent";
import dataElementsState, { dataElementsStateType } from "@src/state/dataElements";

/* global Word, require */

export interface Props {
  title: string;
  isOfficeInitialized: boolean;
}

export default function Taskpane({ title, isOfficeInitialized }: Props) {
  const dataElements = dataElementsState((state) => state as dataElementsStateType);
  React.useEffect(() => {
    dataElements.getAllFromDocument();
  }, [isOfficeInitialized]);

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
      <AddDataElement />
      <ViewVariables />
      <AddComponent />
      <GetFirstParagraph />
      <AddContentControl />
      <ToggleCCDeletable />
      <PrepareCC4Save />
      <Scroll2LastComponent />
    </div>
  );
}
