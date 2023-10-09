import React from "react";
import Components from "@src/components/Components";
import ConditionalComponents from "@src/components/ConditionalComponents";
import GetFirstParagraph from "@src/components/commands/GetFirstParagraph";
import AddContentControl from "@src/components/commands/AddContentControl";
import ToggleCCDeletable from "@src/components/commands/ToggleCCDeletable";
import DataElements from "./DataElements";
import PrepareCC4Save from "@src/components/commands/PrepareCC4Save";
import Scroll2LastComponent from "@src/components/commands/Scroll2LastComponent";
import dataElementsState, { dataElementsStateType } from "@src/state/dataElements";
import conditionalComponentsState, { conditionalComponentsStateType } from "@src/state/conditionalComponentsState";
import componentsState, { componentsStateType } from "@src/state/componentsState";
// import useSelect from "@src/hooks/useSelect";

/* global document, Office, Word, require */

export interface Props {
  title: string;
  isOfficeInitialized: boolean;
}

export default function Taskpane({ title, isOfficeInitialized }: Props) {
  const components: componentsStateType = componentsState((state) => state as componentsStateType);
  const conditionalComponents = conditionalComponentsState((state) => state as conditionalComponentsStateType);
  const dataElements = dataElementsState((state) => state as dataElementsStateType);

  React.useEffect(() => {
    if (isOfficeInitialized) {
      components.loadAll();
      conditionalComponents.loadAll();
      dataElements.loadAll();
    }
  }, [isOfficeInitialized]);

  // if (isOfficeInitialized) {
  //   Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, useSelect);
  // }

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
      <Components />
      <hr />
      <ConditionalComponents />
      <hr />

      <GetFirstParagraph />
      <AddContentControl />
      <ToggleCCDeletable />
      <PrepareCC4Save />
      <Scroll2LastComponent />
    </div>
  );
}
