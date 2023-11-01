import React from "react";
import GetFirstParagraph from "@src/components/AllControls/GetFirstParagraph";
import AddContentControl from "@src/components/AllControls/AddContentControl";
import ToggleAllDeletable from "@src/components/AllControls/ToggleDeletable";
import ToggleAllHidden from "@src/components/AllControls/ToggleAppearanceHidden";
import ToggleAllTags from "@src/components/AllControls/ToggleAppearanceTags";
import ToggleAllBox from "@src/components/AllControls/ToggleAppearanceBox";
import ToggleAllEditable from "@src/components/AllControls/ToggleEditable";
import PrepareCC4Save from "@src/components/AllControls/DeleteFirstComponent";
import Scroll2LastComponent from "@src/components/AllControls/Scroll2LastComponent";
import { Stack } from "@fluentui/react";
import controlsState from "@src/state/controls";
import { controlsStateType } from "@src/state/controls";

/* global window, document, Office, Word, require */

export interface Props {
  title: string;
  isOfficeInitialized: boolean;
}

export default function AllControls() {
  const controls = controlsState((state) => state as controlsStateType);
  return (
    <div>
      <Stack
        horizontal
        style={{ justifyContent: "space-between", alignItems: "center", margin: "0 0 10px", padding: "0" }}
      >
        <h3 style={{ margin: "0", padding: "0" }}>All controls:</h3>
      </Stack>
      <button onClick={controls.loadAll}>sync</button>
      <hr />
      <ToggleAllDeletable />
      <ToggleAllEditable />
      <ToggleAllHidden />
      <ToggleAllTags />
      <ToggleAllBox />
      {/* <hr />
      <GetFirstParagraph />
      <AddContentControl />
      <PrepareCC4Save />
      <Scroll2LastComponent /> */}
    </div>
  );
}
