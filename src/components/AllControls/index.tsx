import React from "react";
import GetFirstParagraph from "@src/components/AllControls/GetFirstParagraph";
import AddContentControl from "@src/components/AllControls/AddContentControl";
import ToggleAllDeletable from "@src/components/AllControls/ToggleDeletable";
import ToggleAllHidden from "@src/components/AllControls/ToggleAppearanceHidden";
import ToggleAllTags from "@src/components/AllControls/ToggleAppearanceTags";
import ToggleAllBox from "@src/components/AllControls/ToggleAppearanceBox";
import ToggleAllEditable from "@src/components/AllControls/ToggleEditable";
import { Stack } from "@fluentui/react";

/* global window, document, Office, Word, require */

export interface Props {
  title: string;
  isOfficeInitialized: boolean;
}

export default function AllControls() {
  return (
    <div>
      <Stack
        horizontal
        style={{ justifyContent: "space-between", alignItems: "center", margin: "0 0 10px", padding: "0" }}
      >
        <h3 style={{ margin: "0", padding: "0" }}>All controls:</h3>
      </Stack>
      <ToggleAllDeletable />
      <ToggleAllEditable />
      <ToggleAllHidden />
      <ToggleAllTags />
      <ToggleAllBox />
      <p>&nbsp;</p>
      <hr />
      {/* <hr />
      <GetFirstParagraph />
      <AddContentControl />
      <PrepareCC4Save />
      <Scroll2LastComponent /> */}
    </div>
  );
}
