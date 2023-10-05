import React from "react";
import { DefaultButton } from "@fluentui/react";

import { contextLoad } from "@src/lib/commandUtils";
/* global Word, require */

const handleClick = () => {
  return Word.run(async (context) => {
    const control = context.document.contentControls.getFirst();
    contextLoad(context, control);
    control.color = "purple";
    control.clear();
    return context.sync();
  });
};

const ReplaceCC4Save = () => {
  return (
    <DefaultButton className={"ms-welcome__action"} iconProps={{ iconName: "ChevronRight" }} onClick={handleClick}>
      Clear CCs
    </DefaultButton>
  );
};

export default ReplaceCC4Save;
