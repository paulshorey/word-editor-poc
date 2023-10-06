import React from "react";
import { DefaultButton } from "@fluentui/react";

import { contextLoad } from "@src/lib/commandUtils";
/* global Word, require */

const handleClick = () => {
  return Word.run(async (context) => {
    const controls = context.document.contentControls;
    await contextLoad(context, controls);
    await context.sync();
    const numberOfContentControls = controls.items.length;
    const control = controls.items[numberOfContentControls - 1];

    control.select("Select");
    return context.sync();
  });
};

const Scroll2LastComponent = () => {
  return (
    <DefaultButton className={"faf-button"} iconProps={{ iconName: "ChevronRight" }} onClick={handleClick}>
      Scroll to Last Component
    </DefaultButton>
  );
};

export default Scroll2LastComponent;
