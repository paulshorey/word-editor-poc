import React from "react";
import { DefaultButton } from "@fluentui/react";

import { contextLoad } from "@src/lib/commandUtils";
/* global Word, require */

const handleClick = () => {
  return Word.run(async (context) => {
    const controls = context.document.contentControls;
    contextLoad(context, controls);
    context
      .sync()
      .then(() => {
        controls.items.forEach((control) => {
          context.load(control);
          context.sync();
          control.color = "purple";
        });
      })
      .catch((e) => {
        // eslint-disable-next-line no-undef
        console.log("===>", e);
      });

    return context.sync();
  });
};

const PrepareCC4Save = () => {
  return (
    <DefaultButton className={"ms-welcome__action"} iconProps={{ iconName: "ChevronRight" }} onClick={handleClick}>
      Toggle CC Deletable
    </DefaultButton>
  );
};

export default PrepareCC4Save;
