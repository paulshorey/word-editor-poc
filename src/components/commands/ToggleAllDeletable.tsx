import React from "react";
import { Checkbox, Stack } from "@fluentui/react";
import { contextLoad } from "@src/lib/commandUtils";
/* global Word, console, require */

const ToggleDeletable = () => {
  const handleClick = (element: any) => {
    const checked = element.target.checked;
    return Word.run(async (context) => {
      const controls = context.document.contentControls;
      contextLoad(context, controls);
      context
        .sync()
        .then(() => {
          controls.items.forEach((control) => {
            context.load(control);
            context.sync();
            control.color = checked ? "red" : "#666666";
            control.cannotDelete = checked;
            control.cannotEdit = !checked;
          });
        })
        .catch((e) => {
          console.log("===>", e);
        });

      return context.sync();
    });
  };

  return (
    <Stack horizontal className={`faf-button`}>
      Toggle All Deletable &nbsp; <Checkbox onChange={handleClick} />
    </Stack>
  );
};

export default ToggleDeletable;
