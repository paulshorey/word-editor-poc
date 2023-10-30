import React from "react";
import { Checkbox, Stack } from "@fluentui/react";
import { contextLoad } from "@src/lib/commandUtils";
/* global Word, console, require */

const ToggleHidden = () => {
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
            control.load(["tag"]);
            context.sync();
            control.color = "#666666";
            if (control.tag === ":") {
              control.appearance = checked ? "Hidden" : "Tags";
            } else {
              control.appearance = checked ? "Hidden" : "Tags";
            }
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
      Toggle All Hidden &nbsp; <Checkbox onChange={handleClick} />
    </Stack>
  );
};

export default ToggleHidden;
