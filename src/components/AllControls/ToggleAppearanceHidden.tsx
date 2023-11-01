import React from "react";
import { Checkbox, Stack } from "@fluentui/react";
import resetControl from "@src/functions/resetControl";
/* global Word, console, require */

const ToggleHidden = () => {
  const handleClick = (element: any) => {
    const checked = element.target.checked;
    return Word.run(async (context) => {
      const controls = context.document.contentControls;
      context.load(controls);
      context
        .sync()
        .then(() => {
          controls.items.forEach((control) => {
            context.load(control);
            control.load(["tag"]);
            context.sync();
            resetControl(control, checked ? "Hidden" : "");
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
      Appearance Hidden &nbsp; <Checkbox onChange={handleClick} />
    </Stack>
  );
};

export default ToggleHidden;
