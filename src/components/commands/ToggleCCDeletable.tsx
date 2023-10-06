import React, { useEffect, useState } from "react";
import { DefaultButton } from "@fluentui/react";

import { contextLoad } from "@src/lib/commandUtils";
/* global Word, require */

const handleClick = (tagName: string, isDeletable: boolean) => {
  return Word.run(async (context) => {
    const controls = context.document.contentControls.getByTag(tagName);
    contextLoad(context, controls);
    context
      .sync()
      .then(() => {
        controls.items.forEach((control) => {
          context.load(control);
          context.sync();
          control.color = isDeletable ? "red" : "green";
          control.cannotDelete = isDeletable;
          control.cannotEdit = isDeletable;
        });
      })
      .catch((e) => {
        // eslint-disable-next-line no-undef
        console.log("===>", e);
      });

    return context.sync();
  });
};

interface ToggleCCDeletableInterface {
  tagName: string;
}
const ToggleCCDeletable = ({ tagName }: ToggleCCDeletableInterface) => {
  const [deletable, setDeletable] = useState<boolean>(true);
  const [toggleStyle, setToggleStyle] = useState("faf-isDeletable");
  useEffect(() => {
    setToggleStyle(deletable ? "faf-isDeletable" : "faf-isNotDeletable");
  }, [deletable]);

  return (
    <DefaultButton
      className={`faf-button ${toggleStyle}`}
      iconProps={{ iconName: "ChevronRight" }}
      onClick={() => {
        handleClick(tagName, deletable);
        setDeletable((prev) => !prev);
      }}
    >
      Toggle CC Deletable
    </DefaultButton>
  );
};

export default ToggleCCDeletable;
