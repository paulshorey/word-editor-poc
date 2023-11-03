import React from "react";
import { DefaultButton } from "@fluentui/react";
/* global Word, require */

const handleClick = () => {
  return Word.run(async (context) => {
    const control = context.document.contentControls.getFirstOrNullObject();
    await context.sync();
    if (control.isNullObject) return;
    control.color = "purple";
    control.cannotEdit = false;
    control.clear();
    await context.sync();
    control.cannotEdit = true;
    control.cannotDelete = false;
    control.load("delete");
    await context.sync();
    control.delete(false);
    return context.sync();
  });
};

const DeleteFirstComponent = () => {
  return (
    <DefaultButton className={"faf-button"} iconProps={{ iconName: "ChevronRight" }} onClick={handleClick}>
      Delete First Component
    </DefaultButton>
  );
};

export default DeleteFirstComponent;
