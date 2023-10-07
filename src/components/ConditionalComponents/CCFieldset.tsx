/* global console */
import React from "react";
import { TextField, Stack, IconButton } from "@fluentui/react";
import wordDocumentState, { wordDocumentStateType } from "@src/state/wordDocument";
import conditionalComponentsState, {
  conditionalComponentsStateType,
  dataElement,
} from "@src/state/conditionalComponentsState";

type Props = {
  control: dataElement;
};

const CCFieldset = ({ control }: Props) => {
  const conditionalComponents: conditionalComponentsStateType = conditionalComponentsState(
    (state) => state as conditionalComponentsStateType
  );
  const wordDocument = wordDocumentState((state) => state as wordDocumentStateType);
  const [tag, setTag] = React.useState(control.tag);
  return (
    <Stack horizontal style={{ width: "100%", justifyContent: "space-between" }}>
      <Stack horizontal wrap>
        <TextField
          style={{
            flexGrow: "1",
          }}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              // conditionalComponents.renameId(control.id, tag);
            }
            if (e.key.length > 1) return;
          }}
          onChange={(_e, value) => {
            setTag(value);
          }}
          value={tag}
        />
        <IconButton
          iconProps={{ iconName: "Accept" }}
          title="Emoji"
          ariaLabel="Emoji"
          onClick={() => {
            // conditionalComponents.renameId(control.id, tag);
          }}
        />
      </Stack>
      <IconButton
        iconProps={{ iconName: "BullseyeTarget" }}
        title="Emoji"
        ariaLabel="Emoji"
        onClick={() => {
          wordDocument.scrollToId(control.id);
        }}
      />
      <IconButton
        iconProps={{ iconName: "ChromeClose" }}
        title="Emoji"
        ariaLabel="Emoji"
        onClick={() => {
          conditionalComponents.deleteId(control.id);
        }}
      />
    </Stack>
  );
};

export default CCFieldset;
