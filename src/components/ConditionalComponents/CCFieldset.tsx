/* global console */
import React, { useEffect, useState } from "react";
import { TextField, Stack, IconButton } from "@fluentui/react";
import wordDocumentState, { wordDocumentStateType } from "@src/state/wordDocument";
import conditionalComponentsState, {
  conditionalComponentsStateType,
  dataElement,
} from "@src/state/conditionalComponentsState";
import Details from "@src/components/ConditionalComponents/Details";

type Props = {
  control: dataElement;
};

const CCFieldset = ({ control }: Props) => {
  const conditionalComponents: conditionalComponentsStateType = conditionalComponentsState(
    (state) => state as conditionalComponentsStateType
  );
  const wordDocument = wordDocumentState((state) => state as wordDocumentStateType);
  const [tag, setTag] = useState(control.title);
  const [rotation, setRotation] = useState("rotate(0deg)");
  const [isOpen, setIsOpen] = useState(false);

  useEffect(() => {
    setRotation(`rotate(${isOpen ? 180 : 0}deg)`);
  }, [isOpen]);

  return (
    <div>
      <Stack horizontal style={{ width: "100%", justifyContent: "space-between", margin: "2px 0" }}>
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
          <div style={{ transform: rotation, transition: ".5s ease-in-out" }}>
            <IconButton
              iconProps={{ iconName: "DrillDown" }}
              title="Emoji"
              ariaLabel="Emoji"
              onClick={() => {
                setIsOpen((prev) => !prev);
                // conditionalComponents.renameId(control.id, tag);
              }}
            />
          </div>
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
      {isOpen && <Details control={control} />}
    </div>
  );
};

export default CCFieldset;
