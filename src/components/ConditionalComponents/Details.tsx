import React from "react";
import { Popup } from "@fluentui/react";
import conditionalComponentsState, {
  conditionalComponentsStateType,
  dataElement,
} from "@src/state/conditionalComponentsState";

type Props = {
  control: dataElement;
};
const Details = ({ control: { id, tag, title } }: Props) => {
  // eslint-disable-next-line no-undef
  console.log("===>", { id, tag, title });
  const conditionalComponents: conditionalComponentsStateType = conditionalComponentsState(
    (state) => state as conditionalComponentsStateType
  );
  const myObj = conditionalComponents.getItemById(id);
  // eslint-disable-next-line no-undef
  console.log("===> KBA");
  // eslint-disable-next-line no-debugger
  debugger;
  return (
    <div
      style={{
        border: "1px dotted black",
        padding: "0 8px",
        margin: "4px 0 8px",
        backgroundColor: "whitesmoke",
        width: "100%",
      }}
    >
      <Popup>
        <h2>{title}</h2>
        <p>Details...TBD</p>
        {myObj && <p>Look: {JSON.stringify(myObj)}</p>}
      </Popup>
    </div>
  );
};

export default Details;
