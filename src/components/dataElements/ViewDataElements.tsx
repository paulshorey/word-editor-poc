import React from "react";
import { DefaultButton } from "@fluentui/react";
import dataElementsState, { dataElementsStateType } from "@src/state/dataElements";

const ViewDataElements = () => {
  const dataElements = dataElementsState((state) => state as dataElementsStateType);
  const [items, set_items] = React.useState([]);

  React.useEffect(() => {
    const list = [];
    for (let key in dataElements.usedInDocument) {
      list.push({
        tag: dataElements.usedInDocument[key].tag,
        id: Math.random() + "",
      });
    }
    set_items(list);
  }, [dataElements.usedInDocument]);

  return (
    <div>
      {!items.length ? (
        <div className="faf-fieldset" style={{ margin: "20px 0 0 10px" }}>
          <DefaultButton
            className="faf-button"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={async () => {
              const dict = await dataElements.getAllFromDocument();
              const list = [];
              for (let key in dict) {
                list.push({
                  tag: dict[key].tag,
                  id: Math.random() + "",
                });
              }
              set_items(list);
            }}
          >
            Load from document
          </DefaultButton>
        </div>
      ) : (
        <div style={{ margin: "0 0 0 10px" }}>
          <h3 style={{ margin: "0", padding: "0" }}>Data Elements in document:</h3>
          {items.map((control) => (
            <div key={control.id} className="faf-fieldset" style={{ justifyContent: "space-between" }}>
              <span>{control.tag}</span>
              <span>
                <button
                  onClick={() => {
                    dataElements.scrollToByName(control.tag);
                  }}
                >
                  scroll to
                </button>
              </span>
              <span>
                <button
                  onClick={() => {
                    dataElements.deleteByName(control.tag);
                  }}
                >
                  x
                </button>
              </span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default ViewDataElements;
