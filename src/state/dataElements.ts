import create from "zustand";

/* global Word, require */

export type dataElement = {
  name: string;
  otherData: any;
};

export type dataElementsStateType = {
  searchResults: dataElement[];
  usedInDocument: Record<string, dataElement>;
  insertToDocument: (element: dataElement) => dataElement | undefined;
  insertToDocumentByName: (elementName: string) => dataElement | undefined;
};

const dataElementsState = create((set, get) => ({
  /**
   * Temporary, to hold the results from an API call. For now it is just dummy data.
   */
  searchResults: [
    {
      name: "TEST_1",
      otherData: "idk",
    },
    {
      name: "TEST_2",
      otherData: "idk",
    },
    {
      name: "TEST_3",
      otherData: "idk",
    },
    {
      name: "TEST_4",
      otherData: "idk",
    },
    {
      name: "TEST_5",
      otherData: "idk",
    },
  ],
  /**
   * All (unique) elements used in the document. We still need to decide how to manage the position of each variable in the document. This is too basic.
   */
  inDocument: {},
  /**
   * Shortcut to add by a name (string) -- not finished -- need to think about how to actually manage state and insert into document.
   */
  insertToDocumentByName: function (elementName: string) {
    // convert varname to uppercase, remove all spaces and special characters
    elementName = elementName
      .toUpperCase()
      .replace(/[^A-Z0-9_]/g, "_")
      .replace(/[_]+/g, "_");
    if (elementName[0] === "_") {
      elementName = elementName.slice(1);
    }
    if (elementName[elementName.length - 1] === "_") {
      elementName = elementName.slice(0, -1);
    }
    let element = { name: elementName, addedDate: new Date().toISOString() };
    // insert into document
    this.insertToDocument(element);
  },
  /**
   * After finding the component you want to use from the search results -- add it here by passing its entire object to this function.
   */
  insertToDocument: function (element: dataElement) {
    // insert into document
    Word.run(async (context) => {
      const contentRange = context.document.getSelection();
      const contentControl = contentRange.insertContentControl();
      contentControl.title = "";
      contentControl.tag = element.name;
      contentControl.color = "#666666";
      contentControl.cannotDelete = false;
      contentControl.cannotEdit = false;
      contentControl.appearance = "Tags";
      contentControl.insertText(element.name, "Replace");
      contentControl.cannotEdit = true;
      context.sync().then(() => {
        const state = get() as dataElementsStateType;
        set({
          usedInDocument: { ...state.usedInDocument, [element.name]: element },
        });
      });
    });
  },
}));

export default dataElementsState;
