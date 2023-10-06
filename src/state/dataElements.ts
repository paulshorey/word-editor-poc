import create from "zustand";

/* global document, Word, require */

export type dataElement = {
  tag: string;
  otherData: any;
};

export type dataElementsStateType = {
  searchResults: dataElement[];
  usedInDocument: Record<string, dataElement>;
  insertToDocument: (element: dataElement) => dataElement | undefined;
  insertToDocumentByName: (name: string) => dataElement | undefined;
};

const dataElementsState = create((set, get) => ({
  /**
   * Temporary, to hold the results from an API call. For now it is just dummy data.
   */
  searchResults: [
    {
      tag: "TEST_1",
      otherData: "idk",
    },
    {
      tag: "TEST_2",
      otherData: "idk",
    },
    {
      tag: "TEST_3",
      otherData: "idk",
    },
    {
      tag: "TEST_4",
      otherData: "idk",
    },
    {
      tag: "TEST_5",
      otherData: "idk",
    },
  ],
  /**
   * All (unique) elements used in the document. We still need to decide how to manage the position of each variable in the document. This is too basic.
   */
  inDocument: {},
  /**
   * Shortcut to add by a tag (string) -- not finished -- need to think about how to actually manage state and insert into document.
   */
  insertToDocumentByName: function (name: string) {
    // convert vartag to uppercase, remove all spaces and special characters
    name = name
      .toUpperCase()
      .replace(/[^A-Z0-9_]/g, "_")
      .replace(/[_]+/g, "_");
    if (name[0] === "_") {
      name = name.slice(1);
    }
    if (name[name.length - 1] === "_") {
      name = name.slice(0, -1);
    }
    name = "DATA_" + name;
    let element = { tag: name, addedDate: new Date().toISOString() };
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
      contentControl.tag = element.tag;
      contentControl.color = "#666666";
      contentControl.cannotDelete = false;
      contentControl.cannotEdit = false;
      contentControl.appearance = "Tags";
      contentControl.insertText(element.tag, "Replace");
      contentControl.cannotEdit = true;
      context.sync().then(async () => {
        const state = get() as dataElementsStateType;
        set({
          usedInDocument: { ...state.usedInDocument, [element.tag]: element },
        });
        await context.sync();
      });
    });
  },
}));

export default dataElementsState;
