/* eslint-disable office-addins/no-context-sync-in-loop */
/* global setTimeout, console, Office, document, Word, require, window */
import { create } from "zustand";
import { TAGNAMES } from "@src/constants/contentControlProperties";
import { ComponentTestData } from "@src/testdata/TestData";
import Don1 from "@src/testdata/Don1";

/**
 * contentControl.id; context.document.contentControls.getById(id)
 */
export type id = number;
/**
 * contentControl.tag; context.document.contentControls.getByTag(tag);
 */
export type tag = string;

type outputOption = {
  id: string;
  title: string;
  condition: string;
};

export type dataElementType = {
  tag: tag;
  id: id;
  title?: string;
  outputOptions?: outputOption[];
};

export type componentsStateType = {
  items: dataElementType[];
  loadAll: () => Promise<Record<id, dataElementType>>;
  insertTag: (documentName: string) => dataElementType | undefined;
  //
  deleteId: (id: id) => Promise<dataElementType[]>;
  // deleteTags: (tag: tag) => Promise<dataElement[]>;
  //
  scrollToId: (id: id) => Promise<void>;
};

const componentsState = create((set, _get) => ({
  /**
   * All dataElements used in the template
   */
  items: [],

  loadAll: function () {
    return new Promise((resolve) => {
      Word.run(async (context) => {
        // 1. Read document
        const contentControls = context.document.contentControls.getByTag(TAGNAMES.component);
        context.load(contentControls, "items");
        await context.sync();
        // 2. Update state
        const all = [];
        for (let item of contentControls.items) {
          all.push({ id: item.id, tag: item.tag, title: item.title });
        }
        set({ items: all });
        resolve(null);
      });
    });
  },

  /**
   * Add a dataElement to the template, into the current cursor selection
   */
  insertTag: function (documentName: string): void {
    setTimeout(async () => {
      Word.run(async (context) => {
        // 0. Get base64 data content
        let base64DataContent;
        switch (documentName) {
          case "comp_with_table":
            base64DataContent = ComponentTestData.comp_with_table.data;
            break;

          case "comp_simple_word":
            base64DataContent = ComponentTestData.comp_simple_word.data;
            break;

          case "don1":
            base64DataContent = Don1;
            break;

          default:
            Promise.reject("ERROR - Document does not exist");
            return;
        }
        const title = formatTag(documentName);

        // 1. Insert into document
        const contentRange = context.document.getSelection();
        const contentControl = contentRange.insertContentControl();
        contentControl.title = title;
        contentControl.tag = TAGNAMES.component;
        contentControl.color = "#666666";
        contentControl.cannotDelete = false;
        contentControl.cannotEdit = false;
        contentControl.appearance = "BoundingBox";
        contentControl.insertText(title, "Replace");
        // contentControl.cannotEdit = true;
        await context.sync();

        // 2. Move cursor outside of the new contentControl
        console.log(["insertHtml"]);
        const rangeAfter = contentControl.getRange("After");
        rangeAfter.load("insertHtml");
        await context.sync();
        rangeAfter.insertHtml(" <br /> ", "Start");
        await context.sync();
        rangeAfter.select("End");
        console.log(["insertHtml done"]);

        // 3. Insert content preview (this does not work - maybe we can do this on the back-end)
        // console.log(["insertFileFromBase64"]);
        // contentControl.load("insertFileFromBase64");
        // await context.sync();
        // contentControl.insertFileFromBase64(base64DataContent, "Replace");
        // await context.sync();
        // console.log(["insertFileFromBase64 done"]);

        // 4. Update app state
        await this.loadAll();
      });
    }, 5);
    setTimeout(async () => {
      Word.run(async (context) => {
        console.log(["context.document.body.load(); after 1000 ms"]);
        context.document.body.load();
        await context.sync();
      });
    }, 1000);
    setTimeout(async () => {
      Word.run(async (context) => {
        console.log(["context.document.body.load(); after 5000 ms"]);
        context.document.body.load();
        await context.sync();
      });
    }, 5000);
    setTimeout(async () => {
      Word.run(async (context) => {
        console.log(["context.document.body.load(); after 10000 ms"]);
        context.document.body.load();
        await context.sync();
      });
    }, 10000);
  },

  deleteId: function (id: id): Promise<dataElementType[]> {
    // eslint-disable-next-line no-undef
    console.log(["dataElementsState.deleteTag()"]);
    return new Promise((resolve) => {
      Word.run(async (context) => {
        // 1. Delete from document
        const contentControl = context.document.contentControls.getById(id);
        await context.sync();
        contentControl.load("delete");
        await context.sync();
        contentControl.delete(false);
        await context.sync();
        // 2. Update state
        const all = await this.loadAll();
        resolve(all);
      });
    });
  },

  // deleteTags: function (tag: tag): Promise<dataElement[]> {
  //   console.log(["dataElementsState.deleteTag()"]);
  //   return new Promise((resolve) => {
  //     Word.run(async (context) => {
  //       // 1. Delete from document
  //       const contentControls = context.document.contentControls.getByTag(tag);
  //       context.load(contentControls, "items");
  //       await context.sync();
  //       for (let item of contentControls.items) {
  //         item.delete(false);
  //       }
  //       await context.sync();
  //       // 2. Update state
  //       const all = await this.loadAll();
  //       resolve(all);
  //     });
  //   });
  // },
}));

export default componentsState;

// HELPERS LIBRARY:
/**
 * convert tag to uppercase, remove all spaces and special characters
 */
function formatTag(tag: tag): tag {
  tag = tag
    .toUpperCase()
    .replace(/[^A-Z0-9_]/g, "_")
    .replace(/[_]+/g, "_");
  if (tag[0] === "_") {
    tag = tag.slice(1);
  }
  if (tag[tag.length - 1] === "_") {
    tag = tag.slice(0, -1);
  }
  return tag;
}
