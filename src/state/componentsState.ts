/* eslint-disable office-addins/no-context-sync-in-loop */
/* global console, setTimeout, Office, document, Word, require */
import { create } from "zustand";
import { TAGNAMES } from "@src/constants/constants";
import { ComponentTestData } from "@src/testdata/TestData";
import Poc2 from "@src/testdata/Poc2";
import Don1 from "@src/testdata/Don1";
import DummyXml from "@src/testdata/DummyXmlTag";
import DummyXmlContent from "@src/testdata/DummyXmlContent";

const wait = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

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
    Word.run(async (context) => {
      // 0. Get base64 data content
      let isXML = false;
      let contentToInsert;
      switch (documentName) {
        case "comp_with_table":
          contentToInsert = ComponentTestData.comp_with_table.data;
          break;
        case "comp_simple_word":
          contentToInsert = ComponentTestData.comp_simple_word.data;
          break;
        case "poc2":
          contentToInsert = Poc2;
          break;
        case "don1":
          contentToInsert = Don1;
          break;
        case "DummyXml":
          isXML = true;
          contentToInsert = DummyXml;
          break;
        case "DummyXmlContent":
          isXML = true;
          contentToInsert = DummyXmlContent;
          break;
        default:
          Promise.reject("ERROR - Document does not exist");
          return;
      }

      // 1. Insert into document
      //
      const contentRange = context.document.getSelection().getRange("Content");
      const contentControl = contentRange.insertContentControl();
      contentControl.tag = TAGNAMES.component; // `COMPONENT#${loadDocument}#${timeStamp}`
      contentControl.title = documentName.toUpperCase();
      contentControl.insertHtml("<div>Loading component content...</div>", "Start");
      await context.sync();
      const range = isXML
        ? contentControl.insertOoxml(contentToInsert, "Replace")
        : contentControl.insertFileFromBase64(contentToInsert, "Replace");
      // contentControl.cannotEdit = true;
      await range.context.sync();
      await context.sync();
      // insert line break if there is no text before
      const rangeBefore = contentControl.getRange("Before");
      const textBefore = rangeBefore.getTextRanges([" "], true).load();
      textBefore.load("items");
      await context.sync();
      if (textBefore.items.length === 0) {
        contentControl.insertBreak("Line", "Before");
        await context.sync();
      }
      // insert line break if there is no text after
      const rangeAfter = contentControl.getRange("After");
      const textAfter = rangeAfter.getTextRanges([" "], true).load();
      textAfter.load("items");
      await context.sync();
      if (textAfter.items.length === 0) {
        contentControl.insertBreak("Line", "After");
        await context.sync();
      }
      // sync
      await context.sync();
      context.document.body.load();
      context.document.load();

      //
      // const contentControl = contentRange.getRange("After").insertContentControl();
      // contentControl.tag = TAGNAMES.component; // `COMPONENT#${loadDocument}#${timeStamp}`
      // contentControl.title = documentName.toUpperCase();
      // contentControl.color = "#666666";
      // contentControl.cannotDelete = false;
      // contentControl.cannotEdit = false;
      // contentControl.appearance = "BoundingBox";
      // contentControl.clear();
      // await context.sync();
      // contentControl.insertHtml("<h1>contentControl</h1>", "Start");
      // await context.sync();
      // contentControl.insertFileFromBase64(base64DataContent, "Start");
      // await context.sync();
      // contentControl.insertText(" ", "Replace");
      // await context.sync();
      // contentControl.clear();
      // await context.sync();
      // contentControl.insertFileFromBase64(base64DataContent, "Replace");
      // await context.sync();
      // contentControl.insertHtml("<hr />", "End");
      // await context.sync();
      // 2. Update state
      this.loadAll();
      return context.sync();
    }).catch(async (error) => {
      console.log("Error: " + error);
      console.log("Debug info: " + JSON.stringify(error?.debugInfo || ""));
    });
  },

  deleteId: function (id: id): Promise<dataElementType[]> {
    // eslint-disable-next-line no-undef
    console.warn("dataElementsState.deleteTag()");
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
  //   console.warn("dataElementsState.deleteTag()");
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
