/* eslint-disable office-addins/no-context-sync-in-loop */
/* global setTimeout, Office, document, Word, require */
import { create } from "zustand";
import { TAGNAMES } from "@src/constants/constants";
import { ComponentTestData } from "@src/testdata/TestData";

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

const componentsState = create((set, get) => ({
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
  insertTag: function (documentName: string) {
    return new Promise((resolve) => {
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

          default:
            Promise.reject("ERROR - Document does not exist");
            return;
        }
        // 1. Insert into document
        const contentRange = context.document.getSelection().getRange();
        const contentControl = contentRange.insertContentControl();
        contentControl.set({
          tag: TAGNAMES.component, // `COMPONENT#${loadDocument}#${timeStamp}`
          title: documentName.toUpperCase(),
        });
        await context.sync();
        contentControl.load("insertFileFromBase64");
        await context.sync();
        contentControl.insertFileFromBase64(base64DataContent, "Replace");
        context
          .sync()
          .then((res) => {
            // eslint-disable-next-line no-undef
            console.log("===> RES:", res);
          })
          .catch((error) => {
            // eslint-disable-next-line no-undef
            console.log("===> Error", error);
          });

        const TIMEOUT = 2000;
        setTimeout(() => {
          // eslint-disable-next-line no-undef
          console.log("===> RESET page", TIMEOUT);
          const body = context.document.body;
          body.load();

          context
            .sync()
            .then(() => {
              // eslint-disable-next-line no-undef
              console.log("===> RELOAD", TIMEOUT);
            })
            .catch((error) => {
              // eslint-disable-next-line no-undef
              console.log("===> Error Clear", error);
            });
        }, TIMEOUT);
        // 2. Update state
        const dataElement: dataElementType = {
          id: contentControl.id,
          tag: contentControl.tag,
          title: documentName.toUpperCase(),
        };
        const state = get() as componentsStateType;
        set({
          items: [dataElement, ...state.items],
        });
        await context.sync();
        await this.loadAll();
        resolve(dataElement);
      });
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
