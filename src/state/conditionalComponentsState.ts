/* eslint-disable office-addins/no-context-sync-in-loop */
/* global console, setTimeout, Office, document, Word, require */
import { create } from "zustand";
import { TAGNAMES } from "@src/constants/constants";

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

export type dataElement = {
  tag: tag;
  id: id;
  title?: string;
  outputOptions?: outputOption[];
};

export type conditionalComponentsStateType = {
  items: dataElement[];
  loadAll: () => Promise<Record<id, dataElement>>;
  insertTag: (tagName: string, condition: string) => dataElement | undefined;
  //
  deleteId: (id: id) => Promise<dataElement[]>;
  // deleteTags: (tag: tag) => Promise<dataElement[]>;
  //
  scrollToId: (id: id) => Promise<void>;
};

const conditionalComponentsState = create((set, get) => ({
  /**
   * All dataElements used in the template
   */
  items: [],

  loadAll: function () {
    console.log("dataElementsState.loadAll()");
    return new Promise((resolve) => {
      Word.run(async (context) => {
        // 1. Read document
        const contentControls = context.document.contentControls.getByTag(TAGNAMES.conditional);
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
  insertTag: function (tagName: string, condition: string) {
    return new Promise((resolve) => {
      const defaultCondition = (context: Word.RequestContext, contentRange: Word.Range, condition: string | null) => {
        const contentControl = contentRange.insertContentControl();
        contentControl.set({
          appearance: "Tags",
          cannotEdit: false,
          cannotDelete: false,
          color: "maroon",
          tag: "Default",
          title: condition || "Un-Conditional",
        });
        return context.sync();
      };

      // 1. Insert into document
      Word.run(async (context) => {
        const contentRange = context.document.getSelection().getRange();
        const contentControl = contentRange.insertContentControl();
        contentControl.set({
          appearance: "Tags",
          cannotEdit: false,
          cannotDelete: false,
          color: "blue",
          tag: tagName,
          title: "CONDITION",
        });

        context.load(contentControl);
        await context.sync();
        defaultCondition(context, contentControl.getRange("Content"), condition);
        context.load(contentControl);

        context.sync().then(async () => {
          // 2. Update state
          const dataElement = {
            id: contentControl.id,
            tag: contentControl.tag,
          };
          const state = get() as conditionalComponentsStateType;
          set({
            items: [dataElement, ...state.items],
          });
          await context.sync();
          resolve(dataElement);
        });
      });
    });
  },

  deleteId: function (id: id): Promise<dataElement[]> {
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

export default conditionalComponentsState;
