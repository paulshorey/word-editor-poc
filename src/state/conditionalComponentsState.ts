/* eslint-disable office-addins/no-context-sync-in-loop */
/* global console, setTimeout, Office, document, Word, require */
import { create } from "zustand";
import { TAGNAMES } from "@src/constants/constants";
import { dataElementType } from "@src/state/componentsState";

/**
 * contentControl.id; context.document.contentControls.getById(id)
 */
export type id = number;
/**
 * contentControl.tag; context.document.contentControls.getByTag(tag);
 */
export type tag = string;

export type outputOption = {
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
  getItemById: (id: id) => dataElementType | undefined;
  insertTag: (tagName: string, displayName: string) => dataElement | undefined;
  //
  deleteId: (id: id) => Promise<dataElement[]>;
  // deleteTags: (tag: tag) => Promise<dataElement[]>;
  //
  scrollToId: (id: id) => Promise<void>;
};

const conditionalComponentsState = create((set, _get) => ({
  /**
   * All dataElements used in the template
   */
  items: [],

  loadAll: function () {
    return new Promise((resolve) => {
      Word.run(async (context) => {
        // 1. Read document
        const contentControls = context.document.contentControls.getByTag(TAGNAMES.conditional);
        context.load(contentControls, "items");
        await context.sync();
        // 2. Update state
        const all = [];

        for (let item of contentControls.items) {
          const itemRange = item.getRange("Content");
          const allOutputOptions = [];
          const scenarios = itemRange.contentControls.getByTag(TAGNAMES.scenario);
          item.context.load(scenarios, "items");
          await context.sync();
          for (let scenario of scenarios.items) {
            allOutputOptions.push({ id: scenario.id, title: scenario.title });
          }

          all.push({
            id: item.id,
            tag: item.tag,
            title: item.title.replace(/^COND: /, ""),
            outputOptions: allOutputOptions,
          });
        }

        set({ items: all });
        resolve(null);
      });
    });
  },

  getItemById: function (id) {
    return this.items.find((item) => item.id === id);
  },

  /**
   * Add a dataElement to the template, into the current cursor selection
   */
  insertTag: function (tagName: string, displayName: string) {
    return new Promise((resolve) => {
      const defaultCondition = async (
        context: Word.RequestContext,
        contentRange: Word.Range,
        displayName: string | null
      ) => {
        const contentControl = contentRange.insertContentControl();
        contentControl.set({
          appearance: "Tags",
          cannotEdit: false,
          cannotDelete: false,
          color: "maroon",
          tag: TAGNAMES.scenario,
          title: displayName || "Un-Conditional",
        });
        await context.sync();
        return contentControl;
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
          title: `COND: ${displayName}`,
        });

        context.load(contentControl);
        await context.sync();
        const defaultTitle = "Default";
        const defaultConditionObj = await defaultCondition(context, contentControl.getRange("Content"), defaultTitle);
        context.load(contentControl);
        context.sync().then(async () => {
          // 2. Update state
          const dataElement = {
            id: contentControl.id,
            tag: contentControl.tag,
            title: displayName,
            outputOptions: [{ id: defaultConditionObj.id, title: defaultTitle, condition: "true" }],
          };
          await context.sync();
          await this.loadAll();
          resolve(dataElement);
        });
      });
    });
  },

  deleteId: function (id: id): Promise<dataElement[]> {
    console.warn("===> dataElementsState.deleteTag()", id);
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
