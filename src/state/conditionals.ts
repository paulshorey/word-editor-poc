/* eslint-disable office-addins/no-context-sync-in-loop */
/* global console, setTimeout, Office, document, Word, require */
import { create } from "zustand";

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
  insertScenario: (id: id, scenarioName: string, rule: string) => dataElement | undefined;
  getItemById: (id: id) => dataElement | undefined;
  insertTag: (tagName: string, displayName: string) => Promise<void>;
  //
  deleteId: (id: id) => Promise<dataElement[]>;
  // deleteTags: (tag: tag) => Promise<dataElement[]>;
  //
  scrollToId: (id: id) => Promise<void>;
};

const defaultCondition = async (context: Word.RequestContext, contentRange: Word.Range, displayName: string | null) => {
  const childContentControl = contentRange.insertContentControl();
  childContentControl.set({
    appearance: "Tags",
    cannotEdit: false,
    cannotDelete: false,
    color: "maroon",
    tag: "SCENARIO",
    title: displayName || "Un-Conditional",
  });
  await context.sync();
  return childContentControl;
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
        const contentControls = context.document.contentControls.getByTag("CONDITIONAL");
        context.load(contentControls, "items");
        await context.sync();
        // 2. Update state
        const all = [];

        for (let item of contentControls.items) {
          const itemRange = item.getRange("Content");
          const allOutputOptions = [];
          const scenarios = itemRange.contentControls.getByTag("SCENARIO");
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

  insertScenario: function (id, scenarioName, _rule) {
    return new Promise((resolve) => {
      // 1. Insert into document after previous scenarios
      Word.run(async (context) => {
        const contentControl = context.document.contentControls.getById(id);
        context.load(contentControl);
        await context.sync();
        await defaultCondition(context, contentControl.getRange("End"), scenarioName);
        context.load(contentControl);

        context.sync().then(async () => {
          await this.loadAll();
          resolve(true);
        });
      });
    });
  },

  getItemById: function (id) {
    return this.items.find((item) => item.id === id);
  },

  /**
   * Add a dataElement to the template, into the current cursor selection
   */
  insertTag: function (tagName: string, displayName: string): Promise<void> {
    return new Promise((resolve) => {
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
        contentControl.insertText(" ", "Start");
        context.load(contentControl);
        await context.sync();

        context.load(contentControl);
        await defaultCondition(context, contentControl.getRange("End"), defaultTitle);
        context.load(contentControl);
        context.sync().then(async () => {
          await this.loadAll();
          resolve();
        });

        // 2. Move cursor outside of the new contentControl
        // insert space before
        const rangeBefore = contentControl.getRange("Before");
        rangeBefore.load(["text", "html", "getHtml"]);
        await context.sync();
        console.log(`text`, rangeBefore.text);
        await context.sync();
        rangeBefore.load("insertHtml");
        await context.sync();
        rangeBefore.insertHtml("&nbsp;", "Start");
        await context.sync();
        rangeBefore.select();
        // insert space after
        const rangeAfter = contentControl.getRange("After");
        rangeAfter.load("insertHtml");
        rangeAfter.load("text");
        await context.sync();
        console.log("rangeAfter", rangeAfter.text);
        await context.sync();
        const afterAdded = rangeAfter.insertHtml("&nbsp;", "End");
        await context.sync();
        afterAdded.select();

        // 3. Update app state
        await this.loadAll();
        resolve();
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
