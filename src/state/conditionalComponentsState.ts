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
  insertScenario: (id: id, scenarioName: string, rule: string) => dataElement | undefined;
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

  insertScenario: function (id, scenarioName, _rule) {
    return new Promise((resolve) => {
      Word.run(async (context) => {
        // 1. Insert into document after previous scenarios
        const parent = context.document.contentControls.getById(id);
        context.load(parent);
        await context.sync();
        // await defaultCondition(context, contentControl.getRange("End"), scenarioName);

        const child = parent.getRange("End").insertContentControl();
        child.set({
          appearance: "Tags",
          cannotEdit: false,
          cannotDelete: false,
          color: "maroon",
          tag: TAGNAMES.scenario,
          title: (scenarioName || "Unconditional") + " Scenario",
        });
        await context.sync();
        child.insertText(
          (scenarioName || "Unconditional") + " Scenario text. Click and edit this content...",
          "Replace"
        );
        await context.sync();

        context.load(parent);

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
   * Add a conditionalComponent to the template, into the current cursor selection
   */
  insertTag: function (tagName: string, displayName: string) {
    return new Promise((resolve) => {
      Word.run(async (context) => {
        const contentRange = context.document.getSelection().getRange();
        const contentControl = contentRange.insertContentControl();
        contentControl.set({
          appearance: "BoundingBox",
          cannotEdit: false,
          cannotDelete: false,
          color: "blue",
          tag: tagName,
          title: `${displayName} Condition`,
        });
        contentControl.load("insertText");
        contentControl.load("insertHtml");
        await context.sync();
        contentControl.insertText(" ", "Replace");
        context.load(contentControl);
        contentControl.getRange("Before").insertHtml("<br />", "Start");
        contentControl.getRange("After").insertHtml("<br />", "Start");
        context.load(contentControl);
        await context.sync();
        await this.insertScenario(contentControl.id, "Default", "");
        context.load(contentControl);
        await contentControl.context.sync();
        await this.loadAll();
        resolve(true);
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
