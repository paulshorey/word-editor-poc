/* eslint-disable office-addins/no-context-sync-in-loop */
/* global console, setTimeout, Office, document, Word, require */
import { create } from "zustand";
import { ComponentTestData } from "@src/Components/testData";

/**
 * contentControl.id; context.document.contentControls.getById(id)
 */
export type id = number;
/**
 * contentControl.tag; context.document.contentControls.getByTag(tag);
 */
export type tag = string;

export type componentType = {
  tag: tag;
  id: id;
  title?: string;
};

export type componentsStateType = {
  items: componentType[];
  loadAll: () => Promise<Record<id, componentType>>;
  add: (documentName: string) => componentType | undefined;
  delete: (id: id) => Promise<componentType[]>;
};

const componentsState = create((set, _get) => ({
  /**
   * All components used in the template
   */
  items: [],

  /**
   * Sync with Word API to make a list of all components in the Word document
   */
  loadAll: function () {
    return new Promise((resolve) => {
      Word.run(async (context) => {
        // 1. Read document
        const contentControls = context.document.contentControls.getByTag("COMPONENT");
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
   * Add a component to the template, into the current cursor selection
   */
  add: function (documentName: string) {
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
          tag: "COMPONENT",
          title: documentName.toUpperCase(),
          appearance: "Hidden",
        });
        await context.sync();
        contentControl.load("insertFileFromBase64");
        await context.sync();
        contentControl.insertFileFromBase64(base64DataContent, "Replace");
        await context.sync();
        //
        setTimeout(() => {
          console.log("context.document.load() after 2000 ms");
          context.document.load();
          return context.sync();
        }, 2000);

        // 2. Move cursor outside of the new contentControl
        // insert space after
        console.warn("insertHtml After");
        const rangeAfter = contentControl.getRange("After");
        rangeAfter.load(["insertHtml", "html"]);
        await context.sync();
        rangeAfter.insertHtml("&nbsp;<br />&nbsp;", "Start");
        await context.sync();
        rangeAfter.select("End");
        console.warn("insertHtml After done");
        // insert space before
        console.warn("insertText Before");
        const rangeBefore = contentControl.getRange("Before");
        rangeBefore.load("text");
        await context.sync();
        console.log("insertText Before done", rangeBefore.text);
        //
        rangeBefore.load("insertHtml Before");
        await context.sync();
        rangeBefore.insertHtml("&nbsp;", "End");
        await context.sync();
        rangeBefore.select("End");
        console.warn("insertHtml Before done");

        // 3. Update app state
        const all = await this.loadAll();
        resolve(all);
      });
    });
  },

  /**
   * Delete one component by ID
   */
  delete: function (id: id): Promise<componentType[]> {
    // eslint-disable-next-line no-undef
    console.warn("componentsState.deleteTag()");
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
}));

export default componentsState;
