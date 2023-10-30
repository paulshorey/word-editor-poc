/* eslint-disable office-addins/no-context-sync-in-loop */
/* global console, setTimeout, Office, document, Word, require */
import { create } from "zustand";
import { createData, TITLES } from "@src/constants/contentControlProperties";
import { selectAndHightlightItem } from "@src/state/wordDocument";
// import * as wordDocument from "@src/state/wordDocument";

/**
 * contentControl.id; context.document.contentControls.getById(id)
 */
export type id = number;

/**
 * contentControl.tag; context.document.contentControls.getByTag(tag);
 */
export type tag = string;

export type dataElement = {
  tag: tag;
  id: id;
};

export type dataElementsStateType = {
  items: dataElement[];
  selectedId: id;
  itemIdsTracked: Record<string, boolean>;
  //
  loadAll: () => Promise<Record<id, dataElement>>;
  insertTag: (tag: tag) => dataElement | undefined;
  //
  renameId: (idToEdit: id, tagRenamed: tag) => dataElement | undefined;
  renameTags: (tagToEdit: tag, tagRenamed: tag) => dataElement | undefined;
  //
  deleteId: (id: id) => Promise<dataElement[]>;
  deleteTags: (tag: tag) => Promise<dataElement[]>;
  //
  scrollToId: (id: id) => Promise<void>;
  selectId: (id: id) => Promise<void>;
};

const dataElementsState = create((set, get) => ({
  /**
   * All dataElements used in the template
   */
  items: [],
  itemIdsTracked: {},
  /**
   * Add a dataElement to the template, into the current cursor selection
   */
  insertTag: function (tag: tag) {
    return new Promise((resolve) => {
      // 1. Insert into document
      Word.run(async (context) => {
        const cc = createData(tag, await this.loadAll());
        const contentRange = context.document.getSelection();
        const contentControl = contentRange.insertContentControl();
        contentControl.title = cc.title;
        contentControl.tag = cc.tag;
        contentControl.color = "#666666";
        contentControl.cannotDelete = false;
        contentControl.cannotEdit = false;
        contentControl.appearance = "Tags";
        contentControl.insertText(cc.tag, "Replace");
        contentControl.cannotEdit = true;
        contentControl.styleBuiltIn = "Strong";
        await context.sync();
        contentControl.cannotEdit = true;

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
        // rangeBefore.select();
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
        const all = await this.loadAll();
        resolve(all);
      });
    });
  },
  /**
   * Edit the item.tag name by Id
   */
  renameId: function (idToEdit: id, tagRenamed: tag): Promise<dataElement[]> {
    tagRenamed = formatTag(tagRenamed);
    console.warn("dataElementsState.deleteTag()");
    return new Promise((resolve) => {
      Word.run(async (context) => {
        // 1. Edit tag name
        const contentControl = context.document.contentControls.getById(idToEdit);
        await context.sync();
        contentControl.cannotEdit = false;
        contentControl.tag = tagRenamed;
        contentControl.insertText(tagRenamed, "Replace");
        contentControl.cannotEdit = true;
        contentControl.select("Start");
        await context.sync();
        // 2. Update state
        const all = await this.loadAll();
        resolve(all);
      });
    });
  },
  /**
   * Edit the item.tag name by Tag (all matching instances)
   */
  renameTags: function (tagToEdit: tag, tagRenamed: tag): Promise<dataElement[]> {
    tagRenamed = formatTag(tagRenamed);
    console.warn("dataElementsState.deleteTag()");
    return new Promise((resolve) => {
      Word.run(async (context) => {
        // 1. Edit tag name
        const contentControls = context.document.contentControls.getByTag(tagToEdit);
        context.load(contentControls, "items");
        await context.sync();
        for (let item of contentControls.items) {
          item.cannotEdit = false;
          item.tag = tagRenamed;
          item.insertText(tagRenamed, "Replace");
          item.cannotEdit = true;
          await context.sync();
        }
        // 2. Update state
        const all = await this.loadAll();
        resolve(all);
      });
    });
  },

  deleteId: function (id: id): Promise<dataElement[]> {
    console.warn("dataElementsState.deleteTag()");
    return new Promise((resolve) => {
      Word.run(async (context) => {
        const contentControl = context.document.contentControls.getById(id);
        // 1. Delete from document
        await context.sync();
        contentControl.load("delete");
        contentControl.cannotDelete = false;
        contentControl.delete(false);
        await context.sync();
        // 2. Update state
        const all = await this.loadAll();
        resolve(all);
      });
    });
  },

  deleteTags: function (tag: tag): Promise<dataElement[]> {
    console.warn("dataElementsState.deleteTag()");
    return new Promise((resolve) => {
      Word.run(async (context) => {
        // 1. Delete from document
        const contentControls = context.document.contentControls.getByTag(tag);
        context.load(contentControls, "items");
        await context.sync();
        for (let item of contentControls.items) {
          item.load("delete");
          item.cannotDelete = false;
          item.delete(false);
        }
        await context.sync();
        // 2. Update state
        const all = await this.loadAll();
        resolve(all);
      });
    });
  },

  /**
   * IMPORTANT! WARNING! It's OK for production, but good to know...
   * During development, "hot reloading", this runs again and again.
   * Each time, itemIdsTracked is reset to empty.
   * So, `item.onEntered` is added multiple times per content control.
   */
  loadAll: function () {
    console.warn("dataElementsState.loadAll()");
    const itemIdsTracked = (get() as dataElementsStateType).itemIdsTracked || {};
    const itemIdsTrackedAdd = {};
    return new Promise((resolve) => {
      Word.run(async (context) => {
        // 1. Read document
        const contentControls = context.document.contentControls.getByTitle(TITLES.data);
        context.load(contentControls, "items");
        await context.sync();
        // 2. Update state
        const all = [];
        for (let item of contentControls.items) {
          item.load("onEntered");
          await context.sync();
          if (!itemIdsTracked[item.id]) {
            console.log(["tracking click event", item.text, item.id]);
            item.track();
            item.color = "#666666";
            item.onEntered.add(this.selectId);
            itemIdsTrackedAdd[item.id] = true;
          } else {
            console.log(["already tracked", item.text, item.id]);
          }
          all.push({ id: item.id, tag: item.tag });
        }
        set({ items: all, itemIdsTracked: { ...itemIdsTracked, ...itemIdsTrackedAdd } });
        resolve(all);
      });
    });
  },
  selectId: async function (target: any) {
    Word.run(async (context) => {
      console.warn("clicked", target);
      if (!target.ids) {
        console.error("no ids", target);
      }
      let id = target.ids[0];
      let item = context.document.contentControls.getById(id);
      item.load("id");
      await context.sync();
      item.load("text");
      await context.sync();
      console.log(item.text);
      set({ selectedId: id });
      await selectAndHightlightItem(item, context);
      await context.sync();
    });
  },
}));

export default dataElementsState;

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
