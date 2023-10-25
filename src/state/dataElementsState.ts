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

export type dataElement = {
  tag: tag;
  id: id;
};

export type dataElementsStateType = {
  items: dataElement[];
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
};

const dataElementsState = create((set, _get) => ({
  /**
   * All dataElements used in the template
   */
  items: [],
  /**
   * Add a dataElement to the template, into the current cursor selection
   */
  insertTag: function (tag: tag) {
    tag = formatTag(tag);
    console.warn("dataElementsState.selectTag()");
    return new Promise((resolve) => {
      // 1. Insert into document
      Word.run(async (context) => {
        const contentRange = context.document.getSelection();
        const contentControl = contentRange.insertContentControl();
        contentControl.title = ":";
        contentControl.tag = tag;
        contentControl.color = "#666666";
        contentControl.cannotDelete = false;
        contentControl.cannotEdit = false;
        contentControl.appearance = "Tags";
        contentControl.insertText(tag, "Replace");
        contentControl.cannotEdit = true;
        await context.sync();

        // 2. Move cursor outside of the new contentControl
        // insert space after
        console.warn("insertAfter");
        const rangeAfter = contentControl.getRange("After");
        rangeAfter.load("insertAfter insertHtml");
        await context.sync();
        rangeAfter.insertHtml("&nbsp;<br />&nbsp;", "Start");
        await context.sync();
        rangeAfter.select("End");
        console.warn("insertAfter done");
        // insert space before
        console.warn("insertBefore");
        const rangeBefore = contentControl.getRange("Before");
        rangeBefore.load("text");
        await context.sync();
        console.log("text before ", rangeBefore.text);

        rangeBefore.load("insertBefore insertHtml");
        await context.sync();
        rangeBefore.insertHtml("&nbsp;", "End");
        await context.sync();
        rangeBefore.select("End");
        console.warn("insertBefore done");

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
          await context.sync();
          item.cannotEdit = false;
          item.tag = tagRenamed;
          item.insertText(tagRenamed, "Replace");
          item.cannotEdit = true;
          await context.sync();
        }
        await context.sync();
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

  loadAll: function () {
    console.warn("dataElementsState.loadAll()");
    return new Promise((resolve) => {
      Word.run(async (context) => {
        // 1. Read document
        const contentControls = context.document.contentControls.getByTitle(":");
        context.load(contentControls, "items");
        await context.sync();
        // 2. Update state
        const all = [];
        for (let item of contentControls.items) {
          all.push({ id: item.id, tag: item.tag });
        }
        set({ items: all });
        resolve(all);
      });
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
