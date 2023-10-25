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

export type variableComponent = {
  tag: tag;
  id: id;
};

export type variableComponentsStateType = {
  items: variableComponent[];
  loadAll: () => Promise<Record<id, variableComponent>>;
  insertTag: (tag: tag) => variableComponent | undefined;
  //
  renameId: (idToEdit: id, tagRenamed: tag) => variableComponent | undefined;
  renameTags: (tagToEdit: tag, tagRenamed: tag) => variableComponent | undefined;
  //
  deleteId: (id: id) => Promise<variableComponent[]>;
  deleteTags: (tag: tag) => Promise<variableComponent[]>;
  //
  scrollToId: (id: id) => Promise<void>;
};

const variableComponentsState = create((set, _get) => ({
  /**
   * All variableComponents used in the template
   */
  items: [],
  /**
   * Add a variableComponent to the template, into the current cursor selection
   */
  insertTag: function (tag: tag) {
    tag = formatTag(tag);
    console.warn("variableComponentsState.selectTag()");
    return new Promise((resolve) => {
      // 1. Insert into document
      Word.run(async (context) => {
        const contentRange = context.document.getSelection();
        const contentControl = contentRange.insertContentControl();
        contentControl.title = ":";
        contentControl.tag = tag;
        contentControl.color = "#666666";
        contentControl.cannotDelete = true;
        contentControl.cannotEdit = false;
        contentControl.appearance = "Tags";
        contentControl.insertText(tag, "Replace");
        contentControl.cannotEdit = true;
        context.sync().then(async () => {
          // 2. Update state
          const all = await this.loadAll();
          resolve(all);
        });
      });
    });
  },
  /**
   * Edit the item.tag name by Id
   */
  renameId: function (idToEdit: id, tagRenamed: tag): Promise<variableComponent[]> {
    tagRenamed = formatTag(tagRenamed);
    console.warn("variableComponentsState.deleteTag()");
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
  renameTags: function (tagToEdit: tag, tagRenamed: tag): Promise<variableComponent[]> {
    tagRenamed = formatTag(tagRenamed);
    console.warn("variableComponentsState.deleteTag()");
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

  deleteId: function (id: id): Promise<variableComponent[]> {
    console.warn("variableComponentsState.deleteTag()");
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

  deleteTags: function (tag: tag): Promise<variableComponent[]> {
    console.warn("variableComponentsState.deleteTag()");
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
    console.warn("variableComponentsState.loadAll()");
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

export default variableComponentsState;

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