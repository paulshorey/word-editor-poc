/* eslint-disable office-addins/no-context-sync-in-loop */
/* global console, setTimeout, Office, document, Word, require */
import { create } from "zustand";
import formatTag from "@src/functions/formatTag";
import selectAndHightlightItem from "@src/functions/selectAndHightlightControl";
import resetControl from "@src/functions/resetControl";

const labels = { DATA: true, TEXT: true, NUMBER: true };
export type label = keyof typeof labels;
export type id = number;
export type tag = string;
export type control = {
  /**
   * For Text/Number, this is string|number. For Data, this is the tag name.
   */
  value: any; // string | number;
  /**
   * Name/key of content, same as the visible text
   */
  tag: tag;
  /**
   * Unique identifier, generated by MS Word
   */
  id: id;
  /**
   * Not used as a title. Displays what type of data element.
   */
  title: label;
};

const debounceData = {
  renamingId: 0,
};

export type controlsStateType = {
  items: control[];
  label: label;
  selectedId: id;
  itemIdsTracked: Record<string, boolean>;
  //
  loadAll: () => Promise<control[]>;
  editValue: (id: id, value: string) => Promise<void>;
  insertTag: (label: label, name: tag, value?: string) => Promise<control | void>;
  //
  renameId: (idToEdit: id, name: tag) => Promise<control | void>;
  renameTags: (tagToEdit: tag, name: tag) => Promise<control | void>;
  rename: (context: any, control: any, name: tag) => Promise<control | void>;
  //
  deleteId: (id: id) => Promise<control[] | void>;
  deleteTags: (tag: tag) => Promise<control[] | void>;
  //
  selectTarget: (target: { ids: id[] }) => Promise<void>;
  clickTarget: (target: { ids: id[] }) => Promise<void>;
  selectId: (target: id) => void;
};
const controlsState = create((set, get) => {
  const state = {
    /**
     * All controls used in the template
     */
    items: [],
    label: "DATA",
    itemIdsTracked: {},
  } as controlsStateType;
  /**
   * Experimental. To use just one state file for several different types of variables.
   * BEFORE adding a new item, always set this to the correct type, as a separate command.
   */
  state.editValue = function (id, text) {
    const that = get() as controlsStateType;
    console.warn("controlsState.deleteTag()", id, text);
    return new Promise((resolve) => {
      Word.run(async (context) => {
        // 1. Edit tag name
        const contentControlParent = context.document.contentControls.getById(id);
        const childContentControls = contentControlParent.getRange("Content").getContentControls();
        await context.sync();
        childContentControls.load("items");
        let child = childContentControls?.getFirstOrNullObject();
        await context.sync();
        child.load("isNullObject");
        await context.sync();
        if (child && !child.isNullObject) {
          child.load("title");
          await context.sync();
          if (child.title === ":") {
            console.log(["CHILD item found", child.title, child]);
          } else {
            console.log(["NOT child item", child.title, child]);
          }
        }
        await context.sync();
        child.cannotEdit = false;
        child.load("insertText");
        await context.sync();
        child.insertText(text, "Replace");
        child.cannotEdit = true;
        await context.sync();
        // 2. Update state
        await that.loadAll();
        resolve();
      });
    });
  };
  /**
   * Add a control to the template, into the current cursor selection
   */
  state.insertTag = function (title, tag, value) {
    const that = get() as controlsStateType;
    return new Promise((resolve) => {
      // 1. Insert into document
      Word.run(async (context) => {
        const [tagName] = formatTag(tag, that.items);
        // const tagName = tag;
        console.log(["insertTag", tagName, title]);
        // parent
        const contentRange = context.document.getSelection();
        const contentControl = contentRange.insertContentControl();
        contentControl.title = title;
        contentControl.tag = tagName;
        resetControl(contentControl);
        contentControl.insertText(" ", "Replace");
        await context.sync();
        // child
        const contentControlRange = contentControl.getRange("Content");
        const contentControlChild = contentControlRange.insertContentControl();
        contentControlChild.tag = ":";
        contentControlChild.title = tagName;
        resetControl(contentControlChild);
        contentControlChild.insertText(value || tagName, "Replace");
        await context.sync();
        // control
        contentControlChild.cannotEdit = true;
        contentControl.cannotEdit = true;
        await context.sync();

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
        await that.loadAll();
        resolve();
      });
    });
  };
  /**
   * Rename tag of control
   */
  state.rename = async function (context: any, parentCC: Word.ContentControl, name: tag): Promise<any> {
    const that = get() as controlsStateType;
    const [tag, tagName, tagNum] = formatTag(name, that.items);
    parentCC.load(["id", "tag"]);
    await context.sync();
    if (tag === parentCC.tag || tagName === parentCC.tag || debounceData.renamingId === parentCC.id) {
      debounceData.renamingId = 0;
    } else {
      debounceData.renamingId = parentCC.id;
    }
    // modify inner ":" tag if nested inside container
    let childCC: Word.ContentControl = undefined;
    let parentRange = parentCC.getRange("Content");
    let childContentControls = parentRange.getContentControls();
    await context.sync();
    childContentControls.load("items");
    let child = childContentControls?.getFirstOrNullObject();
    await context.sync();
    child.load("isNullObject");
    await context.sync();
    if (child && !child.isNullObject) {
      child.load("title");
      await context.sync();
      if (child.title === ":") {
        console.log(["CHILD item found", child.title, child]);
        childCC = child;
      } else {
        console.log(["NOT child item", child.title, child]);
      }
    }
    // modify tags
    if (childCC) {
      // parent
      parentCC.cannotEdit = false;
      parentCC.load("tag");
      parentCC.tag = tag;
      // child
      childCC.cannotEdit = false;
      childCC.load("insertText");
      childCC.insertText(tag, "Replace");
      // sync
      childCC.cannotEdit = true;
      parentCC.cannotEdit = true;
      await context.sync();
    }
    return;
  };
  /**
   * Edit the item.tag name by Id
   */
  state.renameId = async function (idToEdit, name) {
    const that = get() as controlsStateType;
    console.warn("controlsState.deleteTag()");
    return new Promise((resolve) => {
      Word.run(async (context) => {
        // 1. Edit tag name
        const contentControl = context.document.contentControls.getById(idToEdit);
        await that.rename(context, contentControl, name);
        // 2. Update state
        await that.loadAll();
        resolve();
      });
    });
  };
  /**
   * Edit the item.tag name by Tag (all matching instances)
   */
  state.renameTags = async function (tagToEdit, name) {
    const that = get() as controlsStateType;
    console.warn("controlsState.deleteTag()");
    return new Promise((resolve) => {
      Word.run(async (context) => {
        // 1. Edit tag name
        const contentControls = context.document.contentControls.getByTag(tagToEdit);
        context.load(contentControls, "items");
        await context.sync();
        for (let item of contentControls.items) {
          await that.rename(context, item, name);
        }
        // 2. Update state
        await that.loadAll();
        resolve();
      });
    });
  };

  state.deleteId = function (id) {
    const that = get() as controlsStateType;
    console.warn("controlsState.deleteTag()");
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
        await that.loadAll();
        resolve();
      });
    });
  };

  state.deleteTags = function (tag) {
    const that = get() as controlsStateType;
    console.warn("controlsState.deleteTag()");
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
        await that.loadAll();
        resolve();
      });
    });
  };
  state.clickTarget = async function (target) {
    const that = get() as controlsStateType;
    let id = typeof target === "number" ? target : target?.ids?.[0];
    // console.log("controls.clickTarget", id);
    if (!id) {
      console.error("no ids", target);
    }
    if (that.selectedId !== id) {
      set({ selectedId: id });
      that.selectId(id);
    }
  };
  state.selectTarget = async function (target: any) {
    const that = get() as controlsStateType;
    let id = typeof target === "number" ? target : target?.ids?.[0];
    // console.log("controls.selectTarget", id);
    if (!id) {
      console.error("no ids", target);
    }
    if (that.selectedId !== id) {
      set({ selectedId: id });
      that.selectId(id);
    }
  };
  state.selectId = function (id: id): Promise<void> {
    console.log("controls.selectId", id);
    setTimeout(() => {
      set({ selectedId: 0 });
    }, 2000);
    return new Promise((resolve) => {
      Word.run(async (context) => {
        let item = context.document.contentControls.getById(id);
        item.load("id");
        await context.sync();
        item.load("text");
        await context.sync();
        console.log(item.text);
        await selectAndHightlightItem(item, context);
        await context.sync();
        resolve();
      });
    });
  };

  /**
   * IMPORTANT! WARNING! It's OK for production, but good to know...
   * During development, "hot reloading", (get() as controlsStateType) runs again and again.
   * Each time, itemIdsTracked is reset to empty.
   * So, `item.onEntered` is added multiple times per content control.
   */
  state.loadAll = function () {
    const that = get() as controlsStateType;
    console.warn("controlsState.loadAll()");
    const itemIdsTracked = that.itemIdsTracked || {};
    const itemIdsTrackedAdd = {};
    return new Promise((resolve) => {
      Word.run(async (context) => {
        // 1. Read document
        const contentControls = context.document.contentControls;
        context.load(contentControls, "items");
        await context.sync();
        // 2. Update state
        const all = [];
        for (let item of contentControls.items) {
          item.load(["tag", "title", "onEntered"]);
          await context.sync();
          if (item.tag === ":" || !labels[item.title]) {
            continue;
          }

          console.log("Loaded", item.title, item.tag, item.id, item.text);
          all.push({ id: item.id, tag: item.tag, title: item.title, value: item.text });

          if (!that.itemIdsTracked[item.id]) {
            resetControl(item);
            // console.log(["track item", item.text, item.id]);
            item.track();
            item.onEntered.add(that.clickTarget);
            item.onSelectionChanged.add(that.selectTarget);
            itemIdsTrackedAdd[item.id] = true;
            context.load(item);
          } else {
            // console.log(["already tracked", item.text, item.id]);
          }
        }
        set({ items: all, itemIdsTracked: { ...itemIdsTracked, ...itemIdsTrackedAdd } });
        resolve(all);
      });
    });
  };
  return state;
});

export default controlsState;
