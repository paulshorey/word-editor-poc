/* global console, setTimeout, Office, document, Word, require */
import { createState } from "@persevie/statemanjs";
import { id } from "@src/state/dataElementsState";

export type tag = string;

export type stateType = {
  selectedTag: tag | undefined;
};

/**
 * Usage: ```import * as wordDocument from "@src/state/wordDocument";```
 * To render once: ```wordDocument.state.selectedTag;```
 * To set: ```wordDocument.setSelectedTag("myTag");```
 * To re-render whenever the state changes:
 * ```
 * wordDocument.state.subscribe(
 *     (state) => {
 *         console.log("State changed:", state);
 *     },
 *     { properties: ["callback.only.when.this.property.changed"] },
 * );
 * ```
 */
export const state = createState<stateType>({
  /**
   * The tag (any type of variable, data element, component, marker, doesn't matter) that is currently selected
   */
  selectedTag: "",
});

export type wordDocumentType = {
  unsetSelectedTag: () => void;
  scrollToId: (id: id) => Promise<void>;
};

/**
 * Set the selected tag
 */
export const setSelectedTag = (tag: tag): void => {
  state.update((state) => {
    console.log("setSelectedTag", tag);
    state.selectedTag = tag;
  });
};

/**
 * Clear the selected tag
 */
export const unsetSelectedTag = (): void => {
  state.update((state) => {
    state.selectedTag = undefined;
  });
};

/**
 * Selects the content control with the given id
 */
export const scrollToId = (id: id): Promise<void> => {
  return new Promise((resolve) => {
    Word.run(async (context) => {
      // 1. Scroll to item
      const item = context.document.contentControls.getById(id);
      await selectAndHightlightItem(item, context);
      // 2. Update state
      state.update((state) => {
        // Loading id and context.sync in selectAndHightlightItem()
        // eslint-disable-next-line office-addins/call-sync-before-read
        state.selectedTag = item.id + "";
      });
      resolve();
    });
  });
};

const debounceSelectedTag = {
  id: "",
};

export const selectAndHightlightItem = async (item: any, context: any): Promise<void> => {
  item.load("id");
  await context.sync();
  // do not scroll to same item - will start an infinite loop!
  if (debounceSelectedTag.id === item.id) return;
  debounceSelectedTag.id = item.id;
  // if new item, then go ahead and scroll, select, highlight
  item.select("Select");
  item.load("color");
  await context.sync();
  // item.color = "#F5C027";
  item.color = "#08E5FF";
  setTimeout(async () => {
    await context.sync();
    item.color = "#666666";
    context.sync();
  }, 1000);
};
