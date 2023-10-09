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
      await context.sync();
      item.select("End");
      item.load("id");
      item.font.highlightColor = ""; // "KHAKI"
      await context.sync();
      // 2. Update state
      state.update((state) => {
        state.selectedTag = item.id + "";
      });
      resolve();
    });
  });
};
