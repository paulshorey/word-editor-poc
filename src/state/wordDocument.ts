/* global console, setTimeout, Office, document, Word, require */
import { create } from "zustand";
import { id } from "@src/state/dataElements";
/**
 * contentControl.tag;
 */
export type tag = string;

export type wordDocumentStateType = {
  selectedTag: tag | undefined;
  setSelectedTag: (tag: tag) => void;
  unsetSelectedTag: () => void;
  scrollToId: (id: id) => Promise<void>;
};

const wordDocumentState = create((set, _get) => ({
  /**
   * The tag (any type of variable, data element, component, marker, doesn't matter) that is currently selected
   */
  selectedTag: undefined,
  /**
   * Set the selected tag
   */
  setSelectedTag: function (tag: tag) {
    set({ selectedTag: tag });
  },
  /**
   * Clear the selected tag
   */
  unsetSelectedTag: function () {
    set({ selectedTag: undefined });
  },
  /**
   * Selects the content control with the given id
   */
  scrollToId: async function (id: id): Promise<void> {
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
        set({
          selected: item.id,
        });
      });
      resolve();
    });
  },
}));

export default wordDocumentState;
