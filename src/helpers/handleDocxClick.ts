/* eslint-disable office-addins/no-context-sync-in-loop */
/* global console, Word, require */
import { logClear } from "@src/lib/log";
import * as wordDocument from "@src/state/wordDocument";
import { selectAndHightlightItem } from "@src/state/wordDocument";

/**
 * Watch for a click on a content control. Interact with it in the app state.
 * TODO: Whenever this function runs, app state needs to be updated with the new state.
 */
export default function handleDocxClick() {
  return new Promise((resolve) => {
    resolve(true);
    Word.run(async function (context) {
      logClear();
      // Get the current selection as a range.
      const selection = context.document.getSelection();
      selection.load("items");
      await selection.context.sync();
      // await context.sync();
      const selectionRange = selection.getRange("Whole");
      const allComponents = context.document.contentControls;
      allComponents.load("items");
      await context.sync();
      // State -- FOR CONSOLE LOG ONLY -- needs refactor/removed
      const state = {
        selectedParagraphs: [],
        selectedTags: [],
        clickedTag: "",
      };
      // log all paragraphs that the selection touches
      selectionRange.paragraphs.load("text");
      await selectionRange.paragraphs.context.sync();
      state.selectedParagraphs = selectionRange.paragraphs.items.map((p) => p.text);
      // log word under cursor
      const words = selectionRange.getTextRanges([" ", "\t", "\r", "\n"], true); // just get everything including punctuation until nearest whitespace
      words.load("items");
      await context.sync();
      for (let item of words.items) {
        // parse UPPERCASE words from item.text
        const maybeTags = item.text.match(/([0-9A-Z_]{3,})/g);
        if (!maybeTags) continue;
        for (let maybeTag of maybeTags) {
          // check if it's a tag
          const contentControl: Word.ContentControl = context.document.contentControls
            .getByTag(maybeTag)
            .getFirstOrNullObject();
          await context.sync();
          if (!contentControl.isNullObject) {
            state.selectedTags.push(maybeTag);
          }
          // contentControl.load("select");
          contentControl.load("title");
          // contentControl.load("id");
          await context.sync();
          // 1. Scroll to item
          // contentControl.select("Start");
          // await context.sync();
          // 2. Update state
          switch (contentControl.title) {
            case ":":
              wordDocument.setSelectedTag(maybeTag);
              await selectAndHightlightItem(contentControl, context);
              break;
          }
        }
      }
      if (state.selectedTags.length === 1) {
        state.clickedTag = state.selectedTags[0];
      }
      console.log("selected", state);
      // save to application state
      resolve(context.sync());
    });
  });
}
