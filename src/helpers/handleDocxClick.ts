/* eslint-disable office-addins/no-context-sync-in-loop */
/* global console, Word, require */
import { TAGNAMES } from "@src/constants/constants";
import { logClear } from "@src/lib/log";
import * as wordDocument from "@src/state/wordDocument";
import { selectAndHightlightItem } from "@src/state/wordDocument";

/**
 * Watch for a click on a content control. Interact with it in the app state.
 * TODO: Whenever this function runs, app state needs to be updated with the new state.
 */
export default function handleDocxClick() {
  return new Promise((resolve) => {
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
      allComponents.items.forEach(async (item) => {
        try {
          item.load("id");
          item.load("tag");
          item.load("text");
          item.load("title");
          item.load("items");
          await item.context.sync();
          // await context.sync();
          let itemRange = item.getRange("Content");
          itemRange.load("intersectWithOrNullObject");
          itemRange.load("items");
          await itemRange.context.sync();
          // await context.sync();
          let intersection = itemRange.intersectWithOrNullObject(selection.getRange("Content"));
          intersection.load("items");
          intersection.load("isNullObject");
          intersection.load();
          await intersection.context.sync();
          await context.sync();
          console.log([item.tag, item.title, item.text, item.id]);
          if (intersection.isNullObject) {
            console.log(" ===> NO INTERSECTION");
          } else {
            console.log("===> INTERSECTS WITH: ", item.id, item.tag, item.title);
            resolve(context.sync());
            return;
          }
        } catch (e) {
          console.log(" ===> ERROR", e);
        }
      });

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
