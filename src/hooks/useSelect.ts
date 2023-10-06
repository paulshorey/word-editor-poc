/* eslint-disable office-addins/no-context-sync-in-loop */
/* global console, Word, require */
import { logClear } from "@src/lib/log";
import wordDocumentState, { wordDocumentStateType } from "@src/state/wordDocument";

/**
 * Watch for a click on a content control. Interact with it in the app state.
 * NOTE: This does not look like a hook. But documentState() is a hook, so this must be used as a hook.
 */
export default function useSelect() {
  const wordDocument = wordDocumentState((state) => state as wordDocumentStateType);
  Word.run(async function (context) {
    logClear();
    // Get the current selection as a range.
    var selectionRange = context.document.getSelection();

    // State
    var clicked = {
      paragraphs: [],
      words: "",
      tag: "",
    };

    // log all paragraphs that the selection touches
    selectionRange.paragraphs.load("text");
    await selectionRange.paragraphs.context.sync();
    clicked.paragraphs = selectionRange.paragraphs.items.map((p) => p.text);

    // log word under cursor
    var words = selectionRange.getTextRanges([" ", "\t", "\r", "\n"], true); // just get everything including punctuation until nearest whitespace
    words.load("items");
    await context.sync();
    clicked.tag = "";
    for (let item of words.items) {
      if (/[A-Z_]+/.test(item.text)) {
        const contentControl = context.document.contentControls.getByTag(item.text)?.[0];
        if (!contentControl) {
          continue;
        }
        contentControl.load("select");
        contentControl.load("title");
        contentControl.load("id");
        await context.sync();
        // 1. Scroll to item
        contentControl.select("Start");
        // await context.sync();
        // 2. Update state
        switch (contentControl.title) {
          case ":":
            wordDocument.setSelectedTag(item.text);
            break;
        }
        clicked.tag = item.text;
        break;
      }
    }
    // save to application state
    console.log("clicked document", clicked);
    return context.sync();
  });
}
