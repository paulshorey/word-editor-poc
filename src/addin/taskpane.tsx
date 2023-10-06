import App from "@src/components/Taskpane";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { logClear } from "@src/lib/log";
// import dataElementsState, { dataElementsStateType } from "@src/state/dataElements";

/* global document, Word, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "Contoso Task Pane Add-in";

const render = (Component) => {
  ReactDOM.render(
    <AppContainer>
      <ThemeProvider>
        <Component title={title} isOfficeInitialized={isOfficeInitialized} />
      </ThemeProvider>
    </AppContainer>,
    document.getElementById("container")
  );
};

Office.onReady(() => {
  isOfficeInitialized = true;
  render(App);
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelect, function (result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      document.getElementById("message").innerText = result.error.message;
    }
  });
});

function onSelect() {
  // // // const dataElements = dataElementsState((state) => state as dataElementsStateType);
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
    wordItems: for (let item of words.items) {
      if (/[A-Z_]+/.test(item.text)) {
        const contentControls = context.document.contentControls.getByTag(item.text);
        // context.load(contentControls, "select");
        context.load(contentControls, "title");
        context.load(contentControls, "items");
        // delete all from state
        await context.sync();
        for (let control of contentControls.items) {
          context.load(control, "tag");
          // item.select("End");
          // dataElements.updateSelectedTag(control.tag);
          clicked.tag = item.text;
          break wordItems;
        }
      }
    }

    // save to application state
    // eslint-disable-next-line no-undef
    console.log("clicked document", clicked);
    return context.sync();
  }).catch(function (error) {
    // eslint-disable-next-line no-undef
    console.error("Error", error);
  });
}

if ((module as any).hot) {
  (module as any).hot.accept("@src/components/Taskpane", () => {
    const NextApp = require("@src/components/Taskpane").default;
    render(NextApp);
  });
}
