import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { log, logKeys, logClear } from "../lib/log";

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
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, myHandler, function (result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      document.getElementById("message").innerText = result.error.message;
    }
  });
});

function myHandler() {
  Word.run(async function (context) {
    logClear();
    // Get the current selection as a range.
    var selectionRange = context.document.getSelection();

    // State
    var clicked = {
      paragraphs: [],
      words: "",
      variables: [],
    };

    // log all paragraphs that the selection touches
    selectionRange.paragraphs.load("text");
    await selectionRange.paragraphs.context.sync();
    clicked.paragraphs = selectionRange.paragraphs.items.map((p) => p.text);

    // log word under cursor
    var words = selectionRange.getTextRanges([" ", "\t", "\r", "\n"], true); // just get everything including punctuation until nearest whitespace
    words.load("items/text");
    await context.sync();
    clicked.words = words.items.map((p) => p.text).join(" ");

    // parse variables
    clicked.variables = clicked.words.match(/({.+})/g);

    // save to application state
    log("clicked", clicked);

    // done
    return context.sync();
  }).catch(function (error) {
    log("Error", error);
  });
}

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
