import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { log, logKeys } from "../lib/log";

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
    // Get the current selection as a range.
    var range = context.document.getSelection();
    // range.load("paragraphs/items/text");
    await context.sync();
    range.paragraphs.load("text");
    await range.paragraphs.context.sync();
    // Get all words in the range.
    log(
      "range.paragraphs",
      range.paragraphs.items.map((p) => p.text)
    );
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
