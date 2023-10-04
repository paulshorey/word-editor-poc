import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";

/* global document, Word, Office, module, require */

function log(...args) {
  document.getElementById("message").innerHTML = `
    ${args.map((arg) => `<pre><code>${JSON.stringify(arg, null, "  ")}</code></pre>`).join("")}
  `;
}

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
  // eventArgs
  Word.run(function (context) {
    // Get the current selection as a range.
    var range = context.document.getSelection();
    // Load the paragraphs property of the range object.
    range.load("paragraphs");
    return context.sync().then(function () {
      // Get the paragraphs.
      var paragraphs = range.paragraphs;
      // Load the text property for each paragraph.
      context.load(paragraphs, "text");
      return context.sync().then(function () {
        for (var i = 0; i < paragraphs.items.length; i++) {
          // Get the text of the paragraph.
          var paragraphText = paragraphs.items[i].text;
          // Do something with the paragraph text...
          var boldText = "<b>" + paragraphText + "</b>";
          // Replace the paragraph text with the bold version
          // paragraphs.items[i].insertHtml(boldText, "Replace");  // Paul - This causes an infinit loop on WEB - works on Desktop!
        }
        return context.sync();
      });
    });
  }).catch(function (error) {
    log("Error: ", error);
  });
}

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
