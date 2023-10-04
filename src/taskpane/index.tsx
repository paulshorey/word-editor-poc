import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";

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
  Word.run(function (context) {
    // Get the current selection as a range.
    var range = context.document.getSelection();
<<<<<<< HEAD
=======
    // Get all words in the range.
    // var words = range[0].getFirst();
    console.log("range", range);
    // context.load(words, 'items');
    // return context.sync().then(function() {
    //   // Check if there are words in the range.
    //   if (words.items.length > 0) {
    //     // Select the first word in the range.
    //     words.items[0].select('Whole');
    //   }
    // });
    return context.sync();
>>>>>>> 18383bb (work in progress, experimenting, logging)
  }).catch(function (error) {
    console.log("Error: ", error);
  });
}

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
