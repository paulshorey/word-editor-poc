import App from "@src/components/Taskpane";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";
import useSelect from "@src/hooks/useSelect";

/* global document, Word, Office, module, require */

initializeIcons();
// Office.addin.setStartupBehavior(Office.StartupBehavior.load);
// Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
// Office.context.document.settings.saveAsync();

console.log("TASKPANE . TSX");

let isOfficeInitialized = false;

const title = "FAF Task Pane Add-in";

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
  if (isOfficeInitialized) {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, useSelect);
  }
});

if ((module as any).hot) {
  (module as any).hot.accept("@src/components/Taskpane", () => {
    const NextApp = require("@src/components/Taskpane").default;
    render(NextApp);
  });
}
