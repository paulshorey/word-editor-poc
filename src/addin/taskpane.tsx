import App from "@src/Taskpane";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";

/* global setTimeout, console, OfficeExtension, document, Word, Office, module, require */
// OfficeExtension.config.extendedErrorLogging = true;

initializeIcons();

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
  // // auto-open the taskpane
  // Office.addin.setStartupBehavior(Office.StartupBehavior.load);
  // Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
  // Office.context.document.settings.saveAsync();
  // load the app
  isOfficeInitialized = true;
  render(App);
});

if ((module as any).hot) {
  (module as any).hot.accept("@src/Taskpane", () => {
    const NextApp = require("@src/Taskpane").default;
    render(NextApp);
  });
}
