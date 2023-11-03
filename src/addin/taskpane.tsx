import App from "@src/components/Taskpane";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";

/* global setTimeout, console, OfficeExtension, document, Word, Office, module, require */
// OfficeExtension.config.extendedErrorLogging = true;

initializeIcons();
console.log("Office.addin.showAsTaskpane() attempt 1");
Office.addin.showAsTaskpane();

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

Office.onReady(async () => {
  console.log("Office.addin.showAsTaskpane() attempt 2");
  await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
  // console.log("Office.ribbon.requestUpdate attempt");
  // await Office.ribbon.requestUpdate({});
  // let addinState = await Office.addin.getStartupBehavior();
  // // auto-open the taskpane
  // Office.addin.setStartupBehavior(Office.StartupBehavior.load);
  // Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
  // Office.context.document.settings.saveAsync();
  // load the app
  isOfficeInitialized = true;
  render(App);
  if (isOfficeInitialized) {
    // setTimeout(() => {
    //   // auto-open the taskpane
    //   Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    //   Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
    //   Office.context.document.settings.saveAsync();
    // }, 3000);
  }
});

if ((module as any).hot) {
  (module as any).hot.accept("@src/components/Taskpane", () => {
    const NextApp = require("@src/components/Taskpane").default;
    render(NextApp);
  });
}
