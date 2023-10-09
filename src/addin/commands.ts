/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

console.log("COMMANDS . TS");

// Office.addin.setStartupBehavior(Office.StartupBehavior.load);
// Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
// Office.context.document.settings.saveAsync();

Office.onReady(() => {
  console.log("OFFICE . ON READY");
});
Office.initialize = () => {
  console.log("OFFICE . INITIALIZE");
};

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */ 1;
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal() as any;

// The add-in command functions need to be available in global scope
g.action = action;
