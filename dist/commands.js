define(function() { return /******/ (function() { // webpackBootstrap
var __webpack_exports__ = {};
/*!**********************************!*\
  !*** ./src/commands/commands.js ***!
  \**********************************/
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(function () {
  // If needed, Office.js is ready to be called.
  Office.actions.associate("openDialog", function (event) {
    Office.context.ui.displayDialogAsync("https://localhost:3000/mantis-dialog.html", {
      height: 50,
      width: 50
    }, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        var dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
          console.log("Received from dialog:", arg.message);
          dialog.close();
        });
      }
    });
    event.completed();
  });
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  // if (!Office.context.mailbox.item.notificationMessages) {
  //   console.error("NotificationMessages API not available in this context (likely Read mode).");
  //   event.completed();
  //   return;
  // }
  // const message = {
  //   type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
  //   message: "Performed action.",
  //   icon: "Icon.80x80",
  //   persistent: true,
  // };
  // Office.context.mailbox.item.notificationMessages.replaceAsync(
  //   "ActionPerformanceNotification",
  //   message
  // );
  // Be sure to indicate when the add-in command function is complete.
  // event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);
/******/ 	return __webpack_exports__;
/******/ })()
;
});;
//# sourceMappingURL=commands.js.map