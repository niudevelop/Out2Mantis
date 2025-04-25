import sanitizeHtml from "sanitize-html";
import TurndownService from "turndown";






/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */

  const turndown = new TurndownService();

  const itemBody = Office.context.mailbox.item.body.getAsync(
    Office.CoercionType.Html,
    function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const htmlBody = result.value;
        const cleanedHtml = sanitizeHtml(htmlBody, {
          allowedTags: sanitizeHtml.defaults.allowedTags,
          allowedAttributes: {}, // remove all inline style, class, etc.
          transformTags: {
            '*': (tagName, attribs) => {
              // Drop all MS-specific attributes
              return { tagName: tagName, attribs: {} };
            }
          },
        });
        const markdown = turndown.turndown(cleanedHtml);
        console.log("Email Text content:", markdown);
      }
    }
  );
  const item = Office.context.mailbox.item;

  let insertAt = document.getElementById("item-subject");
  let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
  insertAt.appendChild(label);
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(item.subject));
  insertAt.appendChild(document.createElement("br"));
}
