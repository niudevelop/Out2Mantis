sap.ui.define([
    "sap/m/Button",
    "sap/m/TextArea",
    "sap/m/VBox"
  ], function (Button, TextArea, VBox) {
    "use strict";
  
    Office.onReady(() => {
      const textArea = new TextArea({
        growing: true,
        width: "100%",
        height: "300px"
      });
  
      const extractButton = new Button({
        text: "Extract Markdown from Email",
        press: () => {
          Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              const htmlBody = result.value;
  
              const cleanedHtml = window.sanitizeHtml(htmlBody, {
                allowedTags: window.sanitizeHtml.defaults.allowedTags,
                allowedAttributes: {},
                transformTags: {
                  '*': (tagName, attribs) => ({ tagName, attribs: {} })
                }
              });
  
              const turndown = new window.TurndownService();
              const markdown = turndown.turndown(cleanedHtml);
              textArea.setValue(markdown);
            }
          });
        }
      });
  
      const subjectText = new TextArea({
        value: "Subject: " + Office.context.mailbox.item.subject,
        editable: false,
        width: "100%"
      });
  
      new VBox({
        items: [subjectText, extractButton, textArea]
      }).placeAt("content");
    });
  });