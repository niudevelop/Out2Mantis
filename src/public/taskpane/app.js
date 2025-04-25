sap.ui.define(["sap/m/Button", "sap/m/TextArea", "sap/m/VBox"], function (Button, TextArea, VBox) {
  "use strict";

  Office.onReady().then(() => {
    if (!Office.context.mailbox || !Office.context.mailbox.item) {
      console.error("Office context not ready yet.");
      return;
    }

    const textArea = new TextArea({
      growing: true,
      width: "100%",
      height: "300px",
    });

    const extractButton = new Button({
      text: "Extract Markdown from Email",
      press: () => {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const htmlBody = result.value;

            // const cleanedHtml = window.sanitizeHtml(htmlBody, {
            //   allowedTags: window.sanitizeHtml.defaults.allowedTags,
            //   allowedAttributes: {},
            //   transformTags: {
            //     "*": (tagName, attribs) => ({ tagName, attribs: {} }),
            //   },
            // });

            const cleanedHtml = DOMPurify.sanitize(htmlBody, {
              ALLOWED_TAGS: [
                "address",
                "article",
                "aside",
                "footer",
                "header",
                "h1",
                "h2",
                "h3",
                "h4",
                "h5",
                "h6",
                "hgroup",
                "main",
                "nav",
                "section",
                "blockquote",
                "dd",
                "div",
                "dl",
                "dt",
                "figcaption",
                "figure",
                "hr",
                "li",
                "main",
                "ol",
                "p",
                "pre",
                "ul",
                "a",
                "abbr",
                "b",
                "bdi",
                "bdo",
                "br",
                "cite",
                "code",
                "data",
                "dfn",
                "em",
                "i",
                "kbd",
                "mark",
                "q",
                "rb",
                "rp",
                "rt",
                "rtc",
                "ruby",
                "s",
                "samp",
                "small",
                "span",
                "strong",
                "sub",
                "sup",
                "time",
                "u",
                "var",
                "wbr",
                "caption",
                "col",
                "colgroup",
                "table",
                "tbody",
                "td",
                "tfoot",
                "th",
                "thead",
                "tr",
              ],
              ALLOWED_ATTR: ["href", "title", "target"],
            });

            const turndown = new window.TurndownService();
            const markdown = turndown.turndown(cleanedHtml);
            textArea.setValue(markdown);
          }
        });
      },
    });

    const subjectText = new TextArea({
      value: "Subject: " + Office.context.mailbox.item.subject,
      editable: false,
      width: "100%",
    });

    new VBox({
      items: [subjectText, extractButton, textArea],
    }).placeAt("content");
  });
});
