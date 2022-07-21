/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, run)
  }
});

export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */
    const doc = context.document;

    const selection = doc.getSelection();

    context.load(selection, ["text", "paragraphs"]);
    await context.sync();
    const paragraph = selection.paragraphs.getFirst();
    const ParagraphRange = paragraph.getRange("End");
    const Range = selection.getRange("Start").expandToOrNullObject(ParagraphRange).load("text");

    await context.sync();
    console.log("range", Range.text);
  });
}
