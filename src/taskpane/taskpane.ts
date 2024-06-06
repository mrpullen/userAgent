/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, navigator, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    //document.getElementById("run").onclick = run;

    if (navigator.userAgent.indexOf("Trident") !== -1) {
      document.getElementById("trident").style.display = "block";
      /*
      Trident is the webview in use. Do one of the following:
          1. Provide an alternate add-in experience that doesn't use any of the HTML5
          features that aren't supported in Trident (Internet Explorer 11).
          2. Enable the add-in to gracefully fail by adding a message to the UI that
          says something similar to:
          "This add-in won't run in your version of Office. Please upgrade either to
          perpetual Office 2021 (or later) or to a Microsoft 365 account."
      */
    } else if (navigator.userAgent.indexOf("Edge") !== -1) {
      document.getElementById("edge").style.display = "block";
      /*
      EdgeHTML is the browser in use. Do one of the following:
          1. Provide an alternate add-in experience that's supported in EdgeHTML (Microsoft Edge Legacy).
          2. Enable the add-in to gracefully fail by adding a message to the UI that
          says something similar to:
          "This add-in won't run in your version of Office. Please upgrade either to
          perpetual Office 2021 (or later) or to a Microsoft 365 account."
      */
    } else {
      document.getElementById("webview").style.display = "block";
      /* 
      A webview other than Trident or EdgeHTML is in use.
      Provide a full-featured version of the add-in here.
      */
    }
  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */

  const item = Office.context.mailbox.item;
  document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
}
