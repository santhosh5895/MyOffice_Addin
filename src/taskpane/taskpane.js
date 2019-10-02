/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
// var emlformat = require('../../node_modules/jspdf/dist/jspdf.min.js');
import * as jsPDF from 'jspdf';

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("runSendMailContent").onclick = run;
  }
});

export async function run() {
  // Get a reference to the current message
  var item = Office.context.mailbox.item;
  // console.log(JSON.stringify(item.sender));
  // console.log(JSON.stringify(item.to))
  // console.log(JSON.stringify(item.bcc))
  // console.log(JSON.stringify(item.cc))
  // console.log(JSON.stringify(item.dateTimeCreated))
  // console.log(JSON.stringify(item.subject))
  // console.log(JSON.stringify(item.attachments))
  // console.log(JSON.stringify(item.location))
  // console.log(JSON.stringify(item.from))

  var doc = new jsPDF();
  var pdfText = "From: " + item.sender.emailAddress + "\n\nTo: "
    + item.to.reduce((acc, crt) => acc + crt.emailAddress + ";", "") + "\n\nCC: "
    + item.cc.reduce((acc, crt) => acc + crt.emailAddress + ";", "") + "\n\nSubject: "
    + item.subject + "\n\nContent: "
  Office.context.mailbox.item.body.getAsync(
    "text",
    {},
    function callback(result) {
      var value = pdfText + result.value.replace(/(\r\n|\r|\n){2,}/g, '$1\n');
      var splitText = doc.splitTextToSize(value, 180);
      doc.text(splitText, 15, 20);
      var form = new FormData();
      form.append("binary", btoa(doc.output()));

      var settings = {
        "async": true,
        "crossDomain": true,
        "url": "https://appiandev.vuram.com/suite/webapi/uploadDocument",
        "method": "POST",
        "processData": false,
        "contentType": "multipart/form-data",
        "mimeType": "multipart/form-data",
        "headers": {
          "Authorization": "Basic c2FudGhvc2hrdW1hcmFAdnVyYW0uY29tOlNAbnRoMHNo",
          "Appian-Document-Name": item.subject + " | " + item.dateTimeCreated + ".pdf",
          "Access-Control-Allow-Origin":"https://localhost:3000"
        },
        "data": form,
        "success": function () {
          document.getElementById("item-subject").innerHTML = "<b>Status: </b> Mail Content Sent Successfully" + item.subject + "<br/>";
        }
      };

      $.ajax(settings).done(function (response) {
        console.log(response);
      });
      // doc.save('email.pdf');
    });

}