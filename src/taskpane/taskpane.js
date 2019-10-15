/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
// var emlformat = require('../../node_modules/jspdf/dist/jspdf.min.js');
import * as jsPDF from 'jspdf';
import { isNull } from 'util';
var utilities = require("./utilities");
var selectedCaseId;

function setSelectedCaseId(event) {
  selectedCaseId = event.target.id;
  console.log(selectedCaseId);
  $('#runSendMailContent').removeAttr('disabled');
}

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    $('#filters').multiselect(utilities.dropdownSettings);
    $('#filters').multiselect('loadOptions', utilities.dropdownOptions);
    $('#searchCase').submit(fetchCaseDetails);
    new fabric['Spinner'](document.getElementById('loader'));
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
      form.append("data", doc.output('blob'));
      var settings = {
        "url": "https://appiandev.vuram.com/suite/webapi/uploadDocument",
        "method": "POST",
        "processData": false,
        "headers": {
          "Appian-API-Key": apiKey,
          "Appian-Document-Name": item.subject + " | " + item.dateTimeCreated + ".pdf",
          "Access-Control-Allow-Origin": "*"
        },
        "data": form,
        "success": function () {
          document.getElementById("item-subject").innerHTML = "<b>Status: </b> Mail Content Sent Successfully.<br/>";
        }
      };

      $.ajax(settings).done(function (response) {
        console.log(response);
      });
    });

}

export async function fetchCaseDetails(event) {
  event.preventDefault();
  $('#runSendMailContent').attr('disabled',true);
  $('#searchWarning').html("");
  $('#banner').hide();
  $('#caseTable').hide();
  var filters = $('#filters').val();
  var searchText = $('#searchText').val();
  if (filters.length == 0) {
    $('#searchWarning').html("Please select atleast one from the search by filter");
    return;
  }
  else if (isNull(searchText) || searchText.trim() === "") {
    $('#searchWarning').html("Search text is empty. Please enter search text to continue.");
    return;
  }
  else {
    $('#loader').show();
    utilities.http_get("getCaseDetails", {
      searchText: searchText,
      searchFields: filters.join('#'),
      startIndex: 1,
      batchSize: 5
    },
      (response) => {
        $('#loader').hide();
        if (isNull(response) || response.length == 0) {
          $('#banner').show();
        }
        else {
          var tableRow = response.map(item => {
            console.log(item);
            return '<tr><td><input type="radio" id=' + item.caseId + ' name="optradio"/></td><td><label for=' + item.caseId + '>'
              + item.caseIdFormat + '</label></td><td><label for=' + item.caseId + '>' + item.subject + '</label></td></tr>';
          });
          $('#caseTableBody').html(tableRow);
          $('input[name="optradio"]').click(setSelectedCaseId);
          $('#caseTable').show();
        }
      });
  }
}