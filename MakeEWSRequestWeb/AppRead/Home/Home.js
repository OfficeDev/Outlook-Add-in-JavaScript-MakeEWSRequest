/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            var item = Office.context.mailbox.item;
            $('#subject').text(item.normalizedSubject);
            $('#from').text(item.from.displayName);
            $('#to').text(item.to[0].displayName);
        });
    };

})();

function getSoapEnvelope(request) {
    // Wrap an Exchange Web Services request in a SOAP envelope. 
    var result =

    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '  <t:RequestServerVersion Version="Exchange2013"/>' +
    '  </soap:Header>' +
    '  <soap:Body>' +

    request +

    '  </soap:Body>' +
    '</soap:Envelope>';

    return result;
};

function getSubjectRequest(id) {
    // Return a GetItem EWS operation request for the subject of the specified item.  
    var result =

 '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
 '      <ItemShape>' +
 '        <t:BaseShape>IdOnly</t:BaseShape>' +
 '        <t:AdditionalProperties>' +
 '            <t:FieldURI FieldURI="item:Subject"/>' +
 '        </t:AdditionalProperties>' +
 '      </ItemShape>' +
 '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
 '    </GetItem>';

    return result;
};

// Send an EWS request for the message's subject. 
function sendRequest() {
    // Create a local variable that contains the mailbox. 
    var mailbox = Office.context.mailbox;
    var request = getSubjectRequest(mailbox.item.itemId);
    var envelope = getSoapEnvelope(request);

    mailbox.makeEwsRequestAsync(envelope, callback);
};

// Function called when the EWS request is complete. 
function callback(asyncResult) {
    var response = asyncResult.value;
    var context = asyncResult.context;

    // Process the returned response here. 
    var responseSpan = document.getElementById("response");
    responseSpan.innerText = response;
};


// *********************************************************
//
// Outlook-Add-in-Javascript-MakeEWSRequest, https://github.com/OfficeDev/Outlook-Add-in-Javascript-MakeEWSRequest
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************
