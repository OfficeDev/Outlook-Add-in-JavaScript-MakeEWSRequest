# Outlook add-in: Make an Exchange Web Service request from Outlook

**Table of contents**

* [Summary](#summary)
* [Prerequisites](#prerequisites)
* [Key components of the sample](#components)
* [Description of the code](#codedescription)
* [Build and debug](#build)
* [Troubleshooting](#troubleshooting)
* [Questions and comments](#questions)
* [Additional resources](#additional-resources)

<a name="summary"></a>
##Summary
The JavaScript code in this sample shows a simple request for the subject of the current email message. It demonstrates the steps required to create an Exchange Web Service (EWS) request and the best practices for making the request.

<a name="prerequisites"></a>
## Prerequisites ##

This sample requires the following:  

  - Visual Studio 2013 with Update 5 or Visual Studio 2015.  
  - A computer running Exchange 2013 with at least one email account, or an Office 365 account. You can sign up for [an Office 365 Developer subscription](http://aka.ms/ro9c62) and get an Office 365 account through it.
  - Any browser that supports ECMAScript 5.1, HTML5, and CSS3, such as Internet Explorer 9, Chrome 13, Firefox 5, Safari 5.0.6, or a later version of these browsers.
  - Familiarity with JavaScript programming and web services.

<a name="components"></a>
## Key components of the sample
The sample solution contains the following files:

- MakeEwsRequestManifest.xml: The manifest file for the Outlook add-in.
- AppRead\Home\Home.html: The HTML user interface for the mail add-in for Outlook.
- AppRead\Home\Home.js: The JavaScript file that handles requesting and using the EWS request. 

<a name="codedescription"></a>
##Description of the code

The code that creates the EWS XML request includes two methods. The first method, `getSoapEnvelope()`, wraps a SOAP envelope around a web service request. Because the SOAP envelope is standard for all EWS requests, this method can be reused to wrap any EWS request.

The second method, `getSubjectRequest()`, returns the EWS request to get the subject field of an item. The id parameter is the Exchange item identifier for the requested item. Note the following about the request:

- The `ItemShape` element is used to restrict the response to the base shape `IdOnly`. This limits the response to only the item identifier for the item and prevents excessive data from being sent back from the server. 
- The `AdditionalProperties` element is used to add the Subject field to the response. By using the `IdOnly` base shape and a list of additional properties, you can limit the size of the response from the server to just the data that your add-in requires. 

The `sendRequest()` method is called when you click the **Make EWS request** button in the add-in UI. It gets the Exchange identifier of the current item and passes it to the `getSubjectRequest()` and `getSoapEnvelope()` methods, then makes an asynchronous call to the Exchange server by using the  `makeEwsRequestAsync` method. The  `makeEwsRequestAsync` method takes two parameters: the EWS request wrapped in its SOAP envelope, and a callback method that is called when the asynchronous request to EWS is complete. You can add a third optional `userContext` parameter to the  `makeEwsRequestAsync` method if you have to provide additional information to the callback method.

The `callback()` method is called with a single parameter, `asyncResult`. The `asyncResult` object has two members:

- `value` – The contents of the response from EWS. 
- `context` – The `userContext` parameter passed to the [makeEwsRequestAsync](http://msdn.microsoft.com/library/2ec380e0-4a67-4146-92a6-6a39f65dc6f2) method. 

The `callback` method in the sample displays the contents of the response in a scrollable `div` element in the UI, but your code can use the response in more sophisticated ways.

<a name="build"></a>
## Build and debug ##
**Note**: The mail add-in will be activated on any email message in the user's Inbox. You can make it easier to test the add-in by sending one or more email messages to your test account before you run the sample add-in.

1. Open the solution in Visual Studio. Press F5 to build and deploy the sample add-in.
2. Connect to an Exchange account by providing the email address and password for an Exchange 2013 server.
3. Allow the server to configure the mail account.
4. Log on to the email account by entering the account name and password. 
5. Select a message in the Inbox.
6. Wait for the add-in bar to appear over the message.
7. In the add-in bar, click **MakeEWSRequest**.
8. When the mail add-in appears, click the **Make EWS Request** button to request the subject of the current message from the Exchange server.
9. Review the response XML returned by the request.

<a name="troubleshooting"></a>
##Troubleshooting
The following are common errors that can occur when you use Outlook Web App to test a mail add-in for Outlook:

- The add-in bar does not appear when a message is selected. If this occurs, restart the application by selecting **Debug – Stop Debugging** in the Visual Studio window, then press F5 to rebuild and deploy the add-in. 
- Changes to the JavaScript code may not be picked up when you deploy and run the add-in. If the changes are not picked up, clear the cache on the web browser by selecting **Tools – Internet options** and clicking the **Delete…** button. Delete the temporary Internet files and then restart the add-in. 

<a name="questions"></a>
##Questions and comments##

- If you have any trouble running this sample, please [log an issue](https://github.com/OfficeDev/Outlook-Add-in-Javascript-MakeEWSRequest/issues).
- Questions about Office Add-in development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Make sure that your questions or comments are tagged with [office-addins].


<a name="additional-resources"></a>
## Additional resources ##

- [Explore the EWS Managed API, EWS, and web services in Exchange](https://msdn.microsoft.com/library/office/jj536567(v=exchg.150).aspx)
- [makeEwsRequestAsync method](http://msdn.microsoft.com/library/2ec380e0-4a67-4146-92a6-6a39f65dc6f2)

## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.
