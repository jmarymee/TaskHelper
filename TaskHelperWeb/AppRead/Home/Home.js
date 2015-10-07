/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            displayItemDetails();
            //displayTasks();
            //displaySubject();
            //sendGetAllTasks();
            sendCreateTask();
        });
    };

    function displayTasks() {
        var tasks = Office.context.mailbox.item.getEntitiesByType(Office.MailboxEnums.EntityType.TaskSuggestion);

        var s = "count=" + tasks.length + "\n";
        for (t in tasks) {
            s += t + "\n"
        }

        $('#tasks').text(s);
    }

    function displaySubject() {
        sendRequest();
    }

    // Displays the "Subject" and "From" fields, based on the current mail item
    function displayItemDetails() {
        var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
        $('#subject').text(item.subject);

        var from;
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            from = Office.cast.item.toMessageRead(item).from;
        } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            from = Office.cast.item.toAppointmentRead(item).organizer;
        }

        if (from) {
            $('#from').text(from.displayName);
            $('#from').click(function () {
                app.showNotification(from.displayName, from.emailAddress);
            });
        }
    }


    function getAllTasks() {
        var result = 
            '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
            'xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
            'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
            'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '  <soap:Body>' +
            '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
            '    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '      <ItemShape>' +
            '        <t:BaseShape>AllProperties</t:BaseShape>' +
            '      </ItemShape>' +
            '    </GetItem>' +
            '  </soap:Body>' +
            '</soap:Envelope>'
        return result;
    }

    function getCreateTask() {
        var result =
            '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
            '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
            '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
            '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '  <soap:Body>' +
            '    <CreateItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
            '                xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
            '                MessageDisposition="SaveOnly">' +
            '      <Items>' +
            '        <t:Task>' +
            '          <t:Subject>Test EWS TaskHelper</t:Subject>' +
            '          <t:DueDate>2006-10-26T21:32:52</t:DueDate>' +
            '          <t:Status>NotStarted</t:Status>' +
            '        </t:Task>' +
            '      </Items>' +
            '    </CreateItem>' +
            '  </soap:Body>' +
            '</soap:Envelope>';

        return result;
    }

    function getSubjectRequest(id) {
        // Return a GetItem operation request for the subject of the specified item. 
        var result =
     '<?xml version="1.0" encoding="utf-8"?>' +
     '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
     '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
     '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
     '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
     '  <soap:Header>' +
     '    <t:RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
     '  </soap:Header>' +
     '  <soap:Body>' +
     '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
     '      <ItemShape>' +
     '        <t:BaseShape>IdOnly</t:BaseShape>' +
     '        <t:AdditionalProperties>' +
     '            <t:FieldURI FieldURI="item:Subject"/>' +
     '        </t:AdditionalProperties>' +
     '      </ItemShape>' +
     '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
     '    </GetItem>' +
     '  </soap:Body>' +
     '</soap:Envelope>';

        return result;
    }

    function sendRequest() {
        // Create a local variable that contains the mailbox.
        var mailbox = Office.context.mailbox;
        var id = mailbox.item.itemId;
        var soap = getSubjectRequest(id);

        mailbox.makeEwsRequestAsync(soap, callback);
    }

    function sendGetAllTasks() {
        var mailbox = Office.context.mailbox;
        var soap = getAllTasks();

        mailbox.makeEwsRequestAsync(soap, callback);
    }


    function sendCreateTask() {
        var mailbox = Office.context.mailbox;
        var soap = getCreateTask();

        mailbox.makeEwsRequestAsync(soap, callback);
    }

    function callback(asyncResult) {
        var result = asyncResult.value;
        var context = asyncResult.context;

        // Process the returned response here.
        
        Office.context.mailbox.g

        $('#tasks').text("EWS URL: " + Office.context.mailbox.ewsUrl + "\n" + result);
    }





})();