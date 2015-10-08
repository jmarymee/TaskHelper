/// <reference path="../App.js" />


(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

//            displayItemDetails();
            //displayTasks();
            //displaySubject();
            sendGetAllTasks();
            //sendCreateTask();

            $('#createNewTasks').click(getListOfTasksToCreate);
        });
    };

    //This test method creates an array of strings (tasks) and sends to the createAllMethod
    function getListOfTasksToCreate() {
        var arrayofTasks = ['task1', 'task2', 'task3','task4','task5','task6'];
        debugger;
        sendCreateTask(arrayofTasks);
    }

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

    function getProperties() {
        return '        <t:AdditionalProperties>' +
        '          <t:FieldURI FieldURI="item:Subject" />' +
//        '          <t:FieldURI FieldURI="item:Priority" />' +
        '          <t:FieldURI FieldURI="item:Status" />' +
        '        </t:AdditionalProperties>';
    }

    function getAllTasks() {
        var result = getBodyPrefix() +
            '    <m:FindItem Traversal="Shallow">' +
            '      <m:ItemShape>' +
            '        <t:BaseShape>AllProperties</t:BaseShape>' +
//                    getProperties() +
            '      </m:ItemShape>' +
                    getRestriction() +
            '      <m:ParentFolderIds>' +
            '        <t:DistinguishedFolderId Id="tasks"/>' +
            '      </m:ParentFolderIds>' +
            '    </m:FindItem>' +
            getBodyPostfix();

        return result;
    }

    function getRestriction() {
        return '      <m:Restriction>' +
//        '        <t:And>' +
        '          <t:IsNotEqualTo>' +
        '            <t:FieldURI FieldURI="task:Status" />' +
        '            <t:FieldURIOrConstant>' +
        '              <t:Constant Value="2" />' +
        '            </t:FieldURIOrConstant>' +
        '          </t:IsNotEqualTo>' +
//        '        </t:And>' +
        '      </m:Restriction>';
    }

    function addSort() {
        return
            '      <m:SortOrder>' +
            '        <t:FieldOrder Order="Descending">' +
            '          <t:FieldURI FieldURI="item:Priority" />' +
            '        </t:FieldOrder>' +
            '      </m:SortOrder>';
    }

    function getBodyPrefix() {
        return '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
            '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
            '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
            '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
            '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '  <soap:Header>' +
            '    <t:RequestServerVersion Version="Exchange2013" soap:mustUnderstand="0" />' +
            '  </soap:Header>' +
            '  <soap:Body>';
    }

    function getBodyPostfix() {
        return '  </soap:Body>' +
            '</soap:Envelope>';
    }

<<<<<<< HEAD
    //This method creates a single task embedded in a SOAP string (EWS call)
    function getCreateTask(taskObject) {
        //sample date
        //2006-10-26T21:32:52
        //sample object = { taskName: value, startdatString: startDate }
=======


    function getCreateTask() {
>>>>>>> upstream/master
        var result = getBodyPrefix() +
            '    <m:CreateItem MessageDisposition="SaveOnly">' +
            '      <m:Items>' +
<<<<<<< HEAD
            '        <t:Task>' +
            '          <t:Subject>'+taskObject.taskName+'</t:Subject>' +
            '          <t:DueDate>'+taskObject.startDateString+'</t:DueDate>' +
            '          <t:Status>NotStarted</t:Status>' +
            '        </t:Task>' +
=======
                    getTask("Test EWS TaskHelper", "NotStarted") +
>>>>>>> upstream/master
            '      </m:Items>' +
            '    </m:CreateItem>' +
            getBodyPostfix();

        return result;
    }

<<<<<<< HEAD
    //This method creates an array of tasks using a single web call. The array of tasks is formatted as list of tasks in XML format
    function getCreateTaskList(arrayOfTasks) {
        //sample date
        //2006-10-26T21:32:52
        var itemsList = appendListOfTasks(arrayOfTasks);

        var result = getBodyPrefix() +
            '    <m:CreateItem>' +
            '      <m:Items>' +
                    itemsList +
            '      </m:Items>' +
            '    </m:CreateItem>' +
            getBodyPostfix();

        return result;
    }

    //This function takes an array of tasks and creates a soap-formatted list of new tasks to create
    //It needs a SOAP header and tail added before sending the web call
    function appendListOfTasks(arrayOfTasks) {
        var soaptaskList = '';
        var dateString = '2015-10-08T21:32:52';

        for (var loop = 0; loop < arrayOfTasks.length; loop++) {
            soaptaskList = soaptaskList +
            '        <t:Task>' +
            '          <t:Subject>' + arrayOfTasks[loop] + '</t:Subject>' +
            '          <t:DueDate>' + dateString + '</t:DueDate>' +
            '          <t:Status>NotStarted</t:Status>' +
            '        </t:Task>'
        }
        return soaptaskList;
    }
=======

    //        <t:Task>' +
    //          <t:Subject>Test EWS TaskHelper</t:Subject>' +
    //          <t:DueDate>2006-10-26T21:32:52</t:DueDate>' +
    //          <t:Status>NotStarted</t:Status>' +
    //        </t:Task>' +

    function getTask(subject, status) {
        return  '        <t:Task>' +
                '          <t:Subject>' + subject + '</t:Subject>' +
                '          <t:Status>'+ status +'</t:Status>' +
                '        </t:Task>';
    }

>>>>>>> upstream/master

    function getSubjectRequest(id) {
        // Return a GetItem operation request for the subject of the specified item. 
        var result = getBodyPrefix() +
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
             getBodyPostfix();

        return result;
    }

    //This is the generalized call to EWS for task create
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


    //Receives an array of tasks, formats the SOAP header and then makes a single call to create all of the tasks
    function sendCreateTask(arrayOfTasks) {
        var mailbox = Office.context.mailbox;
        var soap = getCreateTaskList(arrayOfTasks);
        mailbox.makeEwsRequestAsync(soap, callback);
    }

    function callback(asyncResult) {
        var result = asyncResult.value;
        var context = asyncResult.context;
        // Process the returned response here.
        
        //Brad's diagnostic call
        //$('#tasks').text("EWS URL: " + Office.context.mailbox.ewsUrl + "\n" + result);
    }
})();