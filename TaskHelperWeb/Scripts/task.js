/// <reference path="/ts/Ews.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#createNewTasks').click(sendCreateAllTasks);
        });
    };

    //Thus function iterates over a task object array and sends a create for each one
    function sendCreateAllTasks() {
        var text = $('#tasklist')[0].value;
        var lines = text.split('\n');
        var tasks = [];
        for (var i = 0; i < lines.length; i++)
            if (lines[i].length > 0)
                tasks.push(new Ews.Task(lines[i]));

        if (tasks.length == 0) {
            $('#result').text("failed: no tasks found");
            return;
        }

        sendCreateTask(tasks);
    }

    function sendCreateTask(tasks) {
        var mailbox = Office.context.mailbox;        
        var ews = new Ews.CreateTaskRequest(tasks);
        var soap = ews.toSoap();
        mailbox.makeEwsRequestAsync(soap, callback);
    }

    function callback(asyncResult) {
        var result = asyncResult.value;
        if (result == null)
            $('#result').text("Failed");
        else {
            if (result.indexOf('ResponseClass="Success"') > -1)
                $('#result').text("Success");
            else
                $('#result').text("Failure:\n" + result);
        }
    }
})();