/// <reference path="../App.js" />

(function () {
    "use strict";

    var bindingName = 'myMatrix';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#bind-selection').click(bindNamedItem);
        });
    };

    // Bind to the named range
    function bindNamedItem() {
        Office.context.document.bindings.addFromSelectionAsync(Office.CoercionType.Matrix,
            { id: bindingName }, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    var message = 'Added new binding. <br />' +
                        'Type: ' + asyncResult.value.type + '; ID: ' + asyncResult.value.id + '.';
                    write(message);
                    addEventHandler();
                }
                else {
                    app.showNotification('Error:', asynchResult.error.message);
                }
            });
    }

    // Add a handler for the SelectionChanged event.
    function addEventHandler() {
        Office.context.document.bindings.getByIdAsync(bindingName, function (result) {
            result.value.addHandlerAsync(Office.EventType.BindingSelectionChanged, myHandler);
        });
    }

    // Display the newly selected row and column indexes.
    function myHandler(bArgs) {
        var message = 'You have selected the following portion of ' + bArgs.binding.id + ':' + '<br />' +
            'Row indexes: from ' + bArgs.startRow + ' to ' + (bArgs.startRow + bArgs.rowCount - 1) +
             '<br />' + 'Column indexes: from ' + bArgs.startColumn + ' to ' + (bArgs.startColumn + bArgs.columnCount - 1) +
            '<br />' + 'Note: Indexes are zero-based.';
        write(message);
    }

    // Write to the UI, and scroll to the bottom of the message:
    function write(message) {
        $('#message').append('<hr />' + message);
        $('#content-main').scrollTop($('#content-main')[0].scrollHeight);
    };

})();