/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

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

// *********************************************************
//
// Excel-Add-in-Javascript-GetSelectionChangeCoordinates, https://github.com/OfficeDev/Excel-Add-in-Javascript-GetSelectionChangeCoordinates
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