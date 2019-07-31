# [ARCHIVED] Excel add-in: Get selection-change coordinates in an Excel table

**Note:** This repo is archived and no longer actively maintained. Security vulnerabilities may exist in the project, or its dependencies. If you plan to reuse or run any code from this repo, be sure to perform appropriate security checks on the code or dependencies first. Do not use this project as the starting point of a production Office Add-in. Always start your production code by using the Office/SharePoint development workload in Visual Studio, or the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), and follow security best practices as you develop the add-in. 

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
This sample add-in shows how to detect when the selection changes in a table (matrix) in Excel 2013, and then how to display the table columns and rows included in the selection.

<a name="prerequisites"></a>
## Prerequisites ##

This sample requires the following:  

  - Visual Studio 2013 with Update 5 or Visual Studio 2015.  
  - Excel 2013
  - Internet Explorer 9 or later, which must be installed but doesn't have to be the default browser. To support Office Add-ins, the Office client that acts as host uses browser components that are part of Internet Explorer 9 or later.
  - One of the following as the default browser: Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13, or a later version of one of these browsers.
  - Familiarity with JavaScript programming and web services.

<a name="components"></a>
## Key components of the sample
The sample solution contains the following files:

- DisplayExcelSelectionCoordinatesManifest.xml: The manifest file for the Excel add-in.
- App\Home\Home.html: The HTML user interface for the Excel add-in.
- App\Home\Home.js. 

<a name="codedescription"></a>
##Description of the code

This add-in opens a blank Excel 2013 file. The user must first enter values in a contiguous rectangular collection of cells, thereby creating a table (matrix); enter values in the table; and then select the table. When the user then chooses **Set Binding**, the app binds to that table.

Choosing **Set Binding** executes the `bindNamedItem` function, in the Home.js file. This function uses the [addFromSelectionAsync](http://msdn.microsoft.com/library/office/apps/fp142282(v=office.15)) method of the [Bindings](http://msdn.microsoft.com/library/fp160966(v=office.15)) object of the JavaScript API for Office to create a new binding of coercion type "matrix." The function then calls the `addEventHandler` function, which uses the [getByIdAsync](http://msdn.microsoft.com/library/fp161008(v=office.15)) method of the [Bindings](http://msdn.microsoft.com/library/fp160966(v=office.15)) object to identify the binding by its ID ("myMatrix") and add a handler for the [bindingSelectionChanged](http://msdn.microsoft.com/library/fp161088(v=office.15)) event of the **Binding** object to the binding. Then, when the user makes a new selection in the table, the event handler uses the [startRow](http://msdn.microsoft.com/library/fp179809), [startColumn](http://msdn.microsoft.com/library/fp179837), [columnCount](http://msdn.microsoft.com/library/fp179813), and [rowCount](http://msdn.microsoft.com/library/fp179805) properties of the [BindingSelectionChangedEventArgs](http://msdn.microsoft.com/library/9b879ce5-e59c-4059-b488-c51eddfdca5b) object to display information about the new selection in the task pane.

<a name="build"></a>
## Build and debug ##

1. Open the solution in Visual Studio. Press F5 to build and deploy the sample add-in and open it in Excel 2013.
2. Create a table by entering values in several contiguous cells occupying at least three rows and three columns.
3. Select the entire table.
4. In the add-in task pane, choose **Set Binding**.
5. Select one or more contiguous cells in the table.

<a name="troubleshooting"></a>
##Troubleshooting

If the app fails to respond as described, try reloading it. (In the task pane, choose the down arrow, and then choose the Reload button.)

<a name="questions"></a>
##Questions and comments##

- If you have any trouble running this sample, please [log an issue](https://github.com/OfficeDev/Excel-Add-in-Javascript-GetSelectionChangeCoordinates/issues).
- Questions about Office Add-in development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Make sure that your questions or comments are tagged with [office-addins].


<a name="additional-resources"></a>
## Additional resources ##

- [More Add-in samples](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Javascript API for Office](http://msdn.microsoft.com/library/fp142185(office.15).aspx)


## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.
