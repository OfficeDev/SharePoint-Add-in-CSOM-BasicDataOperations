---
page_type: sample
products:
- office-sp
- office-365
languages:
- csharp
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 8/17/2015 1:40:30 PM
---
# Basic CRUD operations in SharePoint Add-ins using the client-side object model (CSOM) APIs #

## Summary
Use the SharePoint client-side object model (CSOM) to perform create, read, update, and delete operations on lists and list items from a SharePoint Add-in.

### Applies to ###
-  SharePoint Online and on-premise SharePoint 2013 and later 

----------
## Prerequisites ##
This sample requires the following:


- A SharePoint 2013 development environment that is configured for app isolation and OAuth. (A SharePoint Online Developer Site is automatically configured. For an on premise development environment, see [Set up an on-premises development environment for SharePoint Add-ins](https://msdn.microsoft.com/library/office/fp179923.aspx) and [Use an Office 365 SharePoint site to authorize provider-hosted add-ins on an on-premises SharePoint site](https://msdn.microsoft.com/library/office/dn155905.aspx).) 


- Visual Studio and the Office Developer Tools for Visual Studio installed on your developer computer 


## Description of the code ##
The code that uses the CSOM APIs is located in the Default.aspx.cs file of the SharePoint-Add-in-CSOM-BasicDataOperationsWeb project. The Default.aspx page of the add-in appears after you install and launch the add-in and looks similar to the following.

![The add-in start page with a table listing all the list on the site by name and ID.](/description/fig1.gif) 



The sample demonstrates the following:


- How to read and write data to and from the host web of a SharePoint Add-in.


- How to load the data returned from SharePoint into the client context object and then display the data. 


## To use the sample #

12. Open **Visual Studio** as an administrator.
13. Open the .sln file.
13. In **Solution Explorer**, highlight the SharePoint add-in project and replace the **Site URL** property with the URL of your SharePoint developer site.
14. Press F5.
15. After the app installs, the consent page opens. Click **Trust It**.
16. Enter a string in the text box beside the **Add List** button and click the button. In a moment, the page refreshes and the new list is in the table.
17. Click the ID of the list, and then click **Retrieve List Items**. There will initially be no items on the list. Some additional buttons will appear.
18. Add a string to the text box beside the **Add Item** button and press the button. The new item will appear in the table in the row for the list.
19. Add a string to the text box beside the **Change List Title** button and press the button. The title will change in the table.
20. Press the **Delete the List** button and the list is deleted.

**Do not delete any of the built-in SharePoint lists. If you mistakenly do so, recover the list from the SharePoint Recycle Bin.**

## Troubleshooting

<table border="0" cellspacing="5" cellpadding="5" frame="void" align="left" style="width:601px; height:212px">
<tbody>
<tr style="background-color:#a9a9a9">
<th align="left" scope="col"><strong><span style="font-size:small">Problem </span>
</strong></th>
<th align="left" scope="col"><strong><span style="font-size:small">Solution</span></strong></th>
</tr>
<tr valign="top">
<td><span style="font-size:small">Visual Studio does not open the browser after you press the F5 key.</span></td>
<td><span style="font-size:small">Set the app for SharePoint project as the startup project.</span></td>
</tr>
<tr valign="top">
<td><span style="font-size:small">HTTP error 405 <strong>Method not allowed</strong>.</span></td>
<td><span style="font-size:small">Locate the applicationhost.config file in <em>%userprofile%</em>\Documents\IISExpress\config.</span>
<p><span style="font-size:small">Locate the handler entry for <strong>StaticFile</strong>, and add the verbs
<strong>GET</strong>, <strong>HEAD</strong>, <strong>POST</strong>, <strong>DEBUG</strong>, and
<strong>TRACE</strong>.</span></p>
</td>
</tr>
</tbody>
</table>

## Questions and comments

We'd love to get your feedback on this sample. You can send your questions and suggestions to us in the [Issues](https://github.com/OfficeDev/SharePoint-Add-in-CSOM-BasicDataOperations/issues) section of this repository.
  
<a name="resources"/>
## Additional resources

[SharePoint Add-ins](https://msdn.microsoft.com/library/office/fp179930.aspx)

[Complete basic operations using SharePoint 2013 client library code](https://msdn.microsoft.com/library/office/fp179912.aspx)

### Copyright ###

Copyright (c) Microsoft. All rights reserved.






This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
