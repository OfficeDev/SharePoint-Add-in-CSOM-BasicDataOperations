// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.


using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace SharePoint_Add_in_CSOM_BasicDataOperationsWeb
{
    public partial class Default : System.Web.UI.Page
    {
        SharePointContextToken contextToken;
        string accessToken;
        Uri sharepointUrl;

        protected void Page_PreInit(object sender, EventArgs e)
        {
            Uri redirectUrl;
            switch (SharePointContextProvider.CheckRedirectionStatus(Context, out redirectUrl))
            {
                case RedirectionStatus.Ok:
                    return;
                case RedirectionStatus.ShouldRedirect:
                    Response.Redirect(redirectUrl.AbsoluteUri, endResponse: true);
                    break;
                case RedirectionStatus.CanNotRedirect:
                    Response.Write("An error occurred while processing your request.");
                    Response.End();
                    break;
            }
        }

        // The Page_load method fetches the context token and the access token. The access token is used by all of the data retrieval methods.
        protected void Page_Load(object sender, EventArgs e)
        {
            string contextTokenString = TokenHelper.GetContextTokenFromRequest(Request);

            if (contextTokenString != null)
            {
                contextToken =
                    TokenHelper.ReadAndValidateContextToken(contextTokenString, Request.Url.Authority);

                sharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
                accessToken =
                    TokenHelper.GetAccessToken(contextToken, sharepointUrl.Authority).AccessToken;
                    
                // In a production add-in, you should cache the access token somewhere, such as in a database
                // or ASP.NET Session Cache. (Do not put it in a cookie.) Your code should also check to see 
                // if it is expired before using it (and use the refresh token to get a new one when needed). 
                // For more information, see the MSDN topic at https://msdn.microsoft.com/library/office/dn762763.aspx
                // For simplicity, this sample does not follow these practices.                     
                AddListButton.CommandArgument = accessToken;
                RefreshListButton.CommandArgument = accessToken;
                RetrieveListButton.CommandArgument = accessToken;
                AddItemButton.CommandArgument = accessToken;
                DeleteListButton.CommandArgument = accessToken;
                ChangeListTitleButton.CommandArgument = accessToken;
                RetrieveLists(accessToken);

            }
            else if (!IsPostBack)
            {
                Response.Write("Could not find a context token.");
            }
        }

        //This method retrieves all of the lists on the host Web.
        private void RetrieveLists(string accessToken)
        {
            if (IsPostBack)
            {
                sharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
            }

            AddItemButton.Visible = false;
            AddListItemBox.Visible = false;
            RetrieveListNameBox.Enabled = true;
            DeleteListButton.Visible = false;
            ChangeListTitleButton.Visible = false;
            ChangeListTitleBox.Visible = false;
            ListTable.Rows[0].Cells[1].Text = "List ID";

            //Execute a request for all of the site's lists.
            ClientContext clientContext =
            TokenHelper.GetClientContextWithAccessToken(sharepointUrl.ToString(), accessToken);
            Web web = clientContext.Web;
            ListCollection lists = web.Lists;
            clientContext.Load<ListCollection>(lists);
            clientContext.ExecuteQuery();

            foreach (List list in lists)
            {
                TableRow tableRow = new TableRow();
                TableCell tableCell1 = new TableCell();
                tableCell1.Controls.Add(new LiteralControl(list.Title));
                LiteralControl idClick = new LiteralControl();
                //Use Javascript to populate the RetrieveListNameBox control with the list id.
                string clickScript = "<a onclick=\"document.getElementById(\'RetrieveListNameBox\').value = '" + list.Id.ToString() + "';\" href=\"#\">" + list.Id.ToString() + "</a>";

                idClick.Text = clickScript;
                TableCell tableCell2 = new TableCell();
                tableCell2.Controls.Add(idClick);
                tableRow.Cells.Add(tableCell1);
                tableRow.Cells.Add(tableCell2);
                ListTable.Rows.Add(tableRow);
            }



        }

        //This method retrieves all items from a specified list.
        private void RetrieveListItems(string accessToken, Guid listId)
        {
            if (IsPostBack)
            {
                sharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
            }

            //Adjust the visibility of controls on the page in light of the list-specific context.
            AddItemButton.Visible = true;
            AddListItemBox.Visible = true;
            RetrieveListNameBox.Enabled = false;
            DeleteListButton.Visible = true;
            ChangeListTitleButton.Visible = true;
            ChangeListTitleBox.Visible = true;
            ListTable.Rows[0].Cells[1].Text = "List Items";

            //Execute a request to get the first 100 of the list's items.
            ClientContext clientContext = TokenHelper.GetClientContextWithAccessToken(sharepointUrl.ToString(), accessToken);
            Web web = clientContext.Web;
            ListCollection lists = web.Lists;
            List selectedList = lists.GetById(listId);

            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = "<View><RowLimit>100</RowLimit></View>";

            //Use the fully qualified name to disambiguate the ListItemCollection type.
            Microsoft.SharePoint.Client.ListItemCollection listItems = selectedList.GetItems(camlQuery);
            clientContext.Load<ListCollection>(lists);
            clientContext.Load<List>(selectedList);
            clientContext.Load<Microsoft.SharePoint.Client.ListItemCollection>(listItems);

            clientContext.ExecuteQuery();

            TableRow tableRow = new TableRow();
            TableCell tableCell1 = new TableCell();
            tableCell1.Controls.Add(new LiteralControl(selectedList.Title));
            TableCell tableCell2 = new TableCell();

            foreach (Microsoft.SharePoint.Client.ListItem item in listItems)
            {
                tableCell2.Text += item.FieldValues["Title"] + "<br>";
            }

            tableRow.Cells.Add(tableCell1);
            tableRow.Cells.Add(tableCell2);
            ListTable.Rows.Add(tableRow);


        }


        //This method adds a list with the specified title.
        private void AddList(string accessToken, string newListName)
        {
            if (IsPostBack)
            {
                sharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
            }

            //Execute a request to add a list that has the user-supplied name.
            ClientContext clientContext = TokenHelper.GetClientContextWithAccessToken(sharepointUrl.ToString(), accessToken);
            Web web = clientContext.Web;
            ListCollection lists = web.Lists;
            ListCreationInformation listCreationInfo = new ListCreationInformation();
            listCreationInfo.Title = newListName;
            listCreationInfo.TemplateType = (int)ListTemplateType.GenericList;
            lists.Add(listCreationInfo);
            clientContext.Load<ListCollection>(lists);
            try
            {
                clientContext.ExecuteQuery();
            }
            catch (Exception e)
            {
                AddListNameBox.Text = e.Message;
            }
            RetrieveLists(accessToken);
        }

        //This method adds a list item to the specified list.
        private void AddListItem(string accessToken, Guid listId, string newItemName)
        {
            if (IsPostBack)
            {
                sharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
            }

            //Execute a request to add a list item.
            ClientContext clientContext = TokenHelper.GetClientContextWithAccessToken(sharepointUrl.ToString(), accessToken);
            Web web = clientContext.Web;
            ListCollection lists = web.Lists;
            List selectedList = lists.GetById(listId);
            clientContext.Load<ListCollection>(lists);
            clientContext.Load<List>(selectedList);
            ListItemCreationInformation listItemCreationInfo = new ListItemCreationInformation();
            var listItem = selectedList.AddItem(listItemCreationInfo);
            listItem["Title"] = newItemName;
            listItem.Update();
            clientContext.ExecuteQuery();
            RetrieveListItems(accessToken, listId);



        }

        private void ChangeListTitle(string accessToken, Guid listId, string newListTitle)
        {
            if (IsPostBack)
            {
                sharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
            }

            //Execute a request to change the title of the specified list.
            ClientContext clientContext = TokenHelper.GetClientContextWithAccessToken(sharepointUrl.ToString(), accessToken);
            Web web = clientContext.Web;
            ListCollection lists = web.Lists;
            List selectedList = lists.GetById(listId);
            clientContext.Load<ListCollection>(lists);
            clientContext.Load<List>(selectedList);
            selectedList.Title = newListTitle;
            selectedList.Update();
            clientContext.ExecuteQuery();
            RetrieveListItems(accessToken, listId);

        }

        private void DeleteList(string accessToken, Guid listId)
        {
            if (IsPostBack)
            {
                sharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
            }

            //Execute a request to delete the specified list.
            ClientContext clientContext = TokenHelper.GetClientContextWithAccessToken(sharepointUrl.ToString(), accessToken);
            Web web = clientContext.Web;
            ListCollection lists = web.Lists;
            List selectedList = lists.GetById(listId);
            clientContext.Load<ListCollection>(lists);
            clientContext.Load<List>(selectedList);
            selectedList.DeleteObject();
            clientContext.ExecuteQuery();
            RetrieveListNameBox.Text = "";
            RetrieveLists(accessToken);

        }

        protected void AddList_Click(object sender, EventArgs e)
        {

            string commandAccessToken = ((Button)sender).CommandArgument;
            if (AddListNameBox.Text != "")
            {
                AddList(commandAccessToken, AddListNameBox.Text);
            }
            else
            {
                AddListNameBox.Text = "Enter a list title";
            }
        }

        protected void RefreshList_Click(object sender, EventArgs e)
        {

            string commandAccessToken = ((Button)sender).CommandArgument;
            RetrieveLists(commandAccessToken);
        }

        protected void RetrieveListButton_Click(object sender, EventArgs e)
        {
            string commandAccessToken = ((Button)sender).CommandArgument;

            Guid listId = new Guid();
            if (Guid.TryParse(RetrieveListNameBox.Text, out listId))
            {
                RetrieveListItems(commandAccessToken, listId);
            }
            else
            {
                RetrieveListNameBox.Text = "Enter a List GUID";
            }
        }

        protected void AddItemButton_Click(object sender, EventArgs e)
        {
            string commandAccessToken = ((Button)sender).CommandArgument;
            Guid listId = new Guid(RetrieveListNameBox.Text);
            if (AddListItemBox.Text != "")
            {
                AddListItem(commandAccessToken, listId, AddListItemBox.Text);
            }
            else
            {
                AddListItemBox.Text = "Enter an item title";
            }
        }

        protected void DeleteListButton_Click(object sender, EventArgs e)
        {
            string commandAccessToken = ((Button)sender).CommandArgument;
            Guid listId = new Guid(RetrieveListNameBox.Text);
            DeleteList(commandAccessToken, listId);
        }

        protected void ChangeListTitleButton_Click(object sender, EventArgs e)
        {
            string commandAccessToken = ((Button)sender).CommandArgument;
            Guid listId = new Guid(RetrieveListNameBox.Text);
            if (ChangeListTitleBox.Text != null)
            {
                ChangeListTitle(commandAccessToken, listId, ChangeListTitleBox.Text);
            }
            else
            {
                ChangeListTitleBox.Text = "Enter a new list title";
            }
        }
    }
}

/*
SharePoint-Add-in-CSOM-BasicDataOperations, http://github/officedev/SharePoint-Add-in-CSOM-BasicDataOperations
 
Copyright (c) Microsoft Corporation
All rights reserved. 
 
MIT License:
Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:
 
The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.
 
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.    
  
*/