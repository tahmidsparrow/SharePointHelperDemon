using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System.Collections.Generic;
using System.Linq;

namespace SpListReport.Business
{
    class ListLibrary
    {
        internal ListCollection GetCustomeListLibrary(PnPClientContext clientContext)
        {
            ListCollection spListCollction = clientContext.Web.Lists;
            clientContext.Load(spListCollction, lists => lists.Where(list => list.Hidden == false && (list.BaseTemplate == 100 || list.BaseTemplate == 101)
                                                                                                    && list.AllowDeletion == true && list.IsCatalog == false));
            clientContext.ExecuteQueryRetry();
            return spListCollction;
        }

        internal ListCollection GetCustomeList(PnPClientContext clientContext)
        {
            ListCollection spListCollction = clientContext.Web.Lists;
            clientContext.Load(spListCollction, lists => lists.Where(list => list.Hidden == false && list.BaseTemplate == 100));
            clientContext.ExecuteQueryRetry();
            return spListCollction;
        }

        internal ListCollection GetCustomeLibrary(PnPClientContext clientContext)
        {
            ListCollection spListCollction = clientContext.Web.Lists;
            clientContext.Load(spListCollction, lists => lists.Where(list => list.Hidden == false && list.BaseTemplate == 101
                                                                                                  && list.AllowDeletion == true && list.IsCatalog == false));
            clientContext.ExecuteQueryRetry();
            return spListCollction;
        }

        internal List<string> GetListView(PnPClientContext clientContext, string listName)
        {
            List spList = clientContext.Web.Lists.GetByTitle(listName);
            ViewCollection listViewCollection = spList.Views;
            clientContext.Load(listViewCollection);
            clientContext.ExecuteQueryRetry();

            List<string> listViews = new List<string>();
            foreach (View view in listViewCollection)
            {
                if(view.Title != "Relink Documents" && view.Title != "assetLibTemp" && view.Title != "Merge Documents" && view.Title != "RssView")
                {
                    listViews.Add(view.Title);
                }
            }

            return listViews;
        }

    }
}
