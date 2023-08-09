using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using SpListReport.Business;

namespace SpListReport
{
    class ListDetailsReport
    {
        internal void MainProcess()
        {
            PnPClientContext clientContext = GetClientContext();

            List list1 = clientContext.Web.Lists.GetByTitle("Test2");
            clientContext.ExecuteQuery();
            //new Permission().SetPermission(clientContext, "RndClassic Members", list1, RoleType.Contributor, ObjectType.List);

            ListLibrary listLibrary = new ListLibrary();
            ListCollection spListCollction = listLibrary.GetCustomeListLibrary(clientContext);

            string path = ConfigurationManager.AppSettings["ReportPath"];
            FileStream fileStream = System.IO.File.Create($@"{path}\report.txt");
            MemoryStream reportStream = new MemoryStream();
            StreamWriter report = new StreamWriter(reportStream);
            try
            {
                Permission permissionDetails = new Permission();

                //IQueryable<RoleAssignment> queryForSitePermission = clientContext.Web.RoleAssignments.Include(roleAsg => roleAsg.Member,
                //                                                                                                roleAsg => roleAsg.RoleDefinitionBindings.Include(
                //                                                                                                    roleDef => roleDef.Name));               
                //Dictionary<string, string> sitePermissionCollection = permissionDetails.GetPermissionDetails(clientContext, queryForSitePermission);
                Dictionary<string, string> sitePermissionCollection = permissionDetails.GetPermissionDetails(clientContext, clientContext.Web, ObjectType.Site);
                foreach (var sitePermission in sitePermissionCollection)
                {
                    Console.WriteLine(sitePermission.Key + "   " + sitePermission.Value);
                    report.WriteLine(sitePermission.Key + "   " + sitePermission.Value);
                }

                foreach (List list in spListCollction)
                {
                    Console.WriteLine("");
                    report.WriteLine();
                    Console.WriteLine("List Name: " + list.Title);
                    report.WriteLine("List Name: " + list.Title);
                    
                    clientContext.Load(list, l=>l.HasUniqueRoleAssignments, l=>l.NoCrawl, l=>l.WorkflowAssociations);
                    clientContext.ExecuteQuery();

                    //IQueryable<RoleAssignment> queryForList = list.RoleAssignments.Include(roleAsg => roleAsg.Member, roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name));
                    //Dictionary<string, string> permission = permissionDetails.GetPermissionDetails(clientContext, queryForList);
                    Dictionary<string, string> permission = permissionDetails.GetPermissionDetails(clientContext, clientContext.Web, ObjectType.List);
                    foreach (var item in permission)
                    {
                        Console.WriteLine(item.Key + "   " + item.Value);
                        report.WriteLine(item.Key + "   " + item.Value);
                    }

                    Console.WriteLine("Views Name are given bellow: ");
                    report.WriteLine("Views Name are given bellow: ");

                    List<string> listViews = listLibrary.GetListView(clientContext, list.Title);
                    foreach (var item in listViews)
                    {
                        Console.Write(item + ",");
                        report.Write(item + ", ");
                    }
                    report.WriteLine();

                    CamlQuery camlQuery = new CamlQuery();
                    camlQuery.ViewXml = "<View></View>";
                    ListItemCollection listItems = list.GetItems(camlQuery);
                    clientContext.Load(listItems);
                    clientContext.ExecuteQuery();

                    foreach (ListItem item in listItems)
                    {
                        report.WriteLine("Item Detils: ");
                        report.WriteLine("Item ID: " + item.Id);

                        clientContext.Load(item, i => i.HasUniqueRoleAssignments);
                        clientContext.ExecuteQuery();


                        //IQueryable<RoleAssignment> queryForListItem = item.RoleAssignments.Include(roleAsg => roleAsg.Member, roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name));
                        //Dictionary<string, string> itemPermissionCollection = permissionDetails.GetPermissionDetails(clientContext, queryForListItem);
                        Dictionary<string, string> itemPermissionCollection = permissionDetails.GetPermissionDetails(clientContext, clientContext.Web, ObjectType.ListItem);
                        foreach (var itemPermission in itemPermissionCollection)
                        {
                            Console.WriteLine(itemPermission.Key + "   " + itemPermission.Value);
                            report.WriteLine(itemPermission.Key + "   " + itemPermission.Value);
                        }
                    }

                }

                report.Flush();
                reportStream.Position = 0;
                reportStream.CopyTo(fileStream);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                report.Dispose();
                reportStream.Dispose();
                fileStream.Dispose();
            }
            clientContext.Dispose();
        }

        private PnPClientContext GetClientContext()
        {
            AuthenticationManager authManager = new AuthenticationManager();
            string siteURL = ConfigurationManager.AppSettings["SiteAddress"];
            string userName = ConfigurationManager.AppSettings["UserName"];
            string passWord = ConfigurationManager.AppSettings["PassWord"];
            ClientContext _clientContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(siteURL, userName, passWord);
            PnPClientContext clientContext = PnPClientContext.ConvertFrom(_clientContext);
            return clientContext;
        }
    }
}
