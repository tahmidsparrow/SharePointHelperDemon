using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using OfficeDevPnP.Core;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace SpListReport.Business
{
    public enum ObjectType
    {
        List = 0,
        ListItem = 1,
        Site = 2
    }

    class Permission
    {

        /// <summary>
        /// This funtion get the site/list/list item permission details. And return it by a dictonary.
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="clientObject"></param>
        /// <param name="objectType"></param>
        /// <returns>Dictionary<string, string></returns>
        public Dictionary<string, string> GetPermissionDetails(PnPClientContext clientContext, ClientObject clientObject, ObjectType objectType)
        {
            Dictionary<string, string> permisionDetails = new Dictionary<string, string>();

            if (objectType == ObjectType.List)
            {
                List list = clientObject as List;
                IQueryable<RoleAssignment> queryForList = list.RoleAssignments.Include(roleAsg => roleAsg.Member,
                                                                                                                roleAsg => roleAsg.RoleDefinitionBindings.Include(
                                                                                                                    roleDef => roleDef.Name));
                permisionDetails = GetPermissionDetails(clientContext, queryForList);
            }
            else if (objectType == ObjectType.ListItem)
            {
                ListItem item = clientObject as ListItem;
                IQueryable<RoleAssignment> queryForListItem = item.RoleAssignments.Include(roleAsg => roleAsg.Member, 
                                                                                            roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name));
                permisionDetails = GetPermissionDetails(clientContext, queryForListItem);
            }
            else if (objectType == ObjectType.Site)
            {
                IQueryable<RoleAssignment> queryForSite = clientContext.Web.RoleAssignments.Include(roleAsg => roleAsg.Member,
                                                                                                                roleAsg => roleAsg.RoleDefinitionBindings.Include(
                                                                                                                    roleDef => roleDef.Name));
                permisionDetails = GetPermissionDetails(clientContext, queryForSite);
            }
            return permisionDetails;
        }
        private Dictionary<string, string> GetPermissionDetails(PnPClientContext clientContext, IQueryable<RoleAssignment> queryString)
        {
            IEnumerable roles = clientContext.LoadQuery(queryString);
            clientContext.ExecuteQuery();

            Dictionary<string, string> permisionDetails = new Dictionary<string, string>();

            foreach (RoleAssignment ra in roles)
            {
                var rdc = ra.RoleDefinitionBindings;
                string permission = string.Empty;
                foreach (var rdbc in rdc)
                {
                    permission += rdbc.Name.ToString() + ", ";
                }
                permisionDetails.Add(ra.Member.Title, permission);
            }
            return permisionDetails;
        }

        /// <summary>
        /// This funtion set User or Group permission in Site/list/list item.
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="people"></param>
        /// <param name="clientObject"></param>
        /// <param name="roleType"></param>
        /// <param name="objectType"></param>
        /// <returns></returns>
        public int SetPermission(PnPClientContext clientContext, string people, ClientObject clientObject, RoleType roleType, ObjectType objectType)
        {
            int status = 200;
            string principleType = GetPrincipalType(clientContext, people);
            Principal principalValue = null;
            if (principleType == "SharePointGroup")
            {
                principalValue = clientContext.Web.SiteGroups.GetByName(people);
            }
            else if (principleType == "User")
            {
                principalValue = clientContext.Web.SiteUsers.GetByEmail(people);
            }
            else
            {
                status = 404;
            }
            if(objectType== ObjectType.List)
            {
                List list = clientObject as List;
                SetPermission(clientContext, roleType, principalValue, list);
            }
            else if(objectType == ObjectType.ListItem)
            {
                ListItem item = clientObject as ListItem;
                SetPermission(clientContext, roleType, principalValue, item);
            }
            return status;
        }
        private void SetPermission(PnPClientContext clientContext, RoleType roleType, Principal principalValue, ListItem item)
        {
            RoleDefinitionBindingCollection roleDefBindColl = new RoleDefinitionBindingCollection(clientContext);
            roleDefBindColl.Add(clientContext.Web.RoleDefinitions.GetByType(roleType));
            item.RoleAssignments.Add(principalValue, roleDefBindColl);
            clientContext.ExecuteQuery();
        }
        private void SetPermission(PnPClientContext clientContext, RoleType roleType, Principal principalValue, List list)
        {
            RoleDefinitionBindingCollection roleDefBindColl = new RoleDefinitionBindingCollection(clientContext);
            roleDefBindColl.Add(clientContext.Web.RoleDefinitions.GetByType(roleType));
            list.RoleAssignments.Add(principalValue, roleDefBindColl);
            clientContext.ExecuteQuery();
        }

        /// <summary>
        /// This Funtion is being used to confirm if the passed string is a User or Group of the SharePoint.
        /// </summary>
        /// <Author>Jamali</Author>
        /// <param name="clientContext"></param>
        /// <param name="emailAddress"></param>
        /// <returns>string</returns>
        private string GetPrincipalType(PnPClientContext clientContext, string emailAddress)
        {
            Principal principal = null;
            string principal_type = string.Empty;
            try
            {
                var result = Utility.ResolvePrincipal(clientContext, clientContext.Web, emailAddress, PrincipalType.All, PrincipalSource.All, null, false);
                clientContext.ExecuteQuery();

                if (result.Value.PrincipalType == PrincipalType.User)
                {
                    principal = clientContext.Web.EnsureUser(result.Value.LoginName);
                }
                else if (result.Value.PrincipalType == PrincipalType.SecurityGroup || result.Value.PrincipalType == PrincipalType.SharePointGroup)
                {
                    principal = clientContext.Web.SiteGroups.GetById(result.Value.PrincipalId);
                }
                clientContext.Load(principal);
                clientContext.ExecuteQueryRetry();
                principal_type = principal.PrincipalType.ToString();
            }
            catch (Exception ex)
            {
                principal_type = "NA";
            }
            return principal_type;
        }

        /// <summary>
        /// Thsi function is being used to check if the Object[list/listitem] has unique permission or Inherited permission.
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="clientObject"></param>
        /// <param name="objectType"></param>
        /// <returns></returns>
        public bool HasUniquePermission(PnPClientContext clientContext, ClientObject clientObject, ObjectType objectType)
        {
            bool isUnique = false;
            if (objectType == ObjectType.List)
            {
                List list = clientObject as List;
                isUnique = HasUniquePermission(clientContext, list);
            }
            else if (objectType == ObjectType.ListItem)
            {
                ListItem item = clientObject as ListItem;
                isUnique = HasUniquePermission(clientContext, item);
            }
            else if (objectType == ObjectType.Site)
            {
                ListItem item = clientObject as ListItem;
                isUnique = HasUniquePermission(clientContext);
            }
            return isUnique;
        }
        private bool HasUniquePermission(PnPClientContext clientContext, List list)
        {
            clientContext.Load(list, i => i.HasUniqueRoleAssignments);
            clientContext.ExecuteQuery();
            return list.HasUniqueRoleAssignments;
        }
        private bool HasUniquePermission(PnPClientContext clientContext)
        {
            clientContext.Load(clientContext.Web, i => i.HasUniqueRoleAssignments);
            clientContext.ExecuteQuery();
            return clientContext.Web.HasUniqueRoleAssignments;
        }
        private bool HasUniquePermission(PnPClientContext clientContext, ListItem item)
        {
            clientContext.Load(item, i => i.HasUniqueRoleAssignments);
            clientContext.ExecuteQuery();
            return item.HasUniqueRoleAssignments;
        }

        private string GetUserFieldType(PnPClientContext clientContext, FieldUserValue value)
        {
            var userInfoList = clientContext.Site.RootWeb.SiteUserInfoList;
            var userInfo = userInfoList.GetItemById(value.LookupId);
            clientContext.Load(userInfo, i => i.ContentType);
            clientContext.ExecuteQueryRetry();
            return userInfo.ContentType.Name;
        }

        /// <summary>
        /// This funtion is being used to remove all the permission of the onjectp[list/list item].
        /// But the object has to be unique permission. 
        /// Other wise the funtion will return a null value or error message.
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="clientObject"></param>
        /// <param name="objectType"></param>
        /// <returns>String</returns>
        public string RemoveAllPermission(PnPClientContext clientContext, ClientObject clientObject, ObjectType objectType)
        {
            if(objectType == ObjectType.ListItem)
            {
                ListItem item = clientObject as ListItem;
                return RemoveAllPermission(clientContext, item);
            }
            else if(objectType == ObjectType.List)
            {
                List list = clientObject as List;
                return RemoveAllPermission(clientContext, list);
            }
            else
            {
                return null;
            }
            
        }
        private string RemoveAllPermission(PnPClientContext clientContext, List list)
        {
            clientContext.Load(list, i => i.HasUniqueRoleAssignments, i => i.RoleAssignments);
            clientContext.ExecuteQueryRetry();

            if (list.HasUniqueRoleAssignments)
            {
                foreach (RoleAssignment assignment in list.RoleAssignments)
                {
                    assignment.RoleDefinitionBindings.RemoveAll();
                    assignment.Update();
                    clientContext.ExecuteQuery();
                }
                return "Success";
            }
            else
            {
                return "The obj has no unique permission";
            }
        }
        private string RemoveAllPermission(PnPClientContext clientContext, ListItem item)
        {
            clientContext.Load(item, i => i.HasUniqueRoleAssignments, i => i.RoleAssignments);
            clientContext.ExecuteQueryRetry();

            if (item.HasUniqueRoleAssignments)
            {
                foreach (RoleAssignment assignment in item.RoleAssignments)
                {
                    assignment.RoleDefinitionBindings.RemoveAll();
                    assignment.Update();
                    clientContext.ExecuteQuery();
                }
                return "Success";
            }
            else
            {
                return "The obj has no unique permission";
            }
        }

        /// <summary>
        /// This funtion is being used to remove a specific user/group's permission from a list/list item.
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="displayName"></param>
        /// <param name="clientObject"></param>
        /// <param name="objectType"></param>
        /// <returns>string</returns>
        public string RemovePermission(PnPClientContext clientContext, string displayName, ClientObject clientObject, ObjectType objectType)
        {
            string result = string.Empty;
            ///Permision can't be deleted from the Object which has inherited permission. 
            if (HasUniquePermission(clientContext, clientObject, objectType))
            {
                if (objectType == ObjectType.List)
                {
                    List list = clientObject as List;
                    IQueryable<RoleAssignment> queryString = list.RoleAssignments.Include(roleAsg => roleAsg.Member, roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name));
                    RemovePermission(clientContext, displayName, queryString);
                    result = "Success";
                }
                else if (objectType == ObjectType.ListItem)
                {
                    ListItem item = clientObject as ListItem;
                    IQueryable<RoleAssignment> queryString = item.RoleAssignments.Include(roleAsg => roleAsg.Member, roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name));
                    RemovePermission(clientContext, displayName, queryString);
                    result = "Success";
                }
            }
            else
            {
                result = "The list has inherited permission";
            }
            
            return result;
        }
        private void RemovePermission(PnPClientContext clientContext, string displayName, IQueryable<RoleAssignment> queryString)
        {
            IEnumerable roles = clientContext.LoadQuery(queryString);
            clientContext.ExecuteQuery();
            foreach (RoleAssignment ra in roles)
            {
                if (ra.Member.Title == displayName)
                {
                    ra.RoleDefinitionBindings.RemoveAll();
                    ra.DeleteObject();
                    clientContext.ExecuteQuery();
                }
            }
        }


    }
}
