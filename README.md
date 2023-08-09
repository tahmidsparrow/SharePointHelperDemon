# SharePointHelperDemon

To access a SharePoint Site, SharePoint List, or a particular SharePoint list item the user or the app has to have the proper permission. And these permissions can be managed from the SharePoint site's provided view or programmatically. So, to get a detailed permission report on a site/list/list item, the below code snippet can be very handy.
 
## Get Site Permission
 
In the following code, this function is getting the ClientContext as a parameter and returns a dictionary type as Dictionary<string, string>. From the RoleAssignments property, the necessary permission details can be found. Here once the permission details are loaded, then the result has been iterated for permission Type, and user/group name. Other details can be retrieved if they are needed. Now in the time of iteration in the dictionary, the member (User/Group Name) as key and permission details as values are being set. And the key and value of this dictionary are declared as string type.

```cs
/// <summary>    
/// This funtion get the site permission details. And return it by a dictonary.    
/// </summary>    
/// <param name="clientContext"></param>    
/// <returns>Dictionary<string, string></returns>    
private Dictionary<string, string> GetSitePermissionDetails(ClientContext clientContext){  
    IEnumerable roles = clientContext.LoadQuery(clientContext.Web.RoleAssignments.Include(roleAsg => roleAsg.Member,    
                                                                      roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name)));    
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
```

Now, after filling in the dictionary, to get the details the dictionary has to loop through. Here is the code snippet for calling the GetSitePermissionDetails function and getting the permission details.
```cs
Dictionary<string, string> sitePermissionCollection = GetSitePermissionDetails(clientContext);    
foreach (var sitePermission in sitePermissionCollection)    
{    
        Console.WriteLine(sitePermission.Key + "   " + sitePermission.Value);    
}  
```
## Common Function to Get Site, List, List Item Permission
 
So, to get the list and list item permission the function is a similar process as what we used for the site permission. That’s why we can write a common function for getting permission for all three of them, So, to do that, we have to extract the uncommon part of the code which is the LINQ query, which is load, and send that query as parameter. Here is the code snippet of that common function.
```cs
/// <summary>    
/// This funtion get the site/list/list item permission details. And return it by a dictonary.    
/// </summary>    
/// <param name="clientContext">type ClientContext</param>    
/// <param name="queryString">type IQueryable<RoleAssignment></param>    
/// <returns>return type is Dictionary<string, string></returns>    
private Dictionary<string, string> GetPermissionDetails(ClientContext clientContext, IQueryable<RoleAssignment> queryString)    
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
```
### For Site Permission
 
Here is the code snippet to call this function for the site-level permission.
```cs
IQueryable<RoleAssignment> queryForSitePermission = clientContext.Web.RoleAssignments.Include(roleAsg => roleAsg.Member,   
                                                                   roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name));                   
Dictionary<string, string> sitePermissionCollection = GetPermissionDetails(clientContext, queryForSitePermission);    
  
foreach (var sitePermission in sitePermissionCollection)    
{    
     Console.WriteLine(sitePermission.Key + "   " + sitePermission.Value);        
}
```
Here we only have to send the LINQ query as parameter, otherwise the process is pretty much similar. So, for the other snippet, the iteration will be shown.

### For List Permission
```cs
List list = clientContext.Web.Lists.GetByTitle("List Title");    
clientContext.Load(list);    
clientContext.ExecuteQuery();    
    
IQueryable<RoleAssignment> queryForList = list.RoleAssignments.Include(roleAsg => roleAsg.Member,   
                                                                       roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name));    
Dictionary<string, string> permission = GetPermissionDetails(clientContext, queryForList);    
```
### For List Item Permission
```cs
///ListItem item    
/// The variable type for “item” is ListItem    
IQueryable<RoleAssignment> queryForListItem = item.RoleAssignments.Include(roleAsg => roleAsg.Member,   
                                                                           roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name));    
Dictionary<string, string> itemPermissionCollection = GetPermissionDetails(clientContext, queryForListItem);     
```
### Bonus
 
To check if a list or any list item has unique permission using CSOM (C#), the code snippet is given below. Here the property HasUniqueRoleAssignments returns a Boolean type. Returns “True”, if the list/list item has unique permission and “False” for non-unique permission, which means if the permission is inherited.
```cs
///the variable “list” has to be “List” type.     
clientContext.Load(list, l=>l.HasUniqueRoleAssignments);    
clientContext.ExecuteQuery();    
Console.WriteLine(list. HasUniqueRoleAssignments);    
    
///The variable "item" has to be “ListItem” type.     
clientContext.Load(item, i => i.HasUniqueRoleAssignments);    
clientContext.ExecuteQuery();    
Console.WriteLine(item.HasUniqueRoleAssignments);
```
In this article, I have tried to minimize the description and provide more helpful examples. I hope it will help other SharePoint developers.
