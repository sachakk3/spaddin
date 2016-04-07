using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security;
using System.Web;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;

namespace Algosmart.SharePoint.TimeSheetReceiverWeb.Code
{
    public static class Helper
    {
        public static bool ShouldSecretBeUpdated(IReadOnlyDictionary<string, object> beforeProperties,  IReadOnlyDictionary<string, object> afterProperties)
        {
            // If the property doesn't exist, then the secret should be updated
            if (!beforeProperties.ContainsKey("CMIsSecret") || !afterProperties.ContainsKey("CMIsSecret"))
            {
                return true;
            }
            //// If the value of IsSecret differ, then secret should be updated
            return afterProperties["CMIsSecret"].ToString() != beforeProperties["CMIsSecret"].ToString();
        }
        public static RoleDefinitionBindingCollection GetRoleContribute(ClientContext clientContext)
        {
            RoleDefinitionBindingCollection collRolePMDefinitionBinding = new RoleDefinitionBindingCollection(clientContext);
            //set role type
            collRolePMDefinitionBinding.Add(clientContext.Web.RoleDefinitions.GetByType(RoleType.Contributor));
            return collRolePMDefinitionBinding;
        }
        public static RoleDefinitionBindingCollection GetRoleReader(ClientContext clientContext)
        {
            RoleDefinitionBindingCollection collRoleDefinitionBinding = new RoleDefinitionBindingCollection(clientContext);
            //set role type
            collRoleDefinitionBinding.Add(clientContext.Web.RoleDefinitions.GetByType(RoleType.Reader));
            return collRoleDefinitionBinding;
        }
        public static RoleDefinitionBindingCollection GetRoleFullControl(ClientContext clientContext)
        {
            RoleDefinitionBindingCollection collRoleDefinitionBinding = new RoleDefinitionBindingCollection(clientContext);
            //set role type
            collRoleDefinitionBinding.Add(clientContext.Web.RoleDefinitions.GetByType(RoleType.Administrator));
            return collRoleDefinitionBinding;
        }
        public static bool ExistsInList(this Folder folder, List list)
        {
            try
            {
                list.Context.Load(folder);
                list.Context.ExecuteQuery();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        public static bool IsFolderExists(this List list, ClientContext clientContext, string folderUrl)
        {
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = "<View Scope='RecursiveAll'>"
                                    + "<Query>"
                                        + "<Where>"
                                            + "<And>"
                                            + "<Eq>"
                                                + "<FieldRef Name='FSObjType' />"
                                                + "<Value Type='Integer'>1</Value>"
                                            + "</Eq>"
                                            + "<Eq>"
                                                + "<FieldRef Name='FileRef' />"
                                                + "<Value Type='Integer'>" + folderUrl + "</Value>"
                                            + "</Eq>"
                                            + "</And>"
                                        + "</Where>"
                                    + "</Query>"
                                + "</View>";
            ListItemCollection listItems = list.GetItems(camlQuery);
            clientContext.Load(listItems);
            clientContext.ExecuteQuery();

            if (listItems.Count > 0)
            {
                return true;
            }
            return false;
        }
        public static Group CreateGroup(ClientContext clientContext, string groupName)
        {
            GroupCreationInformation inform = new GroupCreationInformation();
            inform.Title = groupName;
            Group group = clientContext.Web.SiteGroups.Add(inform);            
            clientContext.Load(group);
            clientContext.ExecuteQuery();
            return group;
        }
        public static ClientContext GetO365Context(SPRemoteEventProperties properties, string userName, string password)
        {
            Uri url = GetUrlFromProperties(properties);
            ClientContext context = new ClientContext(url);
            var passWord = new SecureString();
            foreach (char c in password.ToCharArray()) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(userName, passWord);
            var web = context.Web;
            context.Load(web);
            context.ExecuteQuery();
            return context;
        }
        public static Uri GetUrlFromProperties(SPRemoteEventProperties properties)
        {
            Uri sharepointUrl;
            if (properties.ListEventProperties != null)
            {
                sharepointUrl = new Uri(properties.ListEventProperties.WebUrl);
            }
            else if (properties.ItemEventProperties != null)
            {
                sharepointUrl = new Uri(properties.ItemEventProperties.WebUrl);
            }
            else if (properties.WebEventProperties != null)
            {
                sharepointUrl = new Uri(properties.WebEventProperties.FullUrl);
            }
            else
            {
                return null;
            }
            return sharepointUrl;
        }
    }
}