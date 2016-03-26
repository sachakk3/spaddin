using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.SharePoint.Client;

namespace Algosmart.SharePoint.TimeSheetReceiverWeb.Code
{
    public static class Helper
    {
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
            catch (Exception)
            {
                return false;
            }
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
    }
}