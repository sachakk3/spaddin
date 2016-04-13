using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Algosmart.SharePoint.Cmdlets.Code
{
    public static class Helper
    {
        public static ClientContext GetO365Context(string url, string userName, string password)
        {
            ClientContext context = new ClientContext(url);
            var passWord = new SecureString();
            foreach (char c in password.ToCharArray()) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(userName, passWord);
            var web = context.Web;
            context.Load(web);
            context.ExecuteQuery();
            return context;
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
