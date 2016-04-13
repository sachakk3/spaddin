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
    }
}
