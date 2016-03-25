using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Algosmart.SharePoint.TimeSheetReceiverWeb.Code;
using Microsoft.SharePoint.Client;

namespace TestConsoleApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            ClientContext context = GetO365Context("https://spdevone.sharepoint.com/sites/BackOffice/TimeSheet/", "admin@spdevone.onmicrosoft.com", "Rostov2016");
            

            Guid listId = new Guid("{BA735DC0-5237-4DA2-9AF6-83BE894EAB68}");
            int listItemId = 1;
            new RemoteEventReceiverManager().ItemAddedToListEventHandler(context, listId, listItemId);
        }
        private static ClientContext GetO365Context(string url, string userName, string password)
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
