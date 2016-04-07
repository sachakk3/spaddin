using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Algosmart.SharePoint.TimeSheetReceiverWeb.Code;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;

namespace TestConsoleApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            ClientContext context = GetO365Context("https://algosmart.sharepoint.com/teams/dev/BackOffice/TS/", "Developer@algosmart.onmicrosoft.com", "Rostov2016");
            

            Guid listId = new Guid("{593FDB18-601F-4FB1-87C1-D74F9C8C0687}");
            int listItemId = 94;
            new RemoteEventReceiverManager().ItemHandleListEventHandler(context, listId, listItemId, SPRemoteEventType.ItemAdded);
            //new RemoteEventReceiverManager().AssociateRemoteEventsToHostWeb(context);
            //new RemoteEventReceiverManager().RemoveEventReceiversFromHostWeb(context);
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
