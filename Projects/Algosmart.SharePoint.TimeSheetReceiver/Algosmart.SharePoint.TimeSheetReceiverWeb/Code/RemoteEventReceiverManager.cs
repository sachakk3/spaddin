using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Web;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;

namespace Algosmart.SharePoint.TimeSheetReceiverWeb.Code
{
    public class RemoteEventReceiverManager
    {
        private const string LIST_TITLE = "timeboard";

        private const string RECEIVER_ADDED_NAME = "ItemAddedEvent";
        private const string RECEIVER_UPDATED_NAME = "ItemUpdatedEvent";
        private const string LIST_TIMESHEET_ROOTFOLDER = "timeboard";

        public void AssociateRemoteEventsToHostWeb(ClientContext clientContext)
        {
            //clientContext.Load(clientContext.Web.Lists);
            clientContext.Load(clientContext.Web.Lists,
                lists => lists.Include(
                    list => list.Title,
                    list => list.EventReceivers,
                    list => list.RootFolder).Where
                        (list => list.Title == LIST_TITLE));

            clientContext.ExecuteQuery();

            List timeSheetList = clientContext.Web.Lists.FirstOrDefault();
            //List timeSheetList = clientContext.Web.Lists.Where(l=>l.RootFolder.Name)

            if (!IsReseiverExists(timeSheetList, RECEIVER_ADDED_NAME))
            {
                this.AddReceiverToList(timeSheetList, RECEIVER_ADDED_NAME,EventReceiverType.ItemAdded,EventReceiverSynchronization.Synchronous);
            }
            if (!IsReseiverExists(timeSheetList, RECEIVER_UPDATED_NAME))
            {
                this.AddReceiverToList(timeSheetList, RECEIVER_UPDATED_NAME, EventReceiverType.ItemUpdated, EventReceiverSynchronization.Asynchronous);
            }
            clientContext.ExecuteQuery();
        }
        
        public void RemoveEventReceiversFromHostWeb(ClientContext clientContext)
        {
            List myList = clientContext.Web.Lists.GetByTitle(LIST_TITLE);
            clientContext.Load(myList, p => p.EventReceivers);
            clientContext.ExecuteQuery();

            var rerAdded = myList.EventReceivers.Where(
                e => e.ReceiverName == RECEIVER_ADDED_NAME).FirstOrDefault();
            var rerUpdated = myList.EventReceivers.Where(
                e => e.ReceiverName == RECEIVER_UPDATED_NAME).FirstOrDefault();
            try
            {
                rerAdded.DeleteObject();
                rerUpdated.DeleteObject();
                clientContext.ExecuteQuery();

            }
            catch (Exception oops)
            {
                System.Diagnostics.Trace.WriteLine(oops.Message);
            }
        }
        public void ItemUpdatedToListEventHandler(ClientContext clientContext, SPRemoteEventProperties properties)
        {
        }
        public void ItemAddedToListEventHandler(ClientContext clientContext, SPRemoteEventProperties properties)
        {
            try
            {
                Web web = clientContext.Web;
                List timeSheets = web.Lists.GetById(properties.ItemEventProperties.ListId);

                ListItem item = timeSheets.GetItemById(properties.ItemEventProperties.ListItemId);
                clientContext.Load(timeSheets.RootFolder);
                clientContext.Load(item);
                clientContext.ExecuteQuery();
                string folderUrl = string.Format("Project1/{0}/{1}", DateTime.Now.Year, DateTime.Now.Month);


                string folderUrlFull = timeSheets.RootFolder.ServerRelativeUrl + "/" + folderUrl;
                Folder itemTimeSheetsFolder = web.GetFolderByServerRelativeUrl(folderUrlFull);
                if (!itemTimeSheetsFolder.ExistsInList(timeSheets))
                {
                    timeSheets.CreateFolderInList(clientContext, folderUrl);
                }
                string itemPath = item["FileRef"] + "";
                File file = web.GetFileByServerRelativeUrl(itemPath);
                clientContext.Load(file);
                clientContext.ExecuteQuery();
                if (file.Exists)
                {
                    var filePath = string.Format("{0}/{1}/{2}_.000", timeSheets.RootFolder.ServerRelativeUrl, folderUrl, item.Id);
                    file.MoveTo(filePath, MoveOperations.Overwrite);
                }
                clientContext.Load(file);
                clientContext.ExecuteQuery();


            }
            catch (Exception oops)
            {
                System.Diagnostics.Trace.WriteLine(oops.Message);
            }

        }
        public void ItemAddingToListEventHandler(ClientContext clientContext, SPRemoteEventProperties properties) {
            try
            {
                //Web web = clientContext.Web;
                //List timeSheets = web.Lists.GetById(properties.ItemEventProperties.ListId);
                
                //clientContext.Load(timeSheets.RootFolder);                
                //clientContext.ExecuteQuery();
                //string folderUrl = string.Format("Project1/{0}/{1}", DateTime.Now.Year, DateTime.Now.Month);


                //string folderUrlFull = timeSheets.RootFolder.ServerRelativeUrl + "/" + folderUrl;
                //Folder itemTimeSheetsFolder = web.GetFolderByServerRelativeUrl(folderUrlFull);
                //if (!itemTimeSheetsFolder.ExistsInList(timeSheets))
                //{
                //    timeSheets.CreateFolderInList(clientContext, folderUrl);
                //}
                ////string itemPath = properties.ItemEventProperties.BeforeProperties["FileRef"] + "";
                ////var filePath = string.Format("{0}/{1}_.000", folderUrl, item.Id);
                //properties.ItemEventProperties.AfterProperties["FileDirRef"] = folderUrl;
                ////["FileRef"] = filePath;
                ////item.Update();
                //clientContext.ExecuteQuery();

                
            }
            catch (Exception oops)
            {
                System.Diagnostics.Trace.WriteLine(oops.Message);
            }
        }
        private void AddReceiverToList(List list, string receiverName, EventReceiverType type, EventReceiverSynchronization synch)
        {
            EventReceiverDefinitionCreationInformation receiver =
                                new EventReceiverDefinitionCreationInformation();
            receiver.EventType = type;

            //Get WCF URL where this message was handled
            OperationContext op = OperationContext.Current;
            Message msg = op.RequestContext.RequestMessage;
            receiver.ReceiverUrl = msg.Headers.To.ToString();

            receiver.ReceiverName = receiverName;
            receiver.Synchronization = synch;

            //Add the new event receiver to a list in the host web
            list.EventReceivers.Add(receiver);

        }
        private bool IsReseiverExists(List list, string receiverName)
        {
            bool isExists = false;
            foreach (var rer in list.EventReceivers)
            {
                if (rer.ReceiverName == receiverName)
                {
                    isExists = true;
                    break;
                }
            }
            return isExists;
        }
    }
}