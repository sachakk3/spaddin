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
        private const string RECEIVER_ADDING_NAME = "ItemAddingEvent";
        private const string RECEIVER_UPDATED_NAME = "ItemUpdatedEvent";


        public void AssociateRemoteEventsToHostWeb(ClientContext clientContext)
        {
            //Get the Title and EventReceivers lists
            clientContext.Load(clientContext.Web.Lists,
                lists => lists.Include(
                    list => list.Title,
                    list => list.EventReceivers).Where
                        (list => list.Title == LIST_TITLE));

            clientContext.ExecuteQuery();

            List timeSheetList = clientContext.Web.Lists.FirstOrDefault();

            if (!IsReseiverExists(timeSheetList, RECEIVER_ADDED_NAME))
            {
                this.AddReceiverToList(timeSheetList, RECEIVER_ADDED_NAME, EventReceiverType.ItemAdded, EventReceiverSynchronization.Synchronous);
            }
            if (!IsReseiverExists(timeSheetList, RECEIVER_UPDATED_NAME))
            {
                this.AddReceiverToList(timeSheetList, RECEIVER_UPDATED_NAME, EventReceiverType.ItemUpdated, EventReceiverSynchronization.Synchronous);
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
                if (rerAdded != null)
                {
                    rerAdded.DeleteObject();
                }
                if (rerUpdated != null)
                {
                    rerUpdated.DeleteObject();
                }
                clientContext.ExecuteQuery();

            }
            catch (Exception oops)
            {
                System.Diagnostics.Trace.WriteLine(oops.Message);
            }
        }

        public void ItemAddedToListEventHandler(ClientContext clientContext, SPRemoteEventProperties properties)
        {
            try
            {

                List timeSheets = clientContext.Web.Lists.GetById(properties.ItemEventProperties.ListId);
                ListItem item = timeSheets.GetItemById(properties.ItemEventProperties.ListItemId);
                clientContext.Load(timeSheets.RootFolder);
                clientContext.Load(item);
                clientContext.ExecuteQuery();

                string ts_ProjectsLookup = item["ts_ProjectsLookup"] + "";
                if (!string.IsNullOrEmpty(ts_ProjectsLookup))
                {
                    string projectInternalName = GetProjectInternalName(clientContext, ts_ProjectsLookup);
                    string folderUrl = string.Format("{0}/{1}/{2}", projectInternalName, DateTime.Now.Year, DateTime.Now.Month);
                    string folderUrlFull = timeSheets.RootFolder.ServerRelativeUrl + "/" + folderUrl;

                    bool isNewFolder = EnsureFolder(clientContext, timeSheets, item, folderUrl, folderUrlFull);
                    if (isNewFolder)
                    {
                        SetPermissions(clientContext, projectInternalName);
                    }
                    MoveItem(clientContext, timeSheets, item, folderUrl, folderUrlFull);
                }
            }
            catch (Exception oops)
            {
                System.Diagnostics.Trace.WriteLine(oops.Message);
            }

        }
        private bool EnsureFolder(ClientContext clientContext, List list, ListItem item, string folderUrl, string folderUrlFull)
        {
            bool isNewFolder = false;
            Folder itemFolder = clientContext.Web.GetFolderByServerRelativeUrl(folderUrlFull);
            if (!itemFolder.ExistsInList(list))
            {
                list.CreateFolderInList(clientContext, folderUrl);
                isNewFolder = true;
            }
            return isNewFolder;
        }    
        public void ItemUpdatedToListEventHandler(ClientContext clientContext, SPRemoteEventProperties properties) {
            try
            {
                MoveItem(clientContext, properties.ItemEventProperties.ListId, properties.ItemEventProperties.ListItemId);
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
        private void MoveItem(ClientContext clientContext, List timeSheets, ListItem item, string folderUrl,string folderUrlFull)
        {             
            string folderUrlItem = item["FileDirRef"] + "";
            if (folderUrlFull != folderUrlItem)
            {                
                string itemPath = item["FileRef"] + "";
                File file = clientContext.Web.GetFileByServerRelativeUrl(itemPath);
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
        }
        private void SetPermissions(ClientContext clientContext, string projectInternalName)
        {
            string prjPMGroupName = string.Format("{0}-PM", projectInternalName);
            string prjTeamGroupName = string.Format("{0}-Team", projectInternalName);




        }
        private void SetPermissionsToFolder(ClientContext clientContext, string projectInternalName)
        private string GetProjectInternalName(ClientContext clientContext, string ts_ProjectsLookup)
        {
            return "Project1";
        }
    }
}