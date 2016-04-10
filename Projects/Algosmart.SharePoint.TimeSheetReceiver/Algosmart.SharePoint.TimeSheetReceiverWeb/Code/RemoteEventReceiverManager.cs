using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Channels;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;
using TestConsoleApplication;

namespace Algosmart.SharePoint.TimeSheetReceiverWeb.Code
{
    public class RemoteEventReceiverManager
    {
        

        public void ItemHandleListEventHandler(ClientContext clientContext, Guid ListId, int ListItemId, SPRemoteEventType eventType)
        {
            try
            {
                Trace.TraceInformation(string.Format("Начата обработка элемента ID-'{0}'", ListItemId));

                List timeSheets = clientContext.Web.Lists.GetById(ListId);
                ListItem item = timeSheets.GetItemById(ListItemId);
                clientContext.Load(timeSheets.RootFolder);
                clientContext.Load(timeSheets.Fields);
                clientContext.Load(item);
                clientContext.ExecuteQuery();

                FieldLookupValue ts_ProjectsLookup = item[Constants.FIELDS_PROJECTS_LOOKUP] as FieldLookupValue;
                if (ts_ProjectsLookup != null)
                {
                    FieldUserValue userCreated = item[Constants.FIELDS_AUTHOR] as FieldUserValue;
                    Web parentWeb = GetParentWeb(clientContext);
                    string projectInternalName = GetProjectInternalName(clientContext, parentWeb, timeSheets, ts_ProjectsLookup);
                    if (!string.IsNullOrEmpty(projectInternalName))
                    {
                        List<SPTSFolder> folders = InitializeFolders(projectInternalName, userCreated.LookupValue);
                        SPTSFolder dateFolder = folders.Where(f => f.Type == TSFolderType.DateTime).FirstOrDefault();
                        string folderUrlFull = timeSheets.RootFolder.ServerRelativeUrl + "/" + dateFolder.ListRelativeURL;

                        EnsureFolders(clientContext, timeSheets, folders, folderUrlFull);

                        SPTSFolder projectFolder = folders.Where(f => f.Type == TSFolderType.Project).FirstOrDefault();
                        SPTSFolder userFolder = folders.Where(f => f.Type == TSFolderType.User).FirstOrDefault();

                        if (projectFolder.IsNew || userFolder.IsNew)
                        {
                            SetPermissions(clientContext, timeSheets, projectFolder, userFolder, projectInternalName, userCreated.LookupId);
                        }
                        MoveItem(clientContext, timeSheets, item, folderUrlFull);
                        if (eventType == SPRemoteEventType.ItemAdded)
                        {
                            string rate = GetRateFromRatesList(clientContext, parentWeb, ts_ProjectsLookup, userCreated);
                            if (!string.IsNullOrEmpty(rate))
                            {
                                item[Constants.FIELDS_RATE] = rate;                                
                                Trace.TraceInformation("Элементу установлена ставка '{0}'", rate);                                
                            }
                            else {
                                Trace.TraceInformation("Для проекта '{0}', пользователю '{1}' ставка не установлена", ts_ProjectsLookup.LookupValue, userCreated.LookupValue);
                            }
                            item.Update();
                        }
                        if (eventType == SPRemoteEventType.ItemUpdated)
                        {
                            if (item[Constants.FIELDS_TIMEBOARD_STATUS] + "" == "Утверждено")
                            {
                                Microsoft.SharePoint.Client.RecordsRepository.Records.DeclareItemAsRecord(clientContext, item);
                                Trace.TraceInformation("Элемент объявлен запеисью '{0}'", item["Title"]);
                            }
                        }                        
                        clientContext.ExecuteQuery();
                    }
                    else
                    {
                        Trace.TraceError("Для проекта '{0}' не указано internalName", ts_ProjectsLookup.LookupValue);
                    }                                    
                }
                Trace.TraceInformation(string.Format("Окончена обработка элемента ID-'{0}'", ListItemId));
            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.ToString());
            }

        }
        private static Web GetParentWeb(ClientContext clientContext)
        {
            WebInformation webParentInfo = clientContext.Web.ParentWeb;
            clientContext.Load(webParentInfo);
            clientContext.ExecuteQuery();
            Web parentWeb = clientContext.Site.OpenWebById(webParentInfo.Id);
            clientContext.Load(parentWeb);
            clientContext.ExecuteQuery();
            return parentWeb;
        }
        private static string GetRateFromRatesList(ClientContext clientContext, Web parentWeb, FieldLookupValue ts_ProjectsLookup, FieldUserValue userCreated)
        {
            string rate = string.Empty;
            Web web = clientContext.Site.OpenWeb(parentWeb.ServerRelativeUrl + Constants.WEBS_FINANCE_NAME);
            List detailsList = web.Lists.GetByTitle(Constants.LISTS_RATES_TITLE);
            CamlQuery query = new CamlQuery();
            query.ViewXml = "<View>"
                                + "<Query><Where>"
                                    + "<And>"
                                        + "<Eq><FieldRef Name='" + Constants.FIELDS_PROJECTS_LOOKUP + "' LookupId='TRUE'/>"
                                            + "<Value Type='Lookup'>"
                                                + ts_ProjectsLookup.LookupId
                                            + "</Value>"
                                         + "</Eq>"
                                         + "<Eq><FieldRef Name='ts_Employees' LookupId='TRUE'/>"
                                            + "<Value Type='Lookup'>"
                                                + userCreated.LookupId
                                            + "</Value>"
                                         + "</Eq>"
                                    + "</And>"
                                 + "</Where></Query>"
                                + "<ViewFields><FieldRef Name='Title' /><FieldRef Name='ts_Rate' /></ViewFields>"
                            + "</View>";

            ListItemCollection items = detailsList.GetItems(query);
            clientContext.Load(items);
            clientContext.ExecuteQuery();
            if (items.Count > 0)
            {
                ListItem itemDetails = items[0];
                rate = itemDetails[Constants.FIELDS_RATE] + "";
            }
            return rate;
        }
        private static List<SPTSFolder> InitializeFolders(string projectInternalName, string userName)
        {
            Trace.TraceInformation("Формирование структуры папок для проекта '{0}' и пользователя '{1}'", projectInternalName, userName);
            SPTSFolder projectFolder = new SPTSFolder
            {
                Name = projectInternalName,
                Type = TSFolderType.Project
            };
            SPTSFolder userFolder = new SPTSFolder
            {
                Name = userName,
                Type = TSFolderType.User,
                ParentFolder = projectFolder
            };
            SPTSFolder dateFolder = new SPTSFolder
            {
                Name = DateTime.Now.ToString("yyyyMM"),
                Type = TSFolderType.DateTime,
                ParentFolder = userFolder
            };
            return new List<SPTSFolder> { projectFolder, userFolder, dateFolder };
        }
        private static void EnsureFolders(ClientContext clientContext, List list, List<SPTSFolder> folders, string folderUrlFull)
        {
            Trace.TraceInformation("Проверка наличия папок по пути '{0}'", folderUrlFull);
            Folder itemFolder = clientContext.Web.GetFolderByServerRelativeUrl(folderUrlFull);
            if (!itemFolder.ExistsInList(list))            
            {
                CreateFolders(clientContext, list, folders);
            }
        }
        private static void CreateFolders(ClientContext clientContext, List list, List<SPTSFolder> folders)
        {
            string curParentFolderPath = list.RootFolder.ServerRelativeUrl;
            foreach (SPTSFolder folder in folders)
            {
                Folder folderObj = clientContext.Web.GetFolderByServerRelativeUrl(list.RootFolder.ServerRelativeUrl + "/" + folder.ListRelativeURL);                
                if (!folderObj.ExistsInList(list))
                {
                    Trace.TraceInformation("Создание папки по пути '{0}'", curParentFolderPath + "/" + folder.Name);
                    ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
                    listItemCreationInformation.UnderlyingObjectType = FileSystemObjectType.Folder;
                    listItemCreationInformation.LeafName = folder.Name;
                    listItemCreationInformation.FolderUrl = curParentFolderPath;
                    folder.IsNew = true;
                    folder.Folder = list.AddItem(listItemCreationInformation);
                    folder.Folder["Title"] = folder.Name;
                    folder.Folder.Update();
                    clientContext.ExecuteQuery();
                }
                curParentFolderPath += "/" + folder.Name;
            }
        }
        private static bool IsReseiverExists(List list, string receiverName)
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
        private static void MoveItem(ClientContext clientContext, List timeSheets, ListItem item, string folderUrlFull)
        {
            Trace.TraceInformation("Перемещение элемента '{0}'", item["Title"]);
            string folderUrlItem = item["FileDirRef"] + "";
            if (folderUrlFull.ToLower() != folderUrlItem.ToLower())
            {
                string itemPath = item["FileRef"] + "";
                File file = clientContext.Web.GetFileByServerRelativeUrl(itemPath);
                clientContext.Load(file);
                clientContext.ExecuteQuery();
                if (file.Exists)
                {
                    var filePath = string.Format("{0}/{1}_.000", folderUrlFull, item.Id);
                    file.MoveTo(filePath, MoveOperations.Overwrite);
                }
                clientContext.Load(file);
                clientContext.ExecuteQuery();
                Trace.TraceInformation("Элемент перемещен в '{0}'", folderUrlFull);
            }
        }
        private static void SetPermissions(ClientContext clientContext, List list, SPTSFolder projectFolder, SPTSFolder userFolder, string projectInternalName, int userId)
        {
            string prjPMGroupName = string.Format("{0}-PM", projectInternalName);
            string prjTeamGroupName = string.Format("{0}-Team", projectInternalName);

            Trace.TraceInformation("Получение всех групп");
            GroupCollection groups = clientContext.Web.SiteGroups;
            clientContext.Load(groups);
            clientContext.ExecuteQuery();

            Group groupOwner = groups.Where(g => g.Title.ToLower() == Constants.GROUPS_OWNER_TITLE.ToLower()).FirstOrDefault();
            Group groupBoss = groups.Where(g => g.Title.ToLower() == Constants.GROUPS_BOSS_TITLE.ToLower()).FirstOrDefault();

            Group groupPM = groups.Where(g => g.Title.ToLower() == prjPMGroupName.ToLower()).FirstOrDefault();
            Group groupTeam = groups.Where(g => g.Title.ToLower() == prjTeamGroupName.ToLower()).FirstOrDefault();

            if (groupPM == null)
            {
                Trace.TraceInformation("Создание группы '{0}'", prjPMGroupName);
                groupPM = Helper.CreateGroup(clientContext, prjPMGroupName);
            }
            if (groupTeam == null)
            {
                Trace.TraceInformation("Создание группы '{0}'", prjTeamGroupName);
                groupTeam = Helper.CreateGroup(clientContext, prjTeamGroupName);
            }
            if (projectFolder.IsNew)
            {
                Trace.TraceInformation("Назначение прав на группу '{0}'", projectFolder.ListRelativeURL);
                SetPermissionsForProjectFolder(clientContext, projectFolder, groupPM, groupTeam, groupOwner, groupBoss);
            }
            if (userFolder.IsNew)
            {
                Trace.TraceInformation("Назначение прав на группу '{0}'", userFolder.ListRelativeURL);
                SetPermissionsForUserFolder(clientContext, userFolder, groupPM, userId, groupOwner, groupBoss);
            }
        }
        private static void SetPermissionsForProjectFolder(ClientContext clientContext, SPTSFolder folder, Group groupPM, Group groupTeam, Group groupOwner, Group groupBoss)
        {
            RemoveAllRoleAssignments(clientContext, folder);

            RoleDefinitionBindingCollection collRoleOwnerDefinitionBinding = Helper.GetRoleFullControl(clientContext);
            folder.Folder.RoleAssignments.Add(groupOwner, collRoleOwnerDefinitionBinding);
            RoleDefinitionBindingCollection collRoleBossDefinitionBinding = Helper.GetRoleContribute(clientContext);
            folder.Folder.RoleAssignments.Add(groupBoss, collRoleBossDefinitionBinding);

            RoleDefinitionBindingCollection collRolePMDefinitionBinding = Helper.GetRoleContribute(clientContext);
            folder.Folder.RoleAssignments.Add(groupPM, collRolePMDefinitionBinding);
            RoleDefinitionBindingCollection collRoleUserDefinitionBinding = Helper.GetRoleReader(clientContext);
            folder.Folder.RoleAssignments.Add(groupTeam, collRoleUserDefinitionBinding);
            clientContext.ExecuteQuery();
        }
        private static void RemoveAllRoleAssignments(ClientContext clientContext, SPTSFolder folder)
        {
            folder.Folder.BreakRoleInheritance(false, false);
            RoleAssignmentCollection roleAssignments = folder.Folder.RoleAssignments;
            clientContext.Load(roleAssignments);
            clientContext.ExecuteQuery();

            for (int i = folder.Folder.RoleAssignments.Count - 1; i >= 0; i--)
            {
                folder.Folder.RoleAssignments[i].DeleteObject();
            }
        }
        private static void SetPermissionsForUserFolder(ClientContext clientContext, SPTSFolder folder, Group groupPM, int userId, Group groupOwner, Group groupBoss)
        {
            RemoveAllRoleAssignments(clientContext, folder);

            RoleDefinitionBindingCollection collRoleOwnerDefinitionBinding = Helper.GetRoleFullControl(clientContext);
            folder.Folder.RoleAssignments.Add(groupOwner, collRoleOwnerDefinitionBinding);
            RoleDefinitionBindingCollection collRoleBossDefinitionBinding = Helper.GetRoleContribute(clientContext);
            folder.Folder.RoleAssignments.Add(groupBoss, collRoleBossDefinitionBinding);

            Principal user = clientContext.Web.SiteUsers.GetById(userId);
            RoleDefinitionBindingCollection collRoleDefinitionBinding = Helper.GetRoleContribute(clientContext);
            folder.Folder.RoleAssignments.Add(groupPM, collRoleDefinitionBinding);
            folder.Folder.RoleAssignments.Add(user, collRoleDefinitionBinding);
            clientContext.ExecuteQuery();
        }
        private static string GetProjectInternalName(ClientContext clientContext, Web parentWeb, List timeSheets, FieldLookupValue ts_ProjectsLookupValue)
        {
            FieldLookup ts_ProjectsLookup = (FieldLookup)timeSheets.Fields.Where(f => f.InternalName == Constants.FIELDS_PROJECTS_LOOKUP).FirstOrDefault();
            
            List projects = parentWeb.Lists.GetById(new Guid(ts_ProjectsLookup.LookupList));
            ListItem itemProject = projects.GetItemById(ts_ProjectsLookupValue.LookupId);
            clientContext.Load(itemProject);
            clientContext.ExecuteQuery();
            return itemProject[Constants.FIELDS_INTERNAL_NAME] + "";
        }
        public void AssociateRemoteEventsToHostWeb(ClientContext clientContext)
        {            
            clientContext.Load(clientContext.Web.Lists,
                lists => lists.Include(
                    list => list.Title,
                    list => list.EventReceivers).Where
                        (list => list.Title == Constants.LIST_TITLE));

            clientContext.ExecuteQuery();

            List timeSheetList = clientContext.Web.Lists.FirstOrDefault();

            if (!IsReseiverExists(timeSheetList, Constants.RECEIVER_ADDED_NAME))
            {
                this.AddReceiverToList(timeSheetList, Constants.RECEIVER_ADDED_NAME, EventReceiverType.ItemAdded, EventReceiverSynchronization.Asynchronous);
            }
            if (!IsReseiverExists(timeSheetList, Constants.RECEIVER_UPDATED_NAME))
            {
                this.AddReceiverToList(timeSheetList, Constants.RECEIVER_UPDATED_NAME, EventReceiverType.ItemUpdated, EventReceiverSynchronization.Asynchronous);
            }
            clientContext.ExecuteQuery();
        }
        public void RemoveEventReceiversFromHostWeb(ClientContext clientContext)
        {
            List myList = clientContext.Web.Lists.GetByTitle(Constants.LIST_TITLE);
            clientContext.Load(myList, p => p.EventReceivers);
            clientContext.ExecuteQuery();

            var rerAdded = myList.EventReceivers.Where(
                e => e.ReceiverName == Constants.RECEIVER_ADDED_NAME).FirstOrDefault();
            var rerUpdated = myList.EventReceivers.Where(
                e => e.ReceiverName == Constants.RECEIVER_UPDATED_NAME).FirstOrDefault();
            try
            {
                if (rerAdded != null)
                {
                    rerAdded.DeleteObject();
                    Trace.TraceInformation(string.Format("Удален ресивер '{0}'", Constants.RECEIVER_ADDED_NAME));
                }
                if (rerUpdated != null)
                {
                    rerUpdated.DeleteObject();
                    Trace.TraceInformation(string.Format("Удален ресивер '{0}'", Constants.RECEIVER_UPDATED_NAME));

                }
                clientContext.ExecuteQuery();

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
            receiver.SequenceNumber = 1000;
            Trace.TraceInformation(string.Format("Добавление ресивера '{0}' по URL '{1}'", receiver.ReceiverName, receiver.ReceiverUrl));

            //Add the new event receiver to a list in the host web
            list.EventReceivers.Add(receiver);

        }       
    }
}