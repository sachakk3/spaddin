using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;
using TestConsoleApplication;

namespace Algosmart.SharePoint.TimeSheetReceiverWeb.Code
{
    public class RemoteEventReceiverManager
    {
        private const string LIST_TITLE = "timeboard";

        private const string GROUPS_OWNER_TITLE = "Backoffice - Владельцы";
        private const string GROUPS_BOSS_TITLE = "Backoffice-Boss";

        private const string WEBS_FINANCE_URL = "Finance";
        private const string LISTS_RATES_URL = "UserByProjectDetails";


        public void ItemAddedToListEventHandler(ClientContext clientContext, Guid ListId, int ListItemId)
        {
            try
            {
                List timeSheets = clientContext.Web.Lists.GetById(ListId);
                ListItem item = timeSheets.GetItemById(ListItemId);
                clientContext.Load(timeSheets.RootFolder);
                clientContext.Load(timeSheets.Fields);
                clientContext.Load(item);
                clientContext.ExecuteQuery();

                FieldLookupValue ts_ProjectsLookup = item["ts_ProjectsLookup"] as FieldLookupValue;
                if (ts_ProjectsLookup != null)
                {
                    FieldUserValue userCreated = item["Author"] as FieldUserValue;
                    string projectInternalName = GetProjectInternalName(clientContext, timeSheets, item, ts_ProjectsLookup);

                    List<SPTSFolder> folders = InitializeFolders(projectInternalName, userCreated.LookupValue);
                    SPTSFolder dateFolder = folders.Where(f => f.Type == TSFolderType.DateTime).FirstOrDefault();

                    string folderUrlFull = timeSheets.RootFolder.ServerRelativeUrl + "/" + dateFolder.ListRelativeURL;
                    Folder itemFolder = clientContext.Web.GetFolderByServerRelativeUrl(folderUrlFull);
                    EnsureFolder(clientContext, timeSheets, itemFolder , folders);

                    SPTSFolder projectFolder = folders.Where(f => f.Type == TSFolderType.Project).FirstOrDefault();
                    SPTSFolder userFolder = folders.Where(f => f.Type == TSFolderType.User).FirstOrDefault();
                    
                    if (projectFolder.IsNew || userFolder.IsNew)
                    {
                        SetPermissions(clientContext, timeSheets, projectFolder, userFolder, projectInternalName);
                    }
                    MoveItem(clientContext, timeSheets, item, dateFolder.ListRelativeURL, folderUrlFull);
                    string rate = GetRateFromRatesList(clientContext, ts_ProjectsLookup, userCreated);
                    int rateNumber = int.Parse(rate);
                    item["ts_Rate"] = rateNumber;
                    item.Update();
                    clientContext.ExecuteQuery();
                }
            }
            catch (Exception oops)
            {
                System.Diagnostics.Trace.WriteLine(oops.Message);
            }

        }
        private string GetRateFromRatesList(ClientContext clientContext, FieldLookupValue ts_ProjectsLookup, FieldUserValue userCreated)
        {
            string rate = string.Empty;
            Web web = clientContext.Site.OpenWeb(WEBS_FINANCE_URL);
            List detailsList = web.Lists.GetByTitle(LISTS_RATES_URL);
            CamlQuery query = new CamlQuery();
            query.ViewXml = "<View>"
                                + "<Query><Where>"
                                    + "<And>"
                                        + "<Eq><FieldRef Name='ts_ProjectsLookup' LookupId='TRUE'/>"
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
                rate = itemDetails["ts_Rate"] +"" ;
            }
            return rate;
        }
        private List<SPTSFolder> InitializeFolders(string projectInternalName, string userName)
        {
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
        private bool EnsureFolder(ClientContext clientContext, List list, Folder itemFolder, List<SPTSFolder> folders)
        {
            bool isNewFolder = false;
            if (!itemFolder.ExistsInList(list))
            {
                CreateFolders(clientContext, list, folders);
                isNewFolder = true;
            }
            return isNewFolder;
        }
        private void CreateFolders(ClientContext clientContext, List list, List<SPTSFolder> folders)
        {
            string curParentFolderPath = list.RootFolder.ServerRelativeUrl;

            foreach (SPTSFolder folder in folders)
            {                
                Folder folderObj = clientContext.Web.GetFolderByServerRelativeUrl(list.RootFolder.ServerRelativeUrl + "/" + folder.ListRelativeURL);
                if (!folderObj.ExistsInList(list))
                {
                    ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
                    listItemCreationInformation.UnderlyingObjectType = FileSystemObjectType.Folder;
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
        private void SetPermissions(ClientContext clientContext, List list, SPTSFolder projectFolder, SPTSFolder userFolder, string projectInternalName)
        {
            string prjPMGroupName = string.Format("{0}-PM", projectInternalName);
            string prjTeamGroupName = string.Format("{0}-Team", projectInternalName);
            
            GroupCollection groups = clientContext.Web.SiteGroups;
            clientContext.Load(groups);
            clientContext.ExecuteQuery();

            Group groupOwner = groups.Where(g => g.Title.ToLower() == GROUPS_OWNER_TITLE.ToLower()).FirstOrDefault();
            Group groupBoss = groups.Where(g => g.Title.ToLower() == GROUPS_BOSS_TITLE.ToLower()).FirstOrDefault();

            Group groupPM = groups.Where(g => g.Title.ToLower() == prjPMGroupName.ToLower()).FirstOrDefault();
            Group groupTeam = groups.Where(g => g.Title.ToLower() == prjTeamGroupName.ToLower()).FirstOrDefault();
            

            if (groupPM == null)
            {
                groupPM = Helper.CreateGroup(clientContext,prjPMGroupName);                
            }
            if (groupTeam == null)
            {
                groupTeam = Helper.CreateGroup(clientContext,prjTeamGroupName);                
            }
            if (projectFolder.IsNew)
            {
                SetPermissionsForProjectFolder(clientContext, list, projectFolder, groupPM, groupTeam, groupOwner, groupBoss);
            }
            if (userFolder.IsNew)
            {
                SetPermissionsForUserFolder(clientContext, list, userFolder, groupPM, groupTeam, groupOwner, groupBoss);
            }
        }        
        private void SetPermissionsForProjectFolder(ClientContext clientContext, List list, SPTSFolder folder, Group groupPM, Group groupTeam, Group groupOwner, Group groupBoss)
        {
            folder.Folder.BreakRoleInheritance(false, false);

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
        private void SetPermissionsForUserFolder(ClientContext clientContext, List list, SPTSFolder folder, Group groupPM, Group groupTeam, Group groupOwner, Group groupBoss)
        {
            folder.Folder.BreakRoleInheritance(false, false);

            RoleDefinitionBindingCollection collRoleOwnerDefinitionBinding = Helper.GetRoleFullControl(clientContext);
            folder.Folder.RoleAssignments.Add(groupOwner, collRoleOwnerDefinitionBinding);
            RoleDefinitionBindingCollection collRoleBossDefinitionBinding = Helper.GetRoleContribute(clientContext);
            folder.Folder.RoleAssignments.Add(groupBoss, collRoleBossDefinitionBinding);

            RoleDefinitionBindingCollection collRoleDefinitionBinding = Helper.GetRoleContribute(clientContext);
            folder.Folder.RoleAssignments.Add(groupPM, collRoleDefinitionBinding);
            folder.Folder.RoleAssignments.Add(groupTeam, collRoleDefinitionBinding);
            clientContext.ExecuteQuery();
        }
        private string GetProjectInternalName(ClientContext clientContext, List timeSheets, ListItem item, FieldLookupValue ts_ProjectsLookupValue)
        {
            FieldLookup ts_ProjectsLookup = (FieldLookup)timeSheets.Fields.Where(f=>f.InternalName == "ts_ProjectsLookup").FirstOrDefault();
            List projects = clientContext.Site.RootWeb.Lists.GetById(new Guid(ts_ProjectsLookup.LookupList));
            ListItem itemProject = projects.GetItemById(ts_ProjectsLookupValue.LookupId);
            clientContext.Load(itemProject);
            clientContext.ExecuteQuery();
            return itemProject["ts_InternalName"]+"";
        }
           
    }
}