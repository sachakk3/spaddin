using Algosmart.SharePoint.Cmdlets.Code;
using Algosmart.SharePoint.TimeSheetReceiverWeb.Code;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace Algosmart.SharePoint.Cmdlets
{
    [Cmdlet("Update", "Projects")]
    public class UpdateProjects : Cmdlet
    {
        [Parameter(Mandatory = true)]
        public string ProjectSiteURL { get; set; }
        [Parameter(Mandatory = true)]
        public string Login { get; set; }
        [Parameter(Mandatory = true)]
        public string Password { get; set; }  

        /// <summary>
        /// Задание параметров и запуск экспорта
        /// </summary>
        protected override void ProcessRecord()
        {
            Update();
        }
        public void Update()
        {
            ClientContext clientContext = Code.Helper.GetO365Context(this.ProjectSiteURL, this.Login, this.Password);

            string listProjectsTitle = "Проекты";
            Console.WriteLine(string.Format("Получение списка элементов", listProjectsTitle));
            List timeSheets = clientContext.Web.Lists.GetByTitle(listProjectsTitle);
            CamlQuery query = CamlQuery.CreateAllItemsQuery();
            ListItemCollection items = timeSheets.GetItems(query);
            Console.WriteLine("Получение всех групп");
            GroupCollection groups = clientContext.Web.SiteGroups;
            clientContext.Load(groups); 
            clientContext.Load(items);
            clientContext.ExecuteQuery();

            Group groupOwner = groups.Where(g => g.Title.ToLower() == Constants.GROUPS_OWNER_TITLE.ToLower()).FirstOrDefault();
            Group groupBoss = groups.Where(g => g.Title.ToLower() == Constants.GROUPS_BOSS_TITLE.ToLower()).FirstOrDefault();

            Group groupHR = groups.Where(g => g.Title.ToLower() == Constants.GROUPS_HR_TITLE.ToLower()).FirstOrDefault();
            Group groupFin = groups.Where(g => g.Title.ToLower() == Constants.GROUPS_Fin_TITLE.ToLower()).FirstOrDefault();
            Group groupBackOfficePM = groups.Where(g => g.Title.ToLower() == Constants.GROUPS_PM_TITLE.ToLower()).FirstOrDefault();


            foreach (ListItem item in items)
            {

                string projectInternalName = item[Constants.FIELDS_INTERNAL_NAME] + "";
                if (!string.IsNullOrEmpty(projectInternalName))
                {
                    Console.WriteLine("Установка разрешений для элемента '{0}'", item["Title"]);
                    try
                    {
                        SetPermissions(clientContext, item, projectInternalName, groups, groupOwner, groupBoss, groupHR, groupFin, groupBackOfficePM);
                    }
                    catch(Exception ex)
                    {
                        Console.WriteLine("Произошла ошибка установки разрешений для элемента '{0}'. Описание '{1}'", item["Title"], ex.Message);
                    }
                    }
                else {
                    Console.WriteLine("Для проекта '{0}' не указано  internalName", item["Title"]);
                }
            }
        }
        private static void SetPermissions(ClientContext clientContext, ListItem item,string projectInternalName, GroupCollection groups, Group groupOwner, Group groupBoss, Group groupHR, Group groupFin, Group groupBackOfficePM)
        {
            FieldUserValue projectManager = item[Constants.FIELDS_PROJECTS_PM] as FieldUserValue;
            FieldUserValue[] projectMembers = item[Constants.FIELDS_PROJECTS_USERS] as FieldUserValue[];


            string prjPMGroupName = string.Format("{0}-PM", projectInternalName);
            string prjTeamGroupName = string.Format("{0}-Team", projectInternalName);

            
            Group groupPM = groups.Where(g => g.Title.ToLower() == prjPMGroupName.ToLower()).FirstOrDefault();
            Group groupTeam = groups.Where(g => g.Title.ToLower() == prjTeamGroupName.ToLower()).FirstOrDefault();

            if (groupPM == null)
            {
                Console.WriteLine("Создание группы '{0}'", prjPMGroupName);
                groupPM = Helper.CreateGroup(clientContext, prjPMGroupName);
                if (projectManager != null)
                {
                    User user = clientContext.Web.SiteUsers.GetById(projectManager.LookupId);
                    groupPM.Users.AddUser(user);
                }
                else
                {
                    Console.WriteLine("Для проекта '{0}' не указан менеджер проекта", item["Title"]);
                }
            }
            if (groupTeam == null)
            {
                Console.WriteLine("Создание группы '{0}'", prjTeamGroupName);
                groupTeam = Helper.CreateGroup(clientContext, prjTeamGroupName);
                if (projectMembers != null)
                {
                    foreach (FieldUserValue member in projectMembers)
                    {
                        User user = clientContext.Web.SiteUsers.GetById(member.LookupId);
                        groupTeam.Users.AddUser(user);
                    }
                }
                else
                {
                    Console.WriteLine("Для проекта '{0}' не указаны участники проекта", item["Title"]);
                }
            }            
            SetPermissionsForProjectFolder(clientContext, item, groupPM, groupTeam, groupOwner, groupBoss, groupHR, groupFin, groupBackOfficePM);
        }
        private static void SetPermissionsForProjectFolder(ClientContext clientContext, ListItem item, Group groupPM, Group groupTeam, Group groupOwner, Group groupBoss, Group groupHR, Group groupFin, Group groupBackOfficePM)
        {
            RemoveAllRoleAssignments(clientContext, item);

            RoleDefinitionBindingCollection collRoleFullControlDefinitionBinding =Helper.GetRoleFullControl(clientContext);
            item.RoleAssignments.Add(groupOwner, collRoleFullControlDefinitionBinding);

            RoleDefinitionBindingCollection collRoleContributeDefinitionBinding =Helper.GetRoleContribute(clientContext);
            item.RoleAssignments.Add(groupBoss, collRoleContributeDefinitionBinding);
            item.RoleAssignments.Add(groupPM, collRoleContributeDefinitionBinding);

            RoleDefinitionBindingCollection collRoleReaderDefinitionBinding = Helper.GetRoleReader(clientContext);
            item.RoleAssignments.Add(groupHR, collRoleReaderDefinitionBinding);
            item.RoleAssignments.Add(groupFin, collRoleReaderDefinitionBinding);
            item.RoleAssignments.Add(groupBackOfficePM, collRoleReaderDefinitionBinding);
            item.RoleAssignments.Add(groupTeam, collRoleReaderDefinitionBinding);
            clientContext.ExecuteQuery();
        }
        private static void RemoveAllRoleAssignments(ClientContext clientContext, ListItem item)
        {
            item.BreakRoleInheritance(false, false);
            RoleAssignmentCollection roleAssignments = item.RoleAssignments;
            clientContext.Load(roleAssignments);
            clientContext.ExecuteQuery();

            for (int i = item.RoleAssignments.Count - 1; i >= 0; i--)
            {
                item.RoleAssignments[i].DeleteObject();
            }
        }
    }
}
