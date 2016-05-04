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
            ClientContext clientContext = Helper.GetO365Context(this.ProjectSiteURL, this.Login, this.Password);            
            Console.WriteLine("Получение списка проектов");
            List projects = clientContext.Web.Lists.GetByTitle(Constants.LISTS_PROJECTS_TITLE);
            CamlQuery query = CamlQuery.CreateAllItemsQuery();
            ListItemCollection items = projects.GetItems(query);
            Console.WriteLine("Получение всех групп");
            GroupCollection groups = clientContext.Web.SiteGroups;
            clientContext.Load(groups);
            clientContext.Load(items,
                elements => elements.Include(
                    item => item.HasUniqueRoleAssignments,
                    item => item[Constants.FIELDS_TITLE],
                    item => item[Constants.FIELDS_INTERNAL_NAME],
                    item => item[Constants.FIELDS_PROJECTS_PM],
                    item => item[Constants.FIELDS_PROJECTS_USERS]));
            clientContext.ExecuteQuery();
            Console.WriteLine(string.Format("Отправлено на обработку '{0}' проектов",items.Count));
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
                    Console.WriteLine("Установка разрешений для проекта '{0}'", item[Constants.FIELDS_TITLE]);
                    try
                    {
                        SetPermissions(clientContext, item, projectInternalName, groupOwner, groupBoss, groupHR, groupFin, groupBackOfficePM);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Произошла ошибка установки разрешений для проекта '{0}'. Описание '{1}'", item[Constants.FIELDS_TITLE], ex.Message);
                    }
                }
                else
                {
                    Console.WriteLine("Для проекта '{0}' не указано  internalName", item[Constants.FIELDS_TITLE]);
                }
                Console.WriteLine("Завершена установка разрешений для проекта '{0}'", item[Constants.FIELDS_TITLE]);
            }            
        }
        private static void SetPermissions(ClientContext clientContext, ListItem item, string projectInternalName, Group groupOwner, Group groupBoss, Group groupHR, Group groupFin, Group groupBackOfficePM)
        {
            FieldUserValue projectManager = item[Constants.FIELDS_PROJECTS_PM] as FieldUserValue;
            FieldUserValue[] projectMembers = item[Constants.FIELDS_PROJECTS_USERS] as FieldUserValue[];

            string prjPMGroupName = string.Format("{0}-PM", projectInternalName);
            string prjTeamGroupName = string.Format("{0}-Team", projectInternalName);

            Group groupPM = EnsureGroupAndMembers(clientContext, prjPMGroupName, projectManager);
            Group groupTeam = EnsureGroupAndMembers(clientContext, prjTeamGroupName, projectMembers);
            
            SetPermissionsForProjectFolder(clientContext, item, groupPM, groupTeam, groupOwner, groupBoss, groupHR, groupFin, groupBackOfficePM);            
        }
        private static Group EnsureGroupAndMembers(ClientContext clientContext, string groupName, FieldUserValue[] users)
        {
            Group group = EnsureGroup(clientContext, groupName);

            if (users != null)
            {
                foreach (FieldUserValue user in users)
                {
                    AddUserToGroup(clientContext, group, user.LookupId);
                    Console.WriteLine("В группу '{0}' добавлен пользователь '{1}'", groupName, user.LookupValue);
                }
                clientContext.ExecuteQuery();
            }
            else
            {
                Console.WriteLine("Нет пользователей для добавления в группу '{0}'.", groupName);
            }
            return group;
        }
        private static Group EnsureGroupAndMembers(ClientContext clientContext, string groupName, FieldUserValue user)
        {
            Group group = EnsureGroup(clientContext, groupName);

            if (user != null)
            {
                AddUserToGroup(clientContext, group, user.LookupId);
                clientContext.ExecuteQuery();
                Console.WriteLine("В группу '{0}' добавлен пользователь '{1}'", groupName, user.LookupValue);                
            }
            else
            {
                Console.WriteLine("Нет пользователей для добавления в группу '{0}'.", groupName);
            }
            return group;
        }
        private static Group EnsureGroup(ClientContext clientContext, string groupName)
        {
            Group group = clientContext.Web.SiteGroups.Where(g => g.Title.ToLower() == groupName.ToLower()).FirstOrDefault();
            if (group == null)
            {                
                group = Helper.CreateGroup(clientContext, groupName);
                Console.WriteLine("Группа '{0}' добавлена", groupName);
            }
            else
            {                
                RemoveAllUsersFromGroup(clientContext, group);
                Console.WriteLine("Из группы '{0}' удалены все пользователи", groupName);
            }
            return group;
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
        private static void AddUserToGroup(ClientContext clientContext, Group group, int userId)
        {
            User user = clientContext.Web.SiteUsers.GetById(userId);
            group.Users.AddUser(user);
        }
        private static void RemoveAllUsersFromGroup(ClientContext clientContext, Group group)
        {
            UserCollection users = group.Users;
            clientContext.Load(users);
            clientContext.ExecuteQuery();
            if (users.Count > 0)
            {
                for (int i = users.Count - 1; i >= 0; i--)
                {
                    users.Remove(users[i]);
                }
                clientContext.ExecuteQuery();
            }
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
