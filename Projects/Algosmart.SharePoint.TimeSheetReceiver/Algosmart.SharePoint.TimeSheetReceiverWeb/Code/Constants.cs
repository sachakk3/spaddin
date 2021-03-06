﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Algosmart.SharePoint.TimeSheetReceiverWeb.Code
{
    public static class Constants
    {
        public const string LIST_TITLE = "Записи табеля";

        public const string RECEIVER_ADDED_NAME = "ItemAddedEvent";
        public const string RECEIVER_UPDATED_NAME = "ItemUpdatedEvent";

        public const string GROUPS_OWNER_TITLE = "Backoffice - Владельцы";
        public const string GROUPS_BOSS_TITLE = "Backoffice-Boss";

        public const string GROUPS_HR_TITLE = "Backoffice-HR";
        public const string GROUPS_Fin_TITLE = "Backoffice-Fin";
        public const string GROUPS_PM_TITLE = "Backoffice-PM";        

        public const string WEBS_FINANCE_NAME = "/Finance";
        public const string LISTS_RATES_TITLE = "UserByProjectDetails";

        public const string FIELDS_PROJECTS_LOOKUP = "ts_ProjectsLookup";
        public const string FIELDS_AUTHOR = "Author";
        public const string FIELDS_TIMEBOARD_STATUS = "ts_TimeboardStatus";
        public const string FIELDS_INTERNAL_NAME = "ts_InternalName";
        public const string FIELDS_RATE = "ts_Rate";

        public const string FIELDS_PROJECTS_PM = "ts_ProjectManager";
        public const string FIELDS_PROJECTS_USERS = "ts_ProjectUsers";
    }
}