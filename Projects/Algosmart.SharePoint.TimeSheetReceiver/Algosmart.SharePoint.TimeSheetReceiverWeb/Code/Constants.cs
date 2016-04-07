using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Algosmart.SharePoint.TimeSheetReceiverWeb.Code
{
    public static class Constants
    {
        public const string SYSTEM_USER_NAME = "Developer@algosmart.onmicrosoft.com";
        public const string SYSTEM_USER_PASSWORD = "Rostov2016";

        public const string LIST_TITLE = "Записи табеля";

        public const string RECEIVER_ADDED_NAME = "ItemAddedEvent";
        public const string RECEIVER_UPDATED_NAME = "ItemUpdatedEvent";

        public const string GROUPS_OWNER_TITLE = "Backoffice - Владельцы";
        public const string GROUPS_BOSS_TITLE = "Backoffice-Boss";

        public const string WEBS_FINANCE_NAME = "/Finance";
        public const string LISTS_RATES_TITLE = "UserByProjectDetails";

        public const string FIELDS_PROJECTS_LOOKUP = "ts_ProjectsLookup";
        public const string FIELDS_AUTHOR = "Author";
        public const string FIELDS_TIMEBOARD_STATUS = "ts_TimeboardStatus";
        public const string FIELDS_INTERNAL_NAME = "ts_InternalName";
        public const string FIELDS_RATE = "ts_Rate";
    }
}