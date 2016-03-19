using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.SharePoint.Client;

namespace Algosmart.SharePoint.TimeSheetReceiverWeb.Code
{
    public static class Helper
    {
        public static string CreateFolderInList(this List list, ClientContext context, string itemFolderPath)
        {
            Web web = context.Web;
            string[] foldersName = itemFolderPath.Split(new[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
            context.Load(list.RootFolder);
            context.ExecuteQuery();
            string curNewFolderPath = list.RootFolder.ServerRelativeUrl;
            string curParentFolderPath = curNewFolderPath;

            foreach (string folder in foldersName)
            {
                curNewFolderPath += "/" + folder;
                Folder folderObj = web.GetFolderByServerRelativeUrl(curNewFolderPath);
                if (!folderObj.ExistsInList(list))
                {
                    ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
                    listItemCreationInformation.UnderlyingObjectType = FileSystemObjectType.Folder;
                    listItemCreationInformation.FolderUrl = curParentFolderPath;
                    ListItem newFolderItem = list.AddItem(listItemCreationInformation);
                    newFolderItem["Title"] = folder;
                    newFolderItem.Update();
                    context.ExecuteQuery();
                }
                curParentFolderPath += "/" + folder;
            }
            if (itemFolderPath.Substring(0, 1) == "/")
                return list.RootFolder.ServerRelativeUrl + itemFolderPath;
            return list.RootFolder.ServerRelativeUrl + "/" + itemFolderPath;
        }
        public static bool ExistsInList(this Folder folder, List list)
        {
            try
            {
                list.Context.Load(folder);
                list.Context.ExecuteQuery();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}