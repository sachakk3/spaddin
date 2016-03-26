using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace TestConsoleApplication
{
    public class SPTSFolder
    {
        public string Name { get; set; }

        public string ListRelativeURL
        {
            get
            {
                if (this.ParentFolder == null)
                    return Name;
                return ParentFolder.ListRelativeURL + "/" + Name;
            }
        }

        public TSFolderType Type { get; set; }

        public ListItem Folder { get; set; }

        public SPTSFolder ParentFolder { get; set; }

        public bool IsNew { get; set; }
    }
}
