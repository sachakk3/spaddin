using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using First_AppWeb.Code;

namespace First_AppWeb.Controllers
{
    public class HomeController : Controller
    {
        [SharePointContextFilter]
        public ActionResult Index()
        {
            User spUser = null;

            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    spUser = clientContext.Web.CurrentUser;

                    var spWeb = clientContext.Web;
                    var hostListColl = spWeb.Lists;
                    clientContext.Load(spWeb, w => w.Id);
                    clientContext.Load(hostListColl);

                    clientContext.Load(spUser, user => user.Title);

                    clientContext.ExecuteQuery();
                    ViewBag.HostLists = hostListColl.Select(l => new SelectListItem() { Text = l.Title, Value = l.Title });
                    ViewBag.UserName = spUser.Title;
                }
            }

            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        [SharePointContextFilter]
        [HttpPost]
        public ActionResult Subscribe(string listTitle)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (!string.IsNullOrEmpty(listTitle))
                {
                    RERUtility.AddListItemRemoteEventReceiver(
                      clientContext,
                      listTitle,
                      EventReceiverType.ItemAdded,
                      EventReceiverSynchronization.Asynchronous,
                      "RERHostReceiver",
                      "https://addinmvc.azurewebsites.net/Services/RemoteEventReceiverFirst.svc",
                          10);
                }
            }
            return RedirectToAction("Index", new { SPHostUrl = spContext.SPHostUrl.ToString() });
        }
    }
}
