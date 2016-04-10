using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;
using Algosmart.SharePoint.TimeSheetReceiverWeb.Code;

namespace Algosmart.SharePoint.TimeSheetReceiverWeb.Services
{
    public class AppEventReceiver : IRemoteEventService
    {
        /// <summary>
        /// Handles app events that occur after the app is installed or upgraded, or when app is being uninstalled.
        /// </summary>
        /// <param name="properties">Holds information about the app event.</param>
        /// <returns>Holds information returned from the app event.</returns>
        public SPRemoteEventResult ProcessEvent(SPRemoteEventProperties properties)
        {
            System.Diagnostics.Trace.TraceInformation(string.Format("ProcessEvent. Начата обработка события '{0}'", properties.EventType));
            SPRemoteEventResult result = new SPRemoteEventResult();
            try
            {
                switch (properties.EventType)
                {
                    case SPRemoteEventType.AppInstalled:
                        HandleAppInstalled(properties);
                        break;
                    case SPRemoteEventType.AppUninstalling:
                        HandleAppUninstalling(properties);
                        break;               
                }
                System.Diagnostics.Trace.TraceInformation(string.Format("ProcessEvent. Окончена обработка события '{0}'", properties.EventType));

            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError(ex.ToString());
            }
            return result;
        }

        /// <summary>
        /// This method is a required placeholder, but is not used by app events.
        /// </summary>
        /// <param name="properties">Unused.</param>
        public void ProcessOneWayEvent(SPRemoteEventProperties properties)
        {
            System.Diagnostics.Trace.TraceInformation(string.Format("ProcessOneWayEvent. Начата обработка события '{0}'", properties.EventType));

            try
            {
                lock (LockObject.IsLocked)
                {
                    switch (properties.EventType)
                    {
                        case SPRemoteEventType.ItemAdded:
                            HandleTimeSheetEvents(properties);
                            break;
                        case SPRemoteEventType.ItemUpdated:
                            HandleTimeSheetEvents(properties);
                            break;
                    }
                }
                System.Diagnostics.Trace.TraceInformation(string.Format("ProcessOneWayEvent. Окончена обработка события '{0}'", properties.EventType));

            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError(ex.ToString());
            }
        }
        private void HandleAppInstalled(SPRemoteEventProperties properties)
        {
            using (ClientContext clientContext =
                TokenHelper.CreateAppEventClientContext(properties, false))
            {
                if (clientContext != null)
                {
                    new RemoteEventReceiverManager().AssociateRemoteEventsToHostWeb(clientContext);
                }
            }
        }
        private void HandleAppUninstalling(SPRemoteEventProperties properties)
        {
            using (ClientContext clientContext =
                TokenHelper.CreateAppEventClientContext(properties, false))
            {
                if (clientContext != null)
                {
                    new RemoteEventReceiverManager().RemoveEventReceiversFromHostWeb(clientContext);
                }
            }
        }
        private void HandleTimeSheetEvents(SPRemoteEventProperties properties)
        {

            string webUrl = properties.ItemEventProperties.WebUrl;
            Uri webUri = new Uri(webUrl);
            string realm = TokenHelper.GetRealmFromTargetUrl(webUri);
            string accessToken = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, webUri.Authority, realm).AccessToken;
            using (var context = TokenHelper.GetClientContextWithAccessToken(webUrl, accessToken))
            {
                new RemoteEventReceiverManager().ItemHandleListEventHandler(context, properties.ItemEventProperties.ListId, properties.ItemEventProperties.ListItemId, properties.EventType);
            }
        }
    }
}
