using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Common;
using Microsoft.TeamFoundation.Framework.Server;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Server;

namespace WorkItemCopier
{
    public class WorkItemCopier : ISubscriber
    {
        private static readonly IEnumerable<string> fieldsToExclude = new[] { "System.AreaId", "System.IterationId", "System.AreaPath", "System.IterationPath", "System.State" };
        private static readonly string sourceTeamProjectName = "Scrum";
        private static readonly string targetTeamProjectName = "CopyTarget";
        private static readonly string copyState = "Removed";
        private static readonly string expectedWorkItemType = "Bug";

        public string Name
        {
            get { return "Work Item Copier"; }
        }

        public SubscriberPriority Priority
        {
            get
            {
                return SubscriberPriority.Normal;
            }
        }

        public EventNotificationStatus ProcessEvent(TeamFoundationRequestContext requestContext, NotificationType notificationType, object notificationEventArgs, out int statusCode, out string statusMessage, out ExceptionPropertyCollection properties)
        {
            statusCode = 0;
            statusMessage = string.Empty;
            properties = null;
            if (notificationType == NotificationType.Notification)
            {
                var args = notificationEventArgs as WorkItemChangedEvent;

                if (args == null)
                {
                    return EventNotificationStatus.ActionPermitted;
                }
                
                var state = args.ChangedFields.StringFields.FirstOrDefault(f => f.ReferenceName == "System.State");

                if (state == null || args.PortfolioProject != sourceTeamProjectName || state.NewValue != copyState)
                {
                    return EventNotificationStatus.ActionPermitted;
                }

                var tpc = GetTeamProjectCollection(requestContext);

                var workItemStore = tpc.GetService<WorkItemStore>();

                var workItemIdField = args.CoreFields.IntegerFields.FirstOrDefault(cf => cf.ReferenceName == "System.Id");
                if (workItemIdField == null)
                {
                    return EventNotificationStatus.ActionPermitted;
                }

                var workItemId = workItemIdField.OldValue;

                var originalWorkItem = workItemStore.GetWorkItem(workItemId);
                if (originalWorkItem.Type != workItemStore.Projects[sourceTeamProjectName].WorkItemTypes.Cast<WorkItemType>().FirstOrDefault(wit => wit.Name == expectedWorkItemType))
                {
                    return EventNotificationStatus.ActionPermitted;
                }
                var bugWit = workItemStore.Projects[targetTeamProjectName].WorkItemTypes.Cast<WorkItemType>().FirstOrDefault(wit => wit.Name == expectedWorkItemType);
                CopyWit(originalWorkItem, bugWit);
            }
            return EventNotificationStatus.ActionPermitted;
        }

        private static void CopyWit(WorkItem source, WorkItemType target)
        {
            if (target == null)
            {
                return;
            }

            var newWorkItem = target.NewWorkItem();

            foreach (var field in source.Fields.Cast<Field>().Where(field => field.IsEditable && !fieldsToExclude.Contains(field.ReferenceName)))
            {
                newWorkItem.Fields[field.ReferenceName].Value = field.Value;
            }
            newWorkItem.Save();
        }

        public Type[] SubscribedTypes()
        {
            return new[] { typeof(WorkItemChangedEvent) };
        }

        private TfsTeamProjectCollection GetTeamProjectCollection(TeamFoundationRequestContext requestContext)
        {
            var locationService = requestContext.GetService<TeamFoundationLocationService>();
            var accessPoint = locationService.GetServerAccessMapping(requestContext).AccessPoint;
            var serviceHostName = requestContext.ServiceHost.Name;

            return TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(string.Format("{0}/{1}", accessPoint, serviceHostName)));
        }
    }
}
