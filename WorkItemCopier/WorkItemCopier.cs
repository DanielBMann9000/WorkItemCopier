using System;
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
            var sourceTeamProjectName = "Scrum";
            var targetTeamProjectName = "CopyTarget";
            var copyState = "Removed";
            var fieldsToExclude = new[] {"System.AreaId", "System.IterationId", "System.AreaPath", "System.IterationPath", "System.State"};
            statusCode = 0;
            statusMessage = string.Empty;
            properties = null;
            if (notificationType == NotificationType.Notification)
            {
                var args = notificationEventArgs as WorkItemChangedEvent;

                if (args != null)
                {
                    var state = args.ChangedFields.StringFields.FirstOrDefault(f => f.ReferenceName == "System.State");
                    if (state != null && args.PortfolioProject == sourceTeamProjectName && state.NewValue == copyState)
                    {
                        var tpc = GetTeamProjectCollection(requestContext);
                        
                        var workItemStore = tpc.GetService<WorkItemStore>();
                        
                        var firstOrDefault = args.CoreFields.IntegerFields.FirstOrDefault(cf => cf.ReferenceName == "System.Id");
                        if (firstOrDefault != null)
                        {
                            var workItemId = firstOrDefault.OldValue;

                            var originalWorkItem = workItemStore.GetWorkItem(workItemId);
                            var bugWit = workItemStore.Projects[targetTeamProjectName].WorkItemTypes.Cast<WorkItemType>().FirstOrDefault(wit => wit.Name == "Bug");
                            if (bugWit != null)
                            {
                                var newWorkItem = bugWit.NewWorkItem();
                        
                                foreach (var field in originalWorkItem.Fields.Cast<Field>().Where(field => field.IsEditable && !fieldsToExclude.Contains(field.ReferenceName)))
                                {
                                    newWorkItem.Fields[field.ReferenceName].Value = field.Value;
                                }
                                newWorkItem.Save();
                            }
                        }
                    }
                }
            }
            return EventNotificationStatus.ActionPermitted;
        }

        public Type[] SubscribedTypes()
        {
            return new[] { typeof (WorkItemChangedEvent) };
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
