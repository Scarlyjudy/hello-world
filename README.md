using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;

//Description:
// This workflow retrieves the opportunities having Est close date past by 7 days and send out email to corresponsing Pursuit Lead

namespace CXPWorkflows
{
    public class TriggerEmailAlertsForOppsPastCloseDate : CodeActivity
    {
        protected override void Execute(CodeActivityContext executionContext)
        {
            // Get the context service.
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();

            // Use the context service to create an instance of IOrganizationService.
            IOrganizationService objService = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            
            ITracingService objTracer = executionContext.GetExtension<ITracingService>();
            
            WorkflowUtilities wfUtilities = new WorkflowUtilities();
			
			//********** Step 1: Retrieve the opportunities *************
			EntityCollection Opps = RetrieveOppsWithPastCloseDate(objService, objTracer);
            objTracer.Trace("opps retrieved");
            objTracer.Trace("Opps.Entities.Count:" + Opps.Entities.Count);
            string oppNum = "";
           // string orgURL = "cohnreznickdev.crm.dynamics.com";
           string orgURL = "cohnreznick.crm.dynamics.com";  // update the url for dev
            var objectTypeCode = 3; // opportunity entity


            try { 
                //********* Step 2: For each of the retrieved opps - send an email alert to corresponding Pursuit Lead **********
                if ((Opps != null) && (Opps.Entities.Count > 0))
                {
                    foreach (Entity opp in Opps.Entities)
                    {
                        //if (opp.Attributes.Contains("cxp_estclosedateemailalertsent"))
                        //{
                        //bool emailAlreadySent =  opp.GetAttributeValue<bool>("cxp_estclosedateemailalertsent");

                        //if (!emailAlreadySent && opp.Attributes.Contains("ownerid") && opp.Attributes["ownerid"] != null)
                        if (opp.Attributes.Contains("ownerid") && opp.Attributes["ownerid"] != null)
                        {
                            if (opp.Attributes.Contains("cxp_opportunitynumbertext") && opp.Attributes["cxp_opportunitynumbertext"] != null)
                            {
                                oppNum = opp.GetAttributeValue<string>("cxp_opportunitynumbertext");
                            }

                            string oppUrl = "https://" + orgURL + "/main.aspx?etc=" + objectTypeCode + "&id=%7b" + opp.Id + "%7d&pagetype=entityrecord";

                            // Send Email to the pursuit lead                

                            EntityReference receiverReference = opp.GetAttributeValue<EntityReference>("ownerid");
                            EntityReference senderReference = retrieveSender(objService, objTracer);

                            Entity _Email = new Entity("email");
                            //create activityparty
                            Entity from = new Entity("activityparty");
                            Entity to = new Entity("activityparty");
                            from["partyid"] = new EntityReference("systemuser", senderReference.Id);
                            to["partyid"] = new EntityReference("systemuser", receiverReference.Id);

                            _Email["from"] = new Entity[] { from };
                            _Email["to"] = new Entity[] { to };

                            _Email["regardingobjectid"] = new EntityReference("opportunity", opp.Id);
                            _Email["subject"] = "Overdue Opportunity Alert - Action Required";
                            _Email["description"] = @"Hi <br>" +
                                                    "This is a friendly reminder that your opportunity below was expected to close and is impacting the accuracy of the pipeline.<br>" +
                                                    "<b><font color=\"red\">ACTION REQUIRED:</font></b> If this is still a viable opportunity, please extend the close date. Otherwise, please close as won or lost.<br>" +
                                                    "Just click the link <a href=" + oppUrl + ">here </a> to access the opportunity.<br>" +
                                                    "If you need a quick “How to” click this link: <a href= https://cohnreznick.sharepoint.com/CXP/Learn/User%20Guides/Update%20overdue%20opportunities.pdf > How to update overdue Opportunities</a><br><br>Thanks,";

                            Guid EmailID = objService.Create(_Email);

                            if (orgURL == "cohnreznick.crm.dynamics.com")
                            { 
                                SendEmailRequest req = new SendEmailRequest();
                                req.EmailId = EmailID;
                                req.IssueSend = true;
                                req.TrackingToken = "";

                                SendEmailResponse res = (SendEmailResponse)objService.Execute(req);
                            }

                            //update flag on opp once the email is sent
                            Entity modifiedEntity = new Entity("opportunity", opp.Id);
                                modifiedEntity["cxp_estclosedateemailalertsent"] = true;
                                objService.Update(modifiedEntity);

                            }
                        //}
                    }
                }
            }
            catch (Exception ex)
            {   
                throw new InvalidPluginExecutionException("opp number:" + oppNum + "\n" +ex.Message + "\n" + ex.StackTrace);
            }
          
        }
		
        // retrieve opportunities past close date by 7 days 
		protected EntityCollection RetrieveOppsWithPastCloseDate(IOrganizationService objService, ITracingService objTracer) 
		{
            EntityCollection results = null;
            try
            {
                QueryExpression query = new QueryExpression("opportunity");
                query.NoLock = true;
                query.ColumnSet = new ColumnSet(new string[] { "ownerid", "cxp_opportunitynumbertext", "cxp_estclosedateemailalertsent" }); 
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                query.Criteria.AddCondition("estimatedclosedate", ConditionOperator.OlderThanXDays, 7);
                query.Criteria.AddCondition("estimatedclosedate", ConditionOperator.LastXDays, 10);  // no look back
                query.Criteria.AddCondition("cxp_parentopportunity", ConditionOperator.NotNull);
                query.Criteria.AddCondition("cxp_estclosedateemailalertsent", ConditionOperator.NotEqual, true);
                results = objService.RetrieveMultiple(query);
                objTracer.Trace("count of opps retrieved:" + results.Entities.Count);             
            }

            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return results;
		}

        // retrieve the sender
        protected EntityReference retrieveSender(IOrganizationService objService, ITracingService objTracer)
        {
            EntityCollection results = null;
            EntityReference user1 = null;

            try
            {
                QueryExpression query = new QueryExpression("systemuser");
                query.NoLock = true;
                query.ColumnSet = new ColumnSet(new string[] { "fullname" });
                query.Criteria.AddCondition("domainname", ConditionOperator.Equal, "SAM.Admin@CohnReznick.com");

                results = objService.RetrieveMultiple(query);

                if ((results != null) && (results.Entities.Count > 0))
                {
                    foreach (Entity user in results.Entities)
                    {

                        string value = user.Attributes["systemuserid"].ToString();

                        Guid Value = new Guid(value);

                        user1 = new EntityReference("systemuser", Value);
                    }
                }
            }
            catch (Exception ex)
            {
                objTracer.Trace(ex.Message + "\n" + ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
            return user1;
        }

    }
}
