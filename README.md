using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;

namespace CXPWorkflows
{
    public class ReassignLeads : CodeActivity
    {
        protected override void Execute(CodeActivityContext executionContext)
        {
            //Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            //Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            Guid leadId = context.PrimaryEntityId;

            Entity lead = service.Retrieve("lead", leadId, new Microsoft.Xrm.Sdk.Query.ColumnSet(new string[] { "address1_line3" }));

            string userName = lead.GetAttributeValue<string>("address1_line3");

            if (!userName.Equals("Bryan Holmstrom"))
            {
                Guid userId = GetUserId(tracingService, service, userName);
                if (userId != Guid.Empty)
                {
                    AssignLead(tracingService, service, new EntityReference("systemuser", userId), leadId);
                }
            }


            throw new NotImplementedException();
        }

        protected Guid GetUserId(ITracingService tracer, IOrganizationService service, string userName)
        {
            Guid userId = Guid.Empty;

            try
            {
                QueryByAttribute query = new QueryByAttribute("systemuser");
                query.ColumnSet = new ColumnSet(new string[] { "systemuserid" });
                query.AddAttributeValue("fullname", userName);
                EntityCollection results = service.RetrieveMultiple(query);

                userId = results.Entities[0].GetAttributeValue<Guid>("systemuserid");
            }
            catch (Exception e)
            {
                tracer.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw e;
            }

            return userId;
        }

        protected void AssignLead(ITracingService tracer, IOrganizationService service, EntityReference user, Guid leadId)
        {
            try
            {
                Entity lead = new Entity("lead", leadId);
                lead["ownerid"] = user;
                service.Update(lead);
            }
            catch (Exception e)
            {
                tracer.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw e;
            }
        }
    }
}
