using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace CXPPlugins
{
    class CalculateTimeInClosePhase : IPlugin
    {
        //private IOrganizationService service;
        //private ITracingService tracer;

        public void Execute(IServiceProvider serviceProvider)
        {
            // Extract the tracing service for use in debugging sandboxed plug-ins.
            // If you are not registering the plug-in in the sandbox, then you do
            // not have to add any tracing service related code.
            ITracingService ObjTracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            ObjTracer.Trace("tracer created in execute method");

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // Obtain the organization service reference which you will need for
            // web service calls.
            IOrganizationServiceFactory serviceFactory =
                (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService ObjService = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.InputParameters.Contains("OpportunityClose") && context.InputParameters["OpportunityClose"] is Entity)
            {
                Entity opportunityClose = (Entity)context.InputParameters["OpportunityClose"];

                if (opportunityClose.Attributes.Contains("opportunityid") && opportunityClose.Attributes["opportunityid"] != null)
                {
                    EntityReference opportunityRef = (EntityReference)opportunityClose.Attributes["opportunityid"];
                    Entity opportunity = GetOpportunity(opportunityRef.Id, ObjService, ObjTracer);
                    if (opportunity != null)
                    {
                        DateTime dateEnteredClosePhase = opportunity.GetAttributeValue<DateTime>("cxp_dateenteredclosephase");
                        DateTime actualCloseDate = opportunityClose.GetAttributeValue<DateTime>("actualend");

                        TimeSpan timeSpanInClosePhase = actualCloseDate.Subtract(dateEnteredClosePhase);
                        //decimal timeInClosePhase = Decimal.Parse(timeSpanInClosePhase.TotalDays.ToString());
                        decimal timeInClosePhase = 0;
                        Decimal.TryParse(timeSpanInClosePhase.TotalDays.ToString(), out timeInClosePhase);//
                        Entity modifiedOpportunity = new Entity("opportunity");
                        modifiedOpportunity.Id = opportunityRef.Id;
                        modifiedOpportunity.Attributes["cxp_timeinclosephase"] = timeInClosePhase;
                        //tracer.trace("test");
                        ObjService.Update(modifiedOpportunity);
                    }
                }
            }
        }

        protected Entity GetOpportunity(Guid opportunityId, IOrganizationService ObjService, ITracingService ObjTracer)
        {
            Entity opportunity = null;

            try
            {
                opportunity = ObjService.Retrieve("opportunity", opportunityId, new ColumnSet(new string[] { "opportunityid", "cxp_dateenteredclosephase", "actualclosedate" }));
            }
            catch (Exception ex)
            {
                ObjTracer.Trace("CalculateTimeInClosePhase.GetOpportunity: {0}", ex.ToString());
                //throw ex;
                throw new InvalidPluginExecutionException("CalculateTimeInClosePhase.GetOpportunity: {0}" + ex.ToString());
            }

            return opportunity;
        }
    }
}
